using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

// Globe.cs AST(정규식 기반) 분석:
//   Globe.cs 의 (Class, Prop, IsRequired, PropertyType) 카탈로그 추출
//   + Mapping_*.cs 의 LHS 할당 체인 추출
//   → qa/ast_catalog.csv     : 클래스/속성/필수여부/속성타입
//   → qa/ast_assignments.csv : 매퍼 LHS 체인 (a.b.c.Prop 형태)
//   → qa/gaps_reverse_v2.md  : 다중 클래스 충돌 고려한 강화 역방향 갭

class Program
{
    record PropInfo(string ClassName, string PropName, bool IsRequired, string PropType, string XmlName, int LineNo);
    record Assignment(string File, int Line, string LhsChain, string Snippet);

    static int Main(string[] args)
    {
        var repo = args.Length > 0
            ? Path.GetFullPath(args[0])
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));

        var globeCs = Path.Combine(repo, "Resources", "Globe.cs");
        var srcDir  = Path.Combine(repo, "Services");
        var qaDir   = Path.Combine(repo, "qa");
        Directory.CreateDirectory(qaDir);

        var catalog = ParseGlobeCs(globeCs);
        WriteCatalog(catalog, Path.Combine(qaDir, "ast_catalog.csv"));

        var assignments = ParseAssignments(srcDir);
        WriteAssignments(assignments, Path.Combine(qaDir, "ast_assignments.csv"));

        WriteReverseGapV2(catalog, assignments, Path.Combine(qaDir, "gaps_reverse_v2.md"));

        // v3: 타입 추론기 — 부모 타입 컨텍스트로 (Class, Prop) 정밀 매칭
        var resolved = ResolveAssignedClassProps(catalog, srcDir);
        WriteReverseGapV3(catalog, resolved, Path.Combine(qaDir, "gaps_reverse_v3.md"));

        Console.WriteLine("Done.");
        return 0;
    }

    // ───────────────────────── v3: 타입 해석 ─────────────────────────

    record ResolvedAssignment(string ClassName, string PropName, string File, int Line, string Snippet);

    static HashSet<(string Class, string Prop)> ResolveAssignedClassProps(List<PropInfo> catalog, string srcDir)
    {
        // 카탈로그 인덱스: (Class, Prop) → PropTypeName
        // PropType 에서 Globe.* / 제네릭 / 컬렉션 vs 단일 자동 정리
        var propTypeMap = new Dictionary<(string, string), string>();
        var classNames = new HashSet<string>(catalog.Select(p => p.ClassName));
        foreach (var p in catalog)
        {
            propTypeMap[(p.ClassName, p.PropName)] = NormalizeType(p.PropType);
        }

        // 상속 체인 따라가며 (Class+ancestors, Prop) 검색
        string ResolvePropType(string cls, string prop)
        {
            var cur = cls;
            while (cur != null)
            {
                if (propTypeMap.TryGetValue((cur, prop), out var t)) return t;
                _inheritanceMap.TryGetValue(cur, out var parent);
                cur = parent;
            }
            return null;
        }

        // 모든 .cs 파일 (Globe.cs 제외)
        var files = Directory.EnumerateFiles(srcDir, "*.cs")
            .Concat(new[] { Path.Combine(srcDir, "..", "MainForm.cs") })
            .Where(f => File.Exists(f) && !Path.GetFileName(f).Equals("Globe.cs", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // 파일별 변수 타입 컨텍스트 (스코프 무시 — 이름 충돌은 허용 범위 내)
        var resolved = new HashSet<(string, string)>();

        // 정규식
        // var X = new Globe.TypeName(...)  / new Globe.TypeName { ... }  / new TypeName(...)
        var newVarRegex = new Regex(@"\bvar\s+(\w+)\s*=\s*new\s+(?:Globe\.)?([A-Z]\w*)\s*[\(\{]", RegexOptions.Compiled);
        // var X = <anything> ?? new TypeName(...)  / var X = <anything containing 'new TypeName(...)'>;
        // RHS 안 어디서든 new TypeName 이 나오면 X 의 타입 후보로 채택. ; 또는 다음 var 문/문장 끝까지.
        var newVarLooseRegex = new Regex(
            @"\bvar\s+(\w+)\s*=[^;{}]*?\bnew\s+(?:Globe\.)?([A-Z]\w*)\s*[\(\{]",
            RegexOptions.Compiled | RegexOptions.Singleline);
        // TypeName X = new ...  (직접 타입 선언)
        var typedVarRegex = new Regex(@"\b(?:Globe\.)?([A-Z]\w*)\s+(\w+)\s*=\s*new\s+(?:Globe\.)?\1\b", RegexOptions.Compiled);
        // var X = chain.expression   (chain 으로 시작하는 RHS 만)
        var aliasVarRegex = new Regex(@"\bvar\s+(\w+)\s*=\s*([a-z_]\w*(?:\s*\.\s*\w+)+)\s*;", RegexOptions.Compiled);
        // foreach (var x in collection)  → x : collection의 요소 타입
        var foreachRegex = new Regex(@"\bforeach\s*\(\s*var\s+(\w+)\s+in\s+([a-z_]\w*(?:\s*\.\s*\w+)+)\s*\)", RegexOptions.Compiled);
        // 메서드 파라미터: ... Method(Globe.TypeName x, ..., TypeName y, ...)
        // 한 번에 한 파라미터씩 매칭
        var paramRegex = new Regex(@"[\(,]\s*(?:Globe\.)?([A-Z]\w*)\s+(\w+)\s*[,\)]", RegexOptions.Compiled);
        // chain ??= new TypeName(...)  → chain 의 마지막 propType 가 TypeName 임을 확정
        // (이미 catalog 로 알 수 있어서 우리에겐 추가 정보 없음 — 그러나 LHS 체인은 assign 으로 이미 잡힘)
        // 객체 이니셜라이저: new TypeName { ... } — '{' 만 정규식으로 찾고 매칭 '}' 까지 브레이스 카운트
        var initStartRegex = new Regex(@"new\s+(?:Globe\.)?([A-Z]\w*)\s*(?:\([^)]*\))?\s*\{", RegexOptions.Compiled);
        // 이니셜라이저 내부 Prop = ...  (각 Prop 이 직속 멤버인지는 중첩 깊이 0 일 때만 채택)
        var initPropRegex = new Regex(@"(?:^|[\{,])\s*([A-Z]\w*)\s*=", RegexOptions.Compiled);
        // 할당 LHS: a.b.c... = ... (?? = 도 허용)
        var assignRegex = new Regex(@"(?<![=!<>])\b([a-zA-Z_]\w*(?:\s*\.\s*\w+)+)\s*(?:\?\?)?=(?!=)", RegexOptions.Compiled);
        // 컬렉션 접근: chain.Add(...) — 마지막 prop 이 Collection<X>인 경우 사용 처리
        var addCallRegex = new Regex(@"\b([a-zA-Z_]\w*(?:\s*\.\s*\w+)+)\.Add\s*\(", RegexOptions.Compiled);

        foreach (var file in files)
        {
            var text = StripLineComments(File.ReadAllText(file));

            // 위치 기반 변수 선언 이벤트 (charIdx, name, type)
            // 같은 이름이 여러 번 선언되면 위치별로 다른 타입을 가질 수 있음
            var declEvents = new List<(int Pos, string Name, string Type)>();
            // 파일 전역 (위치 무관) 변수 — 메서드 파라미터, 잘 알려진 seed
            var globalVars = new Dictionary<string, string>();
            globalVars["globe"] = "GlobeOecd";

            // 위치 기반 선언 수집기
            void AddDecl(int pos, string name, string typeName)
            {
                if (string.IsNullOrEmpty(typeName)) return;
                if (!classNames.Contains(typeName)) return;
                declEvents.Add((pos, name, typeName));
            }

            // 1) var X = new TypeName(...)
            foreach (Match m in newVarRegex.Matches(text))
                AddDecl(m.Index, m.Groups[1].Value, m.Groups[2].Value);

            // 1.5) var X = ... new TypeName(...) (멀티라인 + ?? 등)
            foreach (Match m in newVarLooseRegex.Matches(text))
                AddDecl(m.Index, m.Groups[1].Value, m.Groups[2].Value);

            // 1.7) X = new Type(...);  (재할당, var 없이)
            var reassignRegex = new Regex(@"(?<![.\w])([a-zA-Z_]\w*)\s*=\s*new\s+(?:Globe\.)?([A-Z]\w*)\s*[\(\{]", RegexOptions.Compiled);
            foreach (Match m in reassignRegex.Matches(text))
                AddDecl(m.Index, m.Groups[1].Value, m.Groups[2].Value);

            // 1.8) LINQ FirstOrDefault 등
            var linqVarRegex = new Regex(
                @"\bvar\s+(\w+)\s*=\s*([a-z_]\w*(?:\s*\.\s*\w+)+)\.(?:First|FirstOrDefault|Single|SingleOrDefault|Last|LastOrDefault)\s*\(",
                RegexOptions.Compiled);
            foreach (Match m in linqVarRegex.Matches(text))
            {
                var chain = m.Groups[2].Value;
                var t = ResolveChainType(chain, globalVars, propTypeMap);
                if (t != null) AddDecl(m.Index, m.Groups[1].Value, ElementType(t));
            }

            // 2) 명시적 타입 선언
            foreach (Match m in typedVarRegex.Matches(text))
                AddDecl(m.Index, m.Groups[2].Value, m.Groups[1].Value);

            // 2.5) 메서드 파라미터 — 파일 전역으로 처리 (메서드 스코프 추적은 생략)
            foreach (Match m in paramRegex.Matches(text))
            {
                var typeName = m.Groups[1].Value;
                if (classNames.Contains(typeName))
                    globalVars[m.Groups[2].Value] = typeName;
            }

            // alias 루프 진입 전 정렬 (LookupVarTypeAt 정확성)
            declEvents.Sort((a, b) => a.Pos.CompareTo(b.Pos));

            // 3) var X = chain  (별칭) — 반복 패스, 위치 정보와 함께 보존
            for (int iter = 0; iter < 5; iter++)
            {
                bool changed = false;
                // 매 iteration 시작 시 정렬 — LookupVarTypeAt 의 정확성 보장
                declEvents.Sort((a, b) => a.Pos.CompareTo(b.Pos));
                foreach (Match m in aliasVarRegex.Matches(text))
                {
                    var chain = m.Groups[2].Value;
                    var firstVar = chain.Split('.')[0].Trim();
                    var ctx = LookupVarTypeAt(declEvents, globalVars, firstVar, m.Index);
                    if (ctx == null) continue;
                    var t = ResolveChainTypeFrom(ctx, chain, propTypeMap);
                    if (t != null)
                    {
                        // 정확히 이 위치에 이미 추가된 적 없으면 추가
                        bool already = declEvents.Any(e => e.Pos == m.Index && e.Name == m.Groups[1].Value);
                        if (!already)
                        {
                            AddDecl(m.Index, m.Groups[1].Value, t);
                            changed = true;
                        }
                    }
                }
                if (!changed) break;
            }

            // 4) foreach 루프 변수 타입 (컬렉션 요소)
            foreach (Match m in foreachRegex.Matches(text))
            {
                var chain = m.Groups[2].Value;
                var ctx = LookupVarTypeAt(declEvents, globalVars, chain.Split('.')[0], m.Index);
                if (ctx == null) continue;
                var t = ResolveChainTypeFrom(ctx, chain, propTypeMap);
                if (t != null) AddDecl(m.Index, m.Groups[1].Value, ElementType(t));
            }

            // alias 후 다시 정렬 (새로 추가된 declEvent 반영)
            declEvents.Sort((a, b) => a.Pos.CompareTo(b.Pos));

            // 위치 기반 변수 조회 위해 scoped wrapper
            string LookupAt(string name, int pos)
                => LookupVarTypeAt(declEvents, globalVars, name, pos);

            // 기존 코드의 varTypes 참조용 (이후 5/5.5/6 단계에서 사용)
            // 단순화를 위해 가장 최근 선언 기준 dict 도 만들어 둠 (전역 fallback)
            var varTypes = new Dictionary<string, string>(globalVars);
            foreach (var (_, name, type) in declEvents)
                varTypes[name] = type; // 최종 상태 (마지막 선언) — fallback 용

            // 5) 객체 이니셜라이저 — 브레이스 카운팅으로 중첩 처리
            //    new TypeName { ... } 의 직속 멤버 Prop 만 (TypeName, Prop) 으로 기록.
            //    중첩 new InnerType { ... } 내부의 Prop 은 별도 매칭에서 처리됨.
            foreach (Match m in initStartRegex.Matches(text))
            {
                var typeName = m.Groups[1].Value;
                if (!classNames.Contains(typeName)) continue;
                int openIdx = m.Index + m.Length - 1; // '{' 위치
                // 매칭 '}' 찾기
                int depth = 1;
                int i = openIdx + 1;
                while (i < text.Length && depth > 0)
                {
                    char c = text[i];
                    if (c == '{') depth++;
                    else if (c == '}') depth--;
                    if (depth == 0) break;
                    i++;
                }
                if (depth != 0) continue;

                // 직속 본문: openIdx+1 ~ i. 단 중첩 new ... { ... } 는 제외해야 함.
                // 본문 내에서 다시 깊이 추적하며 깊이 0 일 때만 'Prop = ' 매칭 채택
                int bodyStart = openIdx + 1;
                int bodyEnd = i;
                int depth2 = 0;
                var topPositions = new List<int>();
                topPositions.Add(bodyStart);
                for (int k = bodyStart; k < bodyEnd; k++)
                {
                    char c = text[k];
                    if (c == '{') depth2++;
                    else if (c == '}') depth2--;
                    else if (c == ',' && depth2 == 0)
                        topPositions.Add(k + 1);
                }
                // 각 top-level segment 의 시작 부분에서 Prop = 매칭
                foreach (var start in topPositions)
                {
                    int end = start;
                    int d = 0;
                    while (end < bodyEnd && (text[end] != ',' || d > 0))
                    {
                        if (text[end] == '{') d++;
                        else if (text[end] == '}') d--;
                        end++;
                    }
                    var seg = text.Substring(start, end - start);
                    var pmatch = Regex.Match(seg, @"^\s*([A-Z]\w*)\s*=");
                    if (!pmatch.Success) continue;
                    var prop = pmatch.Groups[1].Value;
                    var declaring = typeName;
                    while (declaring != null)
                    {
                        if (propTypeMap.ContainsKey((declaring, prop)))
                        {
                            resolved.Add((declaring, prop));
                            break;
                        }
                        _inheritanceMap.TryGetValue(declaring, out var parent);
                        declaring = parent;
                    }
                }
            }

            // 5.5) 컬렉션 .Add() — chain 의 마지막 prop 을 사용 처리
            foreach (Match m in addCallRegex.Matches(text))
            {
                var chain = m.Groups[1].Value;
                var parts = chain.Split('.').Select(s => s.Trim()).ToArray();
                if (parts.Length < 2) continue;
                var rootType = LookupAt(parts[0], m.Index);
                if (rootType == null) continue;
                var ctx = rootType;
                for (int i = 1; i < parts.Length - 1; i++)
                {
                    var t = ResolvePropType(ctx, parts[i]);
                    if (t == null) { ctx = null; break; }
                    ctx = ElementType(t);
                }
                if (ctx == null) continue;
                var lastProp = parts[^1];
                var declaring = ctx;
                while (declaring != null)
                {
                    if (propTypeMap.ContainsKey((declaring, lastProp)))
                    {
                        resolved.Add((declaring, lastProp));
                        break;
                    }
                    _inheritanceMap.TryGetValue(declaring, out var parent);
                    declaring = parent;
                }
            }

            // 6) 할당 LHS 체인 해석 (위치 기반)
            foreach (Match m in assignRegex.Matches(text))
            {
                var chain = m.Groups[1].Value;
                var parts = chain.Split('.').Select(s => s.Trim()).ToArray();
                if (parts.Length < 2) continue;
                var rootType = LookupAt(parts[0], m.Index);
                if (rootType == null) continue;
                var ctx = rootType;
                for (int i = 1; i < parts.Length - 1; i++)
                {
                    var t = ResolvePropType(ctx, parts[i]);
                    if (t == null) { ctx = null; break; }
                    ctx = ElementType(t);
                }
                if (ctx == null) continue;
                var lastProp = parts[^1];
                // 상속 체인 따라가며 (Class, Prop) 의 정확한 declaring class 기록
                var declaring = ctx;
                while (declaring != null)
                {
                    if (propTypeMap.ContainsKey((declaring, lastProp)))
                    {
                        resolved.Add((declaring, lastProp));
                        break;
                    }
                    _inheritanceMap.TryGetValue(declaring, out var parent);
                    declaring = parent;
                }
            }
        }

        return resolved;
    }

    static string ResolveChainType(string chain, Dictionary<string, string> varTypes,
                                    Dictionary<(string, string), string> propTypeMap)
    {
        var parts = chain.Split('.').Select(s => s.Trim()).ToArray();
        if (!varTypes.TryGetValue(parts[0], out var ctx)) return null;
        return ResolveChainFromCtx(ctx, parts, 1, propTypeMap);
    }

    static string ResolveChainTypeFrom(string rootType, string chain,
                                        Dictionary<(string, string), string> propTypeMap)
    {
        var parts = chain.Split('.').Select(s => s.Trim()).ToArray();
        return ResolveChainFromCtx(rootType, parts, 1, propTypeMap);
    }

    static string ResolveChainFromCtx(string ctx, string[] parts, int startIdx,
                                       Dictionary<(string, string), string> propTypeMap)
    {
        for (int i = startIdx; i < parts.Length; i++)
        {
            string t = null;
            var cur = ctx;
            while (cur != null)
            {
                if (propTypeMap.TryGetValue((cur, parts[i]), out t)) break;
                _inheritanceMap.TryGetValue(cur, out var parent);
                cur = parent;
            }
            if (t == null) return null;
            ctx = ElementType(t);
        }
        return ctx;
    }

    // 위치 기반 변수 타입 조회: pos 이전의 가장 마지막 선언 사용. 없으면 globalVars fallback.
    static string LookupVarTypeAt(List<(int Pos, string Name, string Type)> events,
                                    Dictionary<string, string> globalVars,
                                    string name, int pos)
    {
        string best = null;
        foreach (var ev in events)
        {
            if (ev.Pos > pos) break;
            if (ev.Name == name) best = ev.Type;
        }
        if (best != null) return best;
        return globalVars.TryGetValue(name, out var t) ? t : null;
    }

    static string NormalizeType(string t)
    {
        // "Globe." 접두사 제거
        if (t.StartsWith("Globe.")) t = t.Substring(6);
        return t;
    }

    static string ElementType(string t)
    {
        // 컬렉션 → 요소 타입 (Coll<Foo> 또는 List<Foo> 또는 Collection<Foo>)
        var m = Regex.Match(t, @"(?:Collection|List|Coll)<\s*(?:Globe\.)?(\w+)\s*>");
        if (m.Success) return m.Groups[1].Value;
        return NormalizeType(t);
    }

    static void WriteReverseGapV3(List<PropInfo> catalog, HashSet<(string, string)> resolved, string outPath)
    {
        var requiredProps = catalog.Where(p => p.IsRequired).ToList();
        var sb = new StringBuilder();
        sb.AppendLine("# 역방향 갭 리포트 v3 (타입 해석기)");
        sb.AppendLine();
        sb.AppendLine("매퍼 코드의 LHS 체인을 Globe.cs 타입 트리로 해석하여 (Class, Prop) 단위 정밀 매칭.");
        sb.AppendLine();
        sb.AppendLine("- ✓ : (Class, Prop) 가 매퍼에서 직접 또는 간접 할당됨");
        sb.AppendLine("- ✗ : 미해결 — 진짜 갭이거나 해석기가 못 따라간 패턴");
        sb.AppendLine();

        int hit = 0, miss = 0;
        var sbMiss = new StringBuilder();
        sbMiss.AppendLine("## ✗ 미해결");
        sbMiss.AppendLine();
        sbMiss.AppendLine("| 클래스 | 속성 | 타입 | XmlName |");
        sbMiss.AppendLine("|---|---|---|---|");
        foreach (var p in requiredProps.OrderBy(p => p.ClassName).ThenBy(p => p.PropName))
        {
            if (resolved.Contains((p.ClassName, p.PropName))) hit++;
            else
            {
                miss++;
                sbMiss.Append("| ").Append(p.ClassName)
                      .Append(" | `").Append(p.PropName).Append("`")
                      .Append(" | ").Append(EscapeMd(p.PropType))
                      .Append(" | ").Append(p.XmlName).AppendLine(" |");
            }
        }
        sb.AppendLine($"## 결과");
        sb.AppendLine();
        sb.AppendLine($"- ✓ 매칭됨: **{hit}** / {requiredProps.Count}");
        sb.AppendLine($"- ✗ 미해결: **{miss}** / {requiredProps.Count}");
        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine();
        if (miss > 0) sb.Append(sbMiss);

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Reverse gap v3 → {outPath}  (hit={hit}, miss={miss})");
    }

    // ───────────────────────── Globe.cs 파싱 ─────────────────────────

    // public partial class XxxType  또는  public partial class XxxType : Parent
    static readonly Regex ClassRegex = new(
        @"^\s*public\s+partial\s+class\s+(\w+)(?:\s*:\s*(\w+))?\s*$",
        RegexOptions.Compiled | RegexOptions.Multiline);

    // 상속 맵: child → base
    static Dictionary<string, string> _inheritanceMap = new();

    // 속성: public Type Prop { get; set; }   (자동) 또는  public Type Prop  (멀티라인 explicit getter)
    static readonly Regex PropRegex = new(
        @"^\s*public\s+([\w\.<>\[\],\s]+?)\s+(\w+)\s*(?:\{\s*get;\s*(?:private\s+)?set;\s*\}\s*$|$)",
        RegexOptions.Compiled);

    // 어트리뷰트: [System.ComponentModel.DataAnnotations.RequiredAttribute...]
    static readonly Regex RequiredAttrRegex = new(
        @"\bRequiredAttribute\b",
        RegexOptions.Compiled);

    // [System.Xml.Serialization.XmlElementAttribute("Name"...)] 또는 XmlAttributeAttribute
    static readonly Regex XmlNameRegex = new(
        @"\bXml(Element|Attribute)Attribute\s*\(\s*""([^""]+)""",
        RegexOptions.Compiled);

    static List<PropInfo> ParseGlobeCs(string path)
    {
        var lines = File.ReadAllLines(path);
        var result = new List<PropInfo>();
        string currentClass = "";
        var pendingAttrs = new List<string>();

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var t = line.TrimStart();

            // 새 클래스 진입 — 이전 컨텍스트 리셋
            var cm = ClassRegex.Match(line);
            if (cm.Success)
            {
                currentClass = cm.Groups[1].Value;
                if (cm.Groups[2].Success)
                    _inheritanceMap[currentClass] = cm.Groups[2].Value;
                pendingAttrs.Clear();
                continue;
            }

            // enum/struct 선언이 나오면 클래스 컨텍스트 종료
            if (Regex.IsMatch(line, @"^\s*public\s+(enum|struct)\s+\w+"))
            {
                currentClass = "";
                pendingAttrs.Clear();
                continue;
            }

            if (string.IsNullOrEmpty(currentClass)) continue;

            // 어트리뷰트 누적
            if (t.StartsWith("["))
            {
                pendingAttrs.Add(t);
                continue;
            }

            // 속성 매칭
            var pm = PropRegex.Match(line);
            if (pm.Success)
            {
                var propType = pm.Groups[1].Value;
                var propName = pm.Groups[2].Value;
                bool isRequired = pendingAttrs.Any(a => RequiredAttrRegex.IsMatch(a));
                string xmlName = "";
                foreach (var a in pendingAttrs)
                {
                    var xm = XmlNameRegex.Match(a);
                    if (xm.Success) { xmlName = xm.Groups[2].Value; break; }
                }
                result.Add(new PropInfo(currentClass, propName, isRequired, propType, xmlName, i + 1));
                pendingAttrs.Clear();
                continue;
            }

            // 빈 줄/주석/중괄호만은 어트리뷰트 컨텍스트 유지
            if (string.IsNullOrWhiteSpace(t) || t.StartsWith("///") || t.StartsWith("//")
                || t == "{" || t == "}") continue;

            // 그 외는 어트리뷰트 컨텍스트 리셋
            pendingAttrs.Clear();
        }

        return result;
    }

    static void WriteCatalog(List<PropInfo> catalog, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("class,prop,required,prop_type,xml_name,line");
        foreach (var p in catalog.OrderBy(p => p.ClassName).ThenBy(p => p.PropName))
        {
            sb.Append(Csv(p.ClassName)).Append(',')
              .Append(Csv(p.PropName)).Append(',')
              .Append(p.IsRequired ? "R" : "O").Append(',')
              .Append(Csv(p.PropType)).Append(',')
              .Append(Csv(p.XmlName)).Append(',')
              .Append(p.LineNo).AppendLine();
        }
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Catalog: {catalog.Count} props ({catalog.Count(p => p.IsRequired)} required) → {outPath}");
    }

    // ───────────────────────── 매퍼 할당 추출 ─────────────────────────

    // LHS = / ??= ... 형태에서 LHS 의 dotted chain 추출 (단순 식별자 + . 만)
    // 예: "globe.MessageSpec.SendingEntityIn = value"  → "globe.MessageSpec.SendingEntityIn"
    //     "fi.FilingCe ??= new Globe.FilingInfoFilingCe()" → "fi.FilingCe"
    static readonly Regex AssignmentRegex = new(
        @"(?<![=!<>])\b([a-zA-Z_]\w*(?:\s*\.\s*\w+)+)\s*(?:\?\?)?=(?!=)",
        RegexOptions.Compiled);

    // 객체 초기화: { Prop = ... } — Prop 만 추출 (LHS 가 식별자 1개)
    static readonly Regex InitializerRegex = new(
        @"(?<=[\{,]\s*)([A-Z]\w*)\s*=(?!=)",
        RegexOptions.Compiled);

    static List<Assignment> ParseAssignments(string srcDir)
    {
        var list = new List<Assignment>();
        // Mapping_*.cs + MappingOrchestrator.cs + MappingBase.cs + MainForm.cs (전체 .cs 중 Globe.cs 제외)
        var files = Directory.EnumerateFiles(srcDir, "*.cs")
            .Concat(new[] { Path.Combine(srcDir, "..", "MainForm.cs") })
            .Where(f => File.Exists(f) && !Path.GetFileName(f).Equals("Globe.cs", StringComparison.OrdinalIgnoreCase))
            .ToList();
        foreach (var file in files)
        {
            var lines = File.ReadAllLines(file);
            var rawText = File.ReadAllText(file);
            // 라인 끝 주석 제거 (`//.*$`) — InitializerRegex 의 lookbehind 가 주석에 막히는 것을 방지
            // 문자열 리터럴 안의 // 는 무시 (간이: 따옴표 내부 추적)
            var text = StripLineComments(rawText);
            var name = Path.GetFileName(file);

            // 라인 시작 인덱스 미리 계산 (charIndex → line number)
            var lineStarts = new List<int> { 0 };
            for (int k = 0; k < text.Length; k++)
                if (text[k] == '\n') lineStarts.Add(k + 1);

            int LineOf(int charIdx)
            {
                int lo = 0, hi = lineStarts.Count - 1;
                while (lo < hi) { int mid = (lo + hi + 1) / 2; if (lineStarts[mid] <= charIdx) lo = mid; else hi = mid - 1; }
                return lo;
            }

            // dotted chain 할당
            foreach (Match m in AssignmentRegex.Matches(text))
            {
                var chain = m.Groups[1].Value;
                if (chain.Count(c => c == '.') < 1) continue;
                int li = LineOf(m.Index);
                var snip = (li < lines.Length ? lines[li] : "").Trim();
                if (snip.Length > 120) snip = snip.Substring(0, 120) + "…";
                list.Add(new Assignment(name, li + 1, chain, snip));
            }

            // 객체 이니셜라이저: { Prop = ..., Other = ... } — 멀티라인 허용
            // PropName 은 대문자 시작 + (?!=) 로 == 회피
            foreach (Match m in InitializerRegex.Matches(text))
            {
                var prop = m.Groups[1].Value;
                int li = LineOf(m.Index);
                var snip = (li < lines.Length ? lines[li] : "").Trim();
                if (snip.Length > 120) snip = snip.Substring(0, 120) + "…";
                list.Add(new Assignment(name, li + 1, "<init>." + prop, snip));
            }
        }
        return list;
    }

    static void WriteAssignments(List<Assignment> list, string outPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("file,line,lhs_chain,snippet");
        foreach (var a in list)
        {
            sb.Append(Csv(a.File)).Append(',')
              .Append(a.Line).Append(',')
              .Append(Csv(a.LhsChain)).Append(',')
              .Append(Csv(a.Snippet)).AppendLine();
        }
        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Assignments: {list.Count} chains → {outPath}");
    }

    // ───────────────────────── Reverse gap v2 ─────────────────────────

    static void WriteReverseGapV2(List<PropInfo> catalog, List<Assignment> assignments, string outPath)
    {
        // 속성명 다중성 (한 PropName 이 몇 개 클래스에 존재?)
        var multiplicity = catalog
            .GroupBy(p => p.PropName)
            .ToDictionary(g => g.Key, g => g.Select(p => p.ClassName).Distinct().ToList());

        // 모든 LHS 체인을 정규화 — 마지막 식별자만 추출
        var assignedLastProp = assignments
            .Select(a => a.LhsChain.Split('.').Last())
            .ToHashSet();

        // 체인 인덱스 (속성명 → 그 속성명으로 끝나는 모든 체인)
        var chainsByProp = assignments
            .GroupBy(a => a.LhsChain.Split('.').Last())
            .ToDictionary(g => g.Key, g => g.ToList());

        var requiredProps = catalog.Where(p => p.IsRequired).ToList();

        var sb = new StringBuilder();
        sb.AppendLine("# 역방향 갭 리포트 v2 (Globe.cs AST + 매퍼 LHS 체인)");
        sb.AppendLine();
        sb.AppendLine($"- Globe.cs 총 속성: {catalog.Count}개");
        sb.AppendLine($"- 그 중 [Required]: {requiredProps.Count}개");
        sb.AppendLine($"- 매퍼 LHS 체인: {assignments.Count}건");
        sb.AppendLine();
        sb.AppendLine("## 분류 기준");
        sb.AppendLine();
        sb.AppendLine("- 🟢 **고신뢰 ✓**: 속성명이 Globe.cs 에서 유일 (다른 클래스에 없음) AND 매퍼에서 할당됨");
        sb.AppendLine("- 🟡 **중신뢰 ✓?**: 속성명이 여러 클래스에 존재 BUT 매퍼에서 할당된 흔적 있음 — 체인의 부모 컨텍스트로 수동 확인");
        sb.AppendLine("- 🔴 **누락 ✗**: 매퍼 어디에도 할당 흔적 없음");
        sb.AppendLine();

        // 셋업: per-class 카운트
        int green = 0, yellow = 0, red = 0;
        var sbGreen = new StringBuilder();
        var sbYellow = new StringBuilder();
        var sbRed = new StringBuilder();

        sbYellow.AppendLine("## 🟡 중신뢰 (다중 클래스 — 수동 확인)");
        sbYellow.AppendLine();
        sbYellow.AppendLine("| 클래스 | 속성 | 동명 클래스들 | 후보 체인(처음 3개) |");
        sbYellow.AppendLine("|---|---|---|---|");

        sbRed.AppendLine("## 🔴 누락 (매퍼에 할당 없음)");
        sbRed.AppendLine();
        sbRed.AppendLine("| 클래스 | 속성 | 타입 | XmlName |");
        sbRed.AppendLine("|---|---|---|---|");

        foreach (var p in requiredProps.OrderBy(p => p.ClassName).ThenBy(p => p.PropName))
        {
            bool inAssigned = assignedLastProp.Contains(p.PropName);
            var classes = multiplicity[p.PropName];

            if (!inAssigned)
            {
                red++;
                sbRed.Append("| ").Append(p.ClassName)
                     .Append(" | `").Append(p.PropName).Append("`")
                     .Append(" | ").Append(EscapeMd(p.PropType))
                     .Append(" | ").Append(p.XmlName).AppendLine(" |");
            }
            else if (classes.Count == 1)
            {
                green++;
            }
            else
            {
                yellow++;
                var chains = chainsByProp[p.PropName].Take(3).Select(a => $"`{a.LhsChain}`");
                sbYellow.Append("| ").Append(p.ClassName)
                        .Append(" | `").Append(p.PropName).Append("`")
                        .Append(" | ").Append(string.Join(", ", classes))
                        .Append(" | ").Append(string.Join("<br>", chains)).AppendLine(" |");
            }
        }

        sb.AppendLine($"## 종합");
        sb.AppendLine();
        sb.AppendLine($"- 🟢 고신뢰 ✓: **{green}**개");
        sb.AppendLine($"- 🟡 중신뢰 ✓?: **{yellow}**개  (수동 확인 필요)");
        sb.AppendLine($"- 🔴 누락 ✗: **{red}**개  (코드 보강 필요)");
        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine();

        if (red > 0)
        {
            sb.Append(sbRed);
            sb.AppendLine();
        }
        if (yellow > 0)
        {
            sb.Append(sbYellow);
            sb.AppendLine();
        }

        File.WriteAllText(outPath, sb.ToString(), new UTF8Encoding(true));
        Console.WriteLine($"Reverse gap v2 → {outPath}  (green={green}, yellow={yellow}, red={red})");
    }

    // ───────────────────────── 유틸 ─────────────────────────

    static string StripLineComments(string s)
    {
        var sb = new StringBuilder(s.Length);
        bool inStr = false, inChar = false, inVerbatim = false;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            char n = i + 1 < s.Length ? s[i + 1] : '\0';

            if (!inStr && !inChar && c == '@' && n == '"') { inVerbatim = true; inStr = true; sb.Append(c); continue; }
            if (inStr)
            {
                sb.Append(c);
                if (inVerbatim)
                {
                    if (c == '"' && n == '"') { sb.Append(n); i++; continue; }
                    if (c == '"') { inStr = false; inVerbatim = false; }
                }
                else
                {
                    if (c == '\\' && i + 1 < s.Length) { sb.Append(s[i + 1]); i++; continue; }
                    if (c == '"') inStr = false;
                }
                continue;
            }
            if (inChar) { sb.Append(c); if (c == '\'') inChar = false; continue; }

            if (c == '"') { inStr = true; sb.Append(c); continue; }
            if (c == '\'') { inChar = true; sb.Append(c); continue; }
            if (c == '/' && n == '/')
            {
                // 라인 끝까지 스킵 (개행 보존 — 라인 번호 보존을 위해)
                while (i < s.Length && s[i] != '\n') i++;
                if (i < s.Length) sb.Append('\n');
                continue;
            }
            sb.Append(c);
        }
        return sb.ToString();
    }

    static string Csv(string? s)
    {
        if (s == null) return "";
        if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        return s;
    }
    static string EscapeMd(string s) => s.Replace("|", "\\|");
}
