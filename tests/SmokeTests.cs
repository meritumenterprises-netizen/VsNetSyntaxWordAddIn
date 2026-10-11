using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.CodeAnalysis.CSharp;
using VsNetSyntaxWordAddIn;
using VsNetSyntaxWordAddIn.Highlighting;
using Word = Microsoft.Office.Interop.Word;

internal static class SmokeTests
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            Check("CSharp", "public string Name = \"hello\"; // comment\r\treturn Name;");
            Check("Html", "<div class=\"test\">hello &amp; world</div>\r<!-- comment -->");
            Check("TypeScript", "interface Person { name: string; }\rconst p: Person = { name: 'test' };");
            Check("JavaScript", "const value = \"test\";\r// comment\vconsole.log(value);");
            Check("Json", "{\r\"name\": \"test\", \"count\": -1.25e+2, \"enabled\": true, \"items\": [null, false]}\v");
            Check("Xml", "<?xml version=\"1.0\"?>\r<root id=\"test\"><item>hello &amp; world</item></root>\v<!-- comment -->");
            CheckStructuredDataColors();
            CheckSql();
            CheckXmlAttributes();
            Check("CSharp", "");
            Check("JavaScript", "// \uD83D\uDE00\rconst s = 'Za\u017C\u00F3\u0142\u0107';\r\n");
            CheckThemeColors();
            CheckCSharp12();
            CheckBoldNames();
            CheckHtmlBoldTags();
            CheckWordShortcutAvailability();
            if (!args.Contains("--skip-shortcuts")) CheckShortcuts();
            else Console.WriteLine("Shortcut registration checks skipped (hotkeys may be owned by a running Word instance).");
            if (args.Contains("--word")) CheckWord();
            if (args.Contains("--xml-word")) CheckXmlWord();
            if (args.Contains("--shortcut-input")) CheckShortcutInput();
            Console.WriteLine("All smoke tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }

    private static void Check(string language, string source)
    {
        foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
        {
            var highlighter = new SyntaxHighlighter(theme);
            var spans = highlighter.Highlight(source, language);
            int offset = 0;
            foreach (var span in spans)
            {
                Assert(span.Start == offset && span.Length > 0, language + ": invalid coverage");
                offset += span.Length;
            }
            Assert(offset == source.Length, language + ": text length changed");
            if (source.Length > 0)
                Assert(spans.Any(s => s.Color != highlighter.ForegroundColor), language + ": no syntax colors");
            Console.WriteLine(language + " " + theme + ": offsets and colors OK");
        }
    }

    private static void CheckSql()
    {
        const string source = "select [Name], 'it''s SQL' from dbo.Users\rWHERE Id = 42; -- SELECT comment\v/* FROM comment */";
        Check("Sql", source);
        Check("Sql", "");
        foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
        {
            var highlighter = new SyntaxHighlighter(theme);
            var spans = highlighter.Highlight(source, "Sql");
            foreach (string token in new[] { "select", "WHERE", "it''s SQL", "-- SELECT", "/* FROM" })
                Assert(ColorAt(spans, source.IndexOf(token, StringComparison.Ordinal)) != highlighter.ForegroundColor,
                    "SQL: token uncolored: " + token);
            Assert(ColorAt(spans, source.IndexOf("WHERE", StringComparison.Ordinal)) !=
                ColorAt(spans, source.IndexOf("-- SELECT", StringComparison.Ordinal)), "SQL: comment colored as keyword");
            foreach (string keyword in new[] { "select", "from", "WHERE" })
                Assert(BoldAt(spans, source.IndexOf(keyword, StringComparison.Ordinal)), "SQL: keyword not bold: " + keyword);
            foreach (string plain in new[] { "Name", "it''s SQL", "42", "-- SELECT", "/* FROM" })
                Assert(!BoldAt(spans, source.IndexOf(plain, StringComparison.Ordinal)), "SQL: non-keyword incorrectly bold: " + plain);
            Assert(highlighter.Highlight("const value = 1;", "JavaScript").All(s => !s.Bold), "SQL: bold leaked across languages");
        }
        Console.WriteLine("SQL: bold keywords, normal strings/comments and Word line-break offsets OK");
    }

    private static void CheckThemeColors()
    {
        foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
        {
            var highlighter = new SyntaxHighlighter(theme);
            bool dark = theme == SyntaxTheme.Dark;
            Assert(highlighter.ForegroundColor == Ole(dark ? "#DADADA" : "#000000"), "Theme: plain text color");
            Assert(highlighter.BackgroundColor == Ole(dark ? "#1E1E1E" : "#FFFFFF"), "Theme: background color");
            Assert(highlighter.Highlight("unclassified_identifier", "CSharp").All(s => s.Color == highlighter.ForegroundColor),
                "Theme: unclassified text has incorrect foreground");
            var spans = highlighter.Highlight("public string value = \"hello\";", "CSharp");
            Assert(spans.First().Color == Ole(dark ? "#569CD6" : "#0000FF"), "Theme: keyword color");
            Assert(spans.Single(s => s.Start <= 22 && s.Start + s.Length > 22).Color ==
                Ole(dark ? "#D69D85" : "#A31515"), "Theme: string color");
        }
        Console.WriteLine("Themes: light and dark text, background, keywords, and strings OK");
    }

    private static int Ole(string color) => ColorTranslator.ToOle(ColorTranslator.FromHtml(color));

    private static void CheckStructuredDataColors()
    {
        const string json = "{\r\"name\": \"true \\\"quoted\\\" \\u0041\", \"count\": -1.25e+2, \"enabled\": true, \"items\": [null, false]}\v";
        const string xml = "<?xml version=\"1.0\"?>\r<root id=\"test\"><item>hello &amp; world</item><![CDATA[<raw>]]></root>\v<!-- <ignored/> -->";
        foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
        {
            var styles = theme == SyntaxTheme.Dark
                ? ColorCode.Styling.StyleDictionary.DefaultDark : ColorCode.Styling.StyleDictionary.DefaultLight;
            Func<string, int> scopeColor = scope =>
            {
                string value = styles[scope].Foreground;
                return Ole("#" + value.Substring(value.Length - 6));
            };
            var highlighter = new SyntaxHighlighter(theme);
            var jsonSpans = highlighter.Highlight(json, "Json");
            Assert(ColorAt(jsonSpans, json.IndexOf("name", StringComparison.Ordinal)) == Ole("#0000FF"),
                "JSON: property name color incorrect");
            Assert(ColorAt(jsonSpans, json.IndexOf("quoted", StringComparison.Ordinal)) == Ole("#FF0000"),
                "JSON: escaped string color incorrect");
            int number = json.IndexOf("-1.25e+2", StringComparison.Ordinal);
            for (int i = number; i < number + 8; i++)
                Assert(ColorAt(jsonSpans, i) == Ole("#FF0000"), "JSON: exponent number color incorrect");
            foreach (string literal in new[] { "true", "null", "false" })
                Assert(ColorAt(jsonSpans, json.LastIndexOf(literal, StringComparison.Ordinal)) == Ole("#FF0000"),
                    "JSON: constant color incorrect: " + literal);
            Assert(ColorAt(jsonSpans, json.IndexOf("true", StringComparison.Ordinal)) == Ole("#FF0000"),
                "JSON: constant inside a string recolored");
            foreach (string key in new[] { "name", "count", "enabled", "items" })
                Assert(ColorAt(jsonSpans, json.IndexOf(key, StringComparison.Ordinal)) == Ole("#0000FF"),
                    "JSON: property name is not blue: " + key);
            var xmlSpans = highlighter.Highlight(xml, "Xml");
            Assert(ColorAt(xmlSpans, xml.IndexOf("root", StringComparison.Ordinal)) == Ole("#8B0000"),
                "XML: element name color incorrect");
            Assert(ColorAt(xmlSpans, xml.IndexOf("id=", StringComparison.Ordinal)) == Ole("#8B0000"),
                "XML: attribute name color incorrect");
            Assert(ColorAt(xmlSpans, xml.IndexOf("test", StringComparison.Ordinal)) == Ole("#FF0000"),
                "XML: attribute value color incorrect");
            Assert(ColorAt(xmlSpans, xml.IndexOf("raw", StringComparison.Ordinal)) == Ole("#0000FF"),
                "XML: CDATA color incorrect");
            foreach (string text in new[] { "hello", "&amp;", "world" })
                Assert(ColorAt(xmlSpans, xml.IndexOf(text, StringComparison.Ordinal)) == Ole("#0000FF"),
                    "XML: element text is not blue: " + text);
            const string attributes = "<item first='hello &amp; world' second=\"blue\">red</item>";
            var attributeSpans = highlighter.Highlight(attributes, "Xml");
            foreach (int start in new[] { attributes.IndexOf("item", StringComparison.Ordinal), attributes.LastIndexOf("item", StringComparison.Ordinal) })
                for (int i = start; i < start + 4; i++)
                    Assert(ColorAt(attributeSpans, i) == Ole("#8B0000") && BoldAt(attributeSpans, i),
                        "XML: opening/closing element names must be dark red and bold, including tags with attributes");
            foreach (string name in new[] { "first", "second" })
            {
                int start = attributes.IndexOf(name, StringComparison.Ordinal);
                for (int i = start; i < start + name.Length; i++)
                    Assert(ColorAt(attributeSpans, i) == Ole("#8B0000") && !BoldAt(attributeSpans, i),
                        "XML: attribute names must be dark red and regular weight");
            }
            foreach (string value in new[] { "hello", "&amp;", "world", "blue" })
                Assert(ColorAt(attributeSpans, attributes.IndexOf(value, StringComparison.Ordinal)) == Ole("#FF0000"),
                    "XML: quoted attribute value is not red: " + value);
            Assert(ColorAt(attributeSpans, attributes.IndexOf(">red<", StringComparison.Ordinal) + 1) == Ole("#0000FF"),
                "XML: text color confused with attribute value");
            Assert(ColorAt(xmlSpans, xml.IndexOf("ignored", StringComparison.Ordinal)) == scopeColor(ColorCode.Common.ScopeName.HtmlComment),
                "XML: comment content parsed as markup");
            foreach (string value in new[] { "hello", "&amp;", "world", "blue", ">red<" })
                Assert(!BoldAt(attributeSpans, attributes.IndexOf(value, StringComparison.Ordinal)),
                    "XML: element bold leaked into attribute values or text");
            Assert(!BoldAt(xmlSpans, xml.IndexOf("ignored", StringComparison.Ordinal))
                && !BoldAt(xmlSpans, xml.IndexOf("raw", StringComparison.Ordinal)),
                "XML: markup in comments or CDATA must not be bold");
            Assert(jsonSpans.Sum(span => span.Length) == json.Length && xmlSpans.Sum(span => span.Length) == xml.Length,
                "JSON/XML: paragraph and manual line-break offsets changed");
            Assert(jsonSpans.All(span => !span.Bold), "JSON: XML bold rules applied");
        }
        Console.WriteLine("JSON/XML: token colors, escapes, constants, CDATA, comments, and offsets OK in both themes");
    }

    private static void CheckXmlAttributes()
    {
        foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
        {
            foreach (string source in new[] {
                "<item first=\"\" second=\"value\">text</item>",
                "<item first=\"value\" second=\"\">text</item>",
                "<item first='' second='value'>text</item>",
                "<item first=\"line1\rline2\" second=\"\">text</item>",
                "<item\r first=\"\"\r second=\"value\" />",
                "<item first=\u201dvalue\u201d second=\u201c\u201d>text</item>"
            })
            {
                var spans = new SyntaxHighlighter(theme).Highlight(source, "Xml");
                for (int i = 1; i < 5; i++)
                    Assert(ColorAt(spans, i) == Ole("#8B0000") && BoldAt(spans, i),
                        "XML: element with empty or multiline attributes is not dark red/bold: " + source);
                foreach (string attribute in new[] { "first", "second" })
                {
                    int start = source.IndexOf(attribute, StringComparison.Ordinal);
                    Assert(ColorAt(spans, start) == Ole("#8B0000") && !BoldAt(spans, start),
                        "XML: attribute name has incorrect styling: " + source);
                }
                foreach (string value in new[] { "value", "line1", "line2" })
                {
                    int start = source.IndexOf(value, StringComparison.Ordinal);
                    if (start >= 0)
                        Assert(ColorAt(spans, start) == Ole("#FF0000") && !BoldAt(spans, start),
                            "XML: attribute value has incorrect styling: " + source);
                }
                Assert(spans.Sum(span => span.Length) == source.Length, "XML: attribute offsets changed");
            }
            const string smartQuotes = "<a>\r\n<b>\r\n<c name=\u201dYYY\u201d/>\r\n</b>\r\n</a>";
            var sampleSpans = new SyntaxHighlighter(theme).Highlight(smartQuotes, "Xml");
            int element = smartQuotes.IndexOf("<c", StringComparison.Ordinal) + 1;
            Assert(ColorAt(sampleSpans, element) == Ole("#8B0000") && BoldAt(sampleSpans, element),
                "XML: user's smart-quote example element styling failed");
            int sampleValue = smartQuotes.IndexOf("YYY", StringComparison.Ordinal);
            for (int i = sampleValue; i < sampleValue + 3; i++)
                Assert(ColorAt(sampleSpans, i) == Ole("#FF0000") && !BoldAt(sampleSpans, i),
                    "XML: user's smart-quote example value styling failed");
            Assert(sampleSpans.Sum(span => span.Length) == smartQuotes.Length, "XML: smart-quote normalization changed offsets");
        }
        Console.WriteLine("XML: empty/multiline attributes and self-closing elements OK in both themes");
    }

    private static void CheckCSharp12()
    {
        const string source = "using Point = (int X, int Y);\n" +
            "[System.Runtime.CompilerServices.InlineArray(4)] struct Buffer { private int item; }\n" +
            "public record Person(string Name);\n" +
            "public class Example(int size) {\n" +
            "public required string Name { get; init; }\n" +
            "int[] items = [1, 2, .. new int[] { 3 }];\n" +
            "void Read(ref readonly int value) { }\n" +
            "void Run() { var add = (int x = 1) => x + size; }\n" +
            "}\n";
        var options = new CSharpParseOptions(LanguageVersion.CSharp12);
        Assert(!CSharpSyntaxTree.ParseText(source, options).GetDiagnostics().Any(d =>
            d.Severity == Microsoft.CodeAnalysis.DiagnosticSeverity.Error), "C# 12: feature fixture failed to parse");
        foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
        {
            var highlighter = new SyntaxHighlighter(theme);
            int keyword = Ole(theme == SyntaxTheme.Dark ? "#569CD6" : "#0000FF");
            int literal = Ole(theme == SyntaxTheme.Dark ? "#D69D85" : "#A31515");
            var spans = highlighter.Highlight(source, "CSharp");
            Assert(spans.Sum(s => s.Length) == source.Length, "C# 12: feature offsets changed");
            foreach (string word in new[] { "record", "required", "init", "ref", "readonly", "using", "struct", "class", "var" })
                Assert(ColorAt(spans, source.IndexOf(word, StringComparison.Ordinal)) == keyword,
                    "C# 12: missing keyword " + word + " in " + theme);

            const string strings = "string a = \"\"\"record required\"\"\";\n" +
                "var b = \"\"\"\nrecord\n\"\"\"u8;\n" +
                "string c = $$\"\"\"record {{new object()}}\"\"\";";
            var stringSpans = highlighter.Highlight(strings, "CSharp");
            int search = 0;
            while ((search = strings.IndexOf("record", search, StringComparison.Ordinal)) >= 0)
            {
                Assert(ColorAt(stringSpans, search) == literal, "C# 12: raw string colored as keyword");
                search += 6;
            }
            Assert(ColorAt(stringSpans, strings.IndexOf("new", StringComparison.Ordinal)) == keyword,
                "C# 12: interpolation expression colored as string");

            const string identifiers = "class C { int record, required, @class; void M() { record = required; } }";
            var identifierSpans = highlighter.Highlight(identifiers, "CSharp");
            foreach (string name in new[] { "record", "required", "@class" })
                Assert(ColorAt(identifierSpans, identifiers.IndexOf(name, StringComparison.Ordinal)) == highlighter.ForegroundColor,
                    "C# 12: identifier incorrectly colored as keyword: " + name);

            const string comments = "// record\r/// required\v/* init */\rpublic record P;";
            var commentSpans = highlighter.Highlight(comments, "CSharp");
            Assert(ColorAt(commentSpans, comments.IndexOf("record", StringComparison.Ordinal)) != keyword,
                "C# 12: comment colored as keyword");
            Assert(ColorAt(commentSpans, comments.LastIndexOf("record", StringComparison.Ordinal)) == keyword,
                "C# 12: Word line breaks swallowed a declaration");
            Assert(commentSpans.Sum(s => s.Length) == comments.Length, "C# 12: Word offsets changed");
            var incomplete = highlighter.Highlight("public record Person(string Name", "CSharp");
            Assert(incomplete.Sum(s => s.Length) == 32, "C# 12: incomplete snippet lost offsets");
        }
        Console.WriteLine("C# 12: records, contextual keywords, feature syntax, raw strings, interpolation, identifiers and offsets OK");
    }

    private static int ColorAt(System.Collections.Generic.IReadOnlyList<HighlightSpan> spans, int offset) =>
        spans.Single(s => s.Start <= offset && s.Start + s.Length > offset).Color;

    private static void CheckBoldNames()
    {
        const string csharp = "class Person { public string Name { get; set; } public int field; " +
            "public void Greet() { Name = \"Person Greet Name\"; field = 1; } }\r" +
            "class Use { void Run() { Person p = new Person(); p.Greet(); p.Name = \"x\"; } }\v// Person Greet Name";
        const string typescript = "class Person { Name: string = 'Person Greet Name'; field: number = 1; " +
            "Greet<T>(value: T): void { this.Name = 'x'; } get Display(): string { return this.Name; } }\r" +
            "const p = new Person(); p.Greet<number>(1); p.Name = 'x';\v// Person Greet Name";
        foreach (var theme in new[] { SyntaxTheme.Light, SyntaxTheme.Dark })
        {
            foreach (var fixture in new[] {
                new { Language = "CSharp", Source = csharp },
                new { Language = "TypeScript", Source = typescript }
            })
            {
                var spans = new SyntaxHighlighter(theme).Highlight(fixture.Source, fixture.Language);
                Assert(spans.Sum(s => s.Length) == fixture.Source.Length, "Bold: lost source offsets");
                foreach (var name in new[] { "Person", "Name", "Greet" })
                {
                    int first = fixture.Source.IndexOf(name, StringComparison.Ordinal);
                    if (fixture.Language == "TypeScript" && name == "Greet")
                        first = fixture.Source.IndexOf("Greet<T>", StringComparison.Ordinal);
                    int usage = fixture.Source.IndexOf("p." + name, StringComparison.Ordinal);
                    Assert(BoldAt(spans, first), "Bold: declaration not bold: " + fixture.Language + " " + name);
                    if (usage >= 0) Assert(BoldAt(spans, usage + 2), "Bold: member reference not bold: " + name);
                }
                int comment = fixture.Source.LastIndexOf("Person", StringComparison.Ordinal);
                int literal = fixture.Source.IndexOf("Person Greet Name", StringComparison.Ordinal);
                Assert(!BoldAt(spans, comment) && !BoldAt(spans, literal), "Bold: comment or string changed");
                Assert(!BoldAt(spans, 0), "Bold: class keyword changed");
                if (fixture.Language == "CSharp")
                    Assert(!BoldAt(spans, fixture.Source.IndexOf("field", StringComparison.Ordinal)), "Bold: C# field changed");
                else
                    Assert(BoldAt(spans, fixture.Source.IndexOf("Display", StringComparison.Ordinal)), "Bold: TypeScript accessor missing");
            }
            const string signature = "interface Api { value?: string; run<T>(arg: T): void; }";
            var signatures = new SyntaxHighlighter(theme).Highlight(signature, "TypeScript");
            Assert(BoldAt(signatures, signature.IndexOf("value", StringComparison.Ordinal)) &&
                BoldAt(signatures, signature.IndexOf("run", StringComparison.Ordinal)), "Bold: TypeScript interface members missing");
        }
        Console.WriteLine("Bold: C#/TypeScript declarations, references, accessors, signatures and non-name exclusions OK");
    }

    private static bool BoldAt(System.Collections.Generic.IReadOnlyList<HighlightSpan> spans, int offset) =>
        spans.Single(s => s.Start <= offset && s.Start + s.Length > offset).Bold;

    private static void CheckHtmlBoldTags()
    {
        const string source = "<div class=\"sample\">text &amp; <my-widget data-id=\"1\" />\r</div>\v<!-- <div> -->";
        foreach (var theme in new[] { SyntaxTheme.Light, SyntaxTheme.Dark })
        {
            var highlighter = new SyntaxHighlighter(theme);
            var spans = highlighter.Highlight(source, "Html");
            foreach (int offset in new[] { 0, 1, source.IndexOf("my-widget", StringComparison.Ordinal),
                source.IndexOf("</div>", StringComparison.Ordinal) + 2 })
                Assert(BoldAt(spans, offset), "HTML: tag not bold");
            foreach (int offset in new[] { 1, source.IndexOf("my-widget", StringComparison.Ordinal),
                source.IndexOf("</div>", StringComparison.Ordinal) + 2 })
                Assert(ColorAt(spans, offset) == Ole("#8B0000"), "HTML: tag name not dark red");
            foreach (string plain in new[] { "class", "sample", "text", "&amp;", "data-id" })
                Assert(!BoldAt(spans, source.IndexOf(plain, StringComparison.Ordinal)), "HTML: non-tag incorrectly bold: " + plain);
            Assert(!BoldAt(spans, source.LastIndexOf("div", StringComparison.Ordinal)), "HTML: commented tag bold");
            Assert(spans.Sum(s => s.Length) == source.Length, "HTML: bold changed offsets");
            Assert(highlighter.Highlight("const value = 1;", "JavaScript").All(s => !s.Bold), "HTML: bold leaked across language changes");
            const string interpolation = "<span>{{ true ? 'yes' : 42 }}</span><!-- {{ false }} -->";
            var interpolationSpans = highlighter.Highlight(interpolation, "Angular");
            Assert(ColorAt(interpolationSpans, interpolation.IndexOf("true", StringComparison.Ordinal)) != highlighter.ForegroundColor,
                "Angular: interpolation keyword uncolored");
            var expressionSpans = new SyntaxHighlighter(theme).Highlight(" true ? 'yes' : 42 ", "JavaScript");
            Assert(ColorAt(interpolationSpans, interpolation.IndexOf("42", StringComparison.Ordinal)) ==
                ColorAt(expressionSpans, 16), "Angular: interpolation expression palette differs");
            Assert(ColorAt(interpolationSpans, interpolation.IndexOf("false", StringComparison.Ordinal)) ==
                ColorAt(interpolationSpans, interpolation.IndexOf("<!--", StringComparison.Ordinal)), "Angular: interpolation changed comment colors");
            Assert(interpolationSpans.Sum(s => s.Length) == interpolation.Length, "Angular: interpolation offsets changed");
            Assert(BoldAt(interpolationSpans, interpolation.IndexOf("true", StringComparison.Ordinal)),
                "Angular: interpolation keyword not bold");
            foreach (string plain in new[] { "yes", "42", "false" })
                Assert(!BoldAt(interpolationSpans, interpolation.IndexOf(plain, StringComparison.Ordinal)),
                    "Angular: value or comment incorrectly bold: " + plain);
            var attributeSpans = highlighter.Highlight(source, "Angular");
            foreach (string attributeName in new[] { "class", "data-id" })
                Assert(BoldAt(attributeSpans, source.IndexOf(attributeName, StringComparison.Ordinal)),
                    "Angular: standard attribute not bold: " + attributeName);
            foreach (string angular in new[] {
                "<button (click)=\"save()\">Save</button>",
                "<button [disabled]=\"busy\" (click)=\"save()\">Save</button>",
                "<input [(ngModel)]=\"name\" #editor />",
                "<div *ngIf=\"visible\" [attr.aria-label]=\"label\">Text</div>",
                "<button\r(click)=\"save()\"\v[disabled]=\"busy\">Save</button>"
            })
            {
                var angularSpans = highlighter.Highlight(angular, "Angular");
                int nameEnd = angular.IndexOfAny(new[] { ' ', '\r', '\v' });
                for (int offset = 0; offset < nameEnd; offset++)
                    Assert(BoldAt(angularSpans, offset), "HTML: Angular opening tag not bold");
                Assert(ColorAt(angularSpans, 1) != highlighter.ForegroundColor, "HTML: Angular opening tag uncolored");
                int close = angular.IndexOf("</", StringComparison.Ordinal);
                if (close >= 0)
                    Assert(ColorAt(angularSpans, 1) == ColorAt(angularSpans, close + 2) &&
                        BoldAt(angularSpans, close + 2), "HTML: Angular opening/closing tag styles differ");
                Assert(angularSpans.Sum(s => s.Length) == angular.Length, "HTML: Angular offsets changed");
                int attribute = angular.IndexOfAny(new[] { '(', '[', '*', '#' });
                Assert(BoldAt(angularSpans, attribute), "Angular: attribute not bold");
                Assert(ColorAt(angularSpans, attribute) == Ole(theme == SyntaxTheme.Dark ? "#569CD6" : "#0000FF"),
                    "Angular: attribute not blue");
                int value = angular.IndexOf('"') + 1;
                Assert(!BoldAt(angularSpans, value), "Angular: attribute value incorrectly bold");
            }
            Assert(highlighter.Highlight("const value = true;", "JavaScript").All(s => !s.Bold),
                "Angular: keyword bold leaked across language changes");
            const string standalone = "{{ customer.name }}\u00a0\u00a0\r\n" +
                "[disabled]=\"isLoading\"\u00a0\r\n" +
                "(click)=\"save()\"\u00a0\r\n" +
                "[(ngModel)]=\"customer.name\"\u00a0\r\n" +
                "<!--\r\n[hidden]=\"false\"\r\n-->";
            var standaloneSpans = highlighter.Highlight(standalone, "Angular");
            foreach (string binding in new[] { "disabled", "click", "ngModel" })
            {
                int start = standalone.IndexOf(binding, StringComparison.Ordinal);
                for (int i = start; i < start + binding.Length; i++)
                {
                    Assert(BoldAt(standaloneSpans, i), "Angular: standalone binding not bold: " + binding);
                    Assert(ColorAt(standaloneSpans, i) == Ole(theme == SyntaxTheme.Dark ? "#569CD6" : "#0000FF"),
                        "Angular: standalone binding not blue: " + binding);
                }
            }
            foreach (string plain in new[] { "customer.name", "isLoading", "save", "hidden", "false" })
                Assert(!BoldAt(standaloneSpans, standalone.IndexOf(plain, StringComparison.Ordinal)),
                    "Angular: standalone expression or comment incorrectly bold: " + plain);
            Assert(standaloneSpans.Sum(s => s.Length) == standalone.Length,
                "Angular: standalone binding offsets changed");
            Assert(highlighter.Highlight(standalone, "Html").All(s => !s.Bold),
                "Angular: standalone binding bold leaked into HTML");
        }
        Console.WriteLine("HTML tags bold; Angular attributes and interpolation keywords bold; values and comments normal OK");
    }

    private static void CheckShortcuts()
    {
        bool active = true;
        string applied = null;
        using (var shortcuts = new GlobalSyntaxShortcuts(language => applied = language, () => active))
        {
            Assert(shortcuts.RegisteredCount == 8, "Shortcuts: native registration failed: " + shortcuts.Status);
            foreach (var mapping in new[] {
                new { Key = 0x43, Language = "CSharp" },
                new { Key = 0x4A, Language = "JavaScript" },
                new { Key = 0x54, Language = "TypeScript" },
                new { Key = 0x48, Language = "Html" },
                new { Key = 0x41, Language = "Angular" },
                new { Key = 0x53, Language = "Json" },
                new { Key = 0x58, Language = "Xml" },
                new { Key = 0x51, Language = "Sql" }
            })
            {
                SendMessage(shortcuts.Handle, 0x0312, new IntPtr(mapping.Key), new IntPtr((mapping.Key << 16) | 7));
                Assert(applied == mapping.Language, "Shortcuts: native message language mapping failed");
            }
            applied = null;
            SendMessage(shortcuts.Handle, 0x0312, new IntPtr(0x43), new IntPtr((0x43 << 16) | 3));
            Assert(applied == null, "Shortcuts: Ctrl+Alt+C incorrectly dispatched");
            Assert(RegisterHotKey(shortcuts.Handle, 900, 3, 0x43), "Shortcuts: Ctrl+Alt+C was reserved by the add-in");
            UnregisterHotKey(shortcuts.Handle, 900);
            active = false;
            SendMessage(shortcuts.Handle, 0x0312, new IntPtr(0x43), new IntPtr((0x43 << 16) | 7));
            Assert(applied == null, "Shortcuts: formatted outside Word");
            shortcuts.Refresh();
            Assert(shortcuts.RegisteredCount == 0, "Shortcuts: focus loss did not release hotkeys");
            active = true;
            shortcuts.Refresh();
            Assert(shortcuts.RegisteredCount == 8, "Shortcuts: focus return did not register hotkeys");
        }
        using (var shortcuts = new GlobalSyntaxShortcuts(language => { }, () => true))
            Assert(shortcuts.RegisteredCount == 8, "Shortcuts: disposal did not release hotkeys");
        Console.WriteLine("Shortcuts: native registration, dispatch, focus, and cleanup OK");
    }

    private static void CheckWordShortcutAvailability()
    {
        Assert(!GlobalSyntaxShortcuts.MatchesShortcut(0x43, (0x43 << 16) | 3), "Ctrl+Alt+C matched add-in shortcut");
        Assert(GlobalSyntaxShortcuts.MatchesShortcut(0x43, (0x43 << 16) | 7), "Full C# shortcut rejected");
        using (var window = new GlobalSyntaxShortcuts(language => { }, () => false))
        {
            Assert(RegisterHotKey(window.Handle, 901, 3, 0x43), "Ctrl+Alt+C is reserved globally");
            UnregisterHotKey(window.Handle, 901);
        }
        Console.WriteLine("Word shortcut: Ctrl+Alt+C does not match add-in modifiers and is not globally reserved OK");
    }

    private static void CheckWord()
    {
        var application = new Word.Application { Visible = false };
        Word.Document document = null;
        try
        {
            document = application.Documents.Add();
            const string prefix = "Before\r";
            const string code = "public string Name = \"hello\"; // \uD83D\uDE00\r\treturn Name;\v";
            document.Content.Text = prefix + code + "After\r";
            string original = document.Content.Text;
            var range = document.Range(prefix.Length, prefix.Length + code.Length);
            var initialBackground = range.Shading.BackgroundPatternColor;
            range.Select();
            WordSelectionFormatter.Format(application, "CSharp");
            Assert(document.Content.Text == original, "Word: content changed");
            Assert(application.Selection.Range.Start == prefix.Length, "Word: selection start changed");
            Assert(application.Selection.Range.End == prefix.Length + code.Length, "Word: selection end changed");
            Assert(document.Range(prefix.Length, prefix.Length + 6).Font.Color ==
                (Word.WdColor)ColorTranslator.ToOle(Color.Blue), "Word: keyword color incorrect");
            Assert(document.Range(prefix.Length + code.IndexOf("\"hello\"", StringComparison.Ordinal),
                prefix.Length + code.IndexOf("\"hello\"", StringComparison.Ordinal) + 7).Font.Color ==
                (Word.WdColor)ColorTranslator.ToOle(ColorTranslator.FromHtml("#A31515")), "Word: string color incorrect");
            Assert(range.Shading.BackgroundPatternColor == initialBackground, "Word: Normal changed background");
            Assert(document.Range(0, prefix.Length).Font.Name != "Consolas", "Word: prefix formatting changed");
            Assert(document.Undo(), "Word: undo failed");
            Assert(document.Content.Text == original && range.Font.Name != "Consolas", "Word: grouped undo failed");
            var originalBackground = range.Shading.BackgroundPatternColor;
            var prefixBackground = document.Range(0, prefix.Length).Shading.BackgroundPatternColor;
            foreach (var theme in new[] { SyntaxTheme.Dark, SyntaxTheme.Light })
            {
                bool dark = theme == SyntaxTheme.Dark;
                range.Select();
                WordSelectionFormatter.Format(application, "CSharp", theme);
                Assert(range.Shading.BackgroundPatternColor == originalBackground,
                    "Word: theme changed Normal background");
                Assert(document.Range(prefix.Length, prefix.Length + 6).Font.Color ==
                    (Word.WdColor)Ole(dark ? "#569CD6" : "#0000FF"), "Word: themed keyword incorrect");
                int stringStart = prefix.Length + code.IndexOf("\"hello\"", StringComparison.Ordinal);
                Assert(document.Range(stringStart, stringStart + 7).Font.Color ==
                    (Word.WdColor)Ole(dark ? "#D69D85" : "#A31515"), "Word: themed string incorrect");
                int nameStart = prefix.Length + code.IndexOf("Name", StringComparison.Ordinal);
                Assert(document.Range(nameStart, nameStart + 4).Font.Color ==
                    (Word.WdColor)new SyntaxHighlighter(theme).ForegroundColor, "Word: themed plain text incorrect");
                Assert(document.Content.Text == original && application.Selection.Range.Start == prefix.Length &&
                    application.Selection.Range.End == prefix.Length + code.Length, "Word: theme changed content or selection");
                Assert(document.Range(0, prefix.Length).Shading.BackgroundPatternColor == prefixBackground,
                    "Word: theme changed surrounding background");
            }
            Assert(document.Undo() && range.Shading.BackgroundPatternColor == originalBackground,
                "Word: undo did not restore dark theme");
            Assert(document.Undo() && range.Shading.BackgroundPatternColor == originalBackground && range.Font.Name != "Consolas",
                "Word: undo did not restore original formatting");
            foreach (var theme in new[] { SyntaxTheme.Light, SyntaxTheme.Dark })
            {
                foreach (var color in new[] { "#FED9B4", "#EDF5E7" })
                {
                    range.Select();
                    WordSelectionFormatter.Format(application, "CSharp", theme, Ole(color));
                    Assert(range.Shading.BackgroundPatternColor == (Word.WdColor)Ole(color), "Word: custom background incorrect");
                    Assert(document.Range(0, prefix.Length).Shading.BackgroundPatternColor == prefixBackground,
                        "Word: custom background changed prefix");
                    range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                    WordSelectionFormatter.Format(application, "CSharp", theme);
                    Assert(range.Shading.BackgroundPatternColor == (Word.WdColor)Ole(color) &&
                        range.HighlightColorIndex == Word.WdColorIndex.wdYellow, "Word: Normal changed existing background formatting");
                    Assert(document.Undo() && range.Shading.BackgroundPatternColor == (Word.WdColor)Ole(color),
                        "Word: Normal undo changed background");
                }
            }
            Console.WriteLine("Word: Normal preserves shading/highlight; Red and Green work in both themes OK");
            Console.WriteLine("Word: light/dark switching, background, plain text, and themed undo OK");
            Console.WriteLine("Word: text, selection, colors, boundaries, and grouped undo OK");
            const string modernCode = "public record Person(string Name);\rpublic class C(int size) { public required string Name { get; init; } int[] values = [1, 2]; }\r";
            document.Content.Text = modernCode;
            var modernRange = document.Range(0, modernCode.Length);
            foreach (var theme in new[] { SyntaxTheme.Light, SyntaxTheme.Dark })
            {
                modernRange.Select();
                WordSelectionFormatter.Format(application, "CSharp", theme);
                foreach (var word in new[] { "record", "required", "init" })
                {
                    int offset = modernCode.IndexOf(word, StringComparison.Ordinal);
                    Assert(document.Range(offset, offset + word.Length).Font.Color ==
                        (Word.WdColor)Ole(theme == SyntaxTheme.Dark ? "#569CD6" : "#0000FF"), "Word: modern keyword not highlighted: " + word);
                }
                Assert(modernRange.Text == modernCode, "Word: C# 12 formatting changed text");
            }
            Console.WriteLine("Word: C# 12 keywords and unchanged text verified in both themes OK");
            foreach (var language in new[] { "CSharp", "TypeScript", "Html", "Angular" })
            {
                string boldCode = language == "CSharp"
                    ? "class Person { public string Name { get; set; } void Greet() { } }\r"
                    : language == "TypeScript" ? "class Person { Name: string; Greet(): void { } }\r"
                    : "<button (click)=\"save()\">Save</button>\r";
                document.Content.Text = boldCode;
                var boldRange = document.Range(0, boldCode.Length);
                foreach (var theme in new[] { SyntaxTheme.Light, SyntaxTheme.Dark })
                {
                    boldRange.Select();
                    WordSelectionFormatter.Format(application, language, theme);
                    foreach (var name in language == "Html" || language == "Angular" ? new[] { "button" } : new[] { "Person", "Name", "Greet" })
                    {
                        int offset = boldCode.IndexOf(name, StringComparison.Ordinal);
                        Assert(document.Range(offset, offset + name.Length).Font.Bold == -1, "Word: name not bold: " + language + " " + name);
                    }
                    if (language == "Html" || language == "Angular")
                    {
                        Assert(document.Range(8, 15).Font.Bold == (language == "Angular" ? -1 : 0),
                            "Word: attribute bold does not match language");
                        int close = boldCode.IndexOf("</", StringComparison.Ordinal);
                        Assert(document.Range(close, close + 9).Font.Bold == -1, "Word: HTML closing tag not bold");
                        Assert(document.Range(1, 7).Font.Color == document.Range(close + 2, close + 8).Font.Color,
                            "Word: Angular opening/closing tag colors differ");
                    }
                    else
                        Assert(document.Range(0, 5).Font.Bold == 0, "Word: keyword incorrectly bold");
                    Assert(boldRange.Text == boldCode, "Word: bold changed text");
                    Assert(document.Undo(), "Word: bold undo failed");
                }
                boldRange.Select();
                WordSelectionFormatter.Format(application, language);
                WordSelectionFormatter.Format(application, "JavaScript");
                Assert(boldRange.Font.Bold == 0, "Word: bold leaked when changing languages");
            }
            Console.WriteLine("Word: C#/TypeScript bold names and HTML tags, reset and grouped undo verified in both themes OK");
        }
        finally
        {
            if (document != null) document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (application.Documents.Count == 0) application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
    }

    private static void CheckXmlWord()
    {
        var application = new Word.Application { Visible = false };
        Word.Document document = null;
        try
        {
            document = application.Documents.Add();
            const string source = "<a>\r<b>\r<c name=\u201dYYY\u201d/>\r</b>\r</a>\r";
            document.Content.Text = source;
            foreach (SyntaxTheme theme in Enum.GetValues(typeof(SyntaxTheme)))
            {
                var range = document.Range(0, source.Length);
                range.Select();
                WordSelectionFormatter.Format(application, "Xml", theme);
                int element = source.IndexOf("<c", StringComparison.Ordinal) + 1;
                var elementRange = document.Range(element, element + 1);
                Assert(elementRange.Font.Color == (Word.WdColor)Ole("#8B0000") && elementRange.Font.Bold == -1,
                    "Word XML: c is not dark red and bold");
                int attribute = source.IndexOf("name", StringComparison.Ordinal);
                var attributeRange = document.Range(attribute, attribute + 4);
                Assert(attributeRange.Font.Color == (Word.WdColor)Ole("#8B0000") && attributeRange.Font.Bold == 0,
                    "Word XML: name is not dark red and regular weight");
                int value = source.IndexOf("YYY", StringComparison.Ordinal);
                var valueRange = document.Range(value, value + 3);
                Assert(valueRange.Font.Color == (Word.WdColor)Ole("#FF0000") && valueRange.Font.Bold == 0,
                    "Word XML: YYY is not red and regular weight");
                Assert(range.Text == source && application.Selection.Range.Start == 0
                    && application.Selection.Range.End == source.Length,
                    "Word XML: smart quotes, text, or selection changed");
                Assert(document.Undo(), "Word XML: grouped undo failed");
            }
            Console.WriteLine("Word XML: exact smart-quote sample colors, bold, original text, selection, and undo OK in both themes");
        }
        finally
        {
            if (document != null) document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (application.Documents.Count == 0) application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
    }

    private static void CheckShortcutInput()
    {
        IntPtr originalForeground = GetForegroundWindow();
        var application = new Word.Application { Visible = true };
        Word.Document document = null;
        try
        {
            document = application.Documents.Add();
            document.Content.Text = "public string Name = \"hello\";\r";
            var addin = application.COMAddIns.Item("VsNetSyntaxWordAddIn");
            Assert(addin.Connect, "Shortcut input: add-in is not loaded");
            dynamic diagnostics = addin.Object;
            Console.WriteLine("Shortcut input before: " + diagnostics.Status);
            document.Activate();
            IntPtr window = GetAncestor(new IntPtr(application.ActiveWindow.Hwnd), 2);
            Assert(SetForegroundWindow(window), "Shortcut input: could not activate temporary Word window");
            var registrationDeadline = DateTime.UtcNow.AddSeconds(5);
            while (!((string)diagnostics.Status).StartsWith("Registered=8/8") && DateTime.UtcNow < registrationDeadline)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(50);
            }
            Assert(((string)diagnostics.Status).StartsWith("Registered=8/8"), "Shortcut input: hotkeys were not registered");
            foreach (ushort key in new ushort[] { 0x43, 0x4A, 0x54, 0x48, 0x41, 0x53, 0x58, 0x51 })
            {
                var selection = document.Range(0, document.Content.End - 1);
                selection.Font.Name = "Courier New";
                selection.Select();
                var input = new[] { KeyInput(0xA2, false), KeyInput(0xA0, false), KeyInput(0xA4, false), KeyInput(key, false),
                    KeyInput(key, true), KeyInput(0xA4, true), KeyInput(0xA0, true), KeyInput(0xA2, true) };
                Assert(SendInput((uint)input.Length, input, Marshal.SizeOf(typeof(Input))) == input.Length,
                    "Shortcut input: SendInput failed");
                var deadline = DateTime.UtcNow.AddSeconds(5);
                while (selection.Font.Name != "Consolas" && DateTime.UtcNow < deadline)
                {
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(50);
                }
                Console.WriteLine("Shortcut input after: " + diagnostics.Status);
                Assert(selection.Font.Name == "Consolas", "Shortcut input: Ctrl+Shift+Alt+" + (char)key + " did not format text");
                Console.WriteLine("Shortcut input: Ctrl+Shift+Alt+" + (char)key + " formatted temporary Word selection");
            }
        }
        finally
        {
            if (document != null) document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (application.Documents.Count == 0) application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            SetForegroundWindow(originalForeground);
        }
    }

    private static Input KeyInput(ushort key, bool release) => new Input
    {
        Type = 1,
        Data = new InputData { Keyboard = new KeyboardInput { Key = key, Flags = release ? 2u : 0u } }
    };

    [StructLayout(LayoutKind.Sequential)]
    private struct Input { public uint Type; public InputData Data; }
    [StructLayout(LayoutKind.Explicit)]
    private struct InputData
    {
        [FieldOffset(0)] public KeyboardInput Keyboard;
        [FieldOffset(0)] public MouseInput Mouse;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput { public ushort Key, Scan; public uint Flags, Time; public UIntPtr Extra; }
    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput { public int X, Y; public uint Data, Flags, Time; public UIntPtr Extra; }
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint count, Input[] input, int size);
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterHotKey(IntPtr window, int id);
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr window);

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
