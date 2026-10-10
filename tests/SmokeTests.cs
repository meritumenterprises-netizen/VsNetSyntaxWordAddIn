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
            Check("CSharp", "");
            Check("JavaScript", "// \uD83D\uDE00\rconst s = 'Za\u017C\u00F3\u0142\u0107';\r\n");
            CheckThemeColors();
            CheckCSharp12();
            CheckBoldNames();
            if (!args.Contains("--skip-shortcuts")) CheckShortcuts();
            else Console.WriteLine("Shortcut registration checks skipped (hotkeys may be owned by a running Word instance).");
            if (args.Contains("--word")) CheckWord();
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

    private static void CheckShortcuts()
    {
        bool active = true;
        string applied = null;
        using (var shortcuts = new GlobalSyntaxShortcuts(language => applied = language, () => active))
        {
            Assert(shortcuts.RegisteredCount == 4, "Shortcuts: native registration failed: " + shortcuts.Status);
            foreach (var mapping in new[] {
                new { Key = 0x43, Language = "CSharp" },
                new { Key = 0x4A, Language = "JavaScript" },
                new { Key = 0x54, Language = "TypeScript" },
                new { Key = 0x48, Language = "Html" }
            })
            {
                SendMessage(shortcuts.Handle, 0x0312, new IntPtr(mapping.Key), IntPtr.Zero);
                Assert(applied == mapping.Language, "Shortcuts: native message language mapping failed");
            }
            applied = null;
            active = false;
            SendMessage(shortcuts.Handle, 0x0312, new IntPtr(0x43), IntPtr.Zero);
            Assert(applied == null, "Shortcuts: formatted outside Word");
            shortcuts.Refresh();
            Assert(shortcuts.RegisteredCount == 0, "Shortcuts: focus loss did not release hotkeys");
            active = true;
            shortcuts.Refresh();
            Assert(shortcuts.RegisteredCount == 4, "Shortcuts: focus return did not register hotkeys");
        }
        using (var shortcuts = new GlobalSyntaxShortcuts(language => { }, () => true))
            Assert(shortcuts.RegisteredCount == 4, "Shortcuts: disposal did not release hotkeys");
        Console.WriteLine("Shortcuts: native registration, dispatch, focus, and cleanup OK");
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
            foreach (var language in new[] { "CSharp", "TypeScript" })
            {
                string boldCode = language == "CSharp"
                    ? "class Person { public string Name { get; set; } void Greet() { } }\r"
                    : "class Person { Name: string; Greet(): void { } }\r";
                document.Content.Text = boldCode;
                var boldRange = document.Range(0, boldCode.Length);
                foreach (var theme in new[] { SyntaxTheme.Light, SyntaxTheme.Dark })
                {
                    boldRange.Select();
                    WordSelectionFormatter.Format(application, language, theme);
                    foreach (var name in new[] { "Person", "Name", "Greet" })
                    {
                        int offset = boldCode.IndexOf(name, StringComparison.Ordinal);
                        Assert(document.Range(offset, offset + name.Length).Font.Bold == -1, "Word: name not bold: " + language + " " + name);
                    }
                    Assert(document.Range(0, 5).Font.Bold == 0, "Word: keyword incorrectly bold");
                    Assert(boldRange.Text == boldCode, "Word: bold changed text");
                    Assert(document.Undo(), "Word: bold undo failed");
                }
                boldRange.Select();
                WordSelectionFormatter.Format(application, language);
                WordSelectionFormatter.Format(application, "JavaScript");
                Assert(boldRange.Font.Bold == 0, "Word: bold leaked when changing languages");
            }
            Console.WriteLine("Word: C#/TypeScript bold names, reset and grouped undo verified in both themes OK");
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
            while (!((string)diagnostics.Status).StartsWith("Registered=4/4") && DateTime.UtcNow < registrationDeadline)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(50);
            }
            Assert(((string)diagnostics.Status).StartsWith("Registered=4/4"), "Shortcut input: hotkeys were not registered");
            foreach (ushort key in new ushort[] { 0x43, 0x4A, 0x54, 0x48 })
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
