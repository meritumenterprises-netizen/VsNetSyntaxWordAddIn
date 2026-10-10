using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
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
            CheckShortcuts();
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
                foreach (var color in new[] { "#FED9B4", "#E2EFD9" })
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
