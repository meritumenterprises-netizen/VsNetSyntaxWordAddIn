using System;
using System.Drawing;
using System.Linq;
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
            if (args.Contains("--word")) CheckWord();
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
        var spans = new SyntaxHighlighter().Highlight(source, language);
        int offset = 0;
        foreach (var span in spans)
        {
            Assert(span.Start == offset && span.Length > 0, language + ": invalid coverage");
            offset += span.Length;
        }
        Assert(offset == source.Length, language + ": text length changed");
        if (source.Length > 0) Assert(spans.Any(s => s.Color != 0), language + ": no syntax colors");
        Console.WriteLine(language + ": offsets and colors OK");
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
            Assert(document.Range(0, prefix.Length).Font.Name != "Consolas", "Word: prefix formatting changed");
            Assert(document.Undo(), "Word: undo failed");
            Assert(document.Content.Text == original && range.Font.Name != "Consolas", "Word: grouped undo failed");
            Console.WriteLine("Word: text, selection, colors, boundaries, and grouped undo OK");
        }
        finally
        {
            if (document != null) document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
