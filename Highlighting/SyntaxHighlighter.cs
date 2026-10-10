using System;
using System.Collections.Generic;
using System.Drawing;
using ColorCode;
using ColorCode.Common;
using ColorCode.Parsing;
using ColorCode.Styling;

namespace VsNetSyntaxWordAddIn.Highlighting
{
    public enum SyntaxTheme
    {
        Light,
        Dark
    }

    public sealed class HighlightSpan
    {
        public int Start { get; }
        public int Length { get; }
        public int Color { get; }
        public bool Bold { get; }

        public HighlightSpan(int start, int length, int color, bool bold = false)
        {
            Start = start;
            Length = length;
            Color = color;
            Bold = bold;
        }
    }

    public sealed class SyntaxHighlighter : CodeColorizerBase
    {
        private int[] colors;
        private bool[] bold;
        private int position;

        public SyntaxHighlighter(SyntaxTheme theme = SyntaxTheme.Light) : base(GetStyles(theme), null) { }

        public int ForegroundColor => ToWordColor(Styles[ScopeName.PlainText].Foreground);
        public int BackgroundColor => ToWordColor(Styles[ScopeName.PlainText].Background);

        private static StyleDictionary GetStyles(SyntaxTheme theme)
        {
            switch (theme)
            {
                case SyntaxTheme.Light: return StyleDictionary.DefaultLight;
                case SyntaxTheme.Dark: return StyleDictionary.DefaultDark;
                default: throw new ArgumentOutOfRangeException(nameof(theme));
            }
        }

        public IReadOnlyList<HighlightSpan> Highlight(string text, string languageId)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            ILanguage language;
            switch (languageId)
            {
                case "CSharp": language = null; break;
                case "Html": language = Languages.Html; break;
                case "TypeScript": language = Languages.Typescript; break;
                case "JavaScript": language = Languages.JavaScript; break;
                default: throw new ArgumentException("Unsupported language.", nameof(languageId));
            }

            colors = new int[text.Length];
            bold = new bool[text.Length];
            int foreground = ForegroundColor;
            for (int i = 0; i < colors.Length; i++) colors[i] = foreground;
            position = 0;
            // Word paragraph and manual line breaks must retain their original offsets.
            var source = text.Replace('\r', '\n').Replace('\v', '\n');
            if (languageId == "CSharp")
                CSharpSyntaxHighlighter.Highlight(source, ApplyColor, ApplyBold);
            else
            {
                languageParser.Parse(source, language, Write);
                if (position != text.Length) throw new InvalidOperationException("The highlighter did not preserve source offsets.");
                if (languageId == "TypeScript") TypeScriptNameHighlighter.Highlight(source, ApplyBold);
            }

            var spans = new List<HighlightSpan>();
            for (int start = 0; start < colors.Length;)
            {
                int end = start + 1;
                while (end < colors.Length && colors[end] == colors[start] && bold[end] == bold[start]) end++;
                spans.Add(new HighlightSpan(start, end - start, colors[start], bold[start]));
                start = end;
            }
            return spans;
        }

        protected override void Write(string parsedSourceCode, IList<Scope> scopes)
        {
            foreach (var scope in scopes) ApplyScope(scope, parsedSourceCode.Length);
            position += parsedSourceCode.Length;
        }

        private void ApplyScope(Scope scope, int chunkLength)
        {
            if (scope.Index < 0 || scope.Length < 0 || scope.Index + scope.Length > chunkLength)
                throw new InvalidOperationException("The parser produced an invalid token range.");
            if (Styles.Contains(scope.Name))
            {
                string foreground = Styles[scope.Name].Foreground;
                if (!string.IsNullOrEmpty(foreground))
                {
                    int color = ToWordColor(foreground);
                    for (int i = scope.Index; i < scope.Index + scope.Length; i++) colors[position + i] = color;
                }
            }
            foreach (var child in scope.Children) ApplyScope(child, chunkLength);
        }

        private void ApplyColor(int start, int length, string scopeName)
        {
            if (start < 0 || length < 0 || start + length > colors.Length)
                throw new InvalidOperationException("The C# parser produced an invalid token range.");
            if (!Styles.Contains(scopeName)) return;
            string foreground = Styles[scopeName].Foreground;
            if (string.IsNullOrEmpty(foreground)) return;
            int color = ToWordColor(foreground);
            for (int i = start; i < start + length; i++) colors[i] = color;
        }

        private void ApplyBold(int start, int length)
        {
            if (start < 0 || length < 0 || start + length > bold.Length)
                throw new InvalidOperationException("The parser produced an invalid name range.");
            for (int i = start; i < start + length; i++) bold[i] = true;
        }

        private static int ToWordColor(string value)
        {
            string hex = value.TrimStart('#');
            if (hex.Length == 8) hex = hex.Substring(2);
            return ColorTranslator.ToOle(ColorTranslator.FromHtml("#" + hex));
        }
    }
}
