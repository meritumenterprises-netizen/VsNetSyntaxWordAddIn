using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
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
        private bool boldHtmlTags;
        private bool darkRedHtmlTags;
        private bool boldAngularSyntax;
        private bool boldKeywords;
        private readonly SyntaxTheme theme;

        public SyntaxHighlighter(SyntaxTheme theme = SyntaxTheme.Light) : base(GetStyles(theme), null) { this.theme = theme; }

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
            return Highlight(text, languageId, false);
        }

        private IReadOnlyList<HighlightSpan> Highlight(string text, string languageId, bool emphasizeKeywords)
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            ILanguage language;
            switch (languageId)
            {
                case "CSharp": language = null; break;
                case "Html": language = AngularHtmlLanguage.Instance; break;
                case "Angular": language = AngularHtmlLanguage.Instance; break;
                case "TypeScript": language = Languages.Typescript; break;
                case "JavaScript": language = Languages.JavaScript; break;
                default: throw new ArgumentException("Unsupported language.", nameof(languageId));
            }

            colors = new int[text.Length];
            bold = new bool[text.Length];
            boldHtmlTags = languageId == "Html" || languageId == "Angular";
            darkRedHtmlTags = languageId == "Html";
            boldAngularSyntax = languageId == "Angular";
            boldKeywords = emphasizeKeywords || boldAngularSyntax;
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
                if (languageId == "Angular") HighlightAngularFragments(source);
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

        private void HighlightAngularFragments(string source)
        {
            // Consume comments first so examples inside comments remain comment-colored.
            foreach (Match match in Regex.Matches(source,
                @"<!--[\s\S]*?(?:-->|$)|\{\{(?<expression>[\s\S]*?)\}\}|" +
                @"""[^""\n]*""|'[^'\n]*'|" +
                @"(?m:^[\t \u00a0]*(?<binding>\[\([\w:.-]+\)\]|\[[\w:.-]+\]|\([\w:.-]+\))(?=\s*=))"))
            {
                Group binding = match.Groups["binding"];
                if (binding.Success)
                {
                    ApplyColor(binding.Index, binding.Length, ScopeName.Keyword);
                    ApplyBold(binding.Index, binding.Length);
                    continue;
                }
                Group expression = match.Groups["expression"];
                if (!expression.Success) continue;
                ApplyColor(match.Index, 2, ScopeName.HtmlTagDelimiter);
                ApplyColor(match.Index + match.Length - 2, 2, ScopeName.HtmlTagDelimiter);
                var expressionHighlighter = new SyntaxHighlighter(theme);
                foreach (var span in expressionHighlighter.Highlight(expression.Value, "JavaScript", true))
                {
                    for (int i = 0; i < span.Length; i++) colors[expression.Index + span.Start + i] = span.Color;
                    if (span.Bold) ApplyBold(expression.Index + span.Start, span.Length);
                }
            }
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
            if (boldHtmlTags && (scope.Name == ScopeName.HtmlElementName || scope.Name == ScopeName.HtmlTagDelimiter))
                ApplyBold(position + scope.Index, scope.Length);
            if ((boldAngularSyntax && scope.Name == ScopeName.HtmlAttributeName) ||
                (boldKeywords && scope.Name == ScopeName.Keyword))
                ApplyBold(position + scope.Index, scope.Length);
            if (Styles.Contains(scope.Name))
            {
                string foreground = darkRedHtmlTags && scope.Name == ScopeName.HtmlElementName
                    ? "#8B0000" : Styles[scope.Name].Foreground;
                if (!string.IsNullOrEmpty(foreground))
                {
                    int color = ToWordColor(foreground);
                    for (int i = scope.Index; i < scope.Index + scope.Length; i++) colors[position + i] = color;
                }
            }
            if (boldAngularSyntax && scope.Name == ScopeName.HtmlAttributeName)
                ApplyColor(position + scope.Index, scope.Length, ScopeName.Keyword);
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
