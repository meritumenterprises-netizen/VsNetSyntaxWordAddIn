using System;
using System.Collections.Generic;
using ColorCode;

namespace VsNetSyntaxWordAddIn.Highlighting
{
    internal sealed class AngularHtmlLanguage : ILanguage
    {
        internal static readonly ILanguage Instance = new AngularHtmlLanguage();

        public string Id => "html-angular";
        public string Name => "HTML";
        public string CssClassName => Languages.Html.CssClassName;
        public string FirstLinePattern => Languages.Html.FirstLinePattern;
        public IList<LanguageRule> Rules { get; } = new List<LanguageRule>();

        private AngularHtmlLanguage()
        {
            // Extend only attribute names, preserving ColorCode's scopes, script handling and tag grammar.
            foreach (var rule in Languages.Html.Rules)
                Rules.Add(new LanguageRule(rule.Regex.Replace("[a-z0-9-_]+", "[^\\s\\\"'<>/=]+"),
                    new Dictionary<int, string>(rule.Captures)));
        }

        public bool HasAlias(string alias) => string.Equals(alias, Id, StringComparison.OrdinalIgnoreCase);
    }
}
