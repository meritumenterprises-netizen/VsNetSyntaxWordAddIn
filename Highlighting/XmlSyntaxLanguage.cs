using System;
using System.Collections.Generic;
using ColorCode;

namespace VsNetSyntaxWordAddIn.Highlighting
{
    internal sealed class XmlSyntaxLanguage : ILanguage
    {
        internal static readonly ILanguage Instance = new XmlSyntaxLanguage();

        public string Id => "xml-word";
        public string Name => "XML";
        public string CssClassName => Languages.Xml.CssClassName;
        public string FirstLinePattern => Languages.Xml.FirstLinePattern;
        public IList<LanguageRule> Rules { get; } = new List<LanguageRule>();

        private XmlSyntaxLanguage()
        {
            // Permit empty/multiline values without consuming the next attribute's quotes.
            foreach (var rule in Languages.Xml.Rules)
                Rules.Add(new LanguageRule(rule.Regex
                    .Replace("(\")([^\\n]+?)(\")", "(\")([^\"]*)(\")")
                    .Replace("(')([^\\n]+?)(')", "(')([^']*)(')"),
                    new Dictionary<int, string>(rule.Captures)));
        }

        public bool HasAlias(string alias) => string.Equals(alias, Id, StringComparison.OrdinalIgnoreCase);
    }
}
