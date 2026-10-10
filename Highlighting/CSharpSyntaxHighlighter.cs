using System;
using ColorCode.Common;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace VsNetSyntaxWordAddIn.Highlighting
{
    internal static class CSharpSyntaxHighlighter
    {
        private static readonly CSharpParseOptions Options = new CSharpParseOptions(LanguageVersion.CSharp12);
        private static readonly MetadataReference CoreReference = MetadataReference.CreateFromFile(typeof(object).Assembly.Location);

        internal static void Highlight(string source, Action<int, int, string> apply, Action<int, int> applyBold)
        {
            var tree = CSharpSyntaxTree.ParseText(source, Options);
            var root = tree.GetRoot();
            var model = CSharpCompilation.Create("SelectedCode", new[] { tree }, new[] { CoreReference },
                new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary)).GetSemanticModel(tree);
            // Include structured trivia so directives and recoverable skipped tokens are retained.
            foreach (var token in root.DescendantTokens(descendIntoTrivia: true))
            {
                string scope = GetScope(token);
                if (scope != null) apply(token.SpanStart, token.Span.Length, scope);
                if (IsBoldName(token, model)) applyBold(token.SpanStart, token.Span.Length);
            }
            // Trivia overrides XML tokens inside documentation comments.
            foreach (var trivia in root.DescendantTrivia(descendIntoTrivia: true))
            {
                switch (trivia.Kind())
                {
                    case SyntaxKind.SingleLineCommentTrivia:
                    case SyntaxKind.MultiLineCommentTrivia:
                    case SyntaxKind.SingleLineDocumentationCommentTrivia:
                    case SyntaxKind.MultiLineDocumentationCommentTrivia:
                        apply(trivia.SpanStart, trivia.Span.Length, ScopeName.Comment);
                        break;
                    case SyntaxKind.DisabledTextTrivia:
                        apply(trivia.SpanStart, trivia.Span.Length, ScopeName.PlainText);
                        break;
                }
            }
        }

        private static bool IsBoldName(SyntaxToken token, SemanticModel model)
        {
            if (!token.IsKind(SyntaxKind.IdentifierToken)) return false;
            var parent = token.Parent;
            if (parent is ClassDeclarationSyntax type && type.Identifier == token ||
                parent is RecordDeclarationSyntax record && record.Identifier == token ||
                parent is MethodDeclarationSyntax method && method.Identifier == token ||
                parent is LocalFunctionStatementSyntax function && function.Identifier == token ||
                parent is ConstructorDeclarationSyntax constructor && constructor.Identifier == token ||
                parent is PropertyDeclarationSyntax property && property.Identifier == token)
                return true;
            if (!(parent is SimpleNameSyntax)) return false;
            var symbol = model.GetSymbolInfo(parent).Symbol;
            return symbol is IMethodSymbol || symbol is IPropertySymbol ||
                symbol is INamedTypeSymbol namedType && namedType.TypeKind == TypeKind.Class;
        }

        private static string GetScope(SyntaxToken token)
        {
            var kind = token.Kind();
            if (SyntaxFacts.IsKeywordKind(kind) || SyntaxFacts.IsContextualKeyword(kind))
                return ScopeName.Keyword;
            if (SyntaxFacts.IsPreprocessorKeyword(kind))
                return ScopeName.PreprocessorKeyword;
            switch (kind)
            {
                case SyntaxKind.StringLiteralToken:
                case SyntaxKind.CharacterLiteralToken:
                case SyntaxKind.Utf8StringLiteralToken:
                case SyntaxKind.SingleLineRawStringLiteralToken:
                case SyntaxKind.MultiLineRawStringLiteralToken:
                case SyntaxKind.Utf8SingleLineRawStringLiteralToken:
                case SyntaxKind.Utf8MultiLineRawStringLiteralToken:
                case SyntaxKind.InterpolatedStringStartToken:
                case SyntaxKind.InterpolatedVerbatimStringStartToken:
                case SyntaxKind.InterpolatedSingleLineRawStringStartToken:
                case SyntaxKind.InterpolatedMultiLineRawStringStartToken:
                case SyntaxKind.InterpolatedStringTextToken:
                case SyntaxKind.InterpolatedStringEndToken:
                case SyntaxKind.InterpolatedRawStringEndToken:
                    return ScopeName.String;
                case SyntaxKind.NumericLiteralToken:
                    return ScopeName.Number;
                case SyntaxKind.IdentifierToken:
                    if (token.Parent is IdentifierNameSyntax name && name.IsVar &&
                        ((name.Parent is VariableDeclarationSyntax variable && variable.Type == name) ||
                         (name.Parent is DeclarationExpressionSyntax expression && expression.Type == name) ||
                         (name.Parent is ForEachStatementSyntax loop && loop.Type == name)))
                        return ScopeName.Keyword;
                    if (token.Parent is BaseTypeDeclarationSyntax declaration && declaration.Identifier == token)
                        return ScopeName.ClassName;
                    return null;
                default:
                    return null;
            }
        }
    }
}
