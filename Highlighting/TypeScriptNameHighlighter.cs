using System;
using System.Collections.Generic;
using Serenity.TypeScript;

namespace VsNetSyntaxWordAddIn.Highlighting
{
    internal static class TypeScriptNameHighlighter
    {
        internal static void Highlight(string source, Action<int, int> applyBold)
        {
            var root = new Parser().ParseSourceFile("selection.ts", source, scriptKind: ScriptKind.TS);
            var classes = new HashSet<string>(StringComparer.Ordinal);
            Visit(root, node =>
            {
                if (node is ClassDeclaration declaration && declaration.Name != null)
                    classes.Add(declaration.Name.Text);
            });
            Visit(root, node =>
            {
                INode name = null;
                if (node is ClassDeclaration type) name = type.Name;
                else if (node is MethodDeclaration method) name = method.Name;
                else if (node is MethodSignature signature) name = signature.Name;
                else if (node is PropertyDeclaration property) name = property.Name;
                else if (node is PropertySignature propertySignature) name = propertySignature.Name;
                else if (node is GetAccessorDeclaration getter) name = getter.Name;
                else if (node is SetAccessorDeclaration setter) name = setter.Name;
                else if (node is FunctionDeclaration function) name = function.Name;
                else if (node is PropertyAccessExpression access) name = access.Name;
                else if (node is CallExpression call && call.Expression is Identifier) name = call.Expression;
                else if (node is Identifier identifier && classes.Contains(identifier.Text)) name = identifier;
                if (!(name is Identifier) && !(name is PrivateIdentifier)) return;
                if (!name.End.HasValue) return;
                // Node positions include leading trivia; GetText removes it without changing token spelling.
                string text = name.GetText(source);
                int start = name.End.Value - text.Length;
                if (text.Length > 0) applyBold(start, text.Length);
            });
        }

        private static void Visit(INode node, Action<INode> action)
        {
            action(node);
            node.ForEachChild(child =>
            {
                Visit(child, action);
                return null;
            });
        }
    }
}
