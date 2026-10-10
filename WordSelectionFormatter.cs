using System;
using System.Runtime.InteropServices;
using VsNetSyntaxWordAddIn.Highlighting;
using Word = Microsoft.Office.Interop.Word;

namespace VsNetSyntaxWordAddIn
{
    internal static class WordSelectionFormatter
    {
        internal static void Format(Word.Application application, string languageId)
        {
            if (application.Documents.Count == 0) throw new InvalidOperationException("Open a document and select a code block.");
            var selection = application.Selection;
            Word.Range range = null;
            Word.Range tokenRange = null;
            bool undoStarted = false;
            bool screenUpdating = application.ScreenUpdating;
            try
            {
                if (selection.Type != Word.WdSelectionType.wdSelectionNormal)
                    throw new InvalidOperationException("Select a continuous block of plain code text.");
                range = selection.Range.Duplicate;
                string text = range.Text ?? "";
                if (text.Length == 0 || string.IsNullOrWhiteSpace(text))
                    throw new InvalidOperationException("Select a code block before applying highlighting.");
                if (text.Length > 100000)
                    throw new InvalidOperationException("Select no more than 100,000 characters at a time.");
                if (text.IndexOf('\a') >= 0 || text.IndexOf('\u0001') >= 0 ||
                    range.Fields.Count > 0 || range.InlineShapes.Count > 0 || range.ContentControls.Count > 0)
                    throw new InvalidOperationException("Select plain code text without table cell markers, fields, objects, or content controls.");
                if (range.Document.ProtectionType != Word.WdProtectionType.wdNoProtection)
                    throw new InvalidOperationException("The document is protected. Remove protection before formatting code.");

                var spans = new SyntaxHighlighter().Highlight(text, languageId);
                int start = range.Start;
                application.UndoRecord.StartCustomRecord("Syntax highlighting");
                undoStarted = true;
                application.ScreenUpdating = false;
                var font = range.Font;
                try
                {
                    font.Name = "Consolas";
                    font.Size = 10;
                    font.Bold = 0;
                    font.Italic = 0;
                    font.Underline = Word.WdUnderline.wdUnderlineNone;
                    font.Color = Word.WdColor.wdColorBlack;
                }
                finally { Marshal.ReleaseComObject(font); }
                range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
                var shading = range.Shading;
                try
                {
                    shading.Texture = Word.WdTextureIndex.wdTextureNone;
                    shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                }
                finally { Marshal.ReleaseComObject(shading); }

                tokenRange = range.Duplicate;
                foreach (var span in spans)
                {
                    tokenRange.SetRange(start + span.Start, start + span.Start + span.Length);
                    var tokenFont = tokenRange.Font;
                    try { tokenFont.Color = (Word.WdColor)span.Color; }
                    finally { Marshal.ReleaseComObject(tokenFont); }
                }
            }
            finally
            {
                try
                {
                    if (undoStarted) application.UndoRecord.EndCustomRecord();
                }
                finally
                {
                    application.ScreenUpdating = screenUpdating;
                    if (tokenRange != null) Marshal.ReleaseComObject(tokenRange);
                    if (range != null) Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(selection);
                }
            }
        }
    }
}
