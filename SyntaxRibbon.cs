using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using VsNetSyntaxWordAddIn.Highlighting;

namespace VsNetSyntaxWordAddIn
{
    public partial class SyntaxRibbon
    {
        private int selectedLanguageIndex;
        private SyntaxTheme selectedTheme;
        private int selectedBackgroundIndex;
        private static readonly int?[] BackgroundColors = { null, 0xB4D9FE, 0xE7F5ED };
        private Bitmap[] backgroundSwatches;
        private static readonly string[] LanguageIds = { "CSharp", "Html", "TypeScript", "JavaScript", "Angular" };

        private void SyntaxRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (backgroundSwatches == null)
            {
                backgroundSwatches = new Bitmap[BackgroundColors.Length];
                for (int i = 0; i < BackgroundColors.Length; i++)
                {
                    var swatch = new Bitmap(16, 16);
                    using (var graphics = Graphics.FromImage(swatch))
                    {
                        graphics.Clear(Color.Transparent);
                        if (BackgroundColors[i].HasValue)
                        {
                            using (var brush = new SolidBrush(ColorTranslator.FromOle(BackgroundColors[i].Value)))
                                graphics.FillRectangle(brush, 1, 3, 14, 10);
                        }
                        graphics.DrawRectangle(Pens.DimGray, 1, 3, 13, 9);
                        if (!BackgroundColors[i].HasValue)
                            graphics.DrawLine(Pens.DimGray, 2, 11, 13, 4);
                    }
                    backgroundSwatches[i] = swatch;
                    SyntaxBackgroundDropDown.Items[i].Image = swatch;
                }
                Disposed += (s, args) =>
                {
                    foreach (var swatch in backgroundSwatches) swatch.Dispose();
                };
            }
            SyntaxLanguage.SelectedItemIndex = selectedLanguageIndex;
            SyntaxThemeDropDown.SelectedItemIndex = (int)selectedTheme;
            SyntaxBackgroundDropDown.SelectedItemIndex = selectedBackgroundIndex;
            SyntaxBackgroundDropDown.Image = backgroundSwatches[selectedBackgroundIndex];
        }

        private void OnBackgroundSelected(object sender, RibbonControlEventArgs e)
        {
            int index = SyntaxBackgroundDropDown.SelectedItemIndex;
            if (index < 0 || index >= BackgroundColors.Length) return;
            selectedBackgroundIndex = index;
            if (backgroundSwatches != null)
                SyntaxBackgroundDropDown.Image = backgroundSwatches[index];
            Apply();
        }

        private void OnThemeSelected(object sender, RibbonControlEventArgs e)
        {
            int selectedIndex = SyntaxThemeDropDown.SelectedItemIndex;
            if (selectedIndex < 0 || selectedIndex > 1) return;
            selectedTheme = (SyntaxTheme)selectedIndex;
            Apply();
        }

        internal void ApplyLanguage(string languageId)
        {
            int index = Array.IndexOf(LanguageIds, languageId);
            if (index < 0) throw new ArgumentException("Unsupported syntax language.", nameof(languageId));
            selectedLanguageIndex = index;
            SyntaxLanguage.SelectedItemIndex = index;
            Apply();
        }

        private void OnLanguageSelected(object sender, RibbonControlEventArgs e)
        {
            int selectedIndex = SyntaxLanguage.SelectedItemIndex;
            if (selectedIndex < 0 || selectedIndex >= LanguageIds.Length) return;
            selectedLanguageIndex = selectedIndex;
            Apply();
        }

        private void OnApply(object sender, RibbonControlEventArgs e) => Apply();

        private void Apply()
        {
            try
            {
                WordSelectionFormatter.Format(Globals.ThisAddIn.Application, LanguageIds[selectedLanguageIndex], selectedTheme,
                    BackgroundColors[selectedBackgroundIndex]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "VS.NET Syntax", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
