using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace VsNetSyntaxWordAddIn
{
    [ComVisible(true)]
    public sealed class SyntaxRibbon : Office.IRibbonExtensibility
    {
        private int selectedLanguageIndex;
        private static readonly string[] LanguageIds = { "CSharp", "Html", "TypeScript", "JavaScript" };

        public string GetCustomUI(string ribbonId)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("VsNetSyntaxWordAddIn.SyntaxRibbon.xml"))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }

        public int GetSelectedLanguageIndex(Office.IRibbonControl control) => selectedLanguageIndex;

        public void OnLanguageSelected(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex < 0 || selectedIndex >= LanguageIds.Length) return;
            selectedLanguageIndex = selectedIndex;
            Apply();
        }

        public void OnApply(Office.IRibbonControl control) => Apply();

        private void Apply()
        {
            try
            {
                WordSelectionFormatter.Format(Globals.ThisAddIn.Application, LanguageIds[selectedLanguageIndex]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "VS.NET Syntax", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
