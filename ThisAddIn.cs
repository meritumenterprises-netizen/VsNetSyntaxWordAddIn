using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace VsNetSyntaxWordAddIn
{
    public partial class ThisAddIn
    {
        private SyntaxRibbon ribbon;
        private GlobalSyntaxShortcuts shortcuts;

        private SyntaxRibbon Ribbon => ribbon ?? (ribbon = new SyntaxRibbon());

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { Ribbon });
        }

        protected override object RequestComAddInAutomationService()
        {
            return new ShortcutDiagnostics(() => shortcuts?.Status ?? "Shortcuts not initialized");
        }

        private void InternalStartup()
        {
            Startup += OnStartup;
            Shutdown += OnShutdown;
        }

        private void OnStartup(object sender, EventArgs e)
        {
            try { shortcuts = new GlobalSyntaxShortcuts(Ribbon.ApplyLanguage); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "VS.NET Syntax", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OnShutdown(object sender, EventArgs e)
        {
            shortcuts?.Dispose();
            shortcuts = null;
        }
    }
}
