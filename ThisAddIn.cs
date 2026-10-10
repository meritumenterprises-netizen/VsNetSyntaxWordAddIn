using Office = Microsoft.Office.Core;

namespace VsNetSyntaxWordAddIn
{
    public partial class ThisAddIn
    {
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SyntaxRibbon();
        }

        private void InternalStartup() { }
    }
}
