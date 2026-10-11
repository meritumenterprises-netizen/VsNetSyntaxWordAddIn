namespace VsNetSyntaxWordAddIn
{
    partial class SyntaxRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public SyntaxRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem angularLanguage = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem jsonLanguage = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem xmlLanguage = this.Factory.CreateRibbonDropDownItem();
            this.VsNetSyntaxTab = this.Factory.CreateRibbonTab();
            this.SyntaxGroup = this.Factory.CreateRibbonGroup();
            this.ApplySyntax = this.Factory.CreateRibbonButton();
            this.SyntaxLanguage = this.Factory.CreateRibbonDropDown();
            this.SyntaxThemeDropDown = this.Factory.CreateRibbonDropDown();
            this.SyntaxBackgroundDropDown = this.Factory.CreateRibbonDropDown();
            this.VsNetSyntaxTab.SuspendLayout();
            this.SyntaxGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // VsNetSyntaxTab
            // 
            this.VsNetSyntaxTab.Groups.Add(this.SyntaxGroup);
            this.VsNetSyntaxTab.Label = "VS.NET Syntax";
            this.VsNetSyntaxTab.Name = "VsNetSyntaxTab";
            // 
            // SyntaxGroup
            // 
            this.SyntaxGroup.Items.Add(this.ApplySyntax);
            this.SyntaxGroup.Items.Add(this.SyntaxLanguage);
            this.SyntaxGroup.Items.Add(this.SyntaxThemeDropDown);
            this.SyntaxGroup.Items.Add(this.SyntaxBackgroundDropDown);
            this.SyntaxGroup.Label = "Code";
            this.SyntaxGroup.Name = "SyntaxGroup";
            // 
            // ApplySyntax
            // 
            this.ApplySyntax.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ApplySyntax.Label = "Apply";
            this.ApplySyntax.Name = "ApplySyntax";
            this.ApplySyntax.OfficeImageId = "PasteApplyStyle";
            this.ApplySyntax.ScreenTip = "Apply syntax highlighting";
            this.ApplySyntax.ShowImage = true;
            this.ApplySyntax.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnApply);
            // 
            // SyntaxLanguage
            // 
            ribbonDropDownItemImpl1.Label = "C#";
            ribbonDropDownItemImpl2.Label = "HTML";
            ribbonDropDownItemImpl3.Label = "TypeScript";
            ribbonDropDownItemImpl4.Label = "JavaScript";
            angularLanguage.Label = "Angular";
            this.SyntaxLanguage.Items.Add(angularLanguage);
            this.SyntaxLanguage.Items.Add(ribbonDropDownItemImpl1);
            this.SyntaxLanguage.Items.Add(ribbonDropDownItemImpl2);
            this.SyntaxLanguage.Items.Add(ribbonDropDownItemImpl4);
            jsonLanguage.Label = "JSON";
            this.SyntaxLanguage.Items.Add(jsonLanguage);
            this.SyntaxLanguage.Items.Add(ribbonDropDownItemImpl3);
            xmlLanguage.Label = "XML";
            this.SyntaxLanguage.Items.Add(xmlLanguage);
            this.SyntaxLanguage.Label = "Language";
            this.SyntaxLanguage.Name = "SyntaxLanguage";
            this.SyntaxLanguage.SizeString = "TypeScript  ";
            this.SyntaxLanguage.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnLanguageSelected);
            // 
            // SyntaxThemeDropDown
            // 
            ribbonDropDownItemImpl5.Label = "Light";
            ribbonDropDownItemImpl6.Label = "Dark";
            this.SyntaxThemeDropDown.Items.Add(ribbonDropDownItemImpl5);
            this.SyntaxThemeDropDown.Items.Add(ribbonDropDownItemImpl6);
            this.SyntaxThemeDropDown.Label = "Theme";
            this.SyntaxThemeDropDown.Name = "SyntaxThemeDropDown";
            this.SyntaxThemeDropDown.SizeString = "TypeScript Highlight  ";
            this.SyntaxThemeDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnThemeSelected);
            // 
            // SyntaxBackgroundDropDown
            // 
            ribbonDropDownItemImpl7.Label = "Normal";
            ribbonDropDownItemImpl8.Label = "Red";
            ribbonDropDownItemImpl9.Label = "Green";
            this.SyntaxBackgroundDropDown.Items.Add(ribbonDropDownItemImpl7);
            this.SyntaxBackgroundDropDown.Items.Add(ribbonDropDownItemImpl8);
            this.SyntaxBackgroundDropDown.Items.Add(ribbonDropDownItemImpl9);
            this.SyntaxBackgroundDropDown.Label = "Background";
            this.SyntaxBackgroundDropDown.Name = "SyntaxBackgroundDropDown";
            this.SyntaxBackgroundDropDown.ShowImage = true;
            this.SyntaxBackgroundDropDown.SizeString = "TypeScript  ";
            this.SyntaxBackgroundDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnBackgroundSelected);
            // 
            // SyntaxRibbon
            // 
            this.Name = "SyntaxRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.VsNetSyntaxTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SyntaxRibbon_Load);
            this.VsNetSyntaxTab.ResumeLayout(false);
            this.VsNetSyntaxTab.PerformLayout();
            this.SyntaxGroup.ResumeLayout(false);
            this.SyntaxGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab VsNetSyntaxTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SyntaxGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ApplySyntax;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown SyntaxLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown SyntaxThemeDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown SyntaxBackgroundDropDown;
    }

    partial class ThisRibbonCollection
    {
        internal SyntaxRibbon SyntaxRibbon => this.GetRibbon<SyntaxRibbon>();
    }
}
