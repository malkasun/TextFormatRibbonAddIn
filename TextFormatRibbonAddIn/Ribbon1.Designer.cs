namespace TextFormatRibbonAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnSuperscript = this.Factory.CreateRibbonButton();
            this.btnSubscript = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Text Formatting";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSuperscript);
            this.group1.Items.Add(this.btnSubscript);
            this.group1.Label = "Text Format";
            this.group1.Name = "group1";
            // 
            // btnSuperscript
            // 
            this.btnSuperscript.Label = "Superscript";
            this.btnSuperscript.Name = "btnSuperscript";
            this.btnSuperscript.ScreenTip = "Superscript";
            this.btnSuperscript.SuperTip = "Raises the selected text above the normal text line (e.g., for exponents).";
            this.btnSuperscript.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSuperscript_Click);
            // 
            // btnSubscript
            // 
            this.btnSubscript.Label = "Subscript";
            this.btnSubscript.Name = "btnSubscript";
            this.btnSubscript.ScreenTip = "Subscript";
            this.btnSubscript.SuperTip = "Lowers the selected text below the normal text line (e.g., for chemical formulas)" +
    ".";
            this.btnSubscript.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSubscript_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSuperscript;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSubscript;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
