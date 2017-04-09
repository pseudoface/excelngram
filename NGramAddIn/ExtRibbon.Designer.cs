namespace NGramAddIn
{
    partial class ExtRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExtRibbon() : base(Globals.Factory.GetRibbonFactory())
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExtRibbon));
            this.tabExtAddIns = this.Factory.CreateRibbonTab();
            this.grpNGram = this.Factory.CreateRibbonGroup();
            this.cmbNumOfWords = this.Factory.CreateRibbonComboBox();
            this.btnNGram = this.Factory.CreateRibbonButton();
            this.tabExtAddIns.SuspendLayout();
            this.grpNGram.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabExtAddIns
            // 
            this.tabExtAddIns.Groups.Add(this.grpNGram);
            this.tabExtAddIns.Label = "Custom Add Ins";
            this.tabExtAddIns.Name = "tabExtAddIns";
            // 
            // grpNGram
            // 
            this.grpNGram.Items.Add(this.btnNGram);
            this.grpNGram.Items.Add(this.cmbNumOfWords);
            this.grpNGram.Label = "NGram";
            this.grpNGram.Name = "grpNGram";
            // 
            // cmbNumOfWords
            // 
            ribbonDropDownItemImpl1.Label = "1";
            ribbonDropDownItemImpl1.Tag = 1;
            ribbonDropDownItemImpl2.Label = "2";
            ribbonDropDownItemImpl2.Tag = 2;
            ribbonDropDownItemImpl3.Label = "3";
            ribbonDropDownItemImpl3.Tag = 3;
            this.cmbNumOfWords.Items.Add(ribbonDropDownItemImpl1);
            this.cmbNumOfWords.Items.Add(ribbonDropDownItemImpl2);
            this.cmbNumOfWords.Items.Add(ribbonDropDownItemImpl3);
            this.cmbNumOfWords.Label = "# of Words";
            this.cmbNumOfWords.Name = "cmbNumOfWords";
            this.cmbNumOfWords.ScreenTip = "The number of words that each result should contain";
            this.cmbNumOfWords.Text = null;
            this.cmbNumOfWords.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmbNumOfWords_ItemsLoading);
            this.cmbNumOfWords.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cmbNumOfWords_TextChanged);
            // 
            // btnNGram
            // 
            this.btnNGram.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNGram.Image = ((System.Drawing.Image)(resources.GetObject("btnNGram.Image")));
            this.btnNGram.Label = "NGram";
            this.btnNGram.Name = "btnNGram";
            this.btnNGram.ShowImage = true;
            this.btnNGram.SuperTip = "Processes the current sheet\'s values onto a new sheet with the results of the NGr" +
    "am function";
            this.btnNGram.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNGram_Click);
            // 
            // ExtRibbon
            // 
            this.Name = "ExtRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabExtAddIns);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExtRibbon_Load);
            this.tabExtAddIns.ResumeLayout(false);
            this.tabExtAddIns.PerformLayout();
            this.grpNGram.ResumeLayout(false);
            this.grpNGram.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        public Microsoft.Office.Tools.Ribbon.RibbonGroup grpNGram;
        public Microsoft.Office.Tools.Ribbon.RibbonTab tabExtAddIns;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnNGram;
        public Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbNumOfWords;
    }

    partial class ThisRibbonCollection
    {
        internal ExtRibbon Ribbon1
        {
            get { return this.GetRibbon<ExtRibbon>(); }
        }
    }
}
