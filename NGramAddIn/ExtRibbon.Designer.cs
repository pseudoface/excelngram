using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

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
            if(disposing && (components != null))
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExtRibbon));
            this.tabExtAddIns = this.Factory.CreateRibbonTab();
            this.grpNGram = this.Factory.CreateRibbonGroup();
            this.btnNGram = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.twoWordTgl = this.Factory.CreateRibbonToggleButton();
            this.threeWordTgl = this.Factory.CreateRibbonToggleButton();
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
            this.grpNGram.Items.Add(this.separator1);
            this.grpNGram.Items.Add(this.twoWordTgl);
            this.grpNGram.Items.Add(this.threeWordTgl);
            this.grpNGram.Label = "nGram";
            this.grpNGram.Name = "grpNGram";
            // 
            // btnNGram
            // 
            this.btnNGram.Image = ((System.Drawing.Image)(resources.GetObject("btnNGram.Image")));
            this.btnNGram.Label = "NGram";
            this.btnNGram.Name = "btnNGram";
            this.btnNGram.ScreenTip = "Processes the current sheet\'s values onto a new sheet with the results of the NGr" +
    "am function";
            this.btnNGram.ShowImage = true;
            this.btnNGram.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNGram_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // twoWordTgl
            // 
            this.twoWordTgl.Checked = true;
            this.twoWordTgl.Label = "2-word";
            this.twoWordTgl.Name = "twoWordTgl";
            this.twoWordTgl.ScreenTip = "Click to turn on 2-word nGram";
            this.twoWordTgl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.twoWordTgl_Click);
            // 
            // threeWordTgl
            // 
            this.threeWordTgl.Label = "3-word";
            this.threeWordTgl.Name = "threeWordTgl";
            this.threeWordTgl.ScreenTip = "Click to turn on 3-word nGram";
            this.threeWordTgl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.threeWordTgl_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton threeWordTgl;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton twoWordTgl;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal ExtRibbon Ribbon1
        {
            get
            {
                return this.GetRibbon<ExtRibbon>();
            }
        }
    }
}
