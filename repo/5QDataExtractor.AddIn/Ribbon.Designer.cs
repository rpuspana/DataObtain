namespace _5QDataExtractor.AddIn
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.DataExtractor = this.Factory.CreateRibbonTab();
            this.grpDataExtract = this.Factory.CreateRibbonGroup();
            this.btnDynamicExtract = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.DataExtractor.SuspendLayout();
            this.grpDataExtract.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // DataExtractor
            // 
            this.DataExtractor.Groups.Add(this.grpDataExtract);
            this.DataExtractor.Label = "DataExtractor";
            this.DataExtractor.Name = "DataExtractor";
            // 
            // grpDataExtract
            // 
            this.grpDataExtract.Items.Add(this.btnDynamicExtract);
            this.grpDataExtract.Label = "DataExtract";
            this.grpDataExtract.Name = "grpDataExtract";
            // 
            // btnDynamicExtract
            // 
            this.btnDynamicExtract.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDynamicExtract.Label = "DynamicExtract";
            this.btnDynamicExtract.Name = "btnDynamicExtract";
            this.btnDynamicExtract.ShowImage = true;
            this.btnDynamicExtract.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDynamicExtract_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.DataExtractor);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.DataExtractor.ResumeLayout(false);
            this.DataExtractor.PerformLayout();
            this.grpDataExtract.ResumeLayout(false);
            this.grpDataExtract.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab DataExtractor;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDataExtract;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDynamicExtract;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
