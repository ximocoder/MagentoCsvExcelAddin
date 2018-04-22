namespace ExcelAdd4Magento
{
    partial class HusseRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HusseRibbon()
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
            this.groupHusse = this.Factory.CreateRibbonGroup();
            this.btnExportToMagento = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupHusse.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupHusse);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupHusse
            // 
            this.groupHusse.Items.Add(this.btnExportToMagento);
            this.groupHusse.Label = "Husse";
            this.groupHusse.Name = "groupHusse";
            // 
            // btnExportToMagento
            // 
            this.btnExportToMagento.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportToMagento.Image = global::ExcelAdd4Magento.Properties.Resources.magento;
            this.btnExportToMagento.Label = "Export to Magento";
            this.btnExportToMagento.Name = "btnExportToMagento";
            this.btnExportToMagento.ShowImage = true;
            this.btnExportToMagento.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportToMagento_Click);
            // 
            // HusseRibbon
            // 
            this.Name = "HusseRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HusseRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupHusse.ResumeLayout(false);
            this.groupHusse.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupHusse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportToMagento;
    }

    partial class ThisRibbonCollection
    {
        internal HusseRibbon HusseRibbon
        {
            get { return this.GetRibbon<HusseRibbon>(); }
        }
    }
}
