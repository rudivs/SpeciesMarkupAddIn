namespace SpeciesMarkupAddIn
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl15 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl16 = this.Factory.CreateRibbonDropDownItem();
            this.tabTaxonMarkup = this.Factory.CreateRibbonTab();
            this.grpDisplay = this.Factory.CreateRibbonGroup();
            this.checkboxDisplayTaxonPanel = this.Factory.CreateRibbonCheckBox();
            this.cbEditorFontSize = this.Factory.CreateRibbonComboBox();
            this.grpData = this.Factory.CreateRibbonGroup();
            this.btnDeleteCurrent = this.Factory.CreateRibbonButton();
            this.btnClearBatch = this.Factory.CreateRibbonButton();
            this.grpProcess = this.Factory.CreateRibbonGroup();
            this.btnCopyPrevious = this.Factory.CreateRibbonButton();
            this.btnCopyPreviousGenus = this.Factory.CreateRibbonButton();
            this.grpExport = this.Factory.CreateRibbonGroup();
            this.grpBatch = this.Factory.CreateRibbonGroup();
            this.btnBatchCount = this.Factory.CreateRibbonButton();
            this.btnExportExcel = this.Factory.CreateRibbonButton();
            this.tabTaxonMarkup.SuspendLayout();
            this.grpDisplay.SuspendLayout();
            this.grpData.SuspendLayout();
            this.grpProcess.SuspendLayout();
            this.grpExport.SuspendLayout();
            this.grpBatch.SuspendLayout();
            // 
            // tabTaxonMarkup
            // 
            this.tabTaxonMarkup.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabTaxonMarkup.Groups.Add(this.grpDisplay);
            this.tabTaxonMarkup.Groups.Add(this.grpData);
            this.tabTaxonMarkup.Groups.Add(this.grpProcess);
            this.tabTaxonMarkup.Groups.Add(this.grpExport);
            this.tabTaxonMarkup.Groups.Add(this.grpBatch);
            this.tabTaxonMarkup.Label = "Taxon Markup";
            this.tabTaxonMarkup.Name = "tabTaxonMarkup";
            // 
            // grpDisplay
            // 
            this.grpDisplay.Items.Add(this.checkboxDisplayTaxonPanel);
            this.grpDisplay.Items.Add(this.cbEditorFontSize);
            this.grpDisplay.Label = "Display";
            this.grpDisplay.Name = "grpDisplay";
            // 
            // checkboxDisplayTaxonPanel
            // 
            this.checkboxDisplayTaxonPanel.Label = "Show Taxon Panel";
            this.checkboxDisplayTaxonPanel.Name = "checkboxDisplayTaxonPanel";
            this.checkboxDisplayTaxonPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkboxDisplayTaxonPanel_Click);
            // 
            // cbEditorFontSize
            // 
            ribbonDropDownItemImpl1.Label = "8";
            ribbonDropDownItemImpl2.Label = "9";
            ribbonDropDownItemImpl3.Label = "10";
            ribbonDropDownItemImpl4.Label = "11";
            ribbonDropDownItemImpl5.Label = "12";
            ribbonDropDownItemImpl6.Label = "14";
            ribbonDropDownItemImpl7.Label = "16";
            ribbonDropDownItemImpl8.Label = "18";
            ribbonDropDownItemImpl9.Label = "20";
            ribbonDropDownItemImpl10.Label = "22";
            ribbonDropDownItemImpl11.Label = "24";
            ribbonDropDownItemImpl12.Label = "26";
            ribbonDropDownItemImpl13.Label = "30";
            ribbonDropDownItemImpl14.Label = "34";
            ribbonDropDownItemImpl15.Label = "38";
            ribbonDropDownItemImpl16.Label = "42";
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl1);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl2);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl3);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl4);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl5);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl6);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl7);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl8);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl9);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl10);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl11);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl12);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl13);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl14);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl15);
            this.cbEditorFontSize.Items.Add(ribbonDropDownItemImpl16);
            this.cbEditorFontSize.Label = "Editor Font Size";
            this.cbEditorFontSize.Name = "cbEditorFontSize";
            this.cbEditorFontSize.SizeString = "xx";
            this.cbEditorFontSize.Text = null;
            this.cbEditorFontSize.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.cbEditorFontSize_TextChanged);
            // 
            // grpData
            // 
            this.grpData.Items.Add(this.btnDeleteCurrent);
            this.grpData.Items.Add(this.btnClearBatch);
            this.grpData.Label = "Data Management";
            this.grpData.Name = "grpData";
            // 
            // btnDeleteCurrent
            // 
            this.btnDeleteCurrent.Label = "Delete Current Record";
            this.btnDeleteCurrent.Name = "btnDeleteCurrent";
            this.btnDeleteCurrent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteCurrent_Click);
            // 
            // btnClearBatch
            // 
            this.btnClearBatch.Label = "Clear Current Batch";
            this.btnClearBatch.Name = "btnClearBatch";
            this.btnClearBatch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearBatch_Click);
            // 
            // grpProcess
            // 
            this.grpProcess.Items.Add(this.btnCopyPrevious);
            this.grpProcess.Items.Add(this.btnCopyPreviousGenus);
            this.grpProcess.Label = "Process";
            this.grpProcess.Name = "grpProcess";
            // 
            // btnCopyPrevious
            // 
            this.btnCopyPrevious.Label = "Copy Previous Species Name";
            this.btnCopyPrevious.Name = "btnCopyPrevious";
            this.btnCopyPrevious.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyPrevious_Click);
            // 
            // btnCopyPreviousGenus
            // 
            this.btnCopyPreviousGenus.Label = "Copy Previous Genus Name";
            this.btnCopyPreviousGenus.Name = "btnCopyPreviousGenus";
            this.btnCopyPreviousGenus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyPreviousGenus_Click);
            // 
            // grpExport
            // 
            this.grpExport.Items.Add(this.btnExportExcel);
            this.grpExport.Label = "Export";
            this.grpExport.Name = "grpExport";
            // 
            // grpBatch
            // 
            this.grpBatch.Items.Add(this.btnBatchCount);
            this.grpBatch.Label = "Batch";
            this.grpBatch.Name = "grpBatch";
            // 
            // btnBatchCount
            // 
            this.btnBatchCount.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBatchCount.Image = global::SpeciesMarkupAddIn.Properties.Resources.tc;
            this.btnBatchCount.Label = "Taxon List";
            this.btnBatchCount.Name = "btnBatchCount";
            this.btnBatchCount.ShowImage = true;
            this.btnBatchCount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBatchCount_Click);
            // 
            // btnExportExcel
            // 
            this.btnExportExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportExcel.Image = global::SpeciesMarkupAddIn.Properties.Resources.Excel;
            this.btnExportExcel.Label = "Excel";
            this.btnExportExcel.Name = "btnExportExcel";
            this.btnExportExcel.ShowImage = true;
            this.btnExportExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportExcel_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabTaxonMarkup);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabTaxonMarkup.ResumeLayout(false);
            this.tabTaxonMarkup.PerformLayout();
            this.grpDisplay.ResumeLayout(false);
            this.grpDisplay.PerformLayout();
            this.grpData.ResumeLayout(false);
            this.grpData.PerformLayout();
            this.grpProcess.ResumeLayout(false);
            this.grpProcess.PerformLayout();
            this.grpExport.ResumeLayout(false);
            this.grpExport.PerformLayout();
            this.grpBatch.ResumeLayout(false);
            this.grpBatch.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabTaxonMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDisplay;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkboxDisplayTaxonPanel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpExport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteCurrent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearBatch;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyPrevious;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProcess;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyPreviousGenus;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cbEditorFontSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBatchCount;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpBatch;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
