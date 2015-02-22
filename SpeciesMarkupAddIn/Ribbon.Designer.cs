﻿namespace SpeciesMarkupAddIn
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
            this.tabTaxonMarkup = this.Factory.CreateRibbonTab();
            this.grpDisplay = this.Factory.CreateRibbonGroup();
            this.checkboxDisplayTaxonPanel = this.Factory.CreateRibbonCheckBox();
            this.grpData = this.Factory.CreateRibbonGroup();
            this.grpExport = this.Factory.CreateRibbonGroup();
            this.btnDeleteCurrent = this.Factory.CreateRibbonButton();
            this.btnClearBatch = this.Factory.CreateRibbonButton();
            this.btnCopyPrevious = this.Factory.CreateRibbonButton();
            this.btnExportExcel = this.Factory.CreateRibbonButton();
            this.btnCopyPreviousGenus = this.Factory.CreateRibbonButton();
            this.grpProcess = this.Factory.CreateRibbonGroup();
            this.tabTaxonMarkup.SuspendLayout();
            this.grpDisplay.SuspendLayout();
            this.grpData.SuspendLayout();
            this.grpExport.SuspendLayout();
            this.grpProcess.SuspendLayout();
            // 
            // tabTaxonMarkup
            // 
            this.tabTaxonMarkup.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabTaxonMarkup.Groups.Add(this.grpDisplay);
            this.tabTaxonMarkup.Groups.Add(this.grpData);
            this.tabTaxonMarkup.Groups.Add(this.grpProcess);
            this.tabTaxonMarkup.Groups.Add(this.grpExport);
            this.tabTaxonMarkup.Label = "Taxon Markup";
            this.tabTaxonMarkup.Name = "tabTaxonMarkup";
            // 
            // grpDisplay
            // 
            this.grpDisplay.Items.Add(this.checkboxDisplayTaxonPanel);
            this.grpDisplay.Label = "Display";
            this.grpDisplay.Name = "grpDisplay";
            // 
            // checkboxDisplayTaxonPanel
            // 
            this.checkboxDisplayTaxonPanel.Label = "Show Taxon Panel";
            this.checkboxDisplayTaxonPanel.Name = "checkboxDisplayTaxonPanel";
            this.checkboxDisplayTaxonPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkboxDisplayTaxonPanel_Click);
            // 
            // grpData
            // 
            this.grpData.Items.Add(this.btnDeleteCurrent);
            this.grpData.Items.Add(this.btnClearBatch);
            this.grpData.Label = "Data Management";
            this.grpData.Name = "grpData";
            // 
            // grpExport
            // 
            this.grpExport.Items.Add(this.btnExportExcel);
            this.grpExport.Label = "Export";
            this.grpExport.Name = "grpExport";
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
            // btnCopyPrevious
            // 
            this.btnCopyPrevious.Label = "Copy Previous Species Name";
            this.btnCopyPrevious.Name = "btnCopyPrevious";
            this.btnCopyPrevious.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyPrevious_Click);
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
            // btnCopyPreviousGenus
            // 
            this.btnCopyPreviousGenus.Label = "Copy Previous Genus Name";
            this.btnCopyPreviousGenus.Name = "btnCopyPreviousGenus";
            this.btnCopyPreviousGenus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyPreviousGenus_Click);
            // 
            // grpProcess
            // 
            this.grpProcess.Items.Add(this.btnCopyPrevious);
            this.grpProcess.Items.Add(this.btnCopyPreviousGenus);
            this.grpProcess.Label = "Process";
            this.grpProcess.Name = "grpProcess";
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
            this.grpExport.ResumeLayout(false);
            this.grpExport.PerformLayout();
            this.grpProcess.ResumeLayout(false);
            this.grpProcess.PerformLayout();

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
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
