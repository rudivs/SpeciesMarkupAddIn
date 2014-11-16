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
            this.tabTaxonMarkup.SuspendLayout();
            this.grpDisplay.SuspendLayout();
            // 
            // tabTaxonMarkup
            // 
            this.tabTaxonMarkup.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabTaxonMarkup.Groups.Add(this.grpDisplay);
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

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabTaxonMarkup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDisplay;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkboxDisplayTaxonPanel;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
