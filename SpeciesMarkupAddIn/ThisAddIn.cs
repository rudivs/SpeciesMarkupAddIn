using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace SpeciesMarkupAddIn
{
    public partial class ThisAddIn
    {
        private bool markupEnabled;
        private Microsoft.Office.Tools.CustomTaskPane taxonMarkupPanel;
        public TaxonPanel myTaxonPanel;
        public string currentText;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            this.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        internal void setDisplayFlag(bool value)
        {
            markupEnabled = value;
            if (markupEnabled)
            {
                taxonMarkupPanel.Visible = true;
            }
            else
            {
                taxonMarkupPanel.Visible = false;
            }
        }

        void Application_DocumentChange()
        {

            LoadTaxonPanel();
        }

        public void LoadTaxonPanel()
        {
            if (taxonMarkupPanel != null)
            {
                this.CustomTaskPanes.Remove(taxonMarkupPanel);
            }

            myTaxonPanel = new TaxonPanel();
            taxonMarkupPanel = this.CustomTaskPanes.Add(myTaxonPanel, "Taxon Markup Panel");
            taxonMarkupPanel.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            taxonMarkupPanel.Width = 370;
        }

        void Application_WindowSelectionChange(Word.Selection Sel)
        {
            if (markupEnabled)
            {
                currentText = Sel.Text;
            }
        }
    }
}
