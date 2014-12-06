using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace SpeciesMarkupAddIn
{
    public partial class ThisAddIn
    {
        private bool markupEnabled;
        private Microsoft.Office.Tools.CustomTaskPane taxonMarkupPanel;
        public TaxonPanel myTaxonPanel;
        public string currentText;
        List<Taxon> currentTaxa = new List<Taxon>();

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
            Taxon currentTaxon = new Taxon();
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

        internal void ExportToExcel()
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                //set the workbook properties and add a default sheet in it
                SetWorkbookProperties(p);
                //Create a sheet
                ExcelWorksheet ws = CreateSheet(p, "Sample Sheet");

                //Merging cells and create a center heading for out table
                ws.Cells[1, 1].Value = "Taxon Markup Export";

                int rowIndex = 2;
                CreateData(ws, ref rowIndex);

                //Generate A File with Random name
                Byte[] bin = p.GetAsByteArray();
                string file = Guid.NewGuid().ToString() + ".xlsx";
                File.WriteAllBytes(file, bin);

                //These lines will open it in Excel
                ProcessStartInfo pi = new ProcessStartInfo(file);
                Process.Start(pi);
            }
        }

        private static ExcelWorksheet CreateSheet(ExcelPackage p, string sheetName)
        {
            p.Workbook.Worksheets.Add(sheetName);
            ExcelWorksheet ws = p.Workbook.Worksheets[1];
            ws.Name = sheetName; //Setting Sheet's name
            ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

            return ws;
        }

        /// <summary>
        /// Sets the workbook properties and adds a default sheet.
        /// </summary>
        /// <param name="p">The p.</param>
        /// <returns></returns>
        private static void SetWorkbookProperties(ExcelPackage p)
        {
            //Here setting some document properties
            p.Workbook.Properties.Title = "Taxon Markup Export";
        }

        void CreateData(ExcelWorksheet ws, ref int rowIndex)
        {
            int colIndex = 1;
            List<string> tbList = new List<string>() { "textboxGenus", "textboxSpecies", "textboxSpeciesAuthor" };

            foreach (string tbCurrent in tbList)
            {
                var cell = ws.Cells[rowIndex, colIndex];
                cell.Value = ((TextBox)myTaxonPanel.Controls[0].Controls[0].Controls[tbCurrent]).Text;
                colIndex++;
            }
        }
    }
}
