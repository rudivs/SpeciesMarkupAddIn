using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using log4net;

namespace SpeciesMarkupAddIn
{
    class Export
    {
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static object GetPropValue(object src, string propName)
        {
            return src.GetType().GetProperty(propName).GetValue(src, null);
        }

        private static bool Validated()
        {
            if (Globals.ThisAddIn.currentBatch.Count < 1)
                return false;
            else
            {
                return true;
            }
        }

        internal static void ExportToExcel()
        {
            if (Validated())
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();

                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveFileDialog.FilterIndex = 1;
                saveFileDialog.RestoreDirectory = false;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExportToExcelFile(saveFileDialog.FileName);
                }
            }
        }

        /// <summary>
        /// Exports current batch of marked up taxa to Excel.
        /// </summary>
        private static void ExportToExcelFile(string filename)
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                //set the workbook properties and add a default sheet in it
                SetWorkbookProperties(p);
                //Create a sheet
                ExcelWorksheet ws = CreateSheet(p, "Taxon Export");

                // Add column headers
                foreach (KeyValuePair<int, TaxonInternalReferences> entry in CollectionData.columnIndex)
                {
                    var cell = ws.Cells[1, entry.Key];
                    cell.Value = entry.Value.FieldName;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(entry.Value.FillColor);
                    cell.AddComment(entry.Value.Label, "Field description");
                    ws.Column(entry.Key).Width = entry.Value.ColumnWidth;
                }

                int rowIndex = 2;
                CreateData(ws, ref rowIndex);

                //Generate file
                try
                {
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(filename, bin);
                    log.Info("Excel export successful: " + filename);
                }
                catch (IOException ex)
                {
                    log.Error("IOException writing Excel export file", ex);
                    MessageBox.Show("Error writing Excel export file: " + ex.ToString(), "Export error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    log.Error("Unhandled exception writing Excel export file", ex);
                    throw;
                }
                Globals.ThisAddIn.ResetCurrentBatch(CollectionData.DeleteAfterExport);
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
        private static void SetWorkbookProperties(ExcelPackage p)
        {
            //Here setting some document properties
            p.Workbook.Properties.Title = "Taxon Markup Export";
        }

        private static void CreateData(ExcelWorksheet ws, ref int rowIndex)
        {
            for (int i = 0; i < Globals.ThisAddIn.currentBatch.Count; i++ )
            {
                Taxon t = Globals.ThisAddIn.currentBatch.GetByIndex(i);
                for (int j = 1; j <= CollectionData.columnIndex.Keys.Max(); j++ )
                {
                    string prop = CollectionData.columnIndex[j].PropertyName;
                    if (!string.IsNullOrEmpty(prop))
                    {
                        object propertyVal = GetPropValue(t, CollectionData.columnIndex[j].PropertyName);
                        if (!(CollectionData.columnIndex[j].NullableZero && (Convert.ToInt32(propertyVal) == 0)))
                        {
                            var cell = ws.Cells[rowIndex + i, j];
                            cell.Value = propertyVal;
                            cell.Style.WrapText = true;
                        }
                    }
                }
                ws.Row(rowIndex + i).Height = 50;
            }
        }
    }
}
