using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace SpeciesMarkupAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            cbEditorFontSize.Text = Properties.Settings.Default.EditFontSize.ToString();
        }

        private void checkboxDisplayTaxonPanel_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.setDisplayFlag(((RibbonCheckBox)sender).Checked);
        }

        private void btnExportExcel_Click(object sender, RibbonControlEventArgs e)
        {
            Export.ExportToExcel();
        }

        private void btnDeleteCurrent_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult result = MessageBox.Show("Delete current record?", "Confirm delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (result == DialogResult.Yes)
            {
                Globals.ThisAddIn.currentBatch.DeleteCurrent();
                Globals.ThisAddIn.myTaxonPanel.LoadCurrentTaxon();
            }
        }

        private void btnClearBatch_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ResetCurrentBatch(CollectionData.DeleteAdHoc);
        }

        private void btnCopyPrevious_Click(object sender, RibbonControlEventArgs e)
        {
            int index = Globals.ThisAddIn.currentBatch.Index;
            if (index > 0)
            {
                Taxon previousTaxon = Globals.ThisAddIn.currentBatch.GetByIndex(index-1);
                Control taxonPanel = Globals.ThisAddIn.myTaxonPanel.Controls[1].Controls[0];
                taxonPanel.Controls["textboxGenus"].Text = previousTaxon.Genus;
                taxonPanel.Controls["textboxSpecies"].Text = previousTaxon.Species;
                taxonPanel.Controls["textboxSpeciesAuthor"].Text = previousTaxon.SpeciesAuthor;
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void btnCopyPreviousGenus_Click(object sender, RibbonControlEventArgs e)
        {
            int index = Globals.ThisAddIn.currentBatch.Index;
            if (index > 0)
            {
                Taxon previousTaxon = Globals.ThisAddIn.currentBatch.GetByIndex(index - 1);
                Control taxonPanel = Globals.ThisAddIn.myTaxonPanel.Controls[1].Controls[0];
                taxonPanel.Controls["textboxGenus"].Text = previousTaxon.Genus;
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void cbEditorFontSize_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Properties.Settings.Default.EditFontSize = ushort.Parse(cbEditorFontSize.Text);
            Properties.Settings.Default.Save();
        }

        private void btnBatchCount_Click(object sender, RibbonControlEventArgs e)
        {
            using (var form = new TaxonListForm())
            {
                var listbox = form.Controls["tlpBasePanel"].Controls["lbTaxonList"];
                listbox.Font = new Font(listbox.Font.FontFamily, Properties.Settings.Default.EditFontSize);
                form.ShowDialog();
            }
        }

        public void UpdateCount()
        {
            int iCount = Globals.ThisAddIn.currentBatch.Count;
            btnBatchCount.Image = CreateTaxonCountButton(iCount);
        }

        private Image CreateTaxonCountButton(int count)
        {
            Bitmap button = (Bitmap)SpeciesMarkupAddIn.Properties.Resources.ResourceManager.GetObject("tc");
            RectangleF rectf = new RectangleF(0, 1, button.Width, button.Height);
            Graphics g = Graphics.FromImage(button);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            StringFormat format = new StringFormat()
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            count = count < 0 ? 0 : count;
            string sCount = count > 999 ? "999+" : count.ToString();
            Font testFont = new Font("Sans", 25); 
            Font countFont = GetAdjustedFont(g, sCount, testFont, 31, 25, 10, true);

            g.DrawString(sCount, countFont, Brushes.White, rectf, format);

            g.Flush();
            return button;

        }

        public Font GetAdjustedFont(Graphics GraphicRef, string GraphicString, Font OriginalFont, int ContainerWidth, int MaxFontSize, int MinFontSize, bool SmallestOnFail)
        {
            Font testFont = null;
            // We utilize MeasureString which we get via a control instance           
            for (int AdjustedSize = MaxFontSize; AdjustedSize >= MinFontSize; AdjustedSize--)
            {
                testFont = new Font(OriginalFont.Name, AdjustedSize, OriginalFont.Style);

                // Test the string with the new size
                SizeF AdjustedSizeNew = GraphicRef.MeasureString(GraphicString, testFont);

                if (ContainerWidth > Convert.ToInt32(AdjustedSizeNew.Width))
                {
                    // Good font, return it
                    return testFont;
                }
            }

            // If you get here there was no fontsize that worked
            // return MinimumSize or Original?
            if (SmallestOnFail)
            {
                return testFont;
            }
            else
            {
                return OriginalFont;
            }
        }
    }
}
