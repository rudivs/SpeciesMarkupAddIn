using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace SpeciesMarkupAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

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
            }
        }
    }
}
