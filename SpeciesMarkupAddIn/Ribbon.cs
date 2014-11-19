using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

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
            Globals.ThisAddIn.ExportToExcel();
        }
    }
}
