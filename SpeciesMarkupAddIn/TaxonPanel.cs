using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpeciesMarkupAddIn
{
    public partial class TaxonPanel : UserControl
    {
        public TaxonPanel()
        {
            InitializeComponent();
        }

        private void CopySelection(TextBox textboxTarget)
        {
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                textboxTarget.Text = Globals.ThisAddIn.currentText.Trim();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnGenusCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxGenus);
        }

        private void btnSpeciesCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxSpecies);
        }

        private void btnSpeciesAuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxSpeciesAuthor);
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textboxInfraTaxon1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void btnInfraTaxon1Copy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxInfra1Taxon);
        }

        private void btnInfra1AuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxInfra1Author);
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radiobuttonInfra1None.Checked)
            {
                panelInfraspecific1.Visible = false;
            }
            else
            {
                panelInfraspecific1.Visible = true;
            }
        }

        private void btnInfra2TaxonCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxInfra2Taxon);
        }

        private void btnInfra2AuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxInfra2Author);
        }

        private void radiobuttonInfra2None_CheckedChanged(object sender, EventArgs e)
        {
            if (radiobuttonInfra2None.Checked)
            {
                panelInfraspecific2.Visible = false;
            }
            else
            {
                panelInfraspecific2.Visible = true;
            }
        }

        private void TaxonPanel_Load(object sender, EventArgs e)
        {

        }

        private void radiobuttonInfra2None_CheckedChanged_1(object sender, EventArgs e)
        {
            if (radiobuttonInfra2None.Checked)
            {
                panelInfraspecific2.Visible = false;
            }
            else
            {
                panelInfraspecific2.Visible = true;
            }
        }
    }
}
