using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpeciesMarkupAddIn
{
    public partial class TaxonListForm : Form
    {
        public TaxonListForm()
        {
            InitializeComponent();
            lbTaxonList.Items.Clear();
            foreach (Taxon taxon in Globals.ThisAddIn.currentBatch)
            {
                lbTaxonList.Items.Add(taxon.FullName);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.currentBatch.MoveIndex(lbTaxonList.SelectedIndex))
            {
                Globals.ThisAddIn.myTaxonPanel.LoadCurrentTaxon();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Error loading selected taxon.", "Load error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
