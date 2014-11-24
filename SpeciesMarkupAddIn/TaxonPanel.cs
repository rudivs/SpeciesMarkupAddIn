using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace SpeciesMarkupAddIn
{
    public partial class TaxonPanel : UserControl
    {
        public TaxonPanel()
        {
            InitializeComponent();
            this.comboboxFloweringStart.DataSource = new BindingSource(CollectionData.Months,null);
            this.comboboxFloweringStart.ValueMember = "Key";
            this.comboboxFloweringStart.DisplayMember = "Value";
            this.comboboxFloweringEnd.DataSource = new BindingSource(CollectionData.Months, null);
            this.comboboxFloweringEnd.ValueMember = "Key";
            this.comboboxFloweringEnd.DisplayMember = "Value";
        }

        private string FilterText(string inputText)
        {
            string filteredText = inputText;

            // remove unnecessary lines
            Regex regex = new Regex(@"[\r]{3,}");
            filteredText = regex.Replace(filteredText, "\r\r");

            // make newline characters work properly
            filteredText = Regex.Replace(filteredText, @"\r", "\r\n");

            // replace tab with space
            filteredText = Regex.Replace(filteredText, @"\t", " ");

            // replace runs of spaces with a single space
            regex = new Regex(@"[ ]{2,}", RegexOptions.None);
            filteredText = regex.Replace(filteredText, @" ");

            // fix comma space
            filteredText = filteredText.Replace(" , ",", ");
            
            // fix dashes
            filteredText = filteredText.Replace("–", "-");
            filteredText = filteredText.Replace("—", "-");

            // add spaces between digits and metre-based measurements
            filteredText = Regex.Replace(filteredText, @"(\d)m", "$1 m");

            // add spaces between digits and cm measurements
            filteredText = Regex.Replace(filteredText, @"(\d)cm", "$1 cm");

            // replace comma with period as decimal separator (but only for 1 or 2 digits after comma, and no spaces)
            filteredText = Regex.Replace(filteredText, @"([0-9])+\,([0-9]{1,2})(?![0-9])", "$1.$2"); 

            // add spaces between digits and multiplication signs (e.g. 3x5 -> 3 x 5)
            filteredText = Regex.Replace(filteredText, @"(\d)[xX\*](\d)", "$1 x $2");

            // remove invisible characters (optional hyphens, etc.)
            filteredText = Regex.Replace(filteredText, @"[^\w\s.,!@#$%^&*()-=+~`]", "");

            // remove spaces between numbers and hyphens (12 - 15 -> 12-15)
            filteredText = Regex.Replace(filteredText, @"(\d) - (\d)", "$1-$2");
            
            return filteredText;
        }

        private int GetMonthNumber(string lookup)
        {
            int outMonth = 0;
            lookup = lookup.Trim().TrimEnd('?', '!', '.', ',');
            CollectionData.MonthLookup.TryGetValue(lookup, out outMonth);
            return outMonth;
        }

        private void CopySelection(TextBox textboxTarget)
        {
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                textboxTarget.Text = FilterText(Globals.ThisAddIn.currentText).Trim();
            }
        }

        private void AddSelection(TextBox textboxTarget)
        {
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                textboxTarget.Text += FilterText(Globals.ThisAddIn.currentText).Trim();
            }
        }

        private void CopyMonth(ComboBox comboboxTarget)
        {
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                int returnMonth = GetMonthNumber(Globals.ThisAddIn.currentText);
                if (returnMonth > 0)
                {
                    comboboxTarget.SelectedValue = GetMonthNumber(Globals.ThisAddIn.currentText);
                }
            }
        }

        private void EditText(TextBox textboxTarget)
        {
            using (var form = new TextEditForm())
            {
                form.Controls[0].Controls["textboxEditText"].Text = textboxTarget.Text;
                var result = form.ShowDialog();
                if (result == DialogResult.OK)
                {
                    textboxTarget.Text = form.ReturnText;
                }
            }

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
                radiobuttonInfra2None.Checked = true;
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

        private void btnMorphDescriptionCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxMorphDescription);
        }

        private void btnDistributionCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxDistribution);
        }

        private void btnHabitatCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxHabitat);
        }

        private void btnFloweringStartCopy_Click(object sender, EventArgs e)
        {
            CopyMonth(comboboxFloweringStart);
        }

        private void btnFloweringEndCopy_Click(object sender, EventArgs e)
        {
            CopyMonth(comboboxFloweringEnd);
        }

        private void btnNotesCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxNotes);
        }

        private void btnVouchersCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxVouchers);
        }

        private void btnTreatmentAuthorsCopy_Click(object sender, EventArgs e)
        {
            CopySelection(textboxTreatmentAuthors);
        }

        private void btnNotesAdd_Click(object sender, EventArgs e)
        {
            AddSelection(textboxNotes);
        }

        private void btnVouchersAdd_Click(object sender, EventArgs e)
        {
            AddSelection(textboxVouchers);
        }

        private void textboxMorphDescription_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(textboxMorphDescription);
        }

        private void textboxDistribution_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(textboxDistribution);
        }

        private void textboxHabitat_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(textboxHabitat);
        }

        private void textboxNotes_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(textboxNotes);
        }

        private void textboxVouchers_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(textboxVouchers);
        }

    }
}
