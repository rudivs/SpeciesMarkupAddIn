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
using System.Globalization;

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

        public void LoadCurrentTaxon()
        {
            Globals.ThisAddIn.currentTaxon = Globals.ThisAddIn.currentBatch.Current;
            this.textboxGenus.Text = Globals.ThisAddIn.currentTaxon.Genus;
            this.textboxSpecies.Text = Globals.ThisAddIn.currentTaxon.Species;
            this.textboxSpeciesAuthor.Text = Globals.ThisAddIn.currentTaxon.SpeciesAuthor;
            switch (Globals.ThisAddIn.currentTaxon.Infra1Rank)
            {
                case "subsp.":
                    this.radiobuttonInfra1Subspecies.Checked = true;
                    break;
                case "var.":
                    this.radiobuttonInfra1Var.Checked = true;
                    break;
                case "f.":
                    this.radiobuttonInfra1Form.Checked = true;
                    break;
                default:
                    this.radiobuttonInfra1None.Checked = true;
                    break;
            }
            this.textboxInfra1Taxon.Text = Globals.ThisAddIn.currentTaxon.Infra1Taxon;
            this.textboxInfra1Author.Text = Globals.ThisAddIn.currentTaxon.Infra1Author;
            switch (Globals.ThisAddIn.currentTaxon.Infra2Rank)
            {
                case "var.":
                    this.radiobuttonInfra2Var.Checked = true;
                    break;
                case "f.":
                    this.radiobuttonInfra2Form.Checked = true;
                    break;
                default:
                    this.radiobuttonInfra2None.Checked = true;
                    break;
            }
            this.textboxInfra2Taxon.Text = Globals.ThisAddIn.currentTaxon.Infra2Taxon;
            this.textboxInfra2Author.Text = Globals.ThisAddIn.currentTaxon.Infra2Author;
            this.textboxMorphDescription.Text = Globals.ThisAddIn.currentTaxon.MorphDescription;
            this.textboxCommonNames.Text = Globals.ThisAddIn.currentTaxon.CommonNames;
            this.comboboxFloweringStart.SelectedIndex = Globals.ThisAddIn.currentTaxon.FloweringStart;
            this.comboboxFloweringEnd.SelectedIndex = Globals.ThisAddIn.currentTaxon.FloweringEnd;
            this.textboxDistribution.Text = Globals.ThisAddIn.currentTaxon.Distribution;
            this.textboxHabitat.Text = Globals.ThisAddIn.currentTaxon.Habitat;
            this.textboxMinAlt.Text = Globals.ThisAddIn.currentTaxon.MinAlt.ToString();
            this.textboxMaxAlt.Text = Globals.ThisAddIn.currentTaxon.MaxAlt.ToString();
            this.textboxNotes.Text = Globals.ThisAddIn.currentTaxon.Notes;
            this.textboxVouchers.Text = Globals.ThisAddIn.currentTaxon.Vouchers;
        }

        private string FilterText(string inputText)
        {
            string filteredText = inputText;
            string numberPattern = @"(?:^|\s)([1-9](?:\d*|(?:\d{0,2})(?:,\d{3})*)(?:\.\d*[1-9])?|0?\.\d*[1-9]|0)";
            string numberPattern2 = @"([1-9](?:\d*|(?:\d{0,2})(?:,\d{3})*)(?:\.\d*[1-9])?|0?\.\d*[1-9]|0)";
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

            // add spaces between digits and cm, ft, ln measurements
            filteredText = Regex.Replace(filteredText, @"(\d)cm", "$1 cm");
            filteredText = Regex.Replace(filteredText, @"(\d)ft", "$1 ft");
            filteredText = Regex.Replace(filteredText, @"(\d)ln", "$1 ln");

            // replace comma with period as decimal separator (but only for 1 or 2 digits after comma, and no spaces)
            filteredText = Regex.Replace(filteredText, @"([0-9])+\,([0-9]{1,2})(?![0-9])", "$1.$2");

            // remove spaces between numbers and hyphens (12 - 15 -> 12-15)
            filteredText = Regex.Replace(filteredText, @"(\d) ?- ?(\d)", "$1-$2");

            // convert ft to meters
            string pattern = numberPattern + @"(?:-)?" + numberPattern2 + @"?" + @" ft\.?";
            filteredText = Regex.Replace(filteredText, pattern, ft_to_m);

            // convert inches to centimeters
            pattern = numberPattern + @"(?:-)?" + numberPattern2 + @"?" + @" in\.?";
            filteredText = Regex.Replace(filteredText, pattern, in_to_cm);

            // convert lin to mm
            pattern = numberPattern + @"(?:-)?" + numberPattern2 + @"?" + @" lin\.?";
            filteredText = Regex.Replace(filteredText, pattern, lin_to_mm);

            // add spaces between digits and multiplication signs (e.g. 3x5 -> 3 x 5)
            filteredText = Regex.Replace(filteredText, @"(\d) ?[xX\*\u2022\u00D7] ?(\d)", "$1 x $2");

            // remove invisible characters (optional hyphens, etc.)
            filteredText = Regex.Replace(filteredText, @"[^\w\s.,!@#$%^&*()-=+~`]", "");
            
            return filteredText;
        }

        public static decimal DecimalFromString(string inputString)
        {
            NumberStyles style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;
            CultureInfo provider = new CultureInfo("en-US");
            return Decimal.Parse(inputString, style, provider);
        }

        public static string ft_to_m(Match match)
        {
            int groupcount = 0;
            int decimalPlaces = 0;
            string returnString = " ";
            foreach (Group group in match.Groups)
            {
                if (group.ToString().IsNumber())
                {
                    if (groupcount == 1)
                    {
                        returnString += "-";
                    }
                    decimal baseNumber = 0;
                    try
                    {
                        baseNumber = DecimalFromString(group.ToString());
                    }
                    catch (FormatException)
                    {
                        return match.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error converting ft to m: " + ex.ToString(), "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return match.ToString();
                    }
                    if ((baseNumber * CollectionData.ft_to_m < 5) & (groupcount == 0))
                    {
                        decimalPlaces = 1;
                    }
                    returnString += (Math.Round((baseNumber * CollectionData.ft_to_m), decimalPlaces, MidpointRounding.AwayFromZero)).ToString(new CultureInfo("en-US"));
                    groupcount++;
                }
            }
            return returnString + " m";
        }

        public static string in_to_cm(Match match)
        {
            bool is_mm = false;
            int groupcount = 0;
            int decimalPlaces = 0;
            string returnString = " ";
            foreach (Group group in match.Groups)
            {
                if (group.ToString().IsNumber())
                {
                    if (groupcount == 0 )
                    {
                        if (DecimalFromString(group.ToString()) < 4m )
                        {
                            is_mm = true;
                        }
                    }
                    if ( groupcount == 1 )
                    {
                        returnString += "-";
                    }
                    decimal baseNumber = 0;
                    try
                    {
                        baseNumber = DecimalFromString(group.ToString());
                    }
                    catch (FormatException)
                    {
                        return match.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error converting in to cm: " + ex.ToString(), "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return match.ToString();
                    }
                    if (is_mm)
                    {
                        if ((baseNumber * CollectionData.in_to_mm < 5) & (groupcount == 0))
                        {
                            decimalPlaces = 1;
                        }
                        returnString += (Math.Round((baseNumber * CollectionData.in_to_mm), decimalPlaces, MidpointRounding.AwayFromZero)).ToString(new CultureInfo("en-US"));
                    }
                    else
                    {
                        if ((baseNumber * (CollectionData.in_to_mm / 10) < 5) & (groupcount == 0))
                        {
                            decimalPlaces = 1;
                        }
                        returnString += (Math.Round((baseNumber * (CollectionData.in_to_mm / 10)), decimalPlaces, MidpointRounding.AwayFromZero)).ToString(new CultureInfo("en-US"));
                    }
                    groupcount++;
                }
            }
            if (is_mm)
            {
                return returnString + " mm";
            }
            else
            {
                return returnString + " cm";
            }
        }

        public static string lin_to_mm(Match match)
        {
            int groupcount = 0;
            int decimalPlaces = 0;
            string returnString = " ";
            foreach (Group group in match.Groups)
            {
                if (group.ToString().IsNumber())
                {
                    if (groupcount == 1)
                    {
                        returnString += "-";
                    }
                    decimal baseNumber = 0;
                    try
                    {
                        baseNumber = DecimalFromString(group.ToString());
                    }
                    catch (FormatException)
                    {
                        return match.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error converting lin to mm: " + ex.ToString(), "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return match.ToString();
                    }
                    if ((baseNumber * CollectionData.lin_to_mm < 5) & (groupcount == 0))
                    {
                        decimalPlaces = 1;
                    }
                    returnString += (Math.Round((baseNumber * CollectionData.lin_to_mm), decimalPlaces, MidpointRounding.AwayFromZero)).ToString(new CultureInfo("en-US"));
                    groupcount++;
                }
            }
            return returnString + " mm";
        }

        private short GetMonthNumber(string lookup)
        {
            short outMonth = 0;
            lookup = lookup.Trim().TrimEnd('?', '!', '.', ',');
            CollectionData.MonthLookup.TryGetValue(lookup, out outMonth);
            return outMonth;
        }

        private void CopySelection(TextBox textboxTarget, bool capitalise = true)
        {
            string copyText = Globals.ThisAddIn.currentText;
            if (!String.IsNullOrWhiteSpace(copyText))
            {
                if (capitalise)
                {
                    copyText = char.ToUpper(copyText[0]) + copyText.Substring(1);
                }
                textboxTarget.Text = FilterText(copyText).Trim().RemoveInvalidXmlChars();
                if (!textboxTarget.Text.EndsWith("."))
                {
                    textboxTarget.Text += ".";
                }
            }
        }

        private void AddSelection(TextBox textboxTarget, string separator = "", bool capitalise = true, bool addPeriod = true)
        {
            string copyText = Globals.ThisAddIn.currentText;
            if (!String.IsNullOrWhiteSpace(copyText))
            {
                if (capitalise)
                {
                    copyText = char.ToUpper(copyText[0]) + copyText.Substring(1);
                }
                textboxTarget.Text += separator + " " + FilterText(copyText).Trim().RemoveInvalidXmlChars();
                if (addPeriod & !textboxTarget.Text.EndsWith("."))
                {
                    textboxTarget.Text += ".";
                }
            }
        }

        private void CopyNumber(TextBox textboxTarget)
        {
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                bool is_ft = false;
                string inputString = Globals.ThisAddIn.currentText.Trim();
                string cleanString = inputString.Clean();
                if (cleanString.Last(2) == "ft")
                {
                    is_ft = true;
                    cleanString = cleanString.ExceptLast(2);
                }
                else if (cleanString.Last(1) == "m")
                {
                    is_ft = false;
                    cleanString = cleanString.ExceptLast(1);
                }
                int numberValue = Int32.Parse(cleanString.Trim());
                if (is_ft)
                {
                    numberValue = (Int32)Math.Round((numberValue * 0.3048),0,MidpointRounding.AwayFromZero);
                }
                textboxTarget.Text = numberValue.ToString();
            }
        }

        private void CopyMonth(ComboBox comboboxTarget)
        {
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                short returnMonth = GetMonthNumber(Globals.ThisAddIn.currentText);
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
            CopySelection(this.textboxGenus);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnSpeciesCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxSpecies);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnSpeciesAuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxSpeciesAuthor);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnInfraTaxon1Copy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra1Taxon);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnInfra1AuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra1Author);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra1Rank();
            if (this.radiobuttonInfra1None.Checked)
            {
                this.panelInfraspecific1.Visible = false;
                this.radiobuttonInfra2None.Checked = true;
            }
            else
            {
                this.panelInfraspecific1.Visible = true;
            }
        }

        private void UpdateInfra1Rank()
        {
            if (this.radiobuttonInfra1None.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra1Rank = String.Empty;
            }
            if (this.radiobuttonInfra1Subspecies.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra1Rank = "subsp.";
            }
            if (this.radiobuttonInfra1Var.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra1Rank = "var.";
            }
            if (this.radiobuttonInfra1Form.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra1Rank = "f.";
            }
        }

        private void UpdateInfra2Rank()
        {
            if (this.radiobuttonInfra2None.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra2Rank = String.Empty;
            }
            if (this.radiobuttonInfra2Var.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra2Rank = "var.";
            }
            if (this.radiobuttonInfra2Form.Checked == true)
            {
                Globals.ThisAddIn.currentTaxon.Infra2Rank = "f.";
            }
        }

        private void btnInfra2TaxonCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra2Taxon);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnInfra2AuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra2Author);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void radiobuttonInfra2None_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra2Rank();
            if (this.radiobuttonInfra2None.Checked)
            {
                this.panelInfraspecific2.Visible = false;
            }
            else
            {
                this.panelInfraspecific2.Visible = true;
            }
        }

        private void TaxonPanel_Load(object sender, EventArgs e)
        {

        }

        private void btnMorphDescriptionCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxMorphDescription);
        }

        private void btnDistributionCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxDistribution);
        }

        private void btnHabitatCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxHabitat);
        }

        private void btnFloweringStartCopy_Click(object sender, EventArgs e)
        {
            CopyMonth(this.comboboxFloweringStart);
            Globals.ThisAddIn.currentTaxon.FloweringStart = (short)this.comboboxFloweringStart.SelectedIndex;
        }

        private void btnFloweringEndCopy_Click(object sender, EventArgs e)
        {
            CopyMonth(this.comboboxFloweringEnd);
            Globals.ThisAddIn.currentTaxon.FloweringEnd = (short)this.comboboxFloweringEnd.SelectedIndex;
        }

        private void btnNotesCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxNotes);
        }

        private void btnVouchersCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxVouchers);
        }

        private void btnNotesAdd_Click(object sender, EventArgs e)
        {
            AddSelection(this.textboxNotes);
        }

        private void btnVouchersAdd_Click(object sender, EventArgs e)
        {
            AddSelection(this.textboxVouchers);
        }

        private void textboxMorphDescription_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(this.textboxMorphDescription);
        }

        private void textboxDistribution_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(this.textboxDistribution);
        }

        private void textboxHabitat_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(this.textboxHabitat);
        }

        private void textboxNotes_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(this.textboxNotes);
        }

        private void textboxVouchers_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(this.textboxVouchers);
        }

        private void btnMinAltCopy_Click(object sender, EventArgs e)
        {
            CopyNumber(this.textboxMinAlt);
        }

        private void btnMaxAltCopy_Click(object sender, EventArgs e)
        {
            CopyNumber(this.textboxMaxAlt);
        }

        private void textboxGenus_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Genus = this.textboxGenus.Text;
            if (textboxGenus.Focused)
            { 
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void textboxSpecies_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Species = this.textboxSpecies.Text;
            if (textboxSpecies.Focused)
            {
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void textboxSpeciesAuthor_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.SpeciesAuthor = this.textboxSpeciesAuthor.Text;
            if (textboxSpeciesAuthor.Focused)
            {
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void radiobuttonInfra1Subspecies_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra1Rank();
        }

        private void radiobuttonInfra1Var_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra1Rank();
        }

        private void radiobuttonInfra1Form_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra1Rank();
        }

        private void radiobuttonInfra2Var_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra2Rank();
        }

        private void radiobuttonInfra2Form_CheckedChanged(object sender, EventArgs e)
        {
            UpdateInfra2Rank();
        }

        private void textboxInfra1Taxon_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Infra1Taxon = this.textboxInfra1Taxon.Text;
            if (textboxInfra1Taxon.Focused)
            {
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void textboxInfra1Author_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Infra1Author = this.textboxInfra1Author.Text;
            if (textboxInfra1Author.Focused)
            {
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void textboxInfra2Taxon_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Infra2Taxon = this.textboxInfra2Taxon.Text;
            if (textboxInfra2Taxon.Focused)
            {
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void textboxInfra2Author_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Infra2Author = this.textboxInfra2Author.Text;
            if (textboxInfra2Author.Focused)
            {
                Globals.ThisAddIn.UpdateTrackingNumber();
            }
        }

        private void textboxMorphDescription_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.MorphDescription = this.textboxMorphDescription.Text;
        }

        private void textboxDistribution_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Distribution = this.textboxDistribution.Text;
        }

        private void textboxHabitat_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Habitat = this.textboxHabitat.Text;
        }

        private void textboxMinAlt_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.MinAlt = int.Parse(this.textboxMinAlt.Text);
        }

        private void textboxMaxAlt_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.MaxAlt = int.Parse(this.textboxMaxAlt.Text);
        }

        private void textboxNotes_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Notes = this.textboxNotes.Text;
        }

        private void textboxVouchers_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.Vouchers = this.textboxVouchers.Text;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Serialize();
            Globals.ThisAddIn.currentTaxon = new Taxon();
            Globals.ThisAddIn.currentBatch.Add(Globals.ThisAddIn.currentTaxon);
            LoadCurrentTaxon();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Serialize();
            Globals.ThisAddIn.currentBatch.MoveNext();
            LoadCurrentTaxon();
        }

        private void btnPrevious_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Serialize();
            Globals.ThisAddIn.currentBatch.MovePrevious();
            LoadCurrentTaxon();
        }

        private void comboboxFloweringStart_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.FloweringStart = (short)this.comboboxFloweringStart.SelectedIndex;
        }

        private void comboboxFloweringEnd_SelectionChangeCommitted(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.FloweringEnd = (short)this.comboboxFloweringEnd.SelectedIndex;
        }

        private void btnMorphDescriptionAdd_Click(object sender, EventArgs e)
        {
            AddSelection(this.textboxMorphDescription);
        }

        private void btnDistributionAdd_Click(object sender, EventArgs e)
        {
            AddSelection(this.textboxDistribution);
        }

        private void btnHabitatAdd_Click(object sender, EventArgs e)
        {
            AddSelection(this.textboxHabitat);
        }

        private void comboboxFloweringEnd_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnCommonNames_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(this.textboxCommonNames.Text))
            {
                AddSelection(this.textboxCommonNames, ";", true, false);
            }
            else
            {
                AddSelection(this.textboxCommonNames,"",true, false);
            }
        }

        private void textboxCommonNames_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.CommonNames = this.textboxCommonNames.Text;
        }


    }
}
