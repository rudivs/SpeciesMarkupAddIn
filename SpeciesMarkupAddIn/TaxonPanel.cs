﻿using System;
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
using log4net;

namespace SpeciesMarkupAddIn
{
    public partial class TaxonPanel : UserControl
    {
        private static readonly log4net.ILog log = 
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

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
            if (Globals.ThisAddIn.currentBatch.Count > 0)
            {
                Globals.ThisAddIn.currentTaxon = Globals.ThisAddIn.currentBatch.Current;
            }
            else
            {
                Globals.ThisAddIn.currentTaxon = new Taxon();
                Globals.ThisAddIn.currentBatch.Add(Globals.ThisAddIn.currentTaxon);
            }

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
            this.textboxChromosomeNumber.Text = Globals.ThisAddIn.currentTaxon.ChromosomeNumber;
            this.textboxDistribution.Text = Globals.ThisAddIn.currentTaxon.Distribution;
            this.textboxHabitat.Text = Globals.ThisAddIn.currentTaxon.Habitat;
            this.textboxMinAlt.Text = Globals.ThisAddIn.currentTaxon.MinAlt.ToString();
            this.textboxMaxAlt.Text = Globals.ThisAddIn.currentTaxon.MaxAlt.ToString();
            this.textboxNotes.Text = Globals.ThisAddIn.currentTaxon.Notes;
            this.textboxVouchers.Text = Globals.ThisAddIn.currentTaxon.Vouchers;
            Globals.Ribbons.Ribbon.UpdateCount();
        }

        private string FilterText(string inputText)
        {
            string filteredText = inputText;
            string numberPattern = @"(?:^|\s)([1-9](?:\d*|(?:\d{0,2})(?:,\d{3})*)(?:\.\d*[1-9])?|0?\.\d*[1-9]|0)";
            string numberPattern2 = @"([1-9](?:\d*|(?:\d{0,2})(?:,\d{3})*)(?:\.\d*[1-9])?|0?\.\d*[1-9]|0)";

            // replace nonbreaking spaces with regular spaces
            filteredText = Regex.Replace(filteredText, @"\p{Zs}", " ");
            
            // remove unnecessary lines
            filteredText = Regex.Replace(filteredText, @"(\r\n?|\n){3,}", "\r\n\r\n");

            // make newline characters work properly
            filteredText = Regex.Replace(filteredText, @"(\r\n?|\n)", "\r\n");

            // replace tab with space
            filteredText = Regex.Replace(filteredText, @"\t", " ");

            // replace runs of spaces with a single space
            filteredText = Regex.Replace(filteredText, @"(\p{Zs}){2,}", " ");

            // fix comma space
            filteredText = filteredText.Replace(" , ",", ");
            
            // fix dashes
            filteredText = filteredText.Replace("–", "-");
            filteredText = filteredText.Replace("—", "-");

            // remove spaces between numbers - handles cases when spaces used as thousands separator
            //filteredText = Regex.Replace(filteredText, @"(\d+) (\d+)", "$1$2");
            filteredText = Regex.Replace(filteredText, @"(?<=\d)\p{Zs}(?=\d)", "");

            // add spaces between digits and metre-based measurements
            filteredText = Regex.Replace(filteredText, @"(\d)m", "$1 m");

            // add spaces between digits and cm, ft, ln measurements
            filteredText = Regex.Replace(filteredText, @"(\d)cm\b", "$1 cm");
            filteredText = Regex.Replace(filteredText, @"(\d)ft\b", "$1 ft");
            filteredText = Regex.Replace(filteredText, @"(\d)lin\b", "$1 lin");
            filteredText = Regex.Replace(filteredText, @"(\d)in\b", "$1 in");

            // replace comma with period as decimal separator (but only for 1 or 2 digits after comma, and no spaces)
            filteredText = Regex.Replace(filteredText, @"([0-9])+\,([0-9]{1,2})(?![0-9])", "$1.$2");

            // remove spaces between numbers and hyphens (12 - 15 -> 12-15)
            filteredText = Regex.Replace(filteredText, @"(\d) ?- ?(\d)", "$1-$2");

            // convert ft to meters
            string pattern = numberPattern + @"(?:-)?" + numberPattern2 + @"?" + @" ft\.?(?= )";
            filteredText = Regex.Replace(filteredText, pattern, ft_to_m);

            // convert inches to centimeters
            pattern = numberPattern + @"(?:-)?" + numberPattern2 + @"?" + @" in.(?= )";
            filteredText = Regex.Replace(filteredText, pattern, in_to_cm);

            // convert lin to mm
            pattern = numberPattern + @"(?:-)?" + numberPattern2 + @"?" + @" lin\.?(?= )";
            filteredText = Regex.Replace(filteredText, pattern, lin_to_mm);

            // add spaces between digits and multiplication signs (e.g. 3x5 -> 3 x 5)
            filteredText = Regex.Replace(filteredText, @"(\d) ?[xX\*\u2022\u00D7] ?(\d)", "$1 x $2");

            // format chromosome numbers properly
            filteredText = Regex.Replace(filteredText, @"\b(\d+)\p{Zs}?n\p{Zs}?=\p{Zs}?(\d+)\b", "$1n = $2");
            
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
            return String.Format("{0} m [as \"{1}\"]",returnString,match.ToString().Trim());
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
                return String.Format("{0} mm [as \"{1}\"]",returnString,match.ToString().Trim());
            }
            else
            {
                return String.Format("{0} cm [as \"{1}\"]",returnString,match.ToString().Trim());
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
            return String.Format("{0} mm [as \"{1}\"]",returnString,match.ToString().Trim());
        }

        private short GetMonthNumber(string lookup)
        {
            short outMonth = 0;
            lookup = lookup.Trim().TrimEnd('?', '!', '.', ',');
            CollectionData.MonthLookup.TryGetValue(lookup, out outMonth);
            return outMonth;
        }

        private void AuthorFilter(TextBox textboxTarget)
        {
            // replace 'et' with '&' in author names
            textboxTarget.Text = Regex.Replace(textboxTarget.Text, @"\bet\b", "&");
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
            }
        }

        private void AddSelection(TextBox textboxTarget, string separator = "", bool capitalise = true, 
            bool addPeriod = false, bool toLower = false, string comment = null)
        {
            string copyText = Globals.ThisAddIn.currentText;
            string spaceChar = " ";
            var sentenceEndings = new String[] { ".", "?", "!" };
            if (String.IsNullOrWhiteSpace(textboxTarget.Text))
            {
                spaceChar = "";
            }
            if (!String.IsNullOrWhiteSpace(copyText))
            {
                if (toLower)
                {
                    copyText = copyText.ToLower();
                }
                if (capitalise)
                {
                    if (sentenceEndings.Any(c => textboxTarget.Text.EndsWith(c)))
                    {
                        copyText = char.ToUpper(copyText[0]) + copyText.Substring(1);
                    }
                }
                
                string commentText = String.IsNullOrWhiteSpace(comment) ? "" : 
                    String.Format(" ({0})", comment.Trim().RemoveInvalidXmlChars());
                textboxTarget.Text += separator + spaceChar + FilterText(copyText).Trim().RemoveInvalidXmlChars() 
                    + commentText;
                if (addPeriod & !sentenceEndings.Any(c => textboxTarget.Text.EndsWith(c)))
                {
                    textboxTarget.Text += ".";
                }
            }
        }

        private void CopyNumber(TextBox textboxTarget)
        {
            NumberStyles style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;
            CultureInfo provider = new CultureInfo("en-US");

            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentText))
            {
                bool is_ft = false;
                string inputString = Globals.ThisAddIn.currentText.Trim();
                //string cleanString = inputString.Clean();
                string cleanString = Regex.Replace(inputString, @"\s+", "");
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
                try
                {
                    decimal decimalValue = Decimal.Parse(cleanString.Trim(), style, provider);
                    int numberValue = (Int32)Math.Round(decimalValue, 0, MidpointRounding.AwayFromZero);
                    if (is_ft)
                    {
                        numberValue = (Int32)Math.Round((numberValue * 0.3048), 0, MidpointRounding.AwayFromZero);
                    }
                    textboxTarget.Text = numberValue.ToString();
                }
                catch (FormatException)
                {
                    log.Warn("Selected text cannot be converted to number.");
                }
                
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
                var textbox = form.Controls[0].Controls["textboxEditText"];
                textbox.Font = new Font(textbox.Font.FontFamily, Properties.Settings.Default.EditFontSize);
                textbox.Text = textboxTarget.Text;
                var result = form.ShowDialog();
                if (result == DialogResult.OK)
                {
                    textboxTarget.Text = form.ReturnText;
                }
            }

        }

        private void btnGenusCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxGenus,false);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnSpeciesCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxSpecies, false);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnSpeciesAuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxSpeciesAuthor,false);
            AuthorFilter(this.textboxSpeciesAuthor);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnInfraTaxon1Copy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra1Taxon,false);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnInfra1AuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra1Author,false);
            AuthorFilter(this.textboxInfra1Author);
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
            CopySelection(this.textboxInfra2Taxon,false);
            Globals.ThisAddIn.UpdateTrackingNumber();
        }

        private void btnInfra2AuthorCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxInfra2Author,false);
            AuthorFilter(this.textboxInfra2Author);
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
            Globals.ThisAddIn.currentTaxon.MinAlt = this.textboxMinAlt.Text.TryParseNullableInt();
        }

        private void textboxMaxAlt_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.MaxAlt = this.textboxMaxAlt.Text.TryParseNullableInt();
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
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentTaxon.FullName))
            {
                Globals.ThisAddIn.currentTaxon = new Taxon();
                Globals.ThisAddIn.currentBatch.Add(Globals.ThisAddIn.currentTaxon);
                LoadCurrentTaxon();
            }
            Globals.Ribbons.Ribbon.UpdateCount();
        }

        private void btnNewCopyGenus_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Serialize();
            if (!String.IsNullOrWhiteSpace(Globals.ThisAddIn.currentTaxon.FullName))
            {
                String currentGenus = Globals.ThisAddIn.currentTaxon.Genus;
                Globals.ThisAddIn.currentTaxon = new Taxon();
                Globals.ThisAddIn.currentTaxon.Genus = currentGenus;
                Globals.ThisAddIn.currentBatch.Add(Globals.ThisAddIn.currentTaxon);
                LoadCurrentTaxon();
            }
            Globals.Ribbons.Ribbon.UpdateCount();
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

        private void btnCommonNames_Click(object sender, EventArgs e)
        {
            string language = GetLanguageSelect();
            if (language == null)
            {
                return;
            }

            if (language == "Other")
            {
                language = GetLanguageText();
            }

            if (!string.IsNullOrWhiteSpace(this.textboxCommonNames.Text))
            {
                AddSelection(this.textboxCommonNames, ";", false, false, true, language);
            }
            else
            {
                AddSelection(this.textboxCommonNames, "",false, false, true, language);
            }
        }

        private void textboxCommonNames_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.CommonNames = this.textboxCommonNames.Text;
        }


        private void textboxCommonNames_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            EditText(this.textboxCommonNames);
        }

        private void textboxChromosomeNumber_TextChanged(object sender, EventArgs e)
        {
            Globals.ThisAddIn.currentTaxon.ChromosomeNumber = this.textboxChromosomeNumber.Text;
        }

        private void btnChromosomeNumberCopy_Click(object sender, EventArgs e)
        {
            CopySelection(this.textboxChromosomeNumber,false);

            string pattern = @"\b\d+n = \d+\b";
            Match match = Regex.Match(this.textboxChromosomeNumber.Text, pattern);
            if (match.Success)
            {
                this.textboxChromosomeNumber.Text = match.Value;
            }
            else
            {
                this.textboxChromosomeNumber.Text = "";
            }
        }

        /// <summary>
        /// Generate a form with select box that prompts the user for common name language.
        /// </summary>
        private string GetLanguageSelect()
        {
            string language = null;
            using (Form LanguageForm = new Form())
            {
                LanguageForm.Text = @"Language";
                ComboBox LanguageOptions = new ComboBox();
                LanguageOptions.Font = new Font(LanguageOptions.Font.FontFamily, Properties.Settings.Default.EditFontSize);
                LanguageOptions.Location = new Point(5, 5);
                List<string> languages = new List<string>
                {
                    "English", "Afrikaans", "IsiZulu","IsiXhosa", "Sepedi", "Sesotho", "Setswana", "Xitsonga",
                    "Tshivenda", "IsiNdebele", "SiSwati", "Other"
                };
                for (int loop = 0; loop < languages.Count; loop++)
                    LanguageOptions.Items.Add(languages[loop]);
                int def = languages.IndexOf("English");
                if (def >= 0) LanguageOptions.SelectedIndex = def;

                Button Select = new Button();
                Select.Click += new EventHandler(delegate (Object o, EventArgs a)
                {
                    language = LanguageOptions.SelectedItem.ToString();
                    if (LanguageForm != null)
                        LanguageForm.Close();
                });
                Select.Text = "Select";
                Select.Location = new Point(5, LanguageOptions.Location.Y + LanguageOptions.Height + 5); ;
                LanguageForm.Controls.Add(LanguageOptions);
                LanguageForm.Controls.Add(Select);
                LanguageForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                LanguageForm.AutoSize = true;
                LanguageForm.Height = Select.Location.Y + Select.Height + 5;
                LanguageForm.Width = LanguageOptions.Location.X + LanguageOptions.Width + 90;
                LanguageForm.ShowDialog();
            }
            return language;
        }

        /// <summary>
        /// Generate a form with textbox that prompts the user for common name language.
        /// </summary>
        private string GetLanguageText()
        {
            string language;
            using (Form LanguageForm = new Form())
            {
                LanguageForm.Text = @"Language";
                Label Instruction = new Label();
                Instruction.Font = new Font(Instruction.Font.FontFamily, 10);
                Instruction.Text = "Please type in the language:";
                Instruction.Location = new Point(5, 5);
                Instruction.Size = new Size(200, Instruction.Size.Height);


                TextBox LanguageText = new TextBox();
                LanguageText.Font = new Font(LanguageText.Font.FontFamily, Properties.Settings.Default.EditFontSize);
                LanguageText.Width = Math.Max(LanguageText.Width, 200);
                LanguageText.Location = new Point(5, LanguageText.Location.Y + LanguageText.Height + 5);

                Button Select = new Button();
                Select.Click += new EventHandler(delegate (Object o, EventArgs a)
                {
                    if (LanguageForm != null)
                        LanguageForm.Close();
                });
                Select.Text = "Select";
                Select.Location = new Point(5, LanguageText.Location.Y + LanguageText.Height + 5); ;
                LanguageForm.Controls.Add(LanguageText);
                LanguageForm.Controls.Add(Instruction);
                LanguageForm.Controls.Add(Select);
                LanguageForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                LanguageForm.AutoSize = true;
                LanguageForm.Height = Select.Location.Y + Select.Height + 5;
                LanguageForm.Width = new int[]{ LanguageText.Location.X + LanguageText.Width+90, Instruction.Location.X + Instruction.Width+90}.Max();
                LanguageForm.ShowDialog();
                language = LanguageText.Text;
            }
            return language;
        }
    }
}
