using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeciesMarkupAddIn
{
    public class TaxonInternalReferences
    {
        private string _PropertyName;
        private string _FieldName;
        private string _Label;
        private double _ColumnWidth;
        private Color _FillColor;

        public string PropertyName { get {return _PropertyName;}}
        public string FieldName { get { return _FieldName; } }
        public string Label { get { return _Label; } }
        public double ColumnWidth { get { return _ColumnWidth; } }
        public Color FillColor { get { return _FillColor; } }

        public TaxonInternalReferences(string fieldName, string label, Color fillColor, double columnWidth = 8.43, string propertyName = null)
        {
            _PropertyName = propertyName;
            _FieldName = fieldName;
            _Label = label;
            _ColumnWidth = columnWidth;
            _FillColor = fillColor;
        }
    }

    public static class CollectionData
    {
        public const decimal ft_to_m = 0.3048m;
        public const decimal in_to_mm = 25.4m;
        public const decimal lin_to_mm = 2.11667m;

        public const string DeleteAfterExport = "Export complete. Would you like to clear the current taxon markup batch?";
        public const string DeleteAdHoc = "Are you sure you want to clear the current taxon markup batch? This can not be undone.";

        public static string BatchFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CurrentTaxonBatch.xml");

        internal static readonly Dictionary<short, string> Months = new Dictionary<short, string>
        {
            { 0, "Not specified" },
            { 1, "January" },
            { 2, "February" },
            { 3, "March" },
            { 4, "April" },
            { 5, "May" },
            { 6, "June" },
            { 7, "July" },
            { 8, "August" },
            { 9, "September" },
            { 10, "October" },
            { 11, "November" },
            { 12, "December" }
        };


        internal static readonly Dictionary<string, short> MonthLookup = new Dictionary<string, short>
        {
            { "1", 1 },
            { "01", 1 },
            { "I", 1 },
            { "JA", 1 },
            { "Jan", 1 },
            { "January", 1},
            { "2", 2 },
            { "02", 2 },
            { "II", 2 },
            { "FE", 2 },
            { "Feb", 2 },
            { "February", 2 },
            { "3", 3 },
            { "03", 3 },
            { "III", 3 },
            { "MR", 3 },
            { "Mar", 3 },
            { "March", 3 },
            { "4", 4 },
            { "04", 4 },
            { "IV", 4 },
            { "AL", 4 },
            { "AP", 4 },
            { "Apr", 4 },
            { "April", 4 },
            { "5", 5 },
            { "05", 5 },
            { "V", 5 },
            { "MA", 5 },
            { "May", 5 },
            { "6", 6 },
            { "06", 6 },
            { "VI", 6 },
            { "JN", 6 },
            { "Jun", 6 },
            { "June", 6 },
            { "7", 7 },
            { "07", 7 },
            { "VII", 7 },
            { "JL", 7 },
            { "Jul", 7 },
            { "July", 7 },
            { "8", 8 },
            { "08", 8 },
            { "VIII", 8 },
            { "AU", 8 },
            { "Aug", 8 },
            { "August", 8 },
            { "9", 9 },
            { "09", 9 },
            { "IX", 9 },
            { "SE", 9 },
            { "Sep", 9 },
            { "Sept", 9 },
            { "September", 9 },
            { "10", 10 },
            { "X", 10 },
            { "OC", 10 },
            { "Oct", 10 },
            { "October", 10 },
            { "11", 11 },
            { "XI", 11 },
            { "NO", 11 },
            { "NV", 11 },
            { "Nov", 11 },
            { "November", 11},
            { "12", 12 },
            { "XII", 12 },
            { "DE", 12 },
            { "DC", 12 },
            { "Dec", 12 },
            { "December", 12 }
        };

        internal static readonly Dictionary<int, TaxonInternalReferences> columnIndex = new Dictionary<int, TaxonInternalReferences>
        {
            { 1, new TaxonInternalReferences("Full name","Full name",Color.Bisque,30,"FullName") },
            { 2, new TaxonInternalReferences("Tracking number","Tracking number",Color.Bisque,8,"TrackingNumber") },
            { 3, new TaxonInternalReferences("Spnumber","Spnumber",Color.Bisque) },
            { 4, new TaxonInternalReferences("Gename","Genus",Color.LightGreen,15,"Genus") },
            { 5, new TaxonInternalReferences("SP1","Species",Color.LightGreen,15,"Species") },
            { 6, new TaxonInternalReferences("Author1","Author",Color.LightGreen,30,"SpeciesAuthor") },
            { 7, new TaxonInternalReferences("Rank1","Rank",Color.LightGreen,5,"Infra1Rank") },
            { 8, new TaxonInternalReferences("SP2","Infraspecific taxon",Color.LightGreen,15,"Infra1Taxon") },
            { 9, new TaxonInternalReferences("Author2","Author",Color.LightGreen,30,"Infra1Author") },
            { 10, new TaxonInternalReferences("Rank2","Rank",Color.LightGreen,5,"Infra2Rank") },
            { 11, new TaxonInternalReferences("SP3","Infraspecific taxon",Color.LightGreen,15,"Infra2Taxon") },
            { 12, new TaxonInternalReferences("Author3","Author",Color.LightGreen,30,"Infra2Author") },
            { 13, new TaxonInternalReferences("CommonNames","Common names",Color.LightGreen,40,"CommonNames") },
            { 14, new TaxonInternalReferences("Descrip","Morphological description",Color.LightGreen,50,"MorphDescription") },
            { 15, new TaxonInternalReferences("FLRSTART","Flowering time start",Color.LightGreen,7,"FloweringStart") },
            { 16, new TaxonInternalReferences("FLREND","Flowering time end",Color.LightGreen,7,"FloweringEnd") },
            { 17, new TaxonInternalReferences("ChromosomeNumber","Chromosome number",Color.LightGreen,7,"ChromosomeNumber") },
            { 18, new TaxonInternalReferences("Distrib","Distribution",Color.LightGreen,50,"Distribution") },
            { 19, new TaxonInternalReferences("Habitat","Habitat",Color.LightGreen,50,"Habitat") },
            { 20, new TaxonInternalReferences("Minalt", "Minimum altitude",Color.LightGreen,7,"MinAlt") },
            { 21, new TaxonInternalReferences("Maxalt", "Maximum altitude",Color.LightGreen,7,"MaxAlt") },
            { 22, new TaxonInternalReferences("LITNOTE", "Literature note",Color.LightGreen,50,"Notes") },
            { 23, new TaxonInternalReferences("SpecCit", "Voucher specimens",Color.LightGreen,50,"Vouchers") },
            { 24, new TaxonInternalReferences("Admin", "Admin",Color.OrangeRed,13) },
            { 25, new TaxonInternalReferences("MarkupDate", "Date of markup",Color.OrangeRed,13) },
            { 26, new TaxonInternalReferences("PermissionPublisher", "Permission from publishing house",Color.OrangeRed,19) },
            { 27, new TaxonInternalReferences("PermissionAuthor", "Permission from Author",Color.OrangeRed,15) },
            { 28, new TaxonInternalReferences("PublicationCategory","Category",Color.MediumPurple,19) },
            { 29, new TaxonInternalReferences("Title","Journal or book title",Color.MediumPurple,30) },
            { 30, new TaxonInternalReferences("Printyear", "Publication year",Color.MediumPurple,8) },
            { 31, new TaxonInternalReferences("Author(s)", "Author(s)",Color.MediumPurple,30) },
            { 32, new TaxonInternalReferences("PublicationTitle", "Publication title",Color.MediumPurple,30) },
            { 33, new TaxonInternalReferences("Editors","Editor(s)",Color.MediumPurple,30) },
            { 34, new TaxonInternalReferences("Publisher","Publisher",Color.MediumPurple, 22) },
            { 35, new TaxonInternalReferences("PublisherCity","Publisher city",Color.MediumPurple, 22) },
            { 36, new TaxonInternalReferences("Volume","Volume",Color.MediumPurple,6) },
            { 37, new TaxonInternalReferences("Part", "Part",Color.MediumPurple,6) },
            { 38, new TaxonInternalReferences("Fascicle", "Fascicle",Color.MediumPurple,8) },
            { 39, new TaxonInternalReferences("Pagestart", "Page start",Color.MediumPurple,9) },
            { 40, new TaxonInternalReferences("Pageend", "Page end",Color.MediumPurple,9) }
        };
    }
}
