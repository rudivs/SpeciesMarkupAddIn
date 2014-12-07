using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace SpeciesMarkupAddIn
{
    [Serializable()]
    public class Taxon : ISerializable
    {
        public Taxon()
        {

        }

        public Taxon(SerializationInfo info, StreamingContext ctxt)
        {
            this.FullName = (string)info.GetValue("FullName", typeof(string));
            this.TrackingNumber = (string)info.GetValue("TrackingNumber", typeof(string));
            this.Genus = (string)info.GetValue("Genus", typeof(string));
            this.Species = (string)info.GetValue("Species", typeof(string));
            this.SpeciesAuthor = (string)info.GetValue("SpeciesAuthor", typeof(string));
            this.Infra1Rank = (string)info.GetValue("Infra1Rank", typeof(string));
            this.Infra1Taxon = (string)info.GetValue("Infra1Taxon", typeof(string));
            this.Infra1Author = (string)info.GetValue("Infra1Author", typeof(string));
            this.Infra2Rank = (string)info.GetValue("Infra2Rank", typeof(string));
            this.Infra2Taxon = (string)info.GetValue("Infra2Taxon", typeof(string));
            this.Infra2Author = (string)info.GetValue("Infra2Author", typeof(string));
            this.MorphDescription = (string)info.GetValue("MorphDescription", typeof(string));
            this.FloweringStart = (Int16)info.GetValue("FloweringStart", typeof(Int16));
            this.FloweringEnd = (Int16)info.GetValue("FloweringEnd", typeof(Int16));
            this.Distribution = (string)info.GetValue("Distribution", typeof(string));
            this.Habitat = (string)info.GetValue("Habitat", typeof(string));
            this.MinAlt = (int)info.GetValue("MinAlt", typeof(int));
            this.MaxAlt = (int)info.GetValue("MaxAlt", typeof(int));
            this.Notes = (string)info.GetValue("Notes", typeof(string));
            this.Vouchers = (string)info.GetValue("Vouchers", typeof(string));
        }

        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            info.AddValue("FullName", this.FullName);
            info.AddValue("TrackingNumber", this.TrackingNumber);
            info.AddValue("Genus", this.Genus);
            info.AddValue("Species", this.Species);
            info.AddValue("SpeciesAuthor", this.SpeciesAuthor);
            info.AddValue("Infra1Rank", this.Infra1Rank);
            info.AddValue("Infra1Taxon", this.Infra1Taxon);
            info.AddValue("Infra1Author", this.Infra1Author);
            info.AddValue("Infra2Rank", this.Infra2Rank);
            info.AddValue("Infra2Taxon", this.Infra2Taxon);
            info.AddValue("Infra2Author", this.Infra2Author);
            info.AddValue("MorphDescription", this.MorphDescription);
            info.AddValue("FloweringStart", this.FloweringStart);
            info.AddValue("FloweringEnd", this.FloweringEnd);
            info.AddValue("Distribution", this.Distribution);
            info.AddValue("Habitat", this.Habitat);
            info.AddValue("MinAlt", this.MinAlt);
            info.AddValue("MaxAlt", this.MaxAlt);
            info.AddValue("Notes", this.Notes);
            info.AddValue("Vouchers", this.Vouchers);
        }

        public string FullName { get; set; }
        public string TrackingNumber { get; set; }
        public string Genus { get; set; }
        public string Species { get; set; }
        public string SpeciesAuthor { get; set; }
        public string Infra1Rank { get; set; }
        public string Infra1Taxon { get; set; }
        public string Infra1Author { get; set; }
        public string Infra2Rank { get; set; }
        public string Infra2Taxon { get; set; }
        public string Infra2Author { get; set; }
        public string MorphDescription { get; set; }
        public Int16 FloweringStart { get; set; }
        public Int16 FloweringEnd { get; set; }
        public string Distribution { get; set; }
        public string Habitat { get; set; }
        public int MinAlt { get; set; }
        public int MaxAlt { get; set; }
        public string Notes { get; set; }
        public string Vouchers { get; set; }
    }

    [Serializable()]
    public class TaxonList : ISerializable
    {
        private List<Taxon> taxa;
        private int _index;

        public TaxonList()
        {
            taxa = new List<Taxon>();
            _index = -1;
        }

        public TaxonList(SerializationInfo info, StreamingContext ctxt)
        {
            this.taxa = (List<Taxon>)info.GetValue("Taxa", typeof(List<Taxon>));
            this._index = (int)info.GetValue("Index", typeof(int));
        }

        public void GetObjectData(SerializationInfo info, StreamingContext ctxt)
        {
            info.AddValue("Taxa", this.taxa);
            info.AddValue("Index", this._index);
        }

        public void AddTaxon(Taxon taxon)
        {
            taxa.Add(taxon);
            _index = taxa.Count - 1;
        }

        public int Count
        {
            get
            {
                return taxa.Count;
            }
        }

        public Taxon Current
        {
            get
            {
                if (_index < 0 || _index > taxa.Count)
                {
                    throw new InvalidOperationException();
                }
                else
                {
                    return taxa[_index];
                }
            }
        }

        public bool MoveNext()
        {
            if (_index < taxa.Count -1)
            {
                _index++;
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool MovePrevious()
        {
            if (_index <= 0)
            {
                return false;
            }
            else
            {
                _index--;
                return true;
            }
        }
    }
}
