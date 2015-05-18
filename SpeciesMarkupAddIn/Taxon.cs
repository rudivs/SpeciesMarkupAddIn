using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SpeciesMarkupAddIn
{

    public class Taxon
    {
        [XmlIgnore]
        public string FullName 
        {
            get 
            {
                string[] myStrings = new string[]{this.Genus, this.Species, this.SpeciesAuthor, this.Infra1Rank, this.Infra1Taxon, 
                this.Infra1Author, this.Infra2Rank, this.Infra2Taxon,this.Infra2Author};
                return string.Join(" ", myStrings.Where(str => !string.IsNullOrEmpty(str)));
            }
        }
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
        public string CommonNames { get; set; }
        public string MorphDescription { get; set; }
        public short FloweringStart { get; set; }
        public short FloweringEnd { get; set; }
        public string ChromosomeNumber { get; set; }
        public string Distribution { get; set; }
        public string Habitat { get; set; }
        public int? MinAlt { get; set; }
        public int? MaxAlt { get; set; }
        public string Notes { get; set; }
        public string Vouchers { get; set; }

        public Taxon()
        {

        }
    }

    public class TaxonList : IEnumerable<Taxon>
    {
        private List<Taxon> taxa;
        private int _index;

        public TaxonList()
        {
            taxa = new List<Taxon>();
            _index = -1;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public IEnumerator<Taxon> GetEnumerator()
        {
            return (IEnumerator<Taxon>)taxa.GetEnumerator();
        }

        public void Add(Taxon taxon)
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

        public int Index
        {
            get
            {
                return _index;
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

        public Taxon GetByIndex(int index)
        {
            if (index < 0 || index > taxa.Count || taxa.Count == 0)
            {
                throw new InvalidOperationException();
            }
            else
            {
                return taxa[index];
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

        public bool MoveIndex(int index)
        {
            if (index < 0 || index > taxa.Count || taxa.Count == 0)
            {
                return false;
            }
            else
            {
                _index = index;
                return true;
            }
        }

        public bool DeleteCurrent()
        {
            if (taxa.Count >= 1)
            {
                taxa.RemoveAt(_index);
                if (_index > 0)
                {
                    _index--;
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Reset()
        {
            taxa.Clear();
            _index = -1;
        }
    }
}
