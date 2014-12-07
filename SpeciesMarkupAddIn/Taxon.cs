using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeciesMarkupAddIn
{
    public class Taxon
    {
        public Taxon()
        {

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
        public string Description { get; set; }
        public Int16 FloweringStart { get; set; }
        public Int16 FloweringEnd { get; set; }
        public string Distribution { get; set; }
        public string Habitat { get; set; }
        public int MinAlt { get; set; }
        public int MaxAlt { get; set; }
        public string Notes { get; set; }
        public string Vouchers { get; set; }
    }

    public class TaxonList
    {
        private List<Taxon> taxa;
        private int _index;

        public TaxonList()
        {
            taxa = new List<Taxon>();
            _index = -1;
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
