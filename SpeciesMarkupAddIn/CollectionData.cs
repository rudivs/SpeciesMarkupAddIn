using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeciesMarkupAddIn
{
    public static class CollectionData
    {
        public static readonly Dictionary<int, string> Months = new Dictionary<int, string>
        {
            { 0, "Not provided" },
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


        public static readonly Dictionary<string, int> MonthLookup = new Dictionary<string, int>
        {
            { "1", 1 },
            { "01", 1 },
            { "JA", 1 },
            { "Jan", 1 },
            { "January", 1},
            { "2", 2 },
            { "02", 2 },
            { "FE", 2 },
            { "Feb", 2 },
            { "February", 2 },
            { "3", 3 },
            { "03", 3 },
            { "MR", 3 },
            { "Mar", 3 },
            { "March", 3 },
            { "4", 4 },
            { "04", 4 },
            { "AL", 4 },
            { "AP", 4 },
            { "Apr", 4 },
            { "April", 4 },
            { "5", 5 },
            { "05", 5 },
            { "MA", 5 },
            { "May", 5 },
            { "6", 6 },
            { "06", 6 },
            { "JN", 6 },
            { "Jun", 6 },
            { "June", 6 },
            { "7", 7 },
            { "07", 7 },
            { "JL", 7 },
            { "Jul", 7 },
            { "July", 7 },
            { "8", 8 },
            { "08", 8 },
            { "AU", 8 },
            { "Aug", 8 },
            { "August", 8 },
            { "9", 9 },
            { "09", 9 },
            { "SE", 9 },
            { "Sep", 9 },
            { "Sept", 9 },
            { "September", 9 },
            { "10", 10 },
            { "OC", 10 },
            { "Oct", 10 },
            { "October", 10 },
            { "11", 11 },
            { "NO", 11 },
            { "NV", 11 },
            { "Nov", 11 },
            { "November", 11},
            { "12", 12 },
            { "DE", 12 },
            { "DC", 12 },
            { "Dec", 12 },
            { "December", 12 }
        };
    }
}
