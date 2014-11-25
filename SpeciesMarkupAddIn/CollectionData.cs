using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeciesMarkupAddIn
{
    public static class CollectionData
    {
        public const decimal ft_to_m = 0.3048m;
        public const decimal in_to_mm = 25.4m;
        public const decimal lin_to_mm = 2.11667m;

        public static readonly Dictionary<int, string> Months = new Dictionary<int, string>
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


        public static readonly Dictionary<string, int> MonthLookup = new Dictionary<string, int>
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
    }
}
