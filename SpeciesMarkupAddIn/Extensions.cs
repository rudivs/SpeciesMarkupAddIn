using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace SpeciesMarkupAddIn
{
    public static class StringExtension
    {
        public static string Last(this string source, int tail_length)
        {
            if (tail_length >= source.Length)
                return source;
            return source.Substring(source.Length - tail_length);
        }

        public static string ExceptLast(this string source, int tail_length)
        {
            if (tail_length >= source.Length)
                return source;
            return source.Substring(0, source.Length - tail_length);
        }

        public static string Clean(this string source)
        {
            return new string(source.Where(c => !char.IsPunctuation(c)).ToArray());
        }

        public static string RemoveInvalidXmlChars(this string source)
        {
            return new string(source.Where(c=> XmlConvert.IsXmlChar(c)).ToArray());
        }

        public static bool IsNumber(this string source)
        {
            string numberPatern = @"(?:^)([1-9](?:\d*|(?:\d{0,2})(?:,\d{3})*)(?:\.\d*[1-9])?|0?\.\d*[1-9]|0)$";
            if (Regex.IsMatch(source,numberPatern))
            {
                return true;
            }
            else
            {
                return false;
            }
        }


    }
}
