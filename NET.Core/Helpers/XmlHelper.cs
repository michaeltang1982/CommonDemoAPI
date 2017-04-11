using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Sierra.NET.Core
{
    public static class XmlHelper
    {
        /// <summary>
        /// find a given attribute value in the provided xml string
        /// NOTE: this function is case-insensitive
        /// </summary>
        /// <returns>attribute value, or NULL</returns>
        public static string ExtractAttributeFromXml(string xml, string attribute)
        {
            string pattern = "[ ]" + attribute + "[ ]*=[ ]*\"([^\"]+)";
            MatchCollection matches = Regex.Matches(xml, pattern, RegexOptions.IgnoreCase);
            if (matches.Count == 0)
                return null;
            else if (matches.Count == 1)
                return matches[0].Groups[1].Value;
            else
                throw new Exception(string.Format("More than one attribute '{0}' found in the given xml", attribute));
        }
    }
}
