using System;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter.Mssql
{
    public static class SqlConverterTools
    {
        private const string SqlTrueString = "1";
        private const string SqlFalseString = "0";

        public static string CheckedItemNameString(string name)
        {
            if (name.Equals("*",StringComparison.InvariantCultureIgnoreCase))
                return name;
            if (name.Equals("Count(*)", StringComparison.InvariantCultureIgnoreCase))
                return name;

            var stringsToMask = new[] { " ", "'", "-", "+", "*", "\"", "/", @"\", "=" };
            return stringsToMask.Any(name.Contains) ? string.Concat("\"", name, "\"") : name;
        }

        public static string DateString(DateTime date)
        {
            var s = date.ToString("'yyyy-MM-dd HH:mm:ss'");
            return s.Replace(" 00:00:00", string.Empty);
        }

        public static string CheckedSourceNameString(INamedSource source)
        {
            return string.Concat(string.IsNullOrEmpty(source.Schema) ? string.Empty : CheckedItemNameString(source.Schema) + ".", CheckedItemNameString(source.Name));
        }

        public static string GetCheckedBooleanValueString(bool value)
        {
            return value ? SqlTrueString : SqlFalseString;
        }
    }
}