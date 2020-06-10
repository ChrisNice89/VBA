using System;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter.Common.Ansi92
{
    public static class SqlConverterTools
    {
        public static string CheckedItemNameString(string name)
        {
            return name;
        }

        public static string CheckedItemNamedSourceString(INamedSource source)
        {
            return !string.IsNullOrEmpty(source.Schema) ? string.Format("{0}.{1}", source.Schema, source.Name) : source.Name;
        }

        public static string DateString(DateTime date)
        {
            var s = date.ToString("'yyyy-MM-dd HH:mm:ss'");
            return s.Replace(" 00:00:00", string.Empty);
        }
    }
}