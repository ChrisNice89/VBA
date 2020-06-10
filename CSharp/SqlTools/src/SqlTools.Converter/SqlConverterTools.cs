using System.Globalization;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public static class SqlConverterTools
    {
        public static string GetCheckedNumericValueString(string numericValue)
        {
            return numericValue.Replace(",", ".");
        }

        public static string GetCheckedBooleanValueString(bool value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        public static string GetCheckedTextValueString(string textValue)
        {
            return string.Concat("'", textValue.Replace("'", "''"), "'");
        }

        public static string GetCheckedDateTimeValueString(string dateValue)
        {
            return string.Concat("'", dateValue, "'");
        }
    }
}