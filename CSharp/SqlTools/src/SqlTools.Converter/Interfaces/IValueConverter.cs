using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public interface IValueConverter
    {
        string ConvertValueToString(IValue value);
        string ConvertValueToString(ITextValue value, RelationalOperators appendWildCardOperators = 0);
        string ConvertDateTimeToString(System.DateTime dateTime);
        string GetCheckedNumericValueString(string numericValue);
        string GetCheckedTextValueString(string textValue);
        string GetCheckedDateTimeValueString(string dateTimeValue);
    }
}