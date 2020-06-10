using System;
using System.Globalization;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class ValueConverter : IValueConverter
    {
        const string DbNullAsString = "Null";

        protected virtual string WildcardString { get { return "*"; } }
        protected virtual string DateTimeFormat { get { return "yyyy-MM-dd HH:mm:ss"; } }
        protected virtual string DateFormat { get { return "yyyy-MM-dd"; } }

        public virtual string ConvertValueToString(IValue value)
        {
            //TODO: if-Folge (wenn möglich) durch schöneren Ausdruck ersetzen
            var dateTimeValue = value as IDateTimeValue;
            if (dateTimeValue != null)
                return ConvertValueToString(dateTimeValue);

            if (value.TypeOfValue == typeof(Boolean))
                return GetCheckedBooleanValueString(((IBooleanValue)value).Value);
            if (value.TypeOfValue == typeof(int))
                return GetCheckedNumericValueString(((INumericValue<int>)value).Value.ToString(CultureInfo.InvariantCulture));
            if (value.TypeOfValue == typeof(double))
                return GetCheckedNumericValueString(((INumericValue<double>)value).Value.ToString(CultureInfo.InvariantCulture));
            if (value.TypeOfValue == typeof(decimal))
                return GetCheckedNumericValueString(((INumericValue<decimal>)value).Value.ToString(CultureInfo.InvariantCulture));
            if (value.TypeOfValue == typeof(string))
                return ConvertValueToString(((ITextValue)value));
            if (value.TypeOfValue == typeof(IBetweenValue))
                return ConvertValueToString((IBetweenValue) value);
            if (value is INullValue)
                return GetNullValueString();
            if (value.TypeOfValue == typeof(DBNull))
                return GetNullValueString();
         
            throw new NotSupportedTypeException(value.TypeOfValue);
        }

        public virtual string ConvertValueToString(ITextValue value, RelationalOperators appendWildCardOperators = 0)
        {
            return GetCheckedTextValueString(AppendWildcard(value.Value, appendWildCardOperators));
        }

        protected virtual string AppendWildcard(string s, RelationalOperators wildCards)
        {
            if ((wildCards & RelationalOperators.AddWildcardPrefix) == RelationalOperators.AddWildcardPrefix)
            {
                s = string.Concat(WildcardString, s);
            }
            if ((wildCards & RelationalOperators.AddWildcardSuffix) == RelationalOperators.AddWildcardSuffix)
            {
                s = string.Concat(s, WildcardString);
            }

            return s;
        }

        protected virtual string ConvertValueToString(IBetweenValue value)
        {
            return string.Concat(ConvertValueToString(value.FirstValue), " And ", ConvertValueToString(value.SecondValue));
        }

        protected virtual string ConvertValueToString(IDateTimeValue value)
        {
            return ConvertDateTimeToString(value.Value);   
        }

        public virtual string ConvertDateTimeToString(DateTime d)
        {
            var dateFormat = d.TimeOfDay.Ticks == 0 ? DateFormat : DateTimeFormat;

            return GetCheckedDateTimeValueString(d.ToString(dateFormat));
        }

        public virtual string GetCheckedNumericValueString(string numericValue)
        {
            return SqlConverterTools.GetCheckedNumericValueString(numericValue);
        }

        public virtual string GetCheckedBooleanValueString(bool value)
        {
            return SqlConverterTools.GetCheckedBooleanValueString(value);
        }

        public virtual string GetCheckedTextValueString(string textValue)
        {
            return SqlConverterTools.GetCheckedTextValueString(textValue);
        }

        public virtual string GetCheckedDateTimeValueString(string dateTimeValue)
        {
            return dateTimeValue;
        }

        public virtual string GetNullValueString()
        {
            return DbNullAsString;
        }
    }
}
