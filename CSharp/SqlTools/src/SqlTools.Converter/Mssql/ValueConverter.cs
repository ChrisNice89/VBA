namespace AccessCodeLib.Data.SqlTools.Converter.Mssql
{
    public class ValueConverter : Common.Ansi92.ValueConverter
    {
        protected override string DateTimeFormat { get { return "yyyyMMdd HHmmss"; } }
        protected override string DateFormat { get { return "yyyyMMdd"; } }

        public override string GetCheckedDateTimeValueString(string dateTimeValue)
        {
            return string.Concat("'", dateTimeValue, "'");
        }

        public override string GetCheckedBooleanValueString(bool value)
        {
            return SqlConverterTools.GetCheckedBooleanValueString(value);
        }

    }
}
