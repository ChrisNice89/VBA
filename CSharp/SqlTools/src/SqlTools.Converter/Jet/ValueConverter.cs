namespace AccessCodeLib.Data.SqlTools.Converter.Jet
{
    public class ValueConverter : Converter.ValueConverter
    {
        protected override string DateTimeFormat { get { return "yyyy-MM-dd HH:mm:ss"; } }
        protected override string DateFormat { get { return "yyyy-MM-dd"; } }

        public override string GetCheckedDateTimeValueString(string dateTimeValue)
        {
            return string.Concat("#", dateTimeValue, "#");
        }

        public override string GetCheckedBooleanValueString(bool value)
        {
            return SqlConverterTools.GetCheckedBooleanValueString(value);
        }
    }
}
