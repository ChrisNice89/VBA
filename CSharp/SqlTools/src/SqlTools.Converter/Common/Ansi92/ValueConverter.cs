namespace AccessCodeLib.Data.SqlTools.Converter.Common.Ansi92
{
    public class ValueConverter : Converter.ValueConverter
    {
        protected override string WildcardString { get { return "%"; } }
        protected override string DateTimeFormat { get { return "yyyy-MM-dd HH:mm:ss"; } }
        protected override string DateFormat { get { return "yyyy-MM-dd"; } }

        public override string GetCheckedDateTimeValueString(string dateTimeValue)
        {
            return string.Concat("'", dateTimeValue, "'");
        }
    }
}
