namespace AccessCodeLib.Data.SqlTools.Converter.Jet.Dao
{
    public class ConditionConverter : Converter.ConditionConverter
    {
        public ConditionConverter(INameConverter nameConvertor, IValueConverter valueConverter)
            : base(nameConvertor, valueConverter)
        {
        }
    }
}