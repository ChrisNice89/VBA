namespace AccessCodeLib.Data.SqlTools.Converter.Jet.Dao
{
    public class SqlConverter : Converter.SqlConverter
    {
        public SqlConverter()
            : this(new NameConverter(), new ValueConverter())
        {
        }

// ReSharper disable MemberCanBePrivate.Global
        protected SqlConverter(INameConverter nameConverter, IValueConverter valueConverter, IConditionConverter conditionConverter = null)
// ReSharper restore MemberCanBePrivate.Global
            : base(nameConverter, valueConverter, conditionConverter ?? new ConditionConverter(nameConverter, valueConverter))
        {
        }
    }
}