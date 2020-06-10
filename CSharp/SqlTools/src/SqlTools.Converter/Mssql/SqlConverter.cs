namespace AccessCodeLib.Data.SqlTools.Converter.Mssql
{
    public class SqlConverter : Common.Ansi92.SqlConverter
    {
        public SqlConverter()
            : this(new NameConverter(), new ValueConverter())
        {
        }

// ReSharper disable MemberCanBePrivate.Global
        protected SqlConverter(INameConverter nameConverter, IValueConverter valueConverter, IConditionConverter conditionConverter = null)
// ReSharper restore MemberCanBePrivate.Global
            : base(nameConverter, valueConverter, conditionConverter)
        {
        }
    }
}