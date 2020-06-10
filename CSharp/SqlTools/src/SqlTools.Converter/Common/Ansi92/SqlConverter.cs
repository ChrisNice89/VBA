namespace AccessCodeLib.Data.SqlTools.Converter.Common.Ansi92
{
    public class SqlConverter : Converter.SqlConverter
    {
// ReSharper disable MemberCanBeProtected.Global
        public SqlConverter()
// ReSharper restore MemberCanBeProtected.Global
            : this(new NameConverter(), new ValueConverter())
        {
        }

        protected SqlConverter(INameConverter nameConverter, IValueConverter valueConverter, IConditionConverter conditionConverter = null)
            : base(nameConverter, valueConverter, conditionConverter ?? new ConditionConverter(nameConverter, valueConverter))
        {
        }
    }
}