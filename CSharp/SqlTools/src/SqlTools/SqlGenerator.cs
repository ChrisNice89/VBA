using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.Common.Sql.Converter;
using AccessCodeLib.Data.SqlTools.Sql;

namespace AccessCodeLib.Data.SqlTools
{
    public class SqlGenerator : ISqlGenerator
    {
        private readonly ISqlStatement _sqlStatement;
        public virtual ISqlStatement SqlStatement { get { return _sqlStatement; } }

        public SqlGenerator()
        {
            _sqlStatement = new SqlStatement();
        }

        protected SqlGenerator(ISqlStatement sqlStatement)
        {
            _sqlStatement = sqlStatement;
        }

        public SqlGenerator(ISqlConverter sqlConverter) : this()
        {
            Converter = sqlConverter;
        }

        protected SqlGenerator(ISqlConverter sqlConverter, ISqlStatement sqlStatement)
            : this(sqlConverter)
        {
            _sqlStatement = sqlStatement;
        }

        public ISqlConverter Converter { get; set; }

        public ISqlGenerator Select(params IField[] fields)
        {
            _sqlStatement.Add(new SelectStatement(fields));
            return this;
        }

        public ISqlGenerator Select(params string[] fieldNames)
        {
            _sqlStatement.Add(new SelectStatement(fieldNames));
            return this;
        }

        public ISqlGenerator SelectAll()
        {
            _sqlStatement.Add(new SelectStatement("*"));
            return this;
        }

        public ISqlGenerator SelectField(string fieldName, ISource source, string alias)
        {
            _sqlStatement.Add(string.IsNullOrEmpty(alias)
                             ? new SelectStatement(new Field(fieldName, source))
                             : new SelectStatement(new FieldAlias(fieldName, source, alias)));
            return this;
        }

// ReSharper disable UnusedMethodReturnValue.Global
        protected ISqlGenerator SelectField(string fieldName, object source = null, string alias = "")
// ReSharper restore UnusedMethodReturnValue.Global
        {
            ISource fieldSource = null;
            if (source is ISource)
                fieldSource = (ISource)source;
            else if (source is string)
                fieldSource = new NamedSource((string)source);

            return SelectField(fieldName, fieldSource, alias);
        }

        public ISqlGenerator From(string source)
        {
            _sqlStatement.Add(new FromStatement(source));
            return this;
        }

        public ISqlGenerator From(ISource source)
        {
            _sqlStatement.Add(new FromStatement(source));
            return this;
        }

        public ISqlGenerator Join(ISource source, ICondition condition, JoinOperator op = JoinOperator.Inner)
        {
            _sqlStatement.Add(new JoinStatement(source, condition, op));
            return this;
        }

        public ISqlGenerator Where(ICondition condition)
        {
            _sqlStatement.Add(new WhereStatement(condition));
            return this;
        }

        public ISqlGenerator Where(IField field, RelationalOperators relationalOperator, object value)
        {
            return Where(new FieldCondition(field, relationalOperator, value));
        }

        public ISqlGenerator Where(string whereString)
        {
            _sqlStatement.Add(new WhereStatement(whereString));
            return this;
        }

        public ISqlGenerator GroupBy(params IField[] fields)
        {
            _sqlStatement.Add(new GroupByStatement(fields));
            return this;
        }

        public ISqlGenerator GroupBy(params string[] fieldNames)
        {
            _sqlStatement.Add(new GroupByStatement(fieldNames));
            return this;
        }

        public ISqlGenerator Having(ICondition condition)
        {
            _sqlStatement.Add(new HavingStatement(condition));
            return this;
        }

        public ISqlGenerator Having(IField field, RelationalOperators relationalOperator, object value)
        {
            return Having(new FieldCondition(field, relationalOperator, value));
        }

        public ISqlGenerator Having(string havingString)
        {
            _sqlStatement.Add(new HavingStatement(havingString));
            return this;
        }

        public ISqlGenerator OrderBy(params IField[] fields)
        {
            _sqlStatement.Add(new OrderByStatement(fields));
            return this;
        }

        public ISqlGenerator OrderBy(params string[] fieldNames)
        {
            _sqlStatement.Add(new OrderByStatement(fieldNames));
            return this;
        }

        public new string ToString()
        {
            return Converter.GenerateSqlString(_sqlStatement);
        }
    }
}
