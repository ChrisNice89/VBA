using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.Common.Sql.Converter;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class SqlConverter : ISqlConverter
    {
// ReSharper disable MemberCanBePrivate.Global
        protected const string FieldConcatString = ", ";
        protected const string DefaultConditionConcatString = " And ";

        protected const string SqlSelectStringFormat = " Select ";
        protected const string SqlFromStringFormat = " From ";
        protected const string SqlWhereStringFormat = " Where ";
        protected const string SqlGroupByStringFormat = " Group By ";
        protected const string SqlHavingByStringFormat = " Having ";
        protected const string SqlOrderByStringFormat = " Order By ";
// ReSharper restore MemberCanBePrivate.Global

// ReSharper disable MemberCanBePrivate.Global
        protected readonly INameConverter _nameConverter;
        protected readonly IValueConverter _valueConverter;
        protected readonly IConditionConverter _conditionConverter;
// ReSharper restore MemberCanBePrivate.Global

        public SqlConverter(INameConverter newNameConverter, IValueConverter newValueConverter, IConditionConverter newConditionConverter)
        {
            _nameConverter = newNameConverter;
            _valueConverter = newValueConverter;
            _conditionConverter = newConditionConverter;
        }

        protected virtual INameConverter NameConverter
        {
            get { return _nameConverter; }
        }

        protected virtual IValueConverter ValueConverter
        {
            get { return _valueConverter; }
        }

        protected virtual IConditionConverter ConditionConverter
        {
            get { return _conditionConverter; }
        }

        public string GenerateSqlString(ISqlStatement sqlStatement)
        {
            if (sqlStatement == null)
                return null;

            var sql = GenerateFrom(sqlStatement.Find(StatementKeys.From));

            AddSelect(ref sql, sqlStatement.Find(StatementKeys.Select));
            AddWhere(ref sql, sqlStatement.Find(StatementKeys.Where));
            AddGroupBy(ref sql, sqlStatement.Find(StatementKeys.GroupBy));
            AddHaving(ref sql, sqlStatement.Find(StatementKeys.Having));
            AddOrderBy(ref sql, sqlStatement.Find(StatementKeys.OrderBy));

            return sql;
        }

// ReSharper disable MemberCanBePrivate.Global
        protected string GenerateFrom(IEnumerable<IStatement> statements)
// ReSharper restore MemberCanBePrivate.Global
        {
            if (statements == null || !statements.Any())
                return string.Empty;

            var fromString = string.Empty;
            foreach (var statement in statements.OfType<IFromStatement>())
            {
                if (statement is IJoinStatement)
                    ConcatJoinStatment(ref fromString, (IJoinStatement) statement);
                else
                    fromString = string.Concat(fromString, FieldConcatString, GenerateSourceString(statement));
            }
            return string.IsNullOrEmpty(fromString) ? string.Empty : string.Concat(SqlFromStringFormat, fromString.Substring(FieldConcatString.Length)).Trim();
        }

        protected virtual void ConcatJoinStatment(ref string fromString, IJoinStatement statement)
        {
            var lastIndex = fromString.LastIndexOf(',');
            var testString = fromString.Substring(lastIndex + 1).Trim();
            if (testString.Contains("Join"))
            {
                fromString = fromString.Substring(0, lastIndex + 2) + "(" + testString + ")";
            }
            fromString += GenerateSourceString(statement);
        }

        protected virtual string GenerateSourceString(IFromStatement statement)
        {
            return statement is IJoinStatement
                       ? GenerateSourceString((IJoinStatement) statement)
                       : GenerateSourceString(statement.Source);
        }

        protected virtual string GenerateSourceString(IJoinStatement statement)
        {
            return string.Concat(GetJoinOperatorString(statement.JoinOperator), GenerateSourceString(statement.Source),
                                 " On ", ConditionConverter.GenerateSqlString((IConditionGroup)statement.Condition));
            //return string.Concat("(", GenerateSqlString(subSelect.SqlStatement), ")",
            //                     subSelect is IAlias ? " As " + NameConvertor.CheckedItemNameString(((IAlias)subSelect).Alias) : string.Empty);
        }

        protected static string GetJoinOperatorString(JoinOperator op)
        {
            switch (op)
            {
                case  JoinOperator.Inner:
                    return " Inner Join ";
                case JoinOperator.Left:
                    return " Left Join ";
                case JoinOperator.Right:
                    return " Right Join "; 
            }
            throw new NotSupportedException(op + " is not supported.");
        }

// ReSharper disable VirtualMemberNeverOverriden.Global
        protected virtual string GenerateSourceString(ISource source)
        {
            if (source is ISourceAlias && (((ISourceAlias)source).Source) is ISubSelect)
                return GenerateSubSelectSourceString((ISubSelect)((ISourceAlias)source).Source, (IAlias)source);
            return source is ISubSelect
                       ? GenerateSubSelectSourceString((ISubSelect) source)
                       : NameConverter.GenerateSourceNameString(source);
        }

        protected virtual string GenerateSubSelectSourceString(ISubSelect subSelect, IAlias alias = null)
        {
            return string.Concat("(", GenerateSqlString(subSelect.SqlStatement), ")",
                                 (alias != null) ? " As " + NameConverter.GenerateAliasNameString(alias) : string.Empty);
            //return string.Concat("(", GenerateSqlString(subSelect.SqlStatement), ")",
            //                     subSelect is IAlias ? " As " + NameConvertor.CheckedItemNameString(((IAlias)subSelect).Alias) : string.Empty);
        }

        protected virtual void AddSelect(ref string sql, IEnumerable<IStatement> statements)
        {
            if (statements == null || !statements.Any())
                return;

            var fields = AggregateFields(statements);

            var fieldString = AggregateSelectFieldString(fields);
            sql = string.Concat(SqlSelectStringFormat, fieldString, " ", sql).Trim();
        }

        private static IEnumerable<IField> AggregateFields(IEnumerable<IStatement> statements)
        {
            var fields = new List<IField>();
            foreach (var fieldsStatement in statements.OfType<IFieldsStatement>())
            {
                fields.AddRange(fieldsStatement.Fields);
            }
            return fields;
        }

        protected virtual void AddWhere(ref string sql, IEnumerable<IStatement> statements)
        {
            if (statements == null || !statements.Any())
                return;

            var conditionString = AggregateConditionString(statements);

            if (!string.IsNullOrEmpty(conditionString))
                sql = string.Concat(sql, SqlWhereStringFormat, conditionString).Trim();
        }

        protected virtual string AggregateConditionString(IEnumerable<IStatement> statements)
        {
            var conditionString = GenerateConditionString(statements.OfType<IConditionGroup>());

            var stringStatementsString = GenerateConditionString(statements.OfType<IConditionStringStatement>());

            if (String.IsNullOrEmpty(conditionString))
            {
                conditionString = stringStatementsString;
            }
            else if (!String.IsNullOrEmpty(stringStatementsString))
            {
                conditionString = string.Concat(conditionString, DefaultConditionConcatString, stringStatementsString);
            }

            return string.IsNullOrEmpty(conditionString) ? string.Empty : conditionString;
        }

        protected string GenerateConditionString(IEnumerable<IConditionGroup> conditionGroups)
        {
            var conditionString = string.Empty;

            if (conditionGroups == null)
                return string.Empty;

            if (conditionGroups.Any())
            {
                conditionString = ConditionConverter.GenerateSqlString(conditionGroups);
            }

            return string.IsNullOrEmpty(conditionString) ? string.Empty : conditionString;
        }

        public string GenerateConditionString(IConditionGroup conditionGroup)
        {
            if (conditionGroup == null)
                return string.Empty;

            var conditionString = ConditionConverter.GenerateSqlString(conditionGroup);

            return string.IsNullOrEmpty(conditionString) ? string.Empty : conditionString;
        }

        private string GenerateConditionString(IEnumerable<IConditionStringStatement> stringStatements)
        {
            var conditionString = string.Empty;

            if (stringStatements.Any())
            {
                conditionString = stringStatements.Aggregate(conditionString, (current, stringStatement) => string.Concat(current, DefaultConditionConcatString, "(", ConditionConverter.GenerateSqlString(stringStatement.Condition), ")"));
                conditionString = conditionString.Substring(DefaultConditionConcatString.Length);
            }

            return string.IsNullOrEmpty(conditionString) ? string.Empty : conditionString;
        }

        protected virtual void AddGroupBy(ref string sql, IEnumerable<IStatement> statements)
        {
            if (statements == null || !statements.Any())
                return;

            var fields = AggregateFields(statements);

            var fieldString = AggregateFieldString(fields);
            sql = string.Concat(sql, SqlGroupByStringFormat, fieldString).Trim();
        }

        protected virtual void AddHaving(ref string sql, IEnumerable<IStatement> statements)
        {
            if (statements == null || !statements.Any())
                return;

            var conditionString = AggregateConditionString(statements);

            if (!string.IsNullOrEmpty(conditionString))
                sql = string.Concat(sql, SqlHavingByStringFormat, conditionString).Trim();
        }

        protected virtual void AddOrderBy(ref string sql, IEnumerable<IStatement> statements)
        {
            if (statements == null || !statements.Any())
                return;

            var fields = AggregateFields(statements);

            var fieldString = AggregateFieldString(fields);
            sql = string.Concat(sql, SqlOrderByStringFormat, fieldString).Trim();
        }

        protected virtual string AggregateFieldString(IEnumerable<IField> fields)
        {
            return fields.Aggregate("", (current, s) => current + (FieldConcatString + NameConverter.GenerateFieldString(s))).Substring(FieldConcatString.Length);
        }

        protected virtual string AggregateSelectFieldString(IEnumerable<IField> fields)
        {
            var fieldString = fields.Where(f => !string.IsNullOrEmpty(f.Name)).Aggregate(string.Empty, (current, f) => string.Concat(current, FieldConcatString, NameConverter.GenerateSelectFieldString(f)));
            return fieldString.Length > FieldConcatString.Length ? fieldString.Substring(FieldConcatString.Length) : string.Empty;
        }
// ReSharper restore VirtualMemberNeverOverriden.Global
    }
}
