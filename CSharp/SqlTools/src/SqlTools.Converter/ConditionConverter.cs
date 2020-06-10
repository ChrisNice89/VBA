using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;
using System.Collections;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class ConditionConverter : IConditionConverter
    {
        public ConditionConverter(INameConverter nameConverter, IValueConverter valueConverter)
        {
            NameConverter = nameConverter;
            ValueConverter = valueConverter;
        }

// ReSharper disable MemberCanBePrivate.Global
        protected INameConverter NameConverter { get; private set; }
        protected IValueConverter ValueConverter { get; private set; }

        protected string ConditionStringItemPrefix = "";
        protected string ConditionStringItemPostfix = "";

// ReSharper restore MemberCanBePrivate.Global

        public virtual string GenerateSqlString(IEnumerable<IConditionGroup> conditionGroups, LogicalOperator topLevelConcatOperator = LogicalOperator.And)
        {
            var conditionString = string.Empty;
            foreach (var conditionGroup in conditionGroups)
            {
                ConcatConditionGroup(ref conditionString, topLevelConcatOperator, conditionGroup, string.Empty, string.Empty);
            }

            return string.IsNullOrEmpty(conditionString) ? null : GetCheckedConditionString(conditionString);
        }

        public virtual string GenerateSqlString(IConditionGroup conditionGroup)
        {
            var conditionString = string.Empty;
            ConcatConditionGroup(ref conditionString, LogicalOperator.And, conditionGroup, string.Empty, string.Empty);
            return string.IsNullOrEmpty(conditionString) ? null : GetCheckedConditionString(conditionString);
        }

        public virtual string GenerateSqlString(string conditionString)
        {
            return GetCheckedConditionString(conditionString);
        }

        protected virtual string GetCheckedConditionString(string condition)
        {
            return condition;
        }

        private void ConcatConditionGroup(ref string conditionString, LogicalOperator concatOperator, 
                                          IConditionGroup conditionGroup, string groupPrefix = "(", string groupPostfix = ")")
        {
            var groupString = string.Empty;
            foreach (var condition in conditionGroup.Conditions)
            {
                if (condition is IConditionGroup)
                {
                    ConcatConditionGroup(ref groupString, conditionGroup.ConcatOperator, (IConditionGroup)condition);
                }
                else if (condition is IFieldCondition)
                {
                    ConcatFieldCondition(ref groupString, conditionGroup.ConcatOperator, (IFieldCondition)condition);
                }
                else if (condition is IConditionString)
                {
                    ConcatConditionString(ref groupString, conditionGroup.ConcatOperator, (IConditionString) condition);
                }
            }

            if (string.IsNullOrEmpty(groupString))
                return;

            conditionString = string.Concat(conditionString, GetOperatorStringIfStringIsNotEmpty(conditionString, concatOperator), groupPrefix, groupString, groupPostfix);
        }

        private void ConcatFieldCondition(ref string conditionString, LogicalOperator concatOperator, IFieldCondition condition)
        {
            conditionString = string.Concat(conditionString, GetOperatorStringIfStringIsNotEmpty(conditionString, concatOperator), CreateConditionString(condition));
        }

        private void ConcatConditionString(ref string conditionString, LogicalOperator concatOperator, IConditionString condition)
        {
            conditionString = string.Concat(conditionString, GetOperatorStringIfStringIsNotEmpty(conditionString, concatOperator), "(", condition.Value, ")");
        }

        private string GetOperatorStringIfStringIsNotEmpty(string conditionString, LogicalOperator op)
        {
            return conditionString.Length > 0 ? OperatorString(op) : string.Empty;
        }

// ReSharper disable VirtualMemberNeverOverriden.Global
        protected virtual string CreateConditionString(IFieldCondition condition)
        {
            return CreateConditionString(condition.Field, condition.Operator, condition.Value);
        }

        protected virtual string CreateConditionString(IField field, RelationalOperators op, object value)
        {
            if ((op & RelationalOperators.In) == RelationalOperators.In)
                return CreateInConditionString(field, op, value);

            if ((value is IBetweenValue) && (op & RelationalOperators.Between) == RelationalOperators.Between)
                return CreateBetweenConditionString(field, op, (IBetweenValue)value);

            var valueArray = value as IValueArray;
            if (valueArray != null)
                return CreateValueArrayConditionString(field, op, valueArray);
                
            if (value is IEnumerable && !(value is string) )
                return CreateEnumerableConditionString(field, op, value as IEnumerable);

            var fieldName = NameConverter.GenerateFieldString(field);

            var f = value as IField;
            if (f != null)
                return CreateFieldConditionString(fieldName, op, f);
            var v = value as IValue;
            if (v != null)
                return CreateFieldConditionString(fieldName, op, v);
            var s = value as string;
            if (s != null)
                return CreateStringConditionString(fieldName, op, s);
            if (value is DateTime)
                return CreateDateTimeConditionString(fieldName, op, (DateTime)value);

            double number;
            return Double.TryParse(value.ToString(), out number) ? CreateNumericConditionString(fieldName, op, value) : string.Concat(ConditionStringItemPrefix, NameConverter.GenerateFieldString(field), OperatorString(op), value.ToString(), ConditionStringItemPostfix);
        }

        protected virtual string CreateValueArrayConditionString(IField field, RelationalOperators op, IValueArray value)
        {
            return CreateValuesArrayConditionString(field, op, value.Values);
        }

        protected virtual string CreateEnumerableConditionString(IField field, RelationalOperators op, IEnumerable values)
        {
            return string.Concat("(", string.Join(" Or ", (from object v in values select CreateConditionString(field, op, v)).ToArray()), ")");
        }

        protected virtual string CreateValuesArrayConditionString(IField field, RelationalOperators op, IEnumerable<IValue> values)
        {
            return CreateEnumerableConditionString(field, op, values);
        }

        protected virtual string CreateInConditionString(IField field, RelationalOperators op, object value)
        {
            var valueArray = value as IValueArray;
            if (valueArray != null)
                return CreateInConditionString(field, op, valueArray);

            var v = value as IValue;
            if (v != null)
                return CreateInConditionString(field, op, v);

            throw new NotSupportedException();
        }

        protected virtual string CreateInConditionString(IField field, RelationalOperators op, IValueArray valueArray)
        {
            return CreateInConditionString(field, op, valueArray.Values);
        }

        protected virtual string CreateInConditionString(IField field, RelationalOperators op, IValue value)
        {
            return CreateInConditionString(field, op, new[] {value});
        }

        protected virtual string CreateInConditionString(IField field, RelationalOperators op, IValue[] values)
        {
            var valueStrings = new string[values.Length]; 
            
            for (var i = 0; i < values.Length; i++)
            {
                valueStrings[i] = ValueConverter.ConvertValueToString(values[i]);
            }

            return string.Concat(ConditionStringItemPrefix, NameConverter.GenerateFieldString(field), OperatorString(op), "(", string.Join(",", valueStrings), ")", ConditionStringItemPostfix);
        }

        protected virtual string CreateNumericConditionString(string fieldName, RelationalOperators op, object value)
        {
            return string.Concat(ConditionStringItemPrefix, fieldName, OperatorString(op), ValueConverter.GetCheckedNumericValueString(value.ToString()), ConditionStringItemPostfix);
        }

        protected virtual string CreateStringConditionString(string fieldName, RelationalOperators op, string value)
        {
            return string.Concat(ConditionStringItemPrefix, fieldName, OperatorString(op), ValueConverter.GetCheckedTextValueString(value), ConditionStringItemPostfix);
        }

        protected virtual string CreateDateTimeConditionString(string fieldName, RelationalOperators op, DateTime value)
        {
            return string.Concat(ConditionStringItemPrefix, fieldName, OperatorString(op), ValueConverter.ConvertDateTimeToString(value), ConditionStringItemPostfix);
        }

        protected virtual string CreateFieldConditionString(string fieldName, RelationalOperators op, IField value)
        {
            return string.Concat(ConditionStringItemPrefix, fieldName, OperatorString(op), NameConverter.GenerateFieldString(value), ConditionStringItemPostfix);
        }

        protected virtual string CreateFieldConditionString(string fieldName, RelationalOperators op, IValue value)
        {
            RelationalOperators valueOp = 0;

            if ((op & RelationalOperators.AddWildcardPrefix) == RelationalOperators.AddWildcardPrefix)
            {
                op = op ^ RelationalOperators.AddWildcardPrefix;
                valueOp = valueOp | RelationalOperators.AddWildcardPrefix;
            }
            if ((op & RelationalOperators.AddWildcardSuffix) == RelationalOperators.AddWildcardSuffix)
            {
                op = op ^ RelationalOperators.AddWildcardSuffix;
                valueOp = valueOp | RelationalOperators.AddWildcardSuffix;
            }

            var nullValue = value as INullValue;
            return nullValue != null ? CreateFieldConditionString(fieldName, op, nullValue) : string.Concat(ConditionStringItemPrefix, fieldName, OperatorString(op), ((valueOp != 0) && (value is ITextValue)) ? ValueConverter.ConvertValueToString((ITextValue)value, valueOp) : ValueConverter.ConvertValueToString(value), ConditionStringItemPostfix);
        }

        protected virtual string CreateFieldConditionString(string fieldName, RelationalOperators op, INullValue value)
        {

            var cond = string.Empty;

            if ((op & RelationalOperators.Equal) == RelationalOperators.Equal)
            {
                cond = string.Concat(fieldName, " Is Null");
                op = op ^ RelationalOperators.Equal;
                if (op == 0) return string.Concat(ConditionStringItemPrefix, cond, ConditionStringItemPostfix);
            }

            var cond2 = string.Concat(fieldName, OperatorString(op), ValueConverter.ConvertValueToString(value));
            return string.IsNullOrEmpty(cond2) ? cond : string.Concat("(", cond, " Or ", cond2, ")");
        }

        protected virtual string CreateBetweenConditionString(IField field, RelationalOperators op, IBetweenValue value)
        {
            if (op == RelationalOperators.Between)
            {
                var fieldName = NameConverter.GenerateFieldString(field);
                return CreateBetweenConditionString(fieldName, value);
            }
                
            op = (op ^ RelationalOperators.Between);

            if (op == RelationalOperators.AddWildcardSuffix)
            {
                string condString = string.Empty;

                if (!(value.FirstValue is INullValue))
                {
                    condString = CreateConditionString(field, RelationalOperators.Equal | RelationalOperators.GreaterThan, value.FirstValue); 
                }

                if (value.SecondValue is INullValue)
                    return condString;

                if (field.DataType == FieldDataType.DateTime)
                {
                    DateTime d;
                    if (value.SecondValue is IDateTimeValue)
                        d = ((IDateTimeValue)value.SecondValue).Value;
                    else
                        throw new InvalidCastException("IDateTimeValue required");

                    condString = string.Concat("(", condString, " And ", CreateConditionString(field, RelationalOperators.LessThan, d.Date.AddDays(1)), ")");
                    return condString;
                }

                if (field.DataType == FieldDataType.Text)
                {
                    string s;
                    if (value.SecondValue is ITextValue)
                        s = ((ITextValue)value.SecondValue).Value;
                    else
                        throw new InvalidCastException("ITextValue required");

                    if (!string.IsNullOrEmpty(s))
                        s = s.Substring(0, s.Length - 1) + (char)(s[s.Length-1]+1);

                    condString = string.Concat("(", condString, " And ", CreateConditionString(field, RelationalOperators.LessThan, s), ")");
                    return condString;
                }
            }

            throw new NotSupportedRelationalOperatorException(op);
        }


        protected virtual string CreateBetweenConditionString(string fieldName, IBetweenValue value)
        {
            if (!(value.FirstValue is INullValue || value.SecondValue is INullValue))
            {
                return string.Concat("(", fieldName, OperatorString(RelationalOperators.Between), ValueConverter.ConvertValueToString(value), ")");
            }

            if (value.FirstValue is INullValue && value.SecondValue is INullValue)
            {
                return null;
            }

            string valueString = null;
            string operatorString = null;
            if (value.FirstValue is INullValue)
            {
                valueString = ValueConverter.ConvertValueToString(value.SecondValue);
                operatorString = OperatorString(RelationalOperators.Equal | RelationalOperators.LessThan);
            }
            else
            {
                valueString = ValueConverter.ConvertValueToString(value.FirstValue);
                operatorString = OperatorString(RelationalOperators.Equal | RelationalOperators.GreaterThan);
            }

            return string.Concat(ConditionStringItemPrefix, fieldName, operatorString, valueString, ConditionStringItemPostfix);
        }

        protected virtual string OperatorString(RelationalOperators op)
        {
            switch (op)
            {
                case RelationalOperators.Like:
                    return " Like ";
                case RelationalOperators.Between:
                    return " Between ";
                case RelationalOperators.In:
                    return " In ";
            }

            var opString = string.Empty;
            if ((op & RelationalOperators.LessThan) == RelationalOperators.LessThan)
            {
                opString += "<";
                op -= RelationalOperators.LessThan;
            }

            if ((op & RelationalOperators.GreaterThan) == RelationalOperators.GreaterThan)
            {
                opString += ">";
                op -= RelationalOperators.GreaterThan;
            }

            if ((op & RelationalOperators.Equal) == RelationalOperators.Equal)
            {
                opString += "=";
                op -= RelationalOperators.Equal;
            }

            if (op != 0)
                throw new NotSupportedRelationalOperatorException(string.Concat(string.IsNullOrEmpty(opString) ? string.Empty : string.Concat("'", opString, "'", " and "), op, " is not supported."), op);

            return string.Format(" {0} ", opString);
        }

        protected virtual string OperatorString(LogicalOperator op)
        {
            switch (op)
            {
                case LogicalOperator.And:
                    return " And ";
                case LogicalOperator.Or:
                    return " Or ";
            }
            throw new NotSupportedException("ConcatOperator is not supported.");
        }
// ReSharper restore VirtualMemberNeverOverriden.Global
    }
}