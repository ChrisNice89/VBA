using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class ConditionGroup : IConditionGroup
    {
        protected readonly List<ICondition> EmbeddedStatements = new List<ICondition>();

        public ConditionGroup(LogicalOperator concatOperator = LogicalOperator.And)
        {
            ConcatOperator = concatOperator;
        }

        public ConditionGroup(IEnumerable<ICondition> conditions, LogicalOperator concatOperator = LogicalOperator.And)
        {
            ConcatOperator = concatOperator;
            Add(conditions);
        }

        public LogicalOperator ConcatOperator { get; set; }

        public IList<ICondition> Conditions { get { return EmbeddedStatements; } }

        public IConditionGroup Add(ICondition condition)
        {
            EmbeddedStatements.Add(condition);
            return this;
        }

        public IConditionGroup Add(IEnumerable<ICondition> conditions)
        {
            EmbeddedStatements.AddRange(conditions);
            return this;
        }

        public IConditionGroup Add(IField field, RelationalOperators relationalOperator, object value, object ignoreValue = null)
        {

            var enumerableValue = value as IEnumerable;

            if ((relationalOperator & RelationalOperators.Between) == RelationalOperators.Between)
            {
                var bv = value as IBetweenValue;
                if (bv != null)
                    return AddBetweenCondition(field, relationalOperator, bv, ignoreValue);

                if (enumerableValue == null)
                    throw new ArrayTypeMismatchException("RelationalOperators.Between requires 2 values");
                return AddBetweenCondition(field, relationalOperator, enumerableValue, ignoreValue);
            }

            if ((relationalOperator & RelationalOperators.In) == RelationalOperators.In)
            {
                return AddInCondition(field, relationalOperator, value, ignoreValue);
            }

            if (enumerableValue != null && !(value is string))
            {
                return AddEnumerableValueCondition(field, relationalOperator, enumerableValue, ignoreValue);
            }



            /*if (!(value is IValue))
            {
                var enumerableValue = value as IEnumerable;

                if ((relationalOperator & RelationalOperators.Between) == RelationalOperators.Between)
                {
                    if (enumerableValue == null)
                        throw new ArrayTypeMismatchException("RelationalOperators.Between requires 2 values");
                    return AddBetweenCondition(field, relationalOperator, enumerableValue, ignoreValue);
                }

                if ((relationalOperator & RelationalOperators.In) == RelationalOperators.In)
                {
                    return AddInCondition(field, relationalOperator, value, ignoreValue);
                }

                if (enumerableValue != null && !(value is string))
                {
                    return AddEnumerableValueCondition(field, relationalOperator, enumerableValue, ignoreValue);
                }
            }
             * */

            if (!ContainsIgnoreValue(value, ignoreValue))
            {
                EmbeddedStatements.Add(new FieldCondition(field, relationalOperator, GetDataTypeCheckedValue(value, field.DataType)));
            }
            return this;
        }

        private IConditionGroup AddEnumerableValueCondition(IField field, RelationalOperators relationalOperator, IEnumerable values, object ignoreValue = null)
        {
            var conditionGroup = new ConditionGroup(LogicalOperator.Or);
           
            foreach(var value in values)
            {
                conditionGroup.Add(field,relationalOperator,value, ignoreValue);
            }

            Add(conditionGroup);
            return this;
        }

        private IConditionGroup AddBetweenCondition(IField field, RelationalOperators relationalOperator, IBetweenValue values, object ignoreValue)
        {
            
            if (ContainsIgnoreValue(values.FirstValue, ignoreValue) && ContainsIgnoreValue(values.SecondValue, ignoreValue))
            {
                return this;
            }

            Add(new FieldCondition(field, relationalOperator, values));
            return this;
        }

        private IConditionGroup AddBetweenCondition(IField field, RelationalOperators relationalOperator, IEnumerable values, object ignoreValue)
        {
            var i = -1;
            var betweenValues = new object[2];
            foreach (var value in values)
            {
                i++;
                betweenValues[i] = value;
                if (i == 1) break;
            }
            if (i < 1) throw new ArrayTypeMismatchException("RelationalOperators.Between requires 2 values");


            if (ContainsIgnoreValue(betweenValues[0], ignoreValue) && ContainsIgnoreValue(betweenValues[1], ignoreValue))
            {
                return this;
            }

            Add(new FieldCondition(field, relationalOperator, CreateBetweenValue(field.DataType, betweenValues[0], betweenValues[1])));
            return this;
        }

        private IBetweenValue CreateBetweenValue(FieldDataType fieldDataType, object value1, object value2)
        {
            if (fieldDataType == FieldDataType._Unspecified)
            {
                return CreateBetweenValueFromValueTypes(value1, value2);
            }

            var fistValue = ConvertToIValue(fieldDataType, value1);
            var secondValue = ConvertToIValue(fieldDataType, value2);

            return new BetweenValue(fistValue, secondValue);
        }

        private IBetweenValue CreateBetweenValueFromValueTypes(object value1, object value2)
        {
            IValue firstValue;
            IValue secondValue;

            if (DBNull.Value.Equals(value1) || value1 == null)
            {
                firstValue = new NullValue();
                secondValue = GetIValueFromValueType(value2);
                return new BetweenValue(firstValue, secondValue);
            }

            if (DBNull.Value.Equals(value2) || value2 == null)
            {
                secondValue = new NullValue();
                firstValue = GetIValueFromValueType(value1);
                return new BetweenValue(firstValue, secondValue);
            }

            double v1;
            double v2;
            if (double.TryParse(value1.ToString(), out v1) && double.TryParse(value2.ToString(), out v2))
            {
                return new BetweenValue(new NumericValue<double>(v1), new NumericValue<double>(v2));
            }

            DateTime dat1;
            DateTime dat2;
            if (DateTime.TryParse(value1.ToString(), out dat1) && DateTime.TryParse(value2.ToString(), out dat2))
            {
                return new BetweenValue(new DateTimeValue(dat1), new DateTimeValue(dat2));
            }

            return new BetweenValue(new TextValue((string)value1), new TextValue((string)value2));
        }

        private IConditionGroup AddInCondition(IField field, RelationalOperators relationalOperator, object value, object ignoreValue)
        {

            var valueArray = value as IValueArray;
            if (valueArray != null)
                return AddInCondition(field, relationalOperator, valueArray, ignoreValue);

            var enumerable = value as IEnumerable;
            if (enumerable == null)
            {
                var v = TryCreateDataTypeConformValue(value, field.DataType);

                if (v == null)
                    throw new InvalidCastException();

                var values = new ValueArray(new[] { v });

                return AddInCondition(field, relationalOperator, values, ignoreValue);
            }

            var valueList = new List<IValue>();
   
            foreach (var iv in from object v in enumerable select TryCreateDataTypeConformValue(v, field.DataType))
            {
                if (iv == null)
                    throw new InvalidCastException();

                valueList.Add(iv);
            }
            return AddInCondition(field, relationalOperator, new ValueArray(valueList.ToArray()), ignoreValue);
        }

        private IConditionGroup AddInCondition(IField field, RelationalOperators relationalOperator, IValueArray values, object ignoreValue)
        {
            var inValues = values.Values.Where(v => !ContainsIgnoreValue(v, ignoreValue)).ToList();

            if (inValues.Count == 0)
            {
                return this;
            }

            Add(new FieldCondition(field, relationalOperator, new ValueArray(inValues.ToArray())));
            return this;
        }

        private static IValue GetIValueFromValueType(object value)
        {
            var s = value as string;
            if (s != null)
            {
                return new TextValue(s);
            }

            if (DBNull.Value.Equals(value) || value == null)
            {
                return new NullValue();
            }

            double v;
            if (double.TryParse(value.ToString(), out v))
            {
                return new NumericValue<double>(v);
            }

            DateTime d;
            if (DateTime.TryParse(value.ToString(), out d))
            {
                return new DateTimeValue(d);
            }

            return new TextValue(value.ToString());
        }

        IValue ConvertToIValue(FieldDataType fieldDataType, object value)
        {
            var v = TryCreateDataTypeConformValue(value, fieldDataType);

            if (v == null)
                 throw new NotSupportedException("Datatype for between not supported.");

            return v;
        }

        private bool ContainsIgnoreValue(object value, object ignoreValue)
        {
            object checkValue;
            var iv = value as IValue;

            if (iv != null)
            {
                var bv = value as IBetweenValue;
                if (bv != null)
                    return ContainsIgnoreValue(bv, ignoreValue);

                checkValue = iv.Value;
            }
            else
            {
                checkValue = value;
            }

            if (checkValue == null) checkValue = DBNull.Value;
            if (ignoreValue == null) ignoreValue = DBNull.Value;
            
            var enumerable = ignoreValue as IEnumerable;
            return (enumerable != null && !(checkValue is string)) ? ContainsIgnoreValueArray(checkValue, enumerable) : checkValue.Equals(ignoreValue);
        }

        private bool ContainsIgnoreValue(IBetweenValue value, object ignoreValue)
        {
            return ContainsIgnoreValue(value.FirstValue, ignoreValue) || ContainsIgnoreValue(value.SecondValue, ignoreValue);
        }

        private bool ContainsIgnoreValueArray(object value, IEnumerable ignoreValues)
        {
            return ignoreValues != null && ignoreValues.Cast<object>().Any(v => ContainsIgnoreValue(value, v));
        }

        private object GetDataTypeCheckedValue(object value, FieldDataType dataType)
        {
            if (dataType == FieldDataType._Unspecified)
                return value;

            if (!(value is IValue))
                value = TryCreateDataTypeConformValue(value, dataType);

            return value;
        }

        private IValue TryCreateDataTypeConformValue(object value, FieldDataType dataType)
        {

            if (value == null || DBNull.Value.Equals((value)))
                return new NullValue();

            switch (dataType)
            {
                case FieldDataType.Boolean:
                    if (value is bool) return new BooleanValue((bool)value);
                    return null;
                case FieldDataType.DateTime:
                    try
                    {
                        return new DateTimeValue((DateTime) value);
                    }
                    catch (Exception)
                    {
                        return null;
                    }
                case FieldDataType.Numeric:
                    if ((value is int) || (value is Int64))
                        return new NumericValue<int>((int)value);
                    if (value is double)
                        return new NumericValue<double>((double)value);
                    if (value is decimal)
                         return new NumericValue<decimal>((decimal)value);

                    double d;
                    return Double.TryParse(value.ToString(), out d) ? new NumericValue<double>(d) : null;

                case FieldDataType.Text:
                    return new TextValue(value.ToString());
                default:
                    return null;
            }
        }

        public IConditionGroup Add(IField field, RelationalOperators relationalOperator, IField field2)
        {
            EmbeddedStatements.Add(new FieldCondition(field, relationalOperator, field2));
            return this;
        }

        public IConditionGroup Add(IConditionGroup group)
        {
            EmbeddedStatements.Add(group);
            return this;
        }
    }
}