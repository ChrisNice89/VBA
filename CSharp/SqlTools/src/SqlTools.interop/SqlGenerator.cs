using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools.Converter;
using AccessCodeLib.Data.SqlTools.Sql;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("8363614A-E20E-4182-8A34-6B316D3BC24E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".SqlGenerator")]
    public class SqlGenerator : SqlTools.SqlGenerator, ISqlGenerator
    {
        public SqlGenerator() : base(new SqlStatement())
        {
        }

        public SqlGenerator(Common.Sql.Converter.ISqlConverter converter)
            : base(converter, new SqlStatement())
        {
        }

        public new ISqlConverter Converter
        {
            get
            {
                return base.Converter as ISqlConverter;
            }
            set 
            { 
                base.Converter = value;
            }
        }

        public ISqlGenerator Select(
                object field01, object field02 = null, object field03 = null, object field04 = null, object field05 = null,
                object field06 = null, object field07 = null, object field08 = null, object field09 = null, object field10 = null,
                object field11 = null, object field12 = null, object field13 = null, object field14 = null, object field15 = null,
                object field16 = null, object field17 = null, object field18 = null, object field19 = null, object field20 = null
            )
        {
            base.Select(
                    GetFields(field01, field02, field03, field04, field05, field06, field07, field08, field09, field10,
                              field11, field12, field13, field14, field15, field16, field17, field18, field19, field20));

            return this;
        }

        public new ISqlGenerator SelectAll()
        {
            base.SelectAll();
            return this;
        }

        public new ISqlGenerator SelectField(string fieldName, object source, string alias)
        {
            base.SelectField(fieldName, source, alias);
            return this;
        }

        public ISqlGenerator From(object source)
        {
            if (source is ISource)
                base.From((ISource) source);
            else if (source is string)
                base.From((string)source);
            else
                throw new NotSupportedTypeException(source.GetType());

            return this;
        }

        public ISqlGenerator InnerJoin(object leftSource, object rightSource,
                object left01, object right01, RelationalOperators relationalOperator01 = RelationalOperators.Equal,
                object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal)
        {
            return Join(JoinOperator.Inner, leftSource, rightSource,
                        left01, right01, relationalOperator01,
                        left02, right02, relationalOperator02);
        }

        public ISqlGenerator LeftJoin(object leftSource, object rightSource,
                object left01, object right01, RelationalOperators relationalOperator01 = RelationalOperators.Equal,
                object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal)
        {
            return Join(JoinOperator.Left, leftSource, rightSource,
                        left01, right01, relationalOperator01,
                        left02, right02, relationalOperator02);
        }

        public ISqlGenerator RightJoin(object leftSource, object rightSource,
                object left01, object right01, RelationalOperators relationalOperator01 = RelationalOperators.Equal,
                object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal)
        {
            return Join(JoinOperator.Right, leftSource, rightSource,
                        left01, right01, relationalOperator01,
                        left02, right02, relationalOperator02);
        }

        private ISqlGenerator Join(JoinOperator joinOperator, object leftSource, object rightSource,
                object left01, object right01, RelationalOperators relationalOperator01,
                object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal)
        {
            var conditionGenerator = new ConditionGenerator();

            var leftSourceRef = leftSource is ISource ? (ISource)leftSource : new NamedSource((string)leftSource);
            var rightSourceRef = rightSource is ISource ? (ISource)rightSource : new NamedSource((string)rightSource);

            AddCondition(conditionGenerator, leftSourceRef, rightSourceRef, left01, right01, relationalOperator01);
            AddCondition(conditionGenerator, leftSourceRef, rightSourceRef, left02, right02, relationalOperator02);

            Join(GetSourceFromObject(rightSource), conditionGenerator, joinOperator);
            return this;
        }

        private static void AddCondition(IConditionGenerator conditionGenerator, ISource leftSource, ISource rightSource, object leftCondValue, object rightCondValue, RelationalOperators relationalOperator)
        {
            if (leftCondValue == null && rightCondValue == null)
                return;

            if (!(leftCondValue is IField || leftCondValue is string) && (rightCondValue != null && (rightCondValue is IField || rightCondValue is string)))
            {
                var tempCondValue = leftCondValue;
                leftCondValue = rightCondValue;
                rightCondValue = tempCondValue;
                
                var tempSource = leftSource;
                leftSource = rightSource;
                rightSource = tempSource;

                relationalOperator = RevertRelationalOperators(relationalOperator);
            }
            
            var field = leftCondValue is IField ? (IField)leftCondValue : new Field((string)leftCondValue, leftSource);
            if ((rightCondValue is IField) || rightCondValue is string)
                conditionGenerator.Add(field, (Common.Sql.RelationalOperators)relationalOperator, rightCondValue is IField ? rightCondValue : new Field((string)rightCondValue, rightSource));
            else
                conditionGenerator.Add(field, (Common.Sql.RelationalOperators)relationalOperator, rightCondValue);
        }

        private static RelationalOperators RevertRelationalOperators(RelationalOperators op)
        {
            RelationalOperators newOp = 0;
            if ((op & RelationalOperators.GreaterThan) == RelationalOperators.GreaterThan)
            {
                newOp |= RelationalOperators.LessThan;
                op -= RelationalOperators.GreaterThan;
            }

            if ((op & RelationalOperators.LessThan) == RelationalOperators.LessThan)
            {
                newOp |= RelationalOperators.GreaterThan;
                op -= RelationalOperators.LessThan;
            }
            newOp |= op;

            return newOp;
        }

        private static ISource GetSourceFromObject(object source)
        {
            return source is ISource ? (ISource) source : new NamedSource((string) source);
        }

        public ISqlGenerator Where(object field, RelationalOperators relationalOperator, object value)
        {
            Where(new FieldCondition(FieldGenerator.GetFieldFromObject(field), (Common.Sql.RelationalOperators)relationalOperator, value));
            return this;
        }

        public ISqlGenerator WhereBetween(object field, int value1, int value2)
        {
            Where(new FieldCondition(FieldGenerator.GetFieldFromObject(field), Common.Sql.RelationalOperators.Between, new BetweenValue(value1, value2)));
            return this;
        }

        public ISqlGenerator WhereBetween(object field, object value1, object value2)
        {
            var f = FieldGenerator.GetFieldFromObject(field);
            var betweenValue = CreateBetweenValue(f.DataType, value1, value2);
            Where(new FieldCondition(f, Common.Sql.RelationalOperators.Between, betweenValue));
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

        private IValue GetIValueFromValueType(object value)
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
            if (DBNull.Value.Equals(value) || value == null)
            {
                return new NullValue();
            }
                
            switch (fieldDataType)
            {
                case FieldDataType.Numeric:
                    return new NumericValue<double>((double)value);
                case FieldDataType.Text:
                    return new TextValue((string)value);
                case FieldDataType.DateTime:
                    return new DateTimeValue((DateTime)value);
                case FieldDataType.Boolean:
                    return new BooleanValue((bool)value);
                default:
                    throw new NotSupportedException("Datatype for between not supported.");
            }
        }

        public ISqlGenerator WhereCondition(ICondition condition)
        {
            Where(condition);
            return this;
        }

        public ISqlGenerator WhereString(string whereString)
        {
            Where(whereString);
            return this;
        }

        public ISqlGenerator GroupBy(
                object field01, object field02 = null, object field03 = null, object field04 = null, object field05 = null,
                object field06 = null, object field07 = null, object field08 = null, object field09 = null, object field10 = null,
                object field11 = null, object field12 = null, object field13 = null, object field14 = null, object field15 = null,
                object field16 = null, object field17 = null, object field18 = null, object field19 = null, object field20 = null
            )
        {
            base.GroupBy(
                    GetFields(field01, field02, field03, field04, field05, field06, field07, field08, field09, field10,
                              field11, field12, field13, field14, field15, field16, field17, field18, field19, field20));

            return this;
        }

        public ISqlGenerator Having(object field, RelationalOperators relationalOperator, object value)
        {
            Having(new FieldCondition(FieldGenerator.GetFieldFromObject(field), (Common.Sql.RelationalOperators)relationalOperator, value));
            return this;
        }

        public ISqlGenerator HavingCondition(ICondition condition)
        {
            Having(condition);
            return this;
        }

        public ISqlGenerator HavingString(string havingString)
        {
            Having(havingString);
            return this;
        }

        public ISqlGenerator OrderBy(
                object field01, object field02 = null, object field03 = null, object field04 = null, object field05 = null,
                object field06 = null, object field07 = null, object field08 = null, object field09 = null, object field10 = null,
                object field11 = null, object field12 = null, object field13 = null, object field14 = null, object field15 = null,
                object field16 = null, object field17 = null, object field18 = null, object field19 = null, object field20 = null
            )
        {
            base.OrderBy(
                    GetFields(field01, field02, field03, field04, field05, field06, field07, field08, field09, field10,
                              field11, field12, field13, field14, field15, field16, field17, field18, field19, field20));

            return this;
        }

        private static IField[] GetFields(
                object field01, object field02 = null, object field03 = null, object field04 = null, object field05 = null,
                object field06 = null, object field07 = null, object field08 = null, object field09 = null, object field10 = null,
                object field11 = null, object field12 = null, object field13 = null, object field14 = null, object field15 = null,
                object field16 = null, object field17 = null, object field18 = null, object field19 = null, object field20 = null
            )
        {
            var fields = new List<IField>();
            var objectArray = new[]
                                  {
                                      field01, field02, field03, field04, field05, field06, field07, field08, field09, field10,
                                      field11, field12, field13, field14, field15, field16, field17, field18, field19, field20
                                  };

            foreach (var o in objectArray.Where(o => o != null))
            {
                if (o is IField)
                    fields.Add((IField)o);
                else if (o is string)
                    fields.Add(new Field((string) o));
                else if (o is IField[])
                    fields.AddRange((IField[])o);
                else if (o is string[])
                    fields.AddRange(((string[]) o).Select(s => new Field(s)).Cast<IField>());
                else if (o is IFieldList)
                    fields.AddRange(((IFieldList) o).Cast<IField>().ToArray());
                else
                {
                    throw new NotSupportedTypeException(o.GetType());
                }
            }
            return fields.ToArray();
        }

        public new ISqlStatement SqlStatement
        {
            get { return (ISqlStatement)base.SqlStatement; }
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("1F8C5EF1-66D8-43A1-95BC-E4EE5F64F07A")]
    public interface ISqlGenerator : SqlTools.ISqlGenerator
    {
// ReSharper disable UnusedMember.Global
        ISqlConverter Converter { get; set; }
// ReSharper restore UnusedMember.Global

        ISqlGenerator Select(
            object Field1, object Field2 = null, object Field3 = null, object Field4 = null, object Field5 = null,
            object Field6 = null, object Field7 = null, object Field8 = null, object Field9 = null, object Field10 = null,
            object Field11 = null, object Field12 = null, object Field13 = null, object Field14 = null, object Field15 = null,
            object Field16 = null, object Field17 = null, object Field18 = null, object Field19 = null, object Field20 = null
            );
        new ISqlGenerator SelectAll();
        ISqlGenerator SelectField(string FieldName, object Source, string Alias);

        ISqlGenerator From(object Source);

        ISqlGenerator InnerJoin(object leftSource, object rightSource,
            object left01, object right01, RelationalOperators relationalOperator01 = RelationalOperators.Equal,
            object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal);

// ReSharper disable UnusedMethodReturnValue.Global
        ISqlGenerator LeftJoin(object leftSource, object rightSource,
            object left01, object right01, RelationalOperators relationalOperator01 = RelationalOperators.Equal,
            object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal);

        ISqlGenerator RightJoin(object leftSource, object rightSource,
            object left01, object right01, RelationalOperators relationalOperator01 = RelationalOperators.Equal,
            object left02 = null, object right02 = null, RelationalOperators relationalOperator02 = RelationalOperators.Equal);

        ISqlGenerator Where(object field, RelationalOperators RelationalOperator, object value);
        ISqlGenerator WhereBetween(object field, object value1, object value2);
        ISqlGenerator WhereCondition(ICondition condition);
        ISqlGenerator WhereString(string whereString);

        ISqlGenerator GroupBy(
            object Field1, object Field2 = null, object Field3 = null, object Field4 = null, object Field5 = null,
            object Field6 = null, object Field7 = null, object Field8 = null, object Field9 = null, object Field10 = null,
            object Field11 = null, object Field12 = null, object Field13 = null, object Field14 = null, object Field15 = null,
            object Field16 = null, object Field17 = null, object Field18 = null, object Field19 = null, object Field20 = null
            );

        ISqlGenerator Having(object field, RelationalOperators RelationalOperator, object value);
        ISqlGenerator HavingCondition(ICondition condition);
        ISqlGenerator HavingString(string HavingString);

        ISqlGenerator OrderBy(
            object Field1, object Field2 = null, object Field3 = null, object Field4 = null, object Field5 = null,
            object Field6 = null, object Field7 = null, object Field8 = null, object Field9 = null, object Field10 = null,
            object Field11 = null, object Field12 = null, object Field13 = null, object Field14 = null, object Field15 = null,
            object Field16 = null, object Field17 = null, object Field18 = null, object Field19 = null, object Field20 = null
            );
// ReSharper restore UnusedMethodReturnValue.Global

        [DispId(0)]
        new ISqlStatement SqlStatement { get; }
        new string ToString();
    }
}
