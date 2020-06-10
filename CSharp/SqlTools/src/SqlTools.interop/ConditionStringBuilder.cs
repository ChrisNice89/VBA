using System.Collections.Generic;
using System.Runtime.InteropServices;

// ReSharper disable InconsistentNaming
namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("3C7E0ED9-9B0E-4387-BF50-274760F8B087")]
    public interface IConditionStringBuilder : IConditionGenerator
    {
        // IConditionGenerator:
        new LogicalOperator ConcatOperator { get; }

        new IConditionGroup Add(object Field, FieldDataType DataType, RelationalOperators RelationalOperator, object Value, object IgnoreValue = null);
        new IConditionGroup AddConditionString(string ConditionString);

        IConditionGroup AddBetweenCondition(object Field, FieldDataType DataType, object Value, object Value2, object IgnoreValue = null);
        new IConditionGroup BeginGroup(LogicalOperator ConcatOperator = LogicalOperator.And);
        new IEnumerable<ICondition> Conditions { get; }

        // +
        ISqlConverter SqlConverter { get; set; }
        string ToString(LogicalOperator ConcatOperator = LogicalOperator.And);
    }

    [ComVisible(true)]
    [Guid("87B0C8AD-C228-42D1-89F2-DEE1058A1909")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".ConditionStringBuilder")]
    public class ConditionStringBuilder : ConditionGenerator, IConditionStringBuilder
    {
        public ConditionStringBuilder()
        {
        }

        public ConditionStringBuilder(ISqlConverter converter)
        {
            SqlConverter = converter;
        }

        public IConditionGroup AddBetweenCondition(object Field, FieldDataType DataType, object Value, object Value2, object IgnoreValue = null)
        {
            return Add(Field, DataType, RelationalOperators.Between, new[] { Value, Value2 }, IgnoreValue);
        }

        public ISqlConverter SqlConverter { get; set; }

        public string ToString(LogicalOperator concatOperator = LogicalOperator.And)
        {
            ConcatOperator = concatOperator;
            return SqlConverter.GenerateConditionString(this);
        }
    }
}
// ReSharper restore InconsistentNaming