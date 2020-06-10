using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace AccessCodeLib.Data.SqlTools.interop
{
    public class ConditionGroup : Sql.ConditionGroup, IConditionGroup
    {
        public ConditionGroup(LogicalOperator concatOperator = LogicalOperator.And)
            : base((Common.Sql.LogicalOperator)concatOperator)
        {
        }

        public new LogicalOperator ConcatOperator
        {
            get { return (LogicalOperator)base.ConcatOperator; }
            set { base.ConcatOperator = (Common.Sql.LogicalOperator)value; }
        }

        public IConditionGroup AddCondition(ICondition condition)
        {
            EmbeddedStatements.Add(condition);
            return this;
        }

        public IConditionGroup Add(object field, RelationalOperators relationalOperator, object value, object ignoreValue = null)
        {
            Add(field is IField ? (IField)field : new Field((string)field), relationalOperator, value, ignoreValue);
            return this;
        }

        public IConditionGroup Add(object field, FieldDataType dataType, RelationalOperators relationalOperator, object value, object ignoreValue = null)
        {
            Add(field is IField ? (IField)field : new Field((string)field, null , dataType), relationalOperator, value, ignoreValue);
            return this;
        }

// ReSharper disable UnusedMethodReturnValue.Global
        public IConditionGroup Add(IField field, RelationalOperators relationalOperator, object value, object ignoreValue = null)
// ReSharper restore UnusedMethodReturnValue.Global
        {
            Add(field, (Common.Sql.RelationalOperators)((int)relationalOperator), value, ignoreValue);
            return this;
        }

        //private override IList<Common.Sql.ICondition> Conditions { get { return EmbeddedStatements; } }
        public new IEnumerable<ICondition> Conditions
        {
            get
            {
                return EmbeddedStatements.ConvertAll(
                    CommonSqlConditionToInteropCondition);
            }
        }

        public static ICondition CommonSqlConditionToInteropCondition(Common.Sql.ICondition condition)
        {
            return (ICondition) condition;
        }

    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("1668328C-C77E-47C3-A747-46E4485BB28B")]
    public interface IConditionGroup : Common.Sql.IConditionGroup, ICondition
    {
// ReSharper disable UnusedMember.Global
        new LogicalOperator ConcatOperator { get; set; }
// ReSharper restore UnusedMember.Global
        new IEnumerable<ICondition> Conditions { get; }
// ReSharper disable UnusedMember.Global
        IConditionGroup AddCondition(ICondition condition);
// ReSharper restore UnusedMember.Global
        IConditionGroup Add(object Field, FieldDataType DataType, RelationalOperators RelationalOperator, object Value, object ignoreValue = null);
        //IConditionGroup Add(IConditionGroupComInterface group);
        //IConditionGroup Add(IEnumerable<ICondition> conditions);
    }
}