using System.Runtime.InteropServices;
using System.Collections.Generic;
using AccessCodeLib.Data.SqlTools.Sql;

// ReSharper disable InconsistentNaming
namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("5C7E39FF-786C-450C-BC04-AA66B23F1043")]
    public interface IConditionGenerator : IConditionGroup
    {
        // ReSharper disable UnusedMember.Global
        new LogicalOperator ConcatOperator { get; }
        IConditionGroup Add(object Field, FieldDataType DataType, RelationalOperators RelationalOperator, object Value, object ignoreValue = null);
        IConditionGroup AddConditionString(string ConditionString);
        IConditionGroup BeginGroup(LogicalOperator ConcatOperator = LogicalOperator.And);
        new IEnumerable<ICondition> Conditions { get; }
        // ReSharper restore UnusedMember.Global
    }
// ReSharper restore InconsistentNaming

    [ComVisible(true)]
    [Guid("8569A981-D2CA-475F-BD4A-DFB066E87254")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".ConditionGenerator")]
    public class ConditionGenerator : SqlTools.ConditionGenerator, IConditionGenerator
    {
        public new LogicalOperator ConcatOperator
        {
            get { return (LogicalOperator) base.ConcatOperator; }
            set { base.ConcatOperator = (Common.Sql.LogicalOperator) value; }
        }

        public IConditionGroup AddCondition(ICondition condition)
        {
            Add(condition);
            return this;
        }

        public IConditionGroup AddConditionString(string conditionString)
        {
            if (string.IsNullOrEmpty(conditionString))
                return this;

            var condition = new ConditionString(conditionString);
            Add(condition);
            return this;
        }

        public IConditionGroup BeginGroup(LogicalOperator concatOperator = LogicalOperator.And)
        {
            var group = new ConditionGroup(concatOperator);
            Add(group);
            return group;
        }

        public IConditionGroup Add(object field, RelationalOperators relationalOperator, object value, object ignoreValue = null)
        {
            return Add(field, FieldDataType._Unspecified, relationalOperator, value, ignoreValue);
        }

        public IConditionGroup Add(object field, FieldDataType dataType, RelationalOperators relationalOperator,
                                   object value, object ignoreValue = null)
        {
            var fld = FieldGenerator.GetFieldFromObject(field);

            if (fld.DataType != dataType && dataType != FieldDataType._Unspecified)
            {
                fld.DataType = dataType;
            }

            Add(fld, (Common.Sql.RelationalOperators) relationalOperator, value, ignoreValue);
            return this;
        }


        // ReSharper disable MemberCanBePrivate.Global
        // ReSharper disable UnusedMethodReturnValue.Global
        public IConditionGroup Add(IConditionGroup group)
            // ReSharper restore UnusedMethodReturnValue.Global
            // ReSharper restore MemberCanBePrivate.Global
        {
            base.Add(group);
            return this;
        }

        public IConditionGroup BeginGroup(IField field, RelationalOperators relationalOperator, object value,
                                          LogicalOperator concatOperator = LogicalOperator.And)
        {
            var group = new ConditionGroup(concatOperator);
            Add(group);
            group.Add(field, relationalOperator, value);
            return group;
        }

        public new IEnumerable<ICondition> Conditions
        {
            get
            {
                return EmbeddedStatements.ConvertAll(
                    CommonSqlConditionToInteropCondition);
            }
        }

        private static ICondition CommonSqlConditionToInteropCondition(Common.Sql.ICondition condition)
        {
            return (ICondition) condition;
        }
    }
}