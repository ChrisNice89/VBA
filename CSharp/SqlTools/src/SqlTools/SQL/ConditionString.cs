using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class ConditionString : IConditionString
    {
        public ConditionString()
        {
        }

        public ConditionString(string value)
        {
            Value = value;
        }

        public string Value { get; set; }
    }
}