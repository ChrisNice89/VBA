using AccessCodeLib.Data.Common.Sql;
using System;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class ValueArray : IValueArray
    {
        protected ValueArray()
        {
        }

        public ValueArray(IValue[] values)
        {
            Values = values;
        }

        public IValue[] Values { get; private set; }
        object IValue.Value { get { return Values; } }

        public Type TypeOfValue { get { return typeof(IValueArray); } }
    }
}
