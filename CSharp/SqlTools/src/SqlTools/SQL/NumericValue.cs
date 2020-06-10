using AccessCodeLib.Data.Common.Sql;
using System;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class NumericValue<T> : INumericValue<T>
    {
        protected NumericValue()
        {
        }

        public NumericValue(T value)
        {
            Value = value;
        }

        public T Value { get; private set; }
        object IValue.Value { get { return Value; } }

        public Type TypeOfValue { get { return typeof(T); } }
        
    }
}
