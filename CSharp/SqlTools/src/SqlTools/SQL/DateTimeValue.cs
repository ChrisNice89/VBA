using System;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class DateTimeValue : IDateTimeValue
    {
        protected DateTimeValue()
        {
        }

        public DateTimeValue(DateTime value)
        {
            Value = value;
        }

        public DateTime Value { get; private set;}
        object IValue.Value { get { return Value; } }

        public Type TypeOfValue { get { return typeof(DateTime); } }
    }
}
