using AccessCodeLib.Data.Common.Sql;
using System;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class BooleanValue : IBooleanValue
    {
        public BooleanValue(bool value)
        {
            Value = value;
        }

        public bool Value { get; private set; }
        object IValue.Value { get { return Value; }}
    
        public Type TypeOfValue { get { return typeof(bool); } }
        
    }
}
