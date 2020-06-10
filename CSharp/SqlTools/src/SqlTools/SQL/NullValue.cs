using AccessCodeLib.Data.Common.Sql;
using System;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class NullValue : INullValue
    {
        public NullValue()
        {
            Value = System.DBNull.Value;
        }

        public System.DBNull Value { get; private set; }
        object IValue.Value{ get { return Value; } }

        public Type TypeOfValue { get { return typeof(System.DBNull); } }
        
    }
}
