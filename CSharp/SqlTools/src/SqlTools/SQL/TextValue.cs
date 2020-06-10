using AccessCodeLib.Data.Common.Sql;
using System;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class TextValue : ITextValue
    {
        protected TextValue()
        {
        }

        public TextValue(string value)
        {
            Value = value;
        }

        public string Value { get; private set;}
        object IValue.Value { get { return Value; } }

        public Type TypeOfValue { get { return typeof(string); } }
    }
}
