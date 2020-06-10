using System;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class NotSupportedTypeException : NotSupportedException
    {

        public NotSupportedTypeException(Type type)
            : base(string.Format("Type '{0}' is not supported.", type))
        {
            Type = type;
        }

        public NotSupportedTypeException(string message, Type type)
            : base(message)
        {
            Type = type;
        }

        public Type Type { get; private set; }
    }
}
