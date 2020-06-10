using System;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class NotSupportedSourceException : NotSupportedException
    {
        public NotSupportedSourceException(ISource source)
            : base(string.Format("Source '{0}' is not supported.", source))
        {
            SourceObject = source;
        }

        public ISource SourceObject { get; private set; }
    }
}