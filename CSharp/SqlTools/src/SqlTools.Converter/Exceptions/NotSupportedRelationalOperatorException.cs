using System;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class NotSupportedRelationalOperatorException : NotSupportedException
    {

        public NotSupportedRelationalOperatorException(RelationalOperators relationalOperator)
            : base(string.Format("Operator '{0}' is not supported.", relationalOperator))
        {
            RelationalOperator = relationalOperator;
        }

        public NotSupportedRelationalOperatorException(string message, RelationalOperators relationalOperator)
            : base(message)
        {
            RelationalOperator = relationalOperator;
        }

        public RelationalOperators RelationalOperator { get; private set; }
    }
}