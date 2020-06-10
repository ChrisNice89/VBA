using System;

namespace AccessCodeLib.Data.Common.Sql
{
    public interface IValue
    {
        Type TypeOfValue { get; }
        object Value { get; }
    }
}
