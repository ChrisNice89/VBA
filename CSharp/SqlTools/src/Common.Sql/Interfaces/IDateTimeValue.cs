using System;

namespace AccessCodeLib.Data.Common.Sql
{
    public interface IDateTimeValue : IValue
    {
        new DateTime Value { get; }
    }
}
