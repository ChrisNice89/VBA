namespace AccessCodeLib.Data.Common.Sql
{
    public interface IBooleanValue : IValue
    {
        new bool Value { get; }
    }
}