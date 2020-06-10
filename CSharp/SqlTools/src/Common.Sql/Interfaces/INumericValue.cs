namespace AccessCodeLib.Data.Common.Sql
{
    public interface INumericValue<out T> : IValue
    {
        new T Value { get; }
    }
}
