namespace AccessCodeLib.Data.Common.Sql
{
    public interface IValueArray : IValue
    {
        IValue[] Values { get; }
    }
}
