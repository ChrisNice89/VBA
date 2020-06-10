namespace AccessCodeLib.Data.Common.Sql
{
    public interface ITextValue : IValue
    {
        new string Value { get; }
    }
}
