namespace AccessCodeLib.Data.Common.Sql
{
    public interface IField
    {
        string Name { get; }
        ISource Source { get; }
        FieldDataType DataType { get; }
    }
}