namespace AccessCodeLib.Data.Common.Sql
{
    public interface INullValue : IValue
    {
        new System.DBNull Value { get; }
    }
}
