namespace AccessCodeLib.Data.Common.Sql
{
    public interface IBetweenValue : IValue
    {
        IValue FirstValue { get; }
        IValue SecondValue { get; }
    }
}
