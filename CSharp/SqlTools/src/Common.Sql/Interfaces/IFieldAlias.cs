namespace AccessCodeLib.Data.Common.Sql
{
    public interface IFieldAlias : IAlias, IField
    {
// ReSharper disable UnusedMemberInSuper.Global
        IField Field { get; }
// ReSharper restore UnusedMemberInSuper.Global
    }
}