using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("A4657024-448A-443B-89BA-4BE9BBF28233")]
    public enum FieldDataType
    {
        _Unspecified = -1,
        Boolean = 1,
        Numeric = 2,
        Text = 3,
        DateTime = 4
    }
}