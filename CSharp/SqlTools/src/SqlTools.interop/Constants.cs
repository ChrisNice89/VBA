using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("D6FD53DF-5FB9-4451-88A3-11D497C4E46D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(ProgIdLibName + ".Constants")]
    public class Constants : IConstantsComInterface
    {
        private const string SqlToolsDefaultNameSpace = "AccessCodeLib.Data.SqlTools.interop";

        [ComVisible(true)]
        public const string ProgIdLibName = "ACLibSqlTools";

        [ComVisible(true)]
        public string DefaultNameSpace { get { return SqlToolsDefaultNameSpace; } }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A3A18390-A8E7-47B9-8ABA-1E15225CF38A")]
    public interface IConstantsComInterface
    {
// ReSharper disable UnusedMember.Global
        string DefaultNameSpace { get; }
// ReSharper restore UnusedMember.Global
    }
}
