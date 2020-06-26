using System.Runtime.InteropServices;

namespace Skynet.DAO
{
    [ComVisible(true)]
    [Guid("8DBFEFB5-8DA6-414D-9540-8DB23E26788E"), ClassInterface(ClassInterfaceType.None)]
    [ProgId(ProgID + ".Constants")]
    public class Constants : IConstantsComInterface
    {
        [ComVisible(true)]
        public const string ProgID = "Skynet.DAO";

        [ComVisible(true)]
        public string NameSpace { get { return "Skynet.DAO"; } }


    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("83035D4E-1F30-413D-A74E-412DA7D714EC")]
    public interface IConstantsComInterface
    {
        string NameSpace { get; }
    }
}