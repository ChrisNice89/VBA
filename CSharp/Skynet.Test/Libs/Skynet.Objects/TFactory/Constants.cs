using System.Runtime.InteropServices;

namespace Skynet.Objects
{
    [ComVisible(true)]
    [Guid("D6FD53DF-5FB9-4451-88A3-11D497C4E46D"),ClassInterface(ClassInterfaceType.None)]
    [ProgId(ProgID + ".Constants")]
    public class Constants : IConstantsComInterface
    {
        [ComVisible(true)]
        public const string ProgID = "Skynet";

        [ComVisible(true)]
        public string NameSpace { get { return "Skynet.Objects"; } }

  
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A3A18390-A8E7-47B9-8ABA-1E15225CF38A")]
    public interface IConstantsComInterface
    {
        // ReSharper disable UnusedMember.Global
        string NameSpace { get; }
  
        // ReSharper restore UnusedMember.Global
    }
}
