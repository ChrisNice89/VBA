using Skynet.Objects.TObjects.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects.TObjects
{
    [ComVisible(true)]
    [Guid("2FB0CA37-183C-4894-9BD8-2ADB5FCC71B0"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IParameter
    {
        string Name { get; }
        object Value { get; }
    }
}
