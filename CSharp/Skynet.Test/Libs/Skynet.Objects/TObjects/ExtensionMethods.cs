using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects.TObjects.Extensions
{
    public static class ExtensionMethods
    {
        public static IParameter ToParameter(this IValue instance, string Name)
        {
            return (IParameter)instance;
        }
    }


   

}
