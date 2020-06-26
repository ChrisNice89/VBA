using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.DAO.Enums
{

    [ComVisible(true)]
    [Guid("DD8844DD-5AB8-4195-BF17-B902C1324D7D")]
    public enum ConnectionType : int
    {
        [Description("EXCEL")]
        Excel =0 ,
        [Description("CSV")]
        CSV ,
        [Description("MSACCESS")]
        MSACCESS,
        [Description("SQL")]
        SQL
    }
    public static class EnumExtensions
    {
        public static string GetDescription(this Enum value)
        {
            Type type = value.GetType();
            string name = Enum.GetName(type, value);
            if (name != null)
            {
                System.Reflection.FieldInfo field = type.GetField(name);
                if (field != null)
                {
                    DescriptionAttribute attr = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;
                    if (attr != null)
                    {
                        return attr.Description;
                    }
                }
            }
            return null;
        }
    }
}
