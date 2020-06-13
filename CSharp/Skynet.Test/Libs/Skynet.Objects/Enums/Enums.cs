using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects.Enums
{
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
                    DescriptionAttribute attr =Attribute.GetCustomAttribute(field,typeof(DescriptionAttribute)) as DescriptionAttribute;
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
[ComVisible(true)]
[Guid("DCECEB25-9167-4ED7-A0A0-03F1DB63C217")]
public enum CompareResult : int
{
    [Description("IsLower")]
    IsLower = -1,
    [Description("IsGreater")]
    IsGreater = 1,
    [Description("Equals")]
    Equals = 0
}
[ComVisible(true)]
[Guid("79C8F897-FDE8-4AE2-8D7D-71FA83EC307A")]
public enum Connectiontype : int
{
    [Description("EXCEL")]
    EXCEL = 0,
    [Description("CSV")]
    CSV,
    [Description("SQL")]
    SQL,
    [Description("MSACCESS")]
    MSACCESS
}

