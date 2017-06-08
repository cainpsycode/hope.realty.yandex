using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace xlsx.convert.yrl
{
    public static class EnumsHelper
    {
        public static T GetAttributeOfType<T>(this Enum enumVal) where T : Attribute
        {
            var typeInfo = enumVal.GetType().GetTypeInfo();
            var v = typeInfo.DeclaredMembers.First(x => x.Name == enumVal.ToString());
            return v.GetCustomAttribute<T>();
        }

        public static Enum GetEnumValueByAttribute<T>(this T enumType, T attrType, Func<Attribute, bool> comparer) where T: Type
        {
            return Enum.GetValues(enumType)
                .Cast<Enum>()
                .FirstOrDefault(i =>
                {
                    var memInfo = enumType.GetMember(i.ToString());
                    var attr = memInfo[0].GetCustomAttributes(attrType, false);
                    bool re = comparer((Attribute)attr[0]);
                    return re;
                });
        }

        public static string GetDescription(this Enum enumVal)
        {
            var attr = GetAttributeOfType<DescriptionAttribute>(enumVal);
            return attr != null ? attr.Text : string.Empty;
        }
    }

    public static class DateTimeHelper
    {
        public static string ToIso8601(this DateTime date)
        {
            return date.ToString("yyyy-MM-ddTHH\\:mm\\:sszzz");
        }
    }
}
