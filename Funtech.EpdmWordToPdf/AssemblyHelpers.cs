using System;
using System.Reflection;

namespace Funtech.EpdmWordToPdf
{
    public static class AssemblyHelpers
    {
        private static string GetAttributeValue<T>(object inst, Func<T, string> getValue) where T : Attribute
        {
            return getValue((T)inst);
        }

        public static int GetMajorVersion(this Assembly assy)
        {
            return Assembly.GetExecutingAssembly().GetName().Version.Major;
        }

        /// <summary>
        /// Generic Assembly extension method for getting Assembly attribute values
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="assembly"></param>
        /// <returns></returns>
        public static string GetAttributeValue<T>(this Assembly assembly) where T : Attribute
        {
            object[] attributes = assembly.GetCustomAttributes(typeof(T), false);

            string value = null;

            if (attributes.Length > 0)
            {
                if (attributes[0] is AssemblyTitleAttribute)
                {
                    return GetAttributeValue<AssemblyTitleAttribute>(attributes[0], x => { return x.Title; });
                }
                else if (attributes[0] is AssemblyDescriptionAttribute)
                {
                    return GetAttributeValue<AssemblyDescriptionAttribute>(attributes[0], x => { return x.Description; });
                }
                else if (attributes[0] is AssemblyCompanyAttribute)
                {
                    return GetAttributeValue<AssemblyCompanyAttribute>(attributes[0], x => { return x.Company; });
                }
                else if (attributes[0] is AssemblyProductAttribute)
                {
                    return GetAttributeValue<AssemblyProductAttribute>(attributes[0], x => { return x.Product; });
                }
                else if (attributes[0] is AssemblyCopyrightAttribute)
                {
                    return GetAttributeValue<AssemblyCopyrightAttribute>(attributes[0], x => { return x.Copyright; });
                }
                else
                {
                    throw new ArgumentException("type not supported");
                }
            }

            return value;
        }
    }
}
