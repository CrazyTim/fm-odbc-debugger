using System;
using System.ComponentModel;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Extensions
    {
        /// <summary>
        /// Return the "Description" attribute of an enum using Reflection.
        /// Return an empty string if the description isn't set, or if the enum doesn't exist.
        /// See https://stackoverflow.com/a/10986749.
        /// </summary>
        public static string Description(this Enum enumConstant)
        {
            try
            {
                var field = enumConstant.GetType().GetField(enumConstant.ToString());
                var attributes = (DescriptionAttribute[])field.GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attributes.Length > 0) return attributes[0].Description;
            }
            catch
            {
                // enum doesn't exist
            }
            return "";
        }
    }
}
