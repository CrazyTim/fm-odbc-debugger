using System;
using System.ComponentModel;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Extensions {

        /// <summary>
        /// Return the "Description" attribute of an enum using Reflection.
        /// Return an empty string if the description isn't set, or if the enum doesn't exist.
        /// </summary>
        public static string Description(this Enum EnumConstant) {

            // refer: https://stackoverflow.com/a/10986749

            try {

                var field = EnumConstant.GetType().GetField(EnumConstant.ToString());

                DescriptionAttribute[] attributes = (DescriptionAttribute[])field.GetCustomAttributes(typeof(DescriptionAttribute), false);

                if (attributes.Length > 0) {
                    return attributes[0].Description;
                }

                // no description exists
                return "";

            } catch {
                // enum doesnt exist
                return "";
            }

        }

    }

}
