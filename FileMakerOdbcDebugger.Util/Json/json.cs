using System;
using System.Collections;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace FileMakerOdbcDebugger.Util
{
    public static class Json
    {
        /// <summary>
        /// Convert a JSON string to a specific type.
        /// </summary>
        public static object FromJson(this string s, Type ObjType)
        {
            return JsonConvert.DeserializeObject(s, ObjType);
        }

        /// <summary>
        /// Convert an object to a JSON string.
        /// Sort properties on objects alphabetically.
        /// </summary>
        public static string ToJson(this object o, Formatting formatting = Formatting.Indented)
        {
            string s = JsonConvert.SerializeObject(o, formatting);

            if (o as IEnumerable == null)
            {
                // Object: sort.
                var parsedObject = JObject.Parse(s);
                var normalizedObject = JSON_SortPropertiesAlphabetically(parsedObject);
                return JsonConvert.SerializeObject(normalizedObject, formatting);
            }
            else
            {
                // List: don't sort.
                return s;
            }
        }

        private static JObject JSON_SortPropertiesAlphabetically(JObject original)
        {
            var result = new JObject();

            foreach (var property in original.Properties().ToList().OrderBy(p => p.Name))
            {
                JObject value = property.Value as JObject;

                if (value is object)
                {
                    value = JSON_SortPropertiesAlphabetically(value);
                    result.Add(property.Name, value);
                }
                else
                {
                    result.Add(property.Name, property.Value);
                }
            }

            return result;
        }
    }
}
