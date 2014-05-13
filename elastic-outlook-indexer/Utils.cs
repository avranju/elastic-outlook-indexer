using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace elastic_outlook_indexer
{
    public class Utils
    {
        public static void AssignObject<TFrom, TTo>(
            TFrom from, TTo to, Dictionary<Type, Type> typeMapping,
            IList<Type> typesToSkip)
        {
            // get all property names and types from "TTo"
            var props = EnumProperties(typeof(TTo));

            // for every property in props, do the following:
            //  - check if a type mapping exists for that type in "typeMapping"
            //  - if yes, then recursively call AssignObject on that property
            //  - if no, then do a simple assignment
            foreach (var prop in props)
            {
                var propName = prop.Item1;
                var propType = prop.Item2;

                // get value to be assigned from "from"
                var fromVal = from.GetType().InvokeMember(
                    propName,
                    BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.GetProperty,
                    null, from, null);
                var fromType = fromVal.GetType();

                // check if propType is in the list of types to be skipped
                if (typesToSkip.Contains(propType))
                {
                    continue;
                }

                if (typeMapping.ContainsKey(propType))
                {
                    // create an instance of "propType"
                    var toVal = propType.InvokeMember(null,
                        BindingFlags.CreateInstance | BindingFlags.Instance | BindingFlags.Public,
                        null, null, null);

                    // runtime invoke AssignObject; essentially translates to
                    // something like this:
                    //  AssignObject<fromType, propType>(fromVal, toVal, typeMapping, typesToSkip);
                    var method = typeof(Utils).
                                    GetMethod("AssignObject").
                                    MakeGenericMethod(fromVal.GetType(), propType);
                    method.Invoke(null, new object[] { fromVal, toVal, typeMapping, typesToSkip });

                    // assign the new value to "to"
                    typeof(TTo).InvokeMember(propName,
                        BindingFlags.Instance | BindingFlags.SetProperty | BindingFlags.Public,
                        null, to, new object[] { toVal });
                }
                else
                {
                    // if target is a string and source isn't we do a "ToString"
                    if (propType == typeof(String) && fromType != propType)
                    {
                        fromVal = fromVal.ToString();
                    }

                    typeof(TTo).InvokeMember(propName,
                        BindingFlags.Instance | BindingFlags.SetProperty | BindingFlags.Public,
                        null, to, new object[] { fromVal });
                }
            }
        }

        private static IList<Tuple<string, Type>> EnumProperties(Type type)
        {
            var props = new List<Tuple<string, Type>>();
            foreach (var p in type.GetProperties())
            {
                props.Add(new Tuple<string, Type>(p.Name, p.PropertyType));
            }

            return props;
        }
    }
}
