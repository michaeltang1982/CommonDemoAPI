using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Reflection;

namespace Sierra.SharePoint.Library.CSOM
{
    public class ItemProperty
    {
        public string PropertyName { get; private set; }
        public string PropertyValue { get; private set; }
        public Type PropertyType { get; private set; }

        /// <summary>
        /// set string property
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public ItemProperty(string name, string value): this(name, value, "System.string")
        {
        }

        /// <summary>
        /// set property name, value and type
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="type">string should be in the form "AssemblyName:typename" or simply "System.typename" for the built-in types</param>
        public ItemProperty(string name, string value, string type)
        {
            this.PropertyName = name;
            this.PropertyValue = value;
            if (string.IsNullOrEmpty(type)) type = "System.string";
            this.PropertyType = LoadType(type);
            
        }

        public ItemProperty(string name, string value, Type type)
        {
            this.PropertyName = name;
            this.PropertyValue = value;
            this.PropertyType = type;
        }

        private Type LoadType(string type)
        {
            if (type.Contains(":"))
            {
                string[] typeParts = type.Split(':');
                Assembly assembly = Assembly.LoadWithPartialName(typeParts[0]);
                return assembly.GetType(string.Format("{0}.{1}", typeParts[0],typeParts[1]), true, true);
            }
            else
                return Type.GetType(type, true, true);

        }
    }
}
