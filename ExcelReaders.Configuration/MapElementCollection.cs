using System;
using System.Collections.Generic;
using System.Configuration;

namespace ExcelReaders.Configuration
{
    [ConfigurationCollection(typeof(MapElement), AddItemName = "MapElement")]
    public class MapElementCollection : ConfigurationElementCollection , IEnumerable<MapElement>
    {
     


        public MapElement this[int index]
        {
            get { return BaseGet(index) as MapElement; }
            set
            {
                if (BaseGet(index) != null)
                {
                    BaseRemoveAt(index);
                    BaseAdd(index, value);
                }
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new MapElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }
            return ((MapElement) element).Attribute;
        }

        public IEnumerator<MapElement> GetEnumerator()
        {
            int count = Count;
            for (int i = 0; i < count; i++)
            {
                yield return BaseGet(i) as MapElement;
            }
        }
    }
}
