using System;
using System.Collections.Generic;
using System.Configuration;

namespace ExcelReaders.Configuration
{

    [ConfigurationCollection(typeof(MapperElement), AddItemName = "Mapper")]
    public class MapperCollection : ConfigurationElementCollection, IEnumerable<MapperElement>
    {        
        public MapperElement this[int index]
        {
            get { return BaseGet(index) as MapperElement; }
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
            return new MapperElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }
            return ((MapperElement)element).Name;
        }

        public IEnumerator<MapperElement> GetEnumerator()
        {
            int count = Count;
            for (int i = 0; i < count; i++)
            {
                yield return BaseGet(i) as MapperElement;
            }
        }
    }
}
