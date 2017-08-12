using System;
using System.Collections.Generic;
using System.Configuration;

namespace ExcelReaders.Configuration
{
    

    [ConfigurationCollection(typeof(ExcelMappingElement), AddItemName = "ExcelMapping")]
    public class ExcelMappingCollection : ConfigurationElementCollection, IEnumerable<ExcelMappingElement>
    {
        public ExcelMappingElement this[int index]
        {
            get { return BaseGet(index) as ExcelMappingElement; }
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
            return new ExcelMappingElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }
            return ((ExcelMappingElement)element).Name;
        }

        public IEnumerator<ExcelMappingElement> GetEnumerator()
        {
            int count = Count;
            for (int i = 0; i < count; i++)
            {
                yield return BaseGet(i) as ExcelMappingElement;
            }
        }
    }
}
