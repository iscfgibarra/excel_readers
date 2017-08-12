using System;
using System.Collections.Generic;
using System.Configuration;

namespace ExcelReaders.Configuration
{
    

    [ConfigurationCollection(typeof(SheetElement), AddItemName = "Sheet")]
    public class SheetCollection : ConfigurationElementCollection, IEnumerable<SheetElement>
    {
        public SheetElement this[int index]
        {
            get { return BaseGet(index) as SheetElement; }
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
            return new SheetElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }
            return ((SheetElement) element).NoSheet;
        }

        public IEnumerator<SheetElement> GetEnumerator()
        {
            int count = Count;
            for (int i = 0; i < count; i++)
            {
                yield return BaseGet(i) as SheetElement;
            }
        }
    }
}
