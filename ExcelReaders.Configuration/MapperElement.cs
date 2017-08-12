using System.Configuration;

namespace ExcelReaders.Configuration
{
    public class MapperElement : ConfigurationElement
    {
        private const string NameKey = "Name";

        [ConfigurationProperty(NameKey)]
        public string Name
        {
            get { return (string)this[NameKey]; }
            set { this[NameKey] = value; }
        }


        [ConfigurationProperty("MapElements")]
        public MapElementCollection MapElements => (MapElementCollection)this["MapElements"];
    }
}
