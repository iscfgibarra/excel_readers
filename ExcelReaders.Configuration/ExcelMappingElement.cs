using System.Configuration;

namespace ExcelReaders.Configuration
{
    public class ExcelMappingElement: ConfigurationElement
    {
        private const string NameKey = "Name";


        [ConfigurationProperty(NameKey)]
        public string Name
        {
            get
            {
                return this[NameKey].ToString();
            }

            set { this[NameKey] = value; }
        }

        [ConfigurationProperty("Mappers")]
        public MapperCollection Mappers => (MapperCollection)this["Mappers"];

        [ConfigurationProperty("Sheets")]
        public SheetCollection Sheets => (SheetCollection)this["Sheets"];


    }
}
