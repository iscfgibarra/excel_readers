using System.Configuration;

namespace ExcelReaders.Configuration
{
    public class ExcelMappingSection: ConfigurationSection
    {
        [ConfigurationProperty("ExcelMappings")]
        public ExcelMappingCollection ExcelMappings => (ExcelMappingCollection)this["ExcelMappings"];
    }
}
