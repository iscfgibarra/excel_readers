using System.Collections.Generic;

namespace ExcelReaders.Configuration.Helpers
{
    public class MapperConfig
    {
        public string Name { get; set; }

        public List<MapConfig> Maps { get; set; }

        public MapperConfig(string name)
        {
            Name = name;
            Maps = new List<MapConfig>();
        }
    }
}
