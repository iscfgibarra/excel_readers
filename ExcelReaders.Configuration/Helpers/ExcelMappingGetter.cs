using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace ExcelReaders.Configuration.Helpers
{
    public class ExcelMappingGetter
    {
        public List<SheetConfig> Sheets { get; set; }

        public List<MapperConfig> Mappers { get; set; }

        public ExcelMappingGetter(string excelMappingName)
        {
            Sheets = new List<SheetConfig>();
            Mappers = new List<MapperConfig>();

            var section = GetExcelMappingSection();
            var excelMapping = section.ExcelMappings.FirstOrDefault(e => e.Name == excelMappingName);

            if (excelMapping == null)
            {
                throw new ConfigurationErrorsException("No se ha encontrado ningún ExcelMapping con ese nombre.");
            }

            GetSheets(excelMapping);
            GetMappers(excelMapping);

        }

        private ExcelMappingSection GetExcelMappingSection()
        {
            var excelMappingSection = (ExcelMappingSection)ConfigurationManager.GetSection("ExcelMappingSection");

            if (excelMappingSection == null)
            {
                throw new ConfigurationErrorsException("No se ha encontrado la sección ExcelMappingSection en la configuración");
            }

            return excelMappingSection;
        }

        private void GetSheets(ExcelMappingElement excelMappingElement)
        {
            foreach (var sheetElement in excelMappingElement.Sheets)
            {
                Sheets.Add(new SheetConfig
                {
                    NoSheet = sheetElement.NoSheet,
                    SheetName = sheetElement.SheetName,
                    RowNumberStartData = sheetElement.RowNumberStartData,
                    RowNumberStopData = string.IsNullOrEmpty(sheetElement.RowNumberStopData) ? -1 : int.Parse(sheetElement.RowNumberStopData),
                    Map = sheetElement.Map
                });
            }
        }

        private void GetMappers(ExcelMappingElement excelMappingElement)
        {
            foreach (var mapperElement in excelMappingElement.Mappers)
            {
                var mapper = new MapperConfig(mapperElement.Name);

                foreach (var mapElement in mapperElement.MapElements)
                {
                        mapper.Maps.Add(new MapConfig
                        {
                            NoColumn = string.IsNullOrEmpty(mapElement.NoColumn) ? -1: int.Parse(mapElement.NoColumn),
                            Attribute = mapElement.Attribute,
                            Ignore = mapElement.Ignore,
                            Default = mapElement.Default,
                            Format = mapElement.Format
                        });
                }

                Mappers.Add(mapper);
            }
        }
    }
}
