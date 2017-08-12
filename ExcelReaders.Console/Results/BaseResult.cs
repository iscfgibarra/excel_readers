using System;

namespace ExcelReaders.Console.Results
{
    public class BaseResult
    {
        //Esta campo se llena por medio de CalculateFields del ExcelReader
        //es usado en la búsqueda.
        public string Key { get; set; }
        
        public DateTime? FechaInicioVigencia { get; set; }

        public DateTime? FechaFinVigencia { get; set; }

        public string Version { get; set; }
        
    }
}
