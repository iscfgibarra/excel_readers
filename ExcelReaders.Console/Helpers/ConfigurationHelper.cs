namespace ExcelReaders.Console.Helpers
{
    /// <summary>
    /// Permite controlar de manera centralizada los nombres de los directorios, archivos
    /// y secciones de configuración para los ExcelReaders.
    /// </summary>
    public class ConfigurationHelper
    {
        /// <summary>
        /// Directorio en la librería que aloja los archivos de Excel. 
        /// </summary>
        public const string ExcelSourceDirectory = "ExcelSources";
        
        /// <summary>
        ///  Nombre del archivo que servirá de fuente de datos para los Catálogos de CFDI.
        ///  (En este caso es un archivo de contenido que se copia en el directorio de salida de la aplicación).
        /// </summary>
        public const string ExcelSourceCatCfdiFileName = "catCFDI.xls";

        //Estos son los nombres de los ExcelMappings que debe estar contenidos en la
        //sección de configuración.
        public const string AduanaExcelMappingName = "Aduana";
       

  
    }
}
