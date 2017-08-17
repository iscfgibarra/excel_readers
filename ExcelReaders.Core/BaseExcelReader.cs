using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using ExcelReaders.Configuration.Helpers;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace ExcelReaders.Core
{
   
    public class BaseExcelReader<T> : IExcelReader<T> where T : class, new()
    {       
        private static ConcurrentBag<T> _rowDataList;

        protected List<SheetConfig> SheetsConfig;

        protected List<MapperConfig> MappersConfig;

        private DateTimeTypeConvertion _latestTypeConvertion;
        
        private string _xlsfullPath { get; set; }
        
        public ConcurrentBag<T> GetDataList => _rowDataList;

        public static string ExcelSourceDirectory { get; set; }

        public string XlsFilename { get; set; }
        
        

        public string XlsFullPath
        {
            get
            {
                if (!string.IsNullOrEmpty(_xlsfullPath)) return _xlsfullPath;
                
                var xlsDirectory = XlsDirectory;
                var xlsFilename = XlsFilename;

                if (!string.IsNullOrEmpty(xlsFilename))
                {
                    if (!string.IsNullOrEmpty(xlsDirectory))
                    {
                        return $"{xlsDirectory}{xlsFilename}";
                    }
                }

                throw new Exception("No se ha especificado en nombre del archivo de Excel o el directorio del ensamblado esta vacio.");
            }
        }


        protected static string XlsDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return $"{Path.Combine(Path.GetDirectoryName(path), ExcelSourceDirectory)}\\";   ;
            }
        }


        public BaseExcelReader(string excelSourceDirectory, string xlsFilename, string excelMappingName)
        {
            XlsFilename = xlsFilename;                                    
            GetConfiguration(excelMappingName);
            ExcelSourceDirectory = excelSourceDirectory;
        }

        public BaseExcelReader(string pathExcelFile, string excelMappingName)
        {
            _xlsfullPath = pathExcelFile;
            GetConfiguration(excelMappingName);
        }

        /// <summary>
        /// Trae la configuración de la sección especificada.
        /// </summary>
        /// <param name="excelMappingName">Nombre de la sección</param>
        private void GetConfiguration(string excelMappingName)
        {
            var config = new ExcelMappingGetter(excelMappingName);
            SheetsConfig = config.Sheets;
            MappersConfig = config.Mappers;         
        }
        
        /// <summary>
        /// Carga los datos desde la hoja de las hojas de Excel especificadas y las 
        /// almacena en la variable estática _rowDataList.
        /// </summary>
        /// <returns>True si la lista tiene elementos.</returns>
        public bool LoadData()
        {
            if (_rowDataList == null)
            {
                _rowDataList = new ConcurrentBag<T>();
                
                using (var fs = File.OpenRead(XlsFullPath))
                {
                    var workBook = new HSSFWorkbook(fs);
                    var mapperConfig = MappersConfig.FirstOrDefault();
                    bool hasMoreMappers = MappersConfig.Count > 1;

                    foreach (var sheetConfig in SheetsConfig)
                    {                    
                        var sheet = workBook.GetSheet(sheetConfig.SheetName);
                        
                        if (hasMoreMappers)
                        {
                            mapperConfig = MappersConfig.FirstOrDefault(m => m.Name == sheetConfig.Map);
                        }

                        for (int rowIndex = sheetConfig.RowNumberStartData - 1; 
                                    rowIndex < sheetConfig.RowNumberStopData; 
                                            rowIndex++)
                        {
                            var row = sheet.GetRow(rowIndex);                            
                            if(row == null) continue;

                            var obj = FillObject(mapperConfig, row);

                            //Llenar propiedades derivadas de otras
                            CalculateFields(ref obj);
                            
                            _rowDataList.Add(obj);
                        }
                    }
                }
            }
            
            var lista = _rowDataList;            
            return lista.Count > 0;
        }

        private T FillObject(MapperConfig mapperConfig, IRow row)
        {
            var obj = new T();

            foreach (var map in mapperConfig.Maps)
            {
                var propInfo = obj.GetType().GetProperty(map.Attribute);

                if (!string.IsNullOrEmpty(map.Default))
                {
                    propInfo?.SetValue(obj, ConvertToAttributeType(propInfo, map.Default));
                }

                if (map.Ignore) continue;

                //Solo se formatea si la propiedad es String y el formato no esta vació
                if (!string.IsNullOrEmpty(map.Format) && IsString(propInfo))
                {
                    var cell = row.GetCell(map.NoColumn);
                    propInfo?.SetValue(obj, GetValueFormatted(cell, map.Format));
                }
                else
                {
                    var cell = row.GetCell(map.NoColumn);
                    propInfo?.SetValue(obj, ConvertToAttributeType(propInfo, cell));
                }
            }
            return obj;
        }

        /// <summary>
        /// Esta funcion se utiliza cuando hay campos calculados que dependen de valores del objeto.       
        /// </summary>
        /// <param name="obj">El objeto donde se van a modificar los valores</param>
        public virtual void CalculateFields(ref T obj)
        {
            
        }

        /// <summary>
        /// Convierte el valor de la celda al formato proporcionado
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        private string GetValueFormatted(ICell cell, string format)
        {
            double doubleValue;
            
            if (cell.CellType == CellType.Blank) return string.Empty;
            
            string propValue;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    propValue = cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    break;
                case CellType.String:
                    propValue = cell.StringCellValue;
                    break;              
                default:
                    propValue = string.Empty;
                    break;
            }

            double.TryParse(propValue, out doubleValue);
            return doubleValue.ToString(format);
        }
       
        /// <summary>
        /// Permite convertir el valor de la celda en el valor apropiado de 
        /// acuerdo al tipo del atributo.
        /// <remarks>Hay que destacar que el value algunas veces se auto-asigna como String en lugar
        /// de ICell. </remarks>
        /// </summary>
        /// <param name="propertyInfo">Información de la propiedad</param>
        /// <param name="value">Valor de la celda</param>
        /// <returns></returns>
        private object  ConvertToAttributeType (PropertyInfo propertyInfo, object value)
        {
           
            if (value == null) return null;

            try
            {
                //Se pregunta si es un string, algunas veces esta conversion falla
                //sobre todo cuando el campo es de tipo Date, sin embargo no afecta el resultado 
                //de las conversiones.
                if (IsString(propertyInfo)) return value?.ToString();
            }
            catch (Exception e)
            {
                Debug.WriteLine("La celda no tiene formato de texto, se forzara su conversión." +
                    "\n PropertyInfo:" + propertyInfo + "\n Value:" + value + "\n Type:" +
                    typeof(T).FullName + "\n Message:" + e.Message + "\n Stacktrace:"  + e.StackTrace  );                
            }

            object dateTimeConvertion;
            if (ConvertFromDateTime(propertyInfo, value, out dateTimeConvertion)) return dateTimeConvertion;

            return ConvertFromNumericTypes(propertyInfo, value);
        }

        private object ConvertFromNumericTypes(PropertyInfo propertyInfo, object value)
        {
            var valor = (ICell) value;
            var converter = TypeDescriptor.GetConverter(propertyInfo.PropertyType);

            switch (valor.CellType)
            {
                case CellType.String:
                    var stringToConverter = valor.StringCellValue ?? "0";
                    return converter.ConvertFrom(stringToConverter.Replace(",", ""));
                case CellType.Numeric:
                    return converter.ConvertFrom(valor.NumericCellValue.ToString());
                case CellType.Blank:
                case CellType.Unknown:
                case CellType.Formula:
                case CellType.Boolean:
                case CellType.Error:
                    return null;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private bool ConvertFromDateTime(PropertyInfo propertyInfo, object value, out object dateTimeConvertion)
        {
            dateTimeConvertion = null;

            if (propertyInfo.PropertyType == typeof(DateTime?))
            {
                try
                {
                    if (value.ToString() == "null") return true;
                    if (string.IsNullOrEmpty(value.ToString())) return true;                    
                }
                catch (Exception e)
                {
                    Debug.WriteLine(typeof(T).FullName + e.Message + e.StackTrace + e.Source);
                }

                dateTimeConvertion = DateTimeConvertion(value);
                return true;
            }

            if (propertyInfo.PropertyType == typeof(DateTime))
            {
                dateTimeConvertion = DateTimeConvertion(value);
                return true;
            }
            
            return false;
        }

        private bool IsString(PropertyInfo propertyInfo)
        {
            return propertyInfo.PropertyType == typeof(string) || propertyInfo.PropertyType == typeof(String);
        }
        
        
        /// <summary>
        /// Convierte el valor proporcionado en fecha, el intento de conversión directo
        /// solo parseando el string del valor permite ganar velocidad en la conversión,
        /// cuando no es posible hace la conversión usando la excepción, la variable
        /// _latestTypeConvertion permite que el método "recuerde" el ultimo tipo 
        /// de conversión hecha por el método y así ganar tiempo en el proceso 
        /// evitando las excepciones innecesarias.
        /// <remarks>Este método cubre la deficiencia de las hojas de Excel que no tienen
        /// especificado el formato de manera adecuada 
        /// (ejem. que tienen una fecha en formato de cadena).</remarks>
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private DateTime? DateTimeConvertion(object value)
        {
            DateTime? retval;
                       
            try
            {
                if (_latestTypeConvertion != 0)
                {
                    switch (_latestTypeConvertion)
                    {
                        case DateTimeTypeConvertion.WithoutCulture:
                            retval = DateTime.Parse(value.ToString());
                            return retval;
                        case DateTimeTypeConvertion.UsaCulture:
                            retval = DateTime.Parse(value.ToString(), new CultureInfo("en-US", true));
                            return retval;
                        case DateTimeTypeConvertion.NpoiDirect:
                            var valor = (ICell)value;
                            retval = valor.DateCellValue.Date;
                            return retval;
                    }
                }
            }
            catch
            {                   
                _latestTypeConvertion = DateTimeTypeConvertion.WithoutCulture;
            }
            
            
            //Si falla lo anterior hace la conversión basada en excepciones
            try
            {
                _latestTypeConvertion = DateTimeTypeConvertion.WithoutCulture;
                retval = DateTime.Parse(value.ToString());
                return retval;
            }
            catch
            {
                try
                {
                    _latestTypeConvertion = DateTimeTypeConvertion.UsaCulture;
                    retval = DateTime.Parse(value.ToString(), new CultureInfo("en-US", true));
                    return retval;
                }
                catch
                {
                    _latestTypeConvertion = DateTimeTypeConvertion.NpoiDirect;
                    var valor = (ICell)value;
                    retval = valor.DateCellValue.Date;
                    return retval;
                }
            }
        }

        /// <summary>
        /// Limpia los datos almacenados en el ExcelReader
        /// </summary>
        /// <returns>True si la colección no tiene elementos.</returns>
        public bool ClearData()
        {
            _rowDataList = new ConcurrentBag<T>();

            return _rowDataList.Count == 0;
        }
    }
}
