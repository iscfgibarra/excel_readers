using ExcelReaders.Console.Helpers;
using ExcelReaders.Console.Results;
using ExcelReaders.Core;

namespace ExcelReaders.Console.Readers
{    
    public class AduanaExcelReader : BaseExcelReader<AduanaResult>
    {        
        public AduanaExcelReader() : base(
             ConfigurationHelper.ExcelSourceDirectory
            ,  ConfigurationHelper.ExcelSourceCatCfdiFileName
            , ConfigurationHelper.AduanaExcelMappingName)
        {           
            
        }

        public override void CalculateFields(ref AduanaResult obj)
        {            
            obj.Key = obj.Clave;
        }
    }
}
