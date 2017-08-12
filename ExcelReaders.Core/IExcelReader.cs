using System.Collections.Concurrent;

namespace ExcelReaders.Core
{
    public interface IExcelReader<T>
    where T : class 
    {
        string XlsFilename { get; set; }

        bool LoadData();        
        ConcurrentBag<T> GetDataList { get;  }

        bool ClearData();        
    }

}
