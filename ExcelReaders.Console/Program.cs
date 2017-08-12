using ExcelReaders.Console.Readers;

namespace ExcelReaders.Console
{
    class Program
    {
        static void Main(string[] args)
        {
             var reader = new AduanaExcelReader();
            reader.LoadData();
            
            foreach (var item in reader.GetDataList)
            {
                System.Console.WriteLine(item);
            }


            System.Console.Read();
        }
    }
}
