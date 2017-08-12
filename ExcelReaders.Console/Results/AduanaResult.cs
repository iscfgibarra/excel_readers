namespace ExcelReaders.Console.Results
{
    public class AduanaResult : BaseResult
    {
        public string Clave { get; set; }
        
        public override string ToString()
        {
            return $"{Clave}-{FechaInicioVigencia}-{Version}";
        }
    }
}
