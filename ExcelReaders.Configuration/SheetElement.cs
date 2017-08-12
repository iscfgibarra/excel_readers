using System.Configuration;

namespace ExcelReaders.Configuration
{
    public class SheetElement : ConfigurationElement
    {
        private const string NoSheetKey = "NoSheet";
        
        private const string SheetNameKey = "SheetName";

        private const string RowNumberStartDataKey = "RowNumberStartData";

        private const string RowNumberStopDataKey = "RowNumberStopData";
        
        private const string MapKey = "Map";

        [ConfigurationProperty(NoSheetKey)]
        public int NoSheet
        {
            get { return (int)this[NoSheetKey]; }
            set { this[NoSheetKey] = value; }
        }


        [ConfigurationProperty(SheetNameKey)]
        public string SheetName
        {
            get { return (string) this[SheetNameKey]; }
            set { this[SheetNameKey] = value; }
        }

        [ConfigurationProperty(RowNumberStartDataKey)]
        public int RowNumberStartData
        {
            get
            {
                return int.Parse(this[RowNumberStartDataKey].ToString());
                
            }
            set
            {
                this[RowNumberStartDataKey] = value;                 
            }
        }

        [ConfigurationProperty(RowNumberStopDataKey)]
        public string RowNumberStopData
        {
            get
            {
                return  this[RowNumberStopDataKey].ToString();

            }
            set
            {
                this[RowNumberStopDataKey] = value;
            }
        }


        [ConfigurationProperty(MapKey)]
        public string Map
        {
            get { return (string) this[MapKey]; }
            set
            {
                this[MapKey] = value;
            }
        }



    }
}
