using System.Configuration;

namespace ExcelReaders.Configuration
{
    public class MapElement : ConfigurationElement
    {
        private const string AttributeKey = "Attribute";
        
        private const string FormatKey = "Format";

        private const string IgnoreKey = "Ignore";

        private const string NoColumnKey = "NoColumn";

        private const string DefaultKey = "Default";

        [ConfigurationProperty(AttributeKey)]
        public string Attribute
        {
            get { return (string) this[AttributeKey]; }

            set { this[AttributeKey] = value; }
        }

        [ConfigurationProperty(FormatKey)]
        public string Format
        {
            get { return (string)this[FormatKey]; }

            set { this[FormatKey] = value; }
        }

        [ConfigurationProperty(NoColumnKey)]
        public string NoColumn
        {
            get { return (string)this[NoColumnKey]; }

            set { this[NoColumnKey] = value; }
        }

        [ConfigurationProperty(IgnoreKey)]
        public bool Ignore
        {
            get { return (bool)this[IgnoreKey]; }

            set { this[IgnoreKey] = value; }
        }

        [ConfigurationProperty(DefaultKey)]
        public string Default
        {
            get { return (string)this[DefaultKey]; }

            set { this[DefaultKey] = value; }
        }
    }
}
