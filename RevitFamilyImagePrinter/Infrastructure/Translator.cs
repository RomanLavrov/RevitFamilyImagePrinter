using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace RevitFamilyImagePrinter.Infrastructure
{
    class Translator
    {
        private IDictionary<string, string> Dictionary { get; set; }

        public enum Keys
        {
            tabName,
            tabTitle,
            buttonPrint2DSingle_ToolTip,
            buttonPrint2DMulti_ToolTip,
            buttonPrint3DSingle_ToolTip,
            buttonPrint3DMulti_ToolTip,
            buttonPrint2DSingle_Name,
            buttonPrint2DMulti_Name,
            buttonPrint3DSingle_Name,
            buttonPrint3DMulti_Name,
            buttonPrintView_Name,
            buttonPrintView_ToolTip,
            buttonLink_ToolTip
        }

        public Translator(string language)
        {
            Dictionary = GetDictionary(language);
        }

        public string GetValue(Keys key)
        {
            var value = string.Empty;
            string _key = Enum.GetName(typeof(Keys), key);
            if (!String.IsNullOrEmpty(_key))
            {
                Dictionary.TryGetValue(_key, out value);
            }

            value = value.Replace("#", Environment.NewLine);
            return value;
        }

        private IDictionary<string, string> GetDictionary(string language)
        {
            IDictionary<string, string> dictionary = new Dictionary<string, string>();
            var xmlData = ReadXML(language);
            dictionary = XmlToDictionary(xmlData);
            return dictionary;
        }

        private string ReadXML(string filename)
        {
            string result = string.Empty;
            Assembly assembly = Assembly.GetExecutingAssembly();
            var languageFile = assembly.GetName().Name + ".Languages." + filename + ".xml";

            using (Stream stream = assembly.GetManifestResourceStream(languageFile))
            {
                using (StreamReader sr = new StreamReader(stream))
                {
                    result = sr.ReadToEnd();
                }
            }
            return result;
        }

        IDictionary<string, string> XmlToDictionary(string data)
        {
            XElement rootElement = XElement.Parse(data);
            Dictionary<string, string> dict = new Dictionary<string, string>();
            foreach (var el in rootElement.Elements())
            {
                dict.Add(el.Name.LocalName, el.Value);
            }

            return dict;
        }
    }
}
