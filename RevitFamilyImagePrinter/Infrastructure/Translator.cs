using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace RevitFamilyImagePrinter.Infrastructure
{
	/// <summary>
	/// Class to provide multilingualism in entire add in
	/// </summary>
	public class Translator
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
            buttonLink_ToolTip,
	        windowPrintSettingsTitle,
			labelSize_Text,
	        labelScale_Text,
	        labelZoom_Text,
	        labelResolution_Text,
	        labelDetailLevel_Text,
	        labelFormat_Text,
			labelAspectRatio_Text,
	        labelParameters_Text,
			buttonApply_Text,
	        buttonPrint_Text,
			textBlockResolutionWebLow_Text,
			textBlockResolutionWebHigh_Text,
			textBlockResolutionPrintLow_Text,
			textBlockResolutionPrintHigh_Text,
			textBlockDetailLevelCoarse_Text,
			textBlockDetailLevelMedium_Text,
			textBlockDetailLevelFine_Text,
	        errorMessageTitle,
	        errorMessageValuesCorrection,
	        errorMessageValuesRetrieving,
	        errorMessageViewUpdating,
	        errorMessageValuesSaving,
	        errorMessageValuesLoading,
			textBlockProcess_Text,
	        buttonCancel_Text,
	        errorMessageViewCorrecting,
	        errorMessageViewPrinting,
	        errorMessage2dFolderPrinting,
	        errorMessage2dFolderPrintingCycle,
	        errorMessage3dFolderPrintingCycle,
	        errorMessage3dFolderPrinting,
	        errorMessageNoProjectsFolder,
	        errorMessageFamiliesRetrieving,
	        errorMessageProjectProcessing,
	        errorMessageFamilyRemoving,
	        errorMessageFamilyLoading,
	        errorMessageInstanceRemoving,
	        errorMessageInstanceInserting,
	        errorMessageTypeRemoving,
			warningMessageTitle,
	        warningMessageProjectsAmount,
			folderDialogFromTitle,
	        folderDialogToTitle,
	        windowProgressTitle,
	        textBlockProcessPrinting,
	        textBlockProcessLoadingFamilies,
	        textBlockProcessProjectCreated,
			textBlockProcessCreatingProjects,
	        textBlockProcessPreparingPrinting,
		}

        public Translator(string language)
        {
            Dictionary = GetDictionary(language);
        }

        public string GetValue(Keys key)
        {
            var value = string.Empty;
            string _key = Enum.GetName(typeof(Keys), key);
            if (!string.IsNullOrEmpty(_key))
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

        private string ReadXML(string fileName)
        {
            string result = string.Empty;
            Assembly assembly = Assembly.GetExecutingAssembly();
	        var defaultLanguageFile = assembly.GetName().Name + ".Languages.English_USA.xml";
			var languageFile = assembly.GetName().Name + ".Languages." + fileName + ".xml";
	        if (!assembly.GetManifestResourceNames().Contains(languageFile))
		        languageFile = defaultLanguageFile;
			using (Stream stream = assembly.GetManifestResourceStream(languageFile))
			{
				using (StreamReader sr = new StreamReader(stream))
				{
					result = sr.ReadToEnd();
				}
			}
            return result;
        }

	    private IDictionary<string, string> XmlToDictionary(string data)
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
