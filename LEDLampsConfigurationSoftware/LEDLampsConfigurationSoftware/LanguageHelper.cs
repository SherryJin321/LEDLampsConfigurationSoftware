using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace LEDLampsConfigurationSoftware
{
    public class LanguageHelper
    {


        public void LoadLanguageFile(string languageFileName)
        {
            Application.Current.Resources.MergedDictionaries[0] = new ResourceDictionary()
            {
                Source = new Uri(languageFileName, UriKind.RelativeOrAbsolute)
            };
        }
    }
}
