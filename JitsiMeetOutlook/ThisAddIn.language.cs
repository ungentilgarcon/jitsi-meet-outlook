using System;
using System.Text;
using System.Text.Json;
using System.Collections.Generic;
using System.Threading;

namespace JitsiMeetOutlook
{
    public partial class ThisAddIn
    {
        private readonly Lazy<JsonElement> _lazyLanguageJsonRoot = new Lazy<JsonElement>(() =>
        {
            try
            {
                byte[] jsonFile = findJson(Properties.Settings.Default.language);
                if (jsonFile != null)
                {
                    string jsonString = Encoding.UTF8.GetString(jsonFile);
                    JsonDocument document = JsonDocument.Parse(jsonString);
                    return document.RootElement;
                }
            }
            catch { }
            return default;
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        private static byte[] findJson(string language)
        {
            Dictionary<string, byte[]> languages = new Dictionary<string, byte[]>();
            languages.Add("en", Resources.languages.en);
            languages.Add("de", Resources.languages.de);
            languages.Add("fr", Resources.languages.fr);
            languages.Add("ru", Resources.languages.ru);
            languages.Add("es", Resources.languages.es);
            languages.Add("cz", Resources.languages.cz);

            if (languages.ContainsKey(language))
            {
                return languages[language];
            }
            else
            {
                return null;
            }
        }

        public JsonElement getLanguageJsonRoot()
        {
            return _lazyLanguageJsonRoot.Value;
        }

        public string getElementTranslation(string jsonGroup, string jsonElement)
        {
            var root = _lazyLanguageJsonRoot.Value;
            return root.GetProperty(jsonGroup).GetProperty(jsonElement).GetString();
        }

    }

}
