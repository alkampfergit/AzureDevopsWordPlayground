using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.UI.Support
{
    public class StatePersister
    {
        private Dictionary<String, Object> _config;
        private FileInfo _configFile;

        public static StatePersister Instance = new StatePersister();

        private StatePersister()
        {
            var fileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "wordexporter.ui.json";
             _configFile = new FileInfo(fileName);
            if (!_configFile.Exists)
            {
                _config = new Dictionary<string, object>();
            }
            else
            {
                _config = JsonConvert.DeserializeObject<Dictionary<String, Object>>(File.ReadAllText(_configFile.FullName));
            }
        }

        public void Persist()
        {
            File.WriteAllText(_configFile.FullName, JsonConvert.SerializeObject(_config, Formatting.Indented));
        }

        public void Save<T>(String key, T value)
        {
            _config[key] = value;
        }

        public T Load<T>(String key)
        {
            if (_config.TryGetValue(key, out var obj))
                return (T)obj;

            return default(T);
        }
    }
}
