using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace WordExporter.UI.Support
{
    public class StatePersister
    {
        private readonly Dictionary<String, Object> _config;
        private readonly FileInfo _configFile;

        public static readonly StatePersister Instance = new StatePersister();

        private StatePersister()
        {
            var fileName = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "wordexporter.ui.json");
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
            {
                return (T)obj;
            }

            return default(T);
        }
    }
}
