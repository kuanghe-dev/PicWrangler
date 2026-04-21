using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using PicWrangler.Models;

namespace PicWrangler.Services
{
    public class PresetStore
    {
        private readonly string _storePath;

        private static readonly string[] DefaultPresetNames = { "Preset 1", "Preset 2", "Preset 3", "Preset 4" };

        public PresetStore() : this(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PicWrangler")) { }

        public PresetStore(string storageDirectory)
        {
            _storePath = Path.Combine(storageDirectory, "presets.json");
        }

        public Dictionary<string, Preset> LoadAll()
        {
            if (!File.Exists(_storePath))
                return CreateDefaults();

            try
            {
                string json = File.ReadAllText(_storePath);
                var presets = JsonConvert.DeserializeObject<Dictionary<string, Preset>>(json);
                foreach (string name in DefaultPresetNames)
                {
                    if (!presets.ContainsKey(name))
                        presets[name] = null;
                }
                return presets;
            }
            catch
            {
                return CreateDefaults();
            }
        }

        public Preset Load(string presetName)
        {
            var all = LoadAll();
            return all.TryGetValue(presetName, out var preset) ? preset : null;
        }

        public void Save(string presetName, Preset preset)
        {
            var all = LoadAll();
            preset.Name = presetName;
            all[presetName] = preset;

            Directory.CreateDirectory(Path.GetDirectoryName(_storePath));
            File.WriteAllText(_storePath, JsonConvert.SerializeObject(all, Formatting.Indented));
        }

        private Dictionary<string, Preset> CreateDefaults()
        {
            var defaults = new Dictionary<string, Preset>();
            foreach (string name in DefaultPresetNames)
                defaults[name] = null;
            return defaults;
        }
    }
}
