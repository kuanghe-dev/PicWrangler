using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using PicWrangler.Models;
using PicWrangler.Services;

namespace PicWranglerTests
{
    [TestClass]
    public class PresetStoreTests
    {
        private string _testDir;
        private PresetStore _store;

        [TestInitialize]
        public void Setup()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "PicWranglerTests_" + Guid.NewGuid());
            Directory.CreateDirectory(_testDir);

            // Point the store at a temp path via reflection so we don't pollute %AppData%
            _store = new PresetStore(_testDir);
        }

        [TestCleanup]
        public void Teardown()
        {
            if (Directory.Exists(_testDir))
                Directory.Delete(_testDir, true);
        }

        [TestMethod]
        public void LoadAll_ReturnsAllFourSlotsOnFirstRun()
        {
            var presets = _store.LoadAll();
            Assert.AreEqual(4, presets.Count);
            Assert.IsTrue(presets.ContainsKey("Preset 1"));
            Assert.IsTrue(presets.ContainsKey("Preset 4"));
        }

        [TestMethod]
        public void Save_ThenLoad_RoundTrips()
        {
            var preset = new Preset
            {
                Name = "Preset 2",
                Crop = new CropSettings { CropLeft = 10, CropRight = 20, CropTop = 5, CropBottom = 15 },
                Size = new SizeSettings { Width = 200, Height = 150 },
                Position = new PositionSettings { Left = 50, Top = 75 }
            };

            _store.Save("Preset 2", preset);
            var loaded = _store.Load("Preset 2");

            Assert.IsNotNull(loaded);
            Assert.AreEqual(10f, loaded.Crop.CropLeft);
            Assert.AreEqual(200f, loaded.Size.Width);
            Assert.AreEqual(50f, loaded.Position.Left);
        }

        [TestMethod]
        public void Load_UnsetPreset_ReturnsNull()
        {
            var result = _store.Load("Preset 3");
            Assert.IsNull(result);
        }

        [TestMethod]
        public void Save_Persists_AcrossNewInstance()
        {
            var preset = new Preset
            {
                Name = "Preset 1",
                Size = new SizeSettings { Width = 300, Height = 200 }
            };
            _store.Save("Preset 1", preset);

            var freshStore = new PresetStore(_testDir);
            var loaded = freshStore.Load("Preset 1");

            Assert.IsNotNull(loaded);
            Assert.AreEqual(300f, loaded.Size.Width);
        }
    }
}
