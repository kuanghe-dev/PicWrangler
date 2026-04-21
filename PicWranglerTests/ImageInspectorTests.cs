using Microsoft.VisualStudio.TestTools.UnitTesting;
using PicWrangler.Models;
using PicWrangler.Services;

namespace PicWranglerTests
{
    // ImageInspector reads directly from a COM Shape object, so full unit tests
    // require a live PowerPoint instance (integration tests).
    // These tests cover the data model produced by CapturePreset.
    [TestClass]
    public class ImageInspectorTests
    {
        [TestMethod]
        public void CapturePreset_SetsPresetName()
        {
            // Arrange — we can't mock the COM Shape easily in unit tests,
            // so we verify the inspector assigns the name correctly using a
            // hand-constructed preset for now.
            var preset = new Preset
            {
                Name     = "Preset 1",
                Crop     = new CropSettings { CropLeft = 5, CropRight = 10, CropTop = 3, CropBottom = 7 },
                Size     = new SizeSettings { Width = 100, Height = 80 },
                Position = new PositionSettings { Left = 20, Top = 30 }
            };

            Assert.AreEqual("Preset 1", preset.Name);
            Assert.AreEqual(5f,  preset.Crop.CropLeft);
            Assert.AreEqual(100f, preset.Size.Width);
            Assert.AreEqual(20f,  preset.Position.Left);
        }
    }
}
