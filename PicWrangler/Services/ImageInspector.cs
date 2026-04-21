using PicWrangler.Models;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PicWrangler.Services
{
    public class ImageInspector
    {
        public Preset CapturePreset(PPT.Shape shape, string presetName)
        {
            return new Preset
            {
                Name = presetName,
                Crop = new CropSettings
                {
                    CropLeft   = shape.PictureFormat.CropLeft,
                    CropRight  = shape.PictureFormat.CropRight,
                    CropTop    = shape.PictureFormat.CropTop,
                    CropBottom = shape.PictureFormat.CropBottom
                },
                Size = new SizeSettings
                {
                    Width  = shape.Width,
                    Height = shape.Height
                },
                Position = new PositionSettings
                {
                    Left = shape.Left,
                    Top  = shape.Top
                }
            };
        }
    }
}
