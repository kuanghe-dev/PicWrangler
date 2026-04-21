using Microsoft.Office.Core;
using PicWrangler.Models;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PicWrangler.Services
{
    public class ImageApplicator
    {
        public void Apply(PPT.Shape shape, Preset preset, bool applyCrop, bool applySize, bool applyPosition)
        {
            // Crop first — setting crop values changes the frame dimensions,
            // so size must be applied after to guarantee the correct final width/height.
            if (applyCrop && preset.Crop != null)
            {
                shape.PictureFormat.CropLeft   = preset.Crop.CropLeft;
                shape.PictureFormat.CropRight  = preset.Crop.CropRight;
                shape.PictureFormat.CropTop    = preset.Crop.CropTop;
                shape.PictureFormat.CropBottom = preset.Crop.CropBottom;
            }

            if (applySize && preset.Size != null)
            {
                shape.LockAspectRatio = MsoTriState.msoFalse;
                shape.Width  = preset.Size.Width;
                shape.Height = preset.Size.Height;
            }

            if (applyPosition && preset.Position != null)
            {
                shape.Left = preset.Position.Left;
                shape.Top  = preset.Position.Top;
            }
        }
    }
}
