using System.Windows.Forms;
using Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PicWrangler.Helpers
{
    public static class SelectionHelper
    {
        public static PPT.Shape GetSelectedPicture(PPT.Application app)
        {
            var selection = app.ActiveWindow.Selection;

            if (selection.Type != PPT.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select an image first.", "PicWrangler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            PPT.Shape shape = selection.ShapeRange[1];

            if (!IsPictureShape(shape))
            {
                MessageBox.Show("Selected object is not a picture.", "PicWrangler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            return shape;
        }

        public static bool IsPictureShape(PPT.Shape shape)
        {
            if (shape.Type != MsoShapeType.msoPicture &&
                shape.Type != MsoShapeType.msoLinkedPicture &&
                shape.Type != MsoShapeType.msoPlaceholder)
                return false;

            try { _ = shape.PictureFormat.CropLeft; return true; }
            catch { return false; }
        }

        public static PPT.SlideRange GetSelectedSlides(PPT.Application app)
        {
            var selection = app.ActiveWindow.Selection;

            if (selection.Type == PPT.PpSelectionType.ppSelectionSlides)
                return selection.SlideRange;

            MessageBox.Show(
                "No slides selected in the slide panel. Applying to the active slide only.",
                "PicWrangler", MessageBoxButtons.OK, MessageBoxIcon.Information);

            int activeIndex = app.ActiveWindow.View.Slide.SlideIndex;
            return app.ActivePresentation.Slides.Range(activeIndex);
        }
    }
}
