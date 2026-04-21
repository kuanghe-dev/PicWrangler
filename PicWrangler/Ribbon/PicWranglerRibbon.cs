using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PicWrangler.Helpers;
using PicWrangler.Models;
using PicWrangler.Services;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace PicWrangler.Ribbon
{
    [ComVisible(true)]
    public class PicWranglerRibbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;

        private readonly PresetStore _presetStore = new PresetStore();
        private readonly ImageInspector _inspector = new ImageInspector();
        private readonly ImageApplicator _applicator = new ImageApplicator();

        private readonly string[] _presetNames = { "Preset 1", "Preset 2", "Preset 3", "Preset 4" };
        private int _selectedPresetIndex = 0;

        private bool _applyCrop     = true;
        private bool _applySize     = true;
        private bool _applyPosition = true;

        private bool _addTitle = true;
        private bool _addNotes = false;

        // -----------------------------------------------------------------
        // IRibbonExtensibility
        // -----------------------------------------------------------------

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PicWrangler.Ribbon.PicWranglerRibbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // -----------------------------------------------------------------
        // Preset dropdown
        // -----------------------------------------------------------------

        public int ddPreset_GetItemCount(IRibbonControl control) => _presetNames.Length;

        public string ddPreset_GetItemLabel(IRibbonControl control, int index) => _presetNames[index];

        public string ddPreset_GetItemID(IRibbonControl control, int index) => $"preset_{index}";

        public int ddPreset_GetSelectedItemIndex(IRibbonControl control) => _selectedPresetIndex;

        public void ddPreset_SelectionChanged(IRibbonControl control, string selectedId, int selectedIndex)
        {
            _selectedPresetIndex = selectedIndex;
        }

        // -----------------------------------------------------------------
        // Checkbox getters / setters
        // -----------------------------------------------------------------

        public bool chkCrop_GetPressed(IRibbonControl control)         => _applyCrop;
        public bool chkSize_GetPressed(IRibbonControl control)         => _applySize;
        public bool chkPosition_GetPressed(IRibbonControl control)     => _applyPosition;
        public bool chkAddTitle_GetPressed(IRibbonControl control) => _addTitle;
        public bool chkAddNotes_GetPressed(IRibbonControl control) => _addNotes;

        public void chkCrop_Toggle(IRibbonControl control, bool pressed)     => _applyCrop     = pressed;
        public void chkSize_Toggle(IRibbonControl control, bool pressed)     => _applySize     = pressed;
        public void chkPosition_Toggle(IRibbonControl control, bool pressed) => _applyPosition = pressed;
        public void chkAddTitle_Toggle(IRibbonControl control, bool pressed) => _addTitle = pressed;
        public void chkAddNotes_Toggle(IRibbonControl control, bool pressed) => _addNotes = pressed;

        // -----------------------------------------------------------------
        // Set Preset
        // -----------------------------------------------------------------

        public void btnSetPreset_Click(IRibbonControl control)
        {
            var app   = Globals.ThisAddIn.Application;
            var shape = SelectionHelper.GetSelectedPicture(app);
            if (shape == null) return;

            string presetName = _presetNames[_selectedPresetIndex];
            var preset = _inspector.CapturePreset(shape, presetName);
            _presetStore.Save(presetName, preset);

            MessageBox.Show($"Saved settings to \"{presetName}\".", "PicWrangler",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // -----------------------------------------------------------------
        // Apply
        // -----------------------------------------------------------------

        public void btnApply_Click(IRibbonControl control)
        {
            if (!AnyOptionChecked()) return;

            var app   = Globals.ThisAddIn.Application;
            var shape = SelectionHelper.GetSelectedPicture(app);
            if (shape == null) return;

            string presetName = _presetNames[_selectedPresetIndex];
            var preset = _presetStore.Load(presetName);

            if (preset == null)
            {
                MessageBox.Show($"\"{presetName}\" has not been configured yet.", "PicWrangler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _applicator.Apply(shape, preset, _applyCrop, _applySize, _applyPosition);
        }

        // -----------------------------------------------------------------
        // Apply to Slides
        // -----------------------------------------------------------------

        public void btnApplyToSlides_Click(IRibbonControl control)
        {
            if (!AnyOptionChecked()) return;

            var app = Globals.ThisAddIn.Application;

            string presetName = _presetNames[_selectedPresetIndex];
            var preset = _presetStore.Load(presetName);

            if (preset == null)
            {
                MessageBox.Show($"\"{presetName}\" has not been configured yet.", "PicWrangler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PPT.SlideRange slideRange = GetSelectedSlides(app);
            if (slideRange == null) return;

            int imageCount = 0;
            int slideCount = slideRange.Count;

            foreach (PPT.Slide slide in slideRange)
            {
                foreach (PPT.Shape shape in slide.Shapes)
                {
                    if (!SelectionHelper.IsPictureShape(shape)) continue;
                    _applicator.Apply(shape, preset, _applyCrop, _applySize, _applyPosition);
                    imageCount++;
                }
            }

            MessageBox.Show(
                $"Applied \"{presetName}\" to {imageCount} image{(imageCount != 1 ? "s" : "")} across {slideCount} slide{(slideCount != 1 ? "s" : "")}.",
                "PicWrangler", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // -----------------------------------------------------------------
        // Bulk Insert (stub)
        // -----------------------------------------------------------------

        public void btnBulkInsert_Click(IRibbonControl control)
        {
            var app          = Globals.ThisAddIn.Application;
            var presentation = app.ActivePresentation;

            using (var dialog = new OpenFileDialog())
            {
                dialog.Title     = "Select Images to Insert";
                dialog.Filter    = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tiff;*.tif|All Files|*.*";
                dialog.Multiselect = true;

                if (dialog.ShowDialog() != DialogResult.OK || dialog.FileNames.Length == 0)
                    return;

                float slideWidth  = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;

                int insertAfter;
                try   { insertAfter = app.ActiveWindow.View.Slide.SlideIndex; }
                catch { insertAfter = presentation.Slides.Count; }

                foreach (string filePath in dialog.FileNames)
                {
                    var layout = _addTitle
                        ? PPT.PpSlideLayout.ppLayoutTitleOnly
                        : PPT.PpSlideLayout.ppLayoutBlank;

                    var slide = presentation.Slides.Add(++insertAfter, layout);

                    float imgAreaTop    = 0f;
                    float imgAreaHeight = slideHeight;

                    if (_addTitle)
                    {
                        var titleShape = slide.Shapes.Title;
                        titleShape.TextFrame.TextRange.Text = Path.GetFileNameWithoutExtension(filePath);
                        imgAreaTop    = titleShape.Top + titleShape.Height;
                        imgAreaHeight = slideHeight - imgAreaTop;
                    }

                    float aspectRatio;
                    using (var bmp = System.Drawing.Image.FromFile(filePath))
                        aspectRatio = (float)bmp.Width / bmp.Height;

                    float finalWidth, finalHeight;
                    if (slideWidth / imgAreaHeight > aspectRatio)
                    {
                        finalHeight = imgAreaHeight;
                        finalWidth  = finalHeight * aspectRatio;
                    }
                    else
                    {
                        finalWidth  = slideWidth;
                        finalHeight = finalWidth / aspectRatio;
                    }

                    float imgLeft = (slideWidth - finalWidth)    / 2f;
                    float imgTop  = imgAreaTop + (imgAreaHeight - finalHeight) / 2f;

                    slide.Shapes.AddPicture(filePath,
                        MsoTriState.msoFalse, MsoTriState.msoCTrue,
                        imgLeft, imgTop, finalWidth, finalHeight);

                    if (_addNotes)
                        slide.NotesPage.Shapes[2].TextFrame.TextRange.Text =
                            $"Original path: {filePath}";
                }

                int count = dialog.FileNames.Length;
                MessageBox.Show(
                    $"Inserted {count} image{(count != 1 ? "s" : "")} into the presentation.",
                    "PicWrangler", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // -----------------------------------------------------------------
        // Help
        // -----------------------------------------------------------------

        public void btnHelp_Click(IRibbonControl control)
        {
            using (var dialog = new HelpDialog())
                dialog.ShowDialog();
        }

        // -----------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------

        private bool AnyOptionChecked()
        {
            if (!_applyCrop && !_applySize && !_applyPosition)
            {
                MessageBox.Show("Please check at least one of Crop, Size, or Position.", "PicWrangler",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private PPT.SlideRange GetSelectedSlides(PPT.Application app)
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

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            using (var reader = new System.IO.StreamReader(stream))
                return reader.ReadToEnd();
        }
    }
}
