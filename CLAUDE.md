# CLAUDE.md — PicWrangler: Batch Manipulate Pictures (PowerPoint Add-In)

## Project Overview

PicWrangler is a PowerPoint COM Add-In (VSTO) that lets users save and replay image manipulation presets — crop, resize, and reposition — across one or many slides, and bulk-insert images into a presentation. The UI lives in a custom Ribbon tab called **ADD-INS**, grouped under **"PicWrangler: Batch Manipulate Pictures"**.

---

## Technology Stack

- **Language:** C# (.NET Framework 4.7.2+)
- **Framework:** Visual Studio Tools for Office (VSTO)
- **IDE:** Visual Studio 2022 with "Office/SharePoint development" workload
- **Target:** Microsoft PowerPoint 2016/2019/365 (Windows only)
- **Persistence:** JSON file in `%AppData%\PicWrangler\`
- **JSON library:** `Newtonsoft.Json` (NuGet) — `System.Text.Json` is not included in .NET Framework 4.7.2

---

## Architecture

```
PicWrangler/
├── CLAUDE.md
├── PicWrangler.sln
├── PicWrangler/
│   ├── ThisAddIn.cs             # Add-in entry point, lifecycle hooks
│   ├── Ribbon/
│   │   ├── PicWranglerRibbon.cs   # Ribbon XML callbacks (button clicks, dropdown)
│   │   └── PicWranglerRibbon.xml  # Ribbon UI layout definition (EmbeddedResource)
│   ├── Models/
│   │   ├── Preset.cs            # Data model for a single preset
│   │   ├── CropSettings.cs      # Crop offsets (left/right/top/bottom in points)
│   │   ├── SizeSettings.cs      # Width and height in points
│   │   └── PositionSettings.cs  # Left and top position in points
│   ├── Services/
│   │   ├── PresetStore.cs       # Load/save presets to disk (JSON)
│   │   ├── ImageInspector.cs    # Read crop/size/position from a selected shape
│   │   └── ImageApplicator.cs   # Apply crop/size/position from a preset to shape(s)
│   └── Helpers/
│       └── SelectionHelper.cs   # Utilities for getting selected shapes/slides
├── PicWranglerTests/
│   ├── PresetStoreTests.cs
│   ├── ImageInspectorTests.cs
│   └── ImageApplicatorTests.cs
└── assets/
    └── PicWrangler_toolbar.png    # UI mockup reference
```

> **Important:** `PicWranglerRibbon.xml` must have its Build Action set to **Embedded Resource** in the project file. The default for manually added files is `Content`, which causes `GetManifestResourceStream` to return `null` at startup.

---

## Data Models

### `Preset.cs`
```csharp
public class Preset
{
    public string Name { get; set; }          // e.g. "Preset 1"
    public CropSettings Crop { get; set; }    // null if not captured
    public SizeSettings Size { get; set; }    // null if not captured
    public PositionSettings Position { get; set; } // null if not captured
}
```

### `CropSettings.cs`
```csharp
// All values in points (PowerPoint's native unit)
public class CropSettings
{
    public float CropLeft { get; set; }
    public float CropRight { get; set; }
    public float CropTop { get; set; }
    public float CropBottom { get; set; }
}
```

### `SizeSettings.cs`
```csharp
public class SizeSettings
{
    public float Width { get; set; }   // points
    public float Height { get; set; }  // points
}
```

### `PositionSettings.cs`
```csharp
public class PositionSettings
{
    public float Left { get; set; }    // points from slide left edge
    public float Top { get; set; }     // points from slide top edge
}
```

---

## Ribbon UI Specification

### XML Layout (`PicWranglerRibbon.xml`)

The ribbon tab is named **ADD-INS** and contains one group. All action buttons use `size="large"`.

| Control | Type | ID | Label | Default | Notes |
|---|---|---|---|---|---|
| Preset selector | `dropDown` | `ddPreset` | *(none)* | "Preset 1" | Items: "Preset 1"–"Preset 4" |
| Set Preset | `button` | `btnSetPreset` | "Set Preset" | — | Large; saves selected image's settings into chosen preset |
| *(separator)* | | `sep1` | | | |
| Apply | `button` | `btnApply` | "Apply" | — | Large; applies preset to selected image |
| Apply to Slides | `button` | `btnApplyToSlides` | "Apply to Slides" | — | Large; applies preset to all images on selected slides |
| Crop checkbox | `checkBox` | `chkCrop` | "Crop" | checked | Whether to apply crop |
| Size checkbox | `checkBox` | `chkSize` | "Size" | checked | Whether to apply size |
| Position checkbox | `checkBox` | `chkPosition` | "Position" | checked | Whether to apply position |
| *(separator)* | | `sep2` | | | |
| Bulk Insert | `button` | `btnBulkInsert` | "Bulk Insert" | — | Large; opens file dialog to insert images one-per-slide |
| Add Title checkbox | `checkBox` | `chkAddTitle` | "Add Title" | checked | Sets slide title to image filename (no extension) |
| Add Notes checkbox | `checkBox` | `chkAddNotes` | "Add Notes" | unchecked | Adds "Original path: …" to slide notes |
| *(separator)* | | `sep3` | | | |
| Help | `button` | `btnHelp` | "Help" | — | Large; opens help dialog |

---

## Feature Implementation Guide

### 1. Set Preset (`btnSetPreset_Click`)

**Goal:** Read crop, size, and position from the currently selected image and save into the active preset slot.

**Steps:**
1. Get the selected preset name from `ddPreset`.
2. Get the currently selected shape via `Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1]`.
3. Verify the shape can provide picture data (see Shape Type Handling below).
4. Read crop values from `shape.PictureFormat`: `CropLeft`, `CropRight`, `CropTop`, `CropBottom` (all in points).
5. Read size from `shape.Width`, `shape.Height`.
6. Read position from `shape.Left`, `shape.Top`.
7. Construct a `Preset` object and pass to `PresetStore.Save(presetName, preset)`.
8. Show a MessageBox confirming the save.

**Error handling:**
- No shape selected → "Please select an image first."
- Shape type not accepted or `PictureFormat` throws → "Selected object is not a picture."

**Shape type handling (`SelectionHelper.GetSelectedPicture`):**

PowerPoint reports different `MsoShapeType` values depending on how an image was inserted:
- `msoPicture` — embedded picture (Insert > Picture)
- `msoLinkedPicture` — linked picture
- `msoPlaceholder` — image inside a content placeholder

All three are accepted. For `msoPlaceholder`, access to `PictureFormat` is validated with a try/catch since non-picture placeholders (e.g. text boxes) also report this type. The same three-type check is used in the `Apply to Slides` loop.

---

### 2. Apply (`btnApply_Click`)

**Goal:** Apply the active preset to the currently selected image, respecting the Crop/Size/Position checkboxes.

**Steps:**
1. Load the preset via `PresetStore.Load(selectedPresetName)`.
2. Get the selected shape via `SelectionHelper.GetSelectedPicture`.
3. Call `ImageApplicator.Apply(shape, preset, applyCrop, applySize, applyPosition)`.

**`ImageApplicator.Apply` logic — order matters:**
```csharp
// 1. Crop FIRST — setting crop values adjusts the frame dimensions,
//    so size must be set after to guarantee the correct final width/height.
if (applyCrop && preset.Crop != null)
{
    shape.PictureFormat.CropLeft   = preset.Crop.CropLeft;
    shape.PictureFormat.CropRight  = preset.Crop.CropRight;
    shape.PictureFormat.CropTop    = preset.Crop.CropTop;
    shape.PictureFormat.CropBottom = preset.Crop.CropBottom;
}
// 2. Size AFTER crop
if (applySize && preset.Size != null)
{
    shape.LockAspectRatio = MsoTriState.msoFalse;
    shape.Width  = preset.Size.Width;
    shape.Height = preset.Size.Height;
}
// 3. Position last
if (applyPosition && preset.Position != null)
{
    shape.Left = preset.Position.Left;
    shape.Top  = preset.Position.Top;
}
```

**Error handling:**
- Preset is empty → "Preset X has not been configured yet."
- No checkboxes checked → "Please check at least one of Crop, Size, or Position."

---

### 3. Apply to Slides (`btnApplyToSlides_Click`)

**Goal:** Apply the preset to **all picture shapes** on all **selected slides**.

**Steps:**
1. Load the preset.
2. Get selected slides via `app.ActiveWindow.Selection.SlideRange` (falls back to active slide if no slides are selected in the panel).
3. For each slide, iterate `slide.Shapes`. Skip shapes that are not `msoPicture`, `msoLinkedPicture`, or `msoPlaceholder`.
4. Call `ImageApplicator.Apply` for each accepted shape.
5. Report how many images were updated.

---

### 4. Preset Persistence (`PresetStore.cs`)

- Storage path: `%AppData%\PicWrangler\presets.json`
- Format: JSON dictionary keyed by preset name:
  ```json
  {
    "Preset 1": { "Crop": {...}, "Size": {...}, "Position": {...} },
    "Preset 2": null,
    "Preset 3": null,
    "Preset 4": null
  }
  ```
- Initialize with 4 empty presets on first run.
- Load presets on `ThisAddIn_Startup`.

---

### 5. Bulk Insert (`btnBulkInsert_Click`)

**Goal:** Insert one or more image files into the presentation, one image per new slide appended at the end.

**Steps:**
1. Open a multi-select `OpenFileDialog` filtered to common image types (jpg, jpeg, png, gif, bmp, tiff, tif).
2. For each selected file:
   - Add a new slide: `ppLayoutTitleOnly` if Add Title is checked, `ppLayoutBlank` otherwise.
   - If **Add Title** is checked: set `slide.Shapes.Title.TextFrame.TextRange.Text` to the filename without extension (`Path.GetFileNameWithoutExtension`). Use the title shape's bounds to determine the image area (below the title).
   - Load the image with `System.Drawing.Image` to get its pixel dimensions and compute the aspect ratio.
   - Scale the image to fill the available slide area (maintaining aspect ratio) and center it.
   - Insert via `slide.Shapes.AddPicture(filePath, msoFalse, msoCTrue, left, top, width, height)`.
   - If **Add Notes** is checked: set `slide.NotesPage.Shapes[2].TextFrame.TextRange.Text` to `"Original path: {filePath}"`. (`Shapes[2]` is the notes text area; `Shapes[1]` is the slide thumbnail.)
3. Show a completion MessageBox with the count of inserted images.

**Defaults:** Add Title = checked, Add Notes = unchecked.

---

## PowerPoint Object Model — Key APIs

| Task | API |
|---|---|
| Selected shape | `app.ActiveWindow.Selection.ShapeRange[1]` |
| Shape type | `shape.Type` (`MsoShapeType.msoPicture` / `msoLinkedPicture` / `msoPlaceholder`) |
| Crop left/right/top/bottom | `shape.PictureFormat.CropLeft` etc. |
| Width/Height | `shape.Width`, `shape.Height` |
| Position | `shape.Left`, `shape.Top` |
| Lock aspect ratio | `shape.LockAspectRatio = MsoTriState.msoFalse` |
| Shapes on slide | `slide.Shapes` |
| Selected slides | `app.ActiveWindow.Selection.SlideRange` |
| Add slide | `presentation.Slides.Add(index, PPT.PpSlideLayout.ppLayoutTitleOnly)` |
| Insert picture | `slide.Shapes.AddPicture(path, linkToFile, saveWithDoc, left, top, width, height)` |
| Slide dimensions | `presentation.PageSetup.SlideWidth` / `.SlideHeight` |
| Slide notes | `slide.NotesPage.Shapes[2].TextFrame.TextRange.Text` |
| Slide title shape | `slide.Shapes.Title` |

> **Units:** PowerPoint uses **points** (1 inch = 72 points) for all measurements. `System.Drawing.Image` gives pixel dimensions — use the pixel aspect ratio directly; do not convert to points.

---

## Build & Run Instructions

```bash
# Prerequisites
# - Visual Studio 2022
# - "Office/SharePoint development" workload installed
# - PowerPoint 2016 or later installed

# 1. Open solution
start PicWrangler.sln

# 2. Build
# Build > Build Solution (Ctrl+Shift+B)

# 3. Debug
# F5 — launches PowerPoint with the add-in loaded
# Shift+F5 — stops debugging and unloads the add-in from PowerPoint

# 4. Disable debug add-in without VS
# PowerPoint: File > Options > Add-ins > Manage: COM Add-ins > Go > uncheck PicWrangler

# 5. Distribute (ClickOnce)
# Build > Publish PicWrangler → share the entire publish/ folder
# Recipients run setup.exe (installs VSTO runtime if needed)
```

> **Important:** The `PicWrangler` add-in project must be created via the VS template wizard
> (**New Project → PowerPoint VSTO Add-in**), not by hand-authoring a `.csproj`.
> VSTO requires MSBuild targets, COM references, and manifest plumbing that only the
> template generates correctly. Source files can be added to the template-created project
> afterward.
>
> The `PicWranglerTests` project is a standard **Unit Test Project (.NET Framework)** added
> via **File → Add → New Project**, with a project reference to `PicWrangler` and
> `Newtonsoft.Json` installed via NuGet.

---

## Testing Checklist

- [ ] Set Preset captures correct crop values from a cropped image
- [ ] Set Preset captures correct width/height from a resized image
- [ ] Set Preset captures correct left/top from a repositioned image
- [ ] Set Preset works when the image is inside a content placeholder
- [ ] Set Preset shows error when a text placeholder is selected
- [ ] Apply with all checkboxes ON restores all three properties with correct width
- [ ] Apply with only "Size" checked changes only size, not crop or position
- [ ] Apply to Slides processes all pictures on 3 selected slides
- [ ] Apply to Slides skips text boxes and other non-picture shapes
- [ ] Presets survive closing and reopening PowerPoint (persistence check)
- [ ] Error shown when clicking Apply with no selection
- [ ] Error shown when clicking Set Preset on a non-picture shape
- [ ] Bulk Insert inserts one slide per selected image file
- [ ] Bulk Insert titles each slide with the filename (no extension) when Add Title is checked
- [ ] Bulk Insert image is aspect-ratio-correct and centered on the slide
- [ ] Bulk Insert notes contain the full file path when Add Notes is checked
- [ ] Bulk Insert with Add Title unchecked uses blank slide layout (no title placeholder)

---

## Known PowerPoint API Gotchas

1. **Namespace ambiguity** — `Microsoft.Office.Core` and `Microsoft.Office.Interop.PowerPoint` both define `Shape`; `System.Windows.Forms` and `Microsoft.Office.Interop.PowerPoint` both define `Application`. Always use `using PPT = Microsoft.Office.Interop.PowerPoint;` and qualify as `PPT.Shape`, `PPT.Application`, etc.
2. **Crop before size** — setting `PictureFormat.Crop*` values changes the shape's frame dimensions. Always apply crop first, then set `Width`/`Height` last so the final size matches the preset exactly.
3. **`LockAspectRatio` must be set to `msoFalse`** before setting both Width and Height independently.
4. **`Selection.SlideRange`** only works when slides are selected in the slide panel (Normal or Slide Sorter view). In other views, fall back gracefully.
5. **Picture shape types** — pictures report as `msoPicture`, `msoLinkedPicture`, or `msoPlaceholder` depending on how they were inserted. Check all three; validate `PictureFormat` access for placeholders via try/catch since text placeholders share the same type.
6. **Grouped shapes** — `shape.Type == msoGroup`. Optionally recurse into `shape.GroupItems` to find pictures inside groups.
7. **Notes page shape index** — `slide.NotesPage.Shapes[1]` is the slide thumbnail; `Shapes[2]` is the notes text area. These are 1-based indexes.
