# CLAUDE.md — PicWrangler: Batch Manipulate Pictures (PowerPoint Add-In)

## Project Overview

PicWrangler is a PowerPoint COM Add-In (VSTO or Office.js) that lets users save and replay image manipulation presets — crop, resize, and reposition — across one or many slides. The UI lives in a custom Ribbon tab called **ADD-INS**, grouped under **"PicWrangler: Batch Manipulate Pictures"**.

---

## Technology Stack

### Option A — VSTO (Recommended for full feature support)
- **Language:** C# (.NET Framework 4.7.2+)
- **Framework:** Visual Studio Tools for Office (VSTO)
- **IDE:** Visual Studio 2022 with "Office/SharePoint development" workload
- **Target:** Microsoft PowerPoint 2016/2019/365 (Windows only)
- **Persistence:** `app.config` or JSON file in `%AppData%\PicWrangler\`

### Option B — Office.js (Cross-platform, web-based)
- **Language:** TypeScript + React
- **Framework:** Office Add-ins (Office.js)
- **Bundler:** Webpack or Vite
- **Target:** PowerPoint on Windows, Mac, and Web
- **Persistence:** `Office.context.document.settings` or localStorage

> **Recommendation:** Use VSTO (Option A) because Office.js has limited support for advanced image crop/position APIs. All implementation notes below assume VSTO/C# unless otherwise stated.

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
│   │   └── PicWranglerRibbon.xml  # Ribbon UI layout definition
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

The ribbon tab is named **ADD-INS** and contains one group:

| Control | Type | ID | Label | Notes |
|---|---|---|---|---|
| Preset selector | `dropDown` | `ddPreset` | *(none)* | Items: "Preset 1", "Preset 2", "Preset 3", "Preset 4" |
| Set Preset | `button` | `btnSetPreset` | "Set Preset" | Saves current selection's image settings into the chosen preset |
| Apply | `button` | `btnApply` | "Apply" | Applies preset to the currently selected image |
| Apply to Slides | `button` | `btnApplyToSlides` | "Apply to Slides" | Applies preset to all images on selected slides |
| Crop checkbox | `checkBox` | `chkCrop` | "Crop" | Whether to apply crop when using Apply / Apply to Slides |
| Size checkbox | `checkBox` | `chkSize` | "Size" | Whether to apply size |
| Position checkbox | `checkBox` | `chkPosition` | "Position" | Whether to apply position |
| Bulk Insert | `button` | `btnBulkInsert` | "Bulk Insert" | *(explained later — stub for now)* |
| Apply Preset checkbox | `checkBox` | `chkApplyPreset` | "Apply Preset" | *(explained later — stub for now)* |
| Add Title checkbox | `checkBox` | `chkAddTitle` | "Add Title" | *(explained later — stub for now)* |
| Add Notes checkbox | `checkBox` | `chkAddNotes` | "Add Notes" | *(explained later — stub for now)* |
| Help | `button` | `btnHelp` | "Help" | Opens a help dialog |

---

## Feature Implementation Guide

### 1. Set Preset (`btnSetPreset_Click`)

**Goal:** Read crop, size, and position from the currently selected image and save into the active preset slot.

**Steps:**
1. Get the selected preset name from `ddPreset`.
2. Get the currently selected shape via `Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1]`.
3. Verify the shape is a picture (`shape.Type == MsoShapeType.msoPicture`).
4. Read crop values from `shape.PictureFormat`:
   - `CropLeft`, `CropRight`, `CropTop`, `CropBottom` (all in points)
5. Read size from `shape.Width`, `shape.Height`.
6. Read position from `shape.Left`, `shape.Top`.
7. Construct a `Preset` object and pass to `PresetStore.Save(presetName, preset)`.
8. Show a brief status message (e.g., tooltip or MessageBox) confirming the save.

**Error handling:**
- No shape selected → show MessageBox "Please select an image first."
- Shape is not a picture → show MessageBox "Selected object is not a picture."

---

### 2. Apply (`btnApply_Click`)

**Goal:** Apply the active preset to the currently selected image, respecting the Crop/Size/Position checkboxes.

**Steps:**
1. Load the preset via `PresetStore.Load(selectedPresetName)`.
2. Get the selected shape (same as above).
3. Call `ImageApplicator.Apply(shape, preset, applyCrop, applySize, applyPosition)`.

**`ImageApplicator.Apply` logic:**
```csharp
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
```

**Error handling:**
- Preset is empty (never set) → MessageBox "Preset X has not been configured yet."
- No checkboxes checked → MessageBox "Please check at least one of Crop, Size, or Position."

---

### 3. Apply to Slides (`btnApplyToSlides_Click`)

**Goal:** Apply the preset to **all picture shapes** on all **selected slides**.

**Steps:**
1. Load the preset.
2. Get selected slides:
   ```csharp
   var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
   // selection.Type == PpSelectionType.ppSelectionSlides
   var slideRange = selection.SlideRange;
   ```
3. For each slide in `slideRange`, iterate `slide.Shapes`:
   - Skip non-picture shapes.
   - Call `ImageApplicator.Apply(shape, preset, applyCrop, applySize, applyPosition)`.
4. Report how many images were updated (e.g., "Applied preset to 7 images across 3 slides.").

**Error handling:**
- No slides selected in Slide Panel → fall back to the active slide only, with a warning.
- Slide has no pictures → skip silently (count = 0 is fine).

---

### 4. Preset Persistence (`PresetStore.cs`)

- Storage path: `%AppData%\PicWrangler\presets.json`
- Format: a JSON dictionary keyed by preset name:
  ```json
  {
    "Preset 1": { "Crop": {...}, "Size": {...}, "Position": {...} },
    "Preset 2": null,
    "Preset 3": null,
    "Preset 4": null
  }
  ```
- Use `Newtonsoft.Json` (NuGet). `System.Text.Json` is not included in .NET Framework 4.7.2 by default.
- Initialize with 4 empty presets on first run.
- Always load presets on `ThisAddIn_Startup` and populate the dropdown.

---

## PowerPoint Object Model — Key APIs

| Task | API |
|---|---|
| Selected shape | `app.ActiveWindow.Selection.ShapeRange[1]` |
| Shape type check | `shape.Type == MsoShapeType.msoPicture` |
| Crop left | `shape.PictureFormat.CropLeft` |
| Crop right | `shape.PictureFormat.CropRight` |
| Crop top | `shape.PictureFormat.CropTop` |
| Crop bottom | `shape.PictureFormat.CropBottom` |
| Width/Height | `shape.Width`, `shape.Height` |
| Position | `shape.Left`, `shape.Top` |
| Lock aspect ratio | `shape.LockAspectRatio` |
| Shapes on slide | `slide.Shapes` (iterate, check `.Type`) |
| Selected slides | `app.ActiveWindow.Selection.SlideRange` |

> **Units:** PowerPoint uses **points** (1 inch = 72 points) for all measurements. Be consistent — never mix EMUs or pixels.

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

# 4. Install (for end users)
# Publish via ClickOnce: Build > Publish PicWrangler
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
- [ ] Apply with all checkboxes ON restores all three properties
- [ ] Apply with only "Size" checked changes only size, not crop or position
- [ ] Apply to Slides processes all pictures on 3 selected slides
- [ ] Apply to Slides skips text boxes and other non-picture shapes
- [ ] Presets survive closing and reopening PowerPoint (persistence check)
- [ ] Error shown when clicking Apply with no selection
- [ ] Error shown when clicking Set Preset on a non-picture shape

---

## Known PowerPoint API Gotchas

1. **Namespace ambiguity** — `Microsoft.Office.Core` and `Microsoft.Office.Interop.PowerPoint` both define `Shape`; `System.Windows.Forms` and `Microsoft.Office.Interop.PowerPoint` both define `Application`. Always use a namespace alias in files that import both: `using PPT = Microsoft.Office.Interop.PowerPoint;` and qualify as `PPT.Shape`, `PPT.Application`, etc.
2. **Crop values don't reset automatically** — if you apply size before crop, the visual crop region may shift. Always apply crop last, or re-read after resizing.
2. **`LockAspectRatio` must be set to `msoFalse`** before setting both Width and Height independently.
3. **`Selection.SlideRange`** only works when slides are selected in the slide panel (Normal or Slide Sorter view). In other views, fall back gracefully.
4. **Grouped shapes** — `shape.Type == msoGroup`. Optionally recurse into `shape.GroupItems` to find pictures inside groups.
5. **Linked vs. embedded pictures** — `msoPicture` covers both; linked pictures may have different crop behavior.

---

## Stub: Bulk Insert (to be specified later)

The **Bulk Insert** button and its associated **Apply Preset**, **Add Title**, and **Add Notes** checkboxes are reserved for a future feature. For now:
- Wire the button to a no-op handler that shows "Bulk Insert coming soon."
- Keep the checkboxes visible but disabled until the feature is implemented.

---

## File Naming Conventions

- Classes: `PascalCase`
- Methods: `PascalCase`
- Private fields: `_camelCase`
- Constants: `ALL_CAPS` or `PascalCase` (pick one and be consistent)
- JSON keys: `PascalCase` (matches C# property names for easy serialization)
