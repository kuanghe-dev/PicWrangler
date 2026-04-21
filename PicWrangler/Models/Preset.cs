namespace PicWrangler.Models
{
    public class Preset
    {
        public string Name { get; set; }
        public CropSettings Crop { get; set; }
        public SizeSettings Size { get; set; }
        public PositionSettings Position { get; set; }
    }
}
