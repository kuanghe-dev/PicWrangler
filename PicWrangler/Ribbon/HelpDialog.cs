using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace PicWrangler.Ribbon
{
    internal class HelpDialog : Form
    {
        public HelpDialog()
        {
            Text            = "PicWrangler Help";
            Size            = new Size(900, 800);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterScreen;
            Padding         = new Padding(16);

            var rtb = new RichTextBox
            {
                Dock        = DockStyle.Fill,
                ReadOnly    = true,
                BorderStyle = BorderStyle.None,
                BackColor   = SystemColors.Control,
                Font        = new Font("Segoe UI", 9.5f),
                Margin      = new Padding(8),
            };

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            AppendBold(rtb, "PicWrangler — Batch Manipulate Pictures\n");
            rtb.AppendText($"Version {version.Major}.{version.Minor}.{version.Build}\n\n");
            AppendBold(rtb, "PRESETS\n");
            rtb.AppendText(
                "1. Select an image and click \"Set Preset\" to capture its crop, size, and position.\n" +
                "2. Check Crop / Size / Position to choose what gets applied.\n" +
                "3. Click \"Apply\" to apply the preset to the selected image.\n" +
                "4. Select slides in the slide panel and click \"Apply to Slides\" to batch-apply.\n" +
                "Up to 4 presets can be stored simultaneously.\n\n");
            AppendBold(rtb, "BULK INSERT\n");
            rtb.AppendText(
                "5. Click \"Bulk Insert\" to open a file dialog and select one or more images.\n" +
                "   Each image is inserted as a new slide after the current slide.\n" +
                "   Add Title — sets each slide's title to the image filename (checked by default).\n" +
                "   Add Notes — writes the original file path to the slide notes.");

            var okButton = new Button
            {
                Text         = "OK",
                DialogResult = DialogResult.OK,
                Size         = new Size(150, 35),
                Anchor       = AnchorStyles.Bottom | AnchorStyles.Right,
            };

            var panel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            panel.Controls.Add(okButton);
            panel.Resize += (s, e) =>
                okButton.Location = new Point(panel.Width - okButton.Width - 10,
                                              (panel.Height - okButton.Height) / 2);

            Controls.Add(rtb);
            Controls.Add(panel);
            AcceptButton = okButton;
            CancelButton = okButton;
        }

        private static void AppendBold(RichTextBox rtb, string text)
        {
            rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
            rtb.AppendText(text);
            rtb.SelectionFont = new Font(rtb.Font, FontStyle.Regular);
        }
    }
}
