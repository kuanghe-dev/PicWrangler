using System.Collections.Generic;
using Microsoft.Office.Core;
using PicWrangler.Ribbon;
using PicWrangler.Services;

namespace PicWrangler
{
    public partial class ThisAddIn
    {
        private PicWranglerRibbon _ribbon;

        internal PresetStore PresetStore { get; } = new PresetStore();

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new PicWranglerRibbon();
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Ensure the preset file is initialized on first run
            PresetStore.LoadAll();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup   += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown  += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
