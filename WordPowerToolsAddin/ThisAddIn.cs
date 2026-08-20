using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordPowerToolsAddin
{
    public sealed partial class Globals
    {
        private static ThisAddIn _ThisAddIn;
        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set { if (_ThisAddIn == null) _ThisAddIn = value; }
        }
        internal static global::Microsoft.Office.Tools.Factory Factory { get; set; }
    }

    public partial class ThisAddIn : global::Microsoft.Office.Tools.AddInBase
    {
        // הוספנו את המאפיין החסר כדי שהקומפיילר יכיר את Application
        public Word.Application Application { get; private set; }

        public ThisAddIn(global::Microsoft.Office.Tools.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            Globals.Factory = factory;
        }

        protected override void Initialize()
        {
            base.Initialize();
            // עכשיו ההשמה הזו תעבוד ללא שגיאות
            this.Application = this.GetHostItem<global::Microsoft.Office.Interop.Word.Application>(typeof(global::Microsoft.Office.Interop.Word.Application), "Application");
            Globals.ThisAddIn = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
        }

        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        protected override void OnShutdown()
        {
            this.Application = null;
            base.OnShutdown();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PowerRibbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}