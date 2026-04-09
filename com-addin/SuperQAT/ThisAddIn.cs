using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;

namespace SuperQAT
{
    /// <summary>
    /// COM Add-in that works in Word, Excel, and PowerPoint.
    /// Registers itself via the registry so Office loads it on startup.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("SuperQAT.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ThisAddIn : IDTExtensibility2, IRibbonExtensibility
    {
        private object _application;
        private CommandPanel _panel;
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;

        // The Office application object — cast to the right type as needed
        public object Application => _application;

        /// <summary>
        /// Execute any ribbon command by its MSO control ID.
        /// This is the magic method that can trigger ANY Office command.
        /// </summary>
        public void ExecuteCommand(string msoId)
        {
            try
            {
                var app = _application;
                CommandBars bars = null;

                // Get CommandBars from whichever host app we're in
                if (app is Microsoft.Office.Interop.Word.Application wordApp)
                    bars = wordApp.CommandBars;
                else if (app is Microsoft.Office.Interop.Excel.Application excelApp)
                    bars = excelApp.CommandBars;
                else if (app is Microsoft.Office.Interop.PowerPoint.Application pptApp)
                    bars = pptApp.CommandBars;

                if (bars != null)
                {
                    bars.ExecuteMso(msoId);
                }
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Command '{msoId}' is not available right now.\n\n{ex.Message}",
                    "SuperQAT",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
        }

        // ── IDTExtensibility2 ──────────────────────────────────────────

        public void OnConnection(object application, ext_ConnectMode connectMode,
            object addInInst, ref Array custom)
        {
            _application = application;
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _panel = null;
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ── IRibbonExtensibility ───────────────────────────────────────

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SuperQAT.SuperQATRibbon.xml");
        }

        // Ribbon callback: "Open SuperQAT" button clicked
        public void OnOpenPanel(IRibbonControl control)
        {
            if (_panel == null || _panel.IsDisposed)
            {
                _panel = new CommandPanel(this);
            }

            _panel.Show();
            _panel.BringToFront();
        }

        // Ribbon callback: get button image
        public System.Drawing.Bitmap GetImage(IRibbonControl control)
        {
            return null; // Uses built-in Office icon specified in XML
        }

        // ── Helpers ────────────────────────────────────────────────────

        private static string GetResourceText(string resourceName)
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return "";
                using (var reader = new System.IO.StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        // ── COM Registration ──────────────────────────────────────────
        // These methods register/unregister the add-in in the registry
        // so Word, Excel, and PowerPoint all load it.

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            string[] apps = {
                @"Software\Microsoft\Office\Word\Addins\SuperQAT.Connect",
                @"Software\Microsoft\Office\Excel\Addins\SuperQAT.Connect",
                @"Software\Microsoft\Office\PowerPoint\Addins\SuperQAT.Connect"
            };

            foreach (var subkey in apps)
            {
                var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(subkey);
                if (key != null)
                {
                    key.SetValue("FriendlyName", "SuperQAT");
                    key.SetValue("Description", "All 649 QAT commands in one panel");
                    key.SetValue("LoadBehavior", 3); // Load on startup
                    key.Close();
                }
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            string[] apps = {
                @"Software\Microsoft\Office\Word\Addins\SuperQAT.Connect",
                @"Software\Microsoft\Office\Excel\Addins\SuperQAT.Connect",
                @"Software\Microsoft\Office\PowerPoint\Addins\SuperQAT.Connect"
            };

            foreach (var subkey in apps)
            {
                try { Microsoft.Win32.Registry.CurrentUser.DeleteSubKey(subkey, false); }
                catch { /* ignore if not present */ }
            }
        }
    }
}
