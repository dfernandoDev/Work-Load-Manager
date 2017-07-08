using System;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Diagnostics;

namespace WorkLoadManager
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }

        private IntPtr _hwnd;
    }
    public partial class WorkLoadRibbon
    {
        private string sessionID;
        private bool authenticated = false;

        private void RallyConnect(object sender, RibbonControlEventArgs e)
        {
            RallyAuthenticator.frmAuthenticate fm = new RallyAuthenticator.frmAuthenticate();
            fm.WebURL = Properties.Settings.Default.Rally_SSO_Url;

            IntPtr hwnd = Process.GetCurrentProcess().MainWindowHandle;
            IWin32Window win = Control.FromHandle(hwnd);
            DialogResult result = fm.ShowDialog(win);

            if (result == DialogResult.OK)
            {
                authenticated = true;
                sessionID = fm.SessionID;

                Rally.RestApi.RallyRestApi restApi = new Rally.RestApi.RallyRestApi(webServiceVersion: "v2.0");
                // Rally.RestApi.Auth.ApiAuthManager authMan = new Rally.RestApi.Auth.ApiAuthManager();

            }
            /*
            //Excel.Workbook wb = ((Workbook)Globals.ThisAddIn.Application.ActiveWorkbook);
            //Excel.Worksheet ws = ((Worksheet)wb.Worksheets[1]);
            //Excel.Worksheet ws = ((Worksheet)wb.ActiveSheet);
            Excel.Worksheet ws = ((Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            
            var range = ((Range) Globals.ThisAddIn.Application.Cells[1, 1]);
            var range = ((Range)ws.Cells[1, 1]);
            range.Value = DateTime.Now;

            range = ((Range)ws.Cells[1, 2]);
            range.Value = sessionID;
            */
        }
    }
}
