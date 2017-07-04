using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace WorkLoad
{
    public partial class WorkLoadRibbon
    {
        private string sessionID;
        private bool authenticated = false;

        private void RallyConnect(object sender, RibbonControlEventArgs e)
        {
            frmAuthenticate fm = new frmAuthenticate();
            fm.ShowDialog();
            sessionID = fm.SessionID ;

            if (sessionID.Length>0)
            {
                authenticated = true;

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
