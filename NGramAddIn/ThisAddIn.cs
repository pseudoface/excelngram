using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace NGramAddIn
{
    public partial class ThisAddIn
    {
        private ExcelModule utilities;
        protected override object RequestComAddInAutomationService()
        {
            return utilities ?? (utilities = new ExcelModule());
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Microsoft.Office.Tools.Excel.Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);

            if (Application.Version == "12.0")
            {
                // 2010-specific code.
            }
            else
            {
                // 2013/2016-specific code.
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        //add this as the last method (a new one)
        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook wb, bool saveAsUI, ref bool cancel)
        {
            wb.Application.StatusBar = "NGram function";
            DialogResult msg = MessageBox.Show(@"This current Sheet's NGram results will be processed into a new Sheet", @"NGram functionality", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop);
            if (msg == DialogResult.Cancel)
            {
                cancel = true;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
