using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace NGramAddIn
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAddinVSTO
    {
        void ConvertIntoNewNGramSheet();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ExcelModule : IAddinVSTO
    {
        #region IAddinVSTO Members

        public void ConvertIntoNewNGramSheet()
        {
            Excel.Worksheet wks = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (wks != null)
            {
                Excel.Range myRange = wks.Range["A1", System.Type.Missing];
                myRange.Value2 = "NGram";
            }
        }

        #endregion IAddinVSTO Members
    }
}
