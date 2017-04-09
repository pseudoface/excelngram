using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace NGramAddIn
{
    public partial class ExtRibbon
    {
        private int wsCount;
        private byte size = 1;
        public Excel.Application app;
        public Excel.Workbook wkbk;
        public object missing = Type.Missing;

        //TODO: should project's Property Pages -> Build -> Com Visible Object be checked?

        /// <summary>
        /// Determine whether the current Excel Application is in Edit Mode.
        /// </summary>
        /// <param name="exapp">The current sheet's <see cref="Excel.Application"/> object.</param>
        /// <returns><see cref="bool"/></returns>
        public bool IsInEditMode(Excel.Application exapp)
        {
            if(exapp.Interactive == false)
            {
                return false;
            }
            try
            {
                exapp.Interactive = false;
                exapp.Interactive = true;

                return false;
            }
            catch
            {
                return true;
            }
        }

        public void testNamedRangeFind()
        {
            wkbk = Globals.ThisAddIn.Application.ActiveWorkbook;
            int i = wkbk.Names.Count;
            string address = "";
            string sheetName = "";

            if(i != 0)
            {
                foreach(Excel.Name name in wkbk.Names)
                {
                    string value = name.Value;
                    //Sheet and Cell e.g. =Sheet1!$A$1 or =#REF!#REF! if refers to nothing
                    string linkName = name.Name;
                    //gives the name of the link e.g. sales
                    if(value != "=#REF!#REF!")
                    {
                        address = name.RefersToRange.Cells.Address[true, true, Excel.XlReferenceStyle.xlA1, missing, missing];
                        sheetName = name.RefersToRange.Cells.Worksheet.Name;
                    }
                    Debug.WriteLine("" + value + ", " + linkName + " ," + address + ", " + sheetName);
                }
            }

        }

        private void btnNGram_Click(object sender, RibbonControlEventArgs e)
        {
            //Microsoft.Office.Tools.Excel.Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);
            Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            string buttonName = "btnNGram";

            if (((RibbonButton)sender).Enabled)
            {
                //if there are selections to choose from in the dropdown
                if(cmbNumOfWords.Items.Count > 0)
                {
                    //check that the current dropdown selection is a digit, if it is then convert it, otherwise use 1 as a default.
                    size = IsAllDigits(cmbNumOfWords.Text) ? Convert.ToByte(cmbNumOfWords.Text) : Convert.ToByte(1);
                }

                //TODO: if a cell is currently being edited, update the cell to whatever value sits in it.
                if (IsInEditMode(Globals.ThisAddIn.Application))
                {
                    
                }

                Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
                Excel.Range newRange = worksheet.Range["A1", "D5"].Copy();
                newRange.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, Missing.Value, Missing.Value);

                Excel.Range rng = Globals.ThisAddIn.Application.Selection as Excel.Range;

                if (rng != null)
                {
                    if (rng.Value != null && rng.Count > 0)
                    {
                        //string[] cellValues = rng.Value;
                        //List<string> result = ((IEnumerable)cellValues).Cast<string>().ToList();
                        app = Globals.ThisAddIn.Application;
                        app.AddCustomList(ListArray:rng, ByRow: false);


                        List<string> result = new List<string>();

                        if (rng.Count > 1)
                        {
                            foreach (dynamic cell in rng.Value)
                            {
                                if (cell != null)
                                {
                                    result.Add(cell.ToString());
                                }
                            }
                        }
                        else
                        {
                            result.Add(rng.Value);
                        }

                        if (result.Count > 0)
                        {
                            IList<object> nGramColl = new List<object>();

                            foreach (string phrase in result)
                            {
                                if (!string.IsNullOrWhiteSpace(phrase))
                                {
                                    //IEnumerable<string> nGram = makeNgrams(phrase, size);
                                    IList<string> nGram = makeNgrams(phrase, size).ToList();

                                    if(nGram.Any())
                                    {
                                        nGramColl.Add(nGram.ToList());
                                    }
                                }
                            }

                            Workbook nGramWorkbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);

                            //Adding a new worksheet to our workbook.
                            //nGramWorkbook.Sheets.Add(System.Type.Missing, System.Type.Missing, 1, Excel.XlSheetType.xlWorksheet);
                            //var activeSheet = nGramWorkbook.ActiveSheet;

                            nGramWorkbook.Sheets.Add(After: nGramWorkbook.ActiveSheet, Count: 1, Type: Excel.XlSheetType.xlWorksheet);

                            foreach(Excel.Worksheet ws in nGramWorkbook.Sheets)
                            {
                                if(ws == nGramWorkbook.ActiveSheet)
                                {
                                    ws.Name = "Source";
                                }
                                else
                                {
                                    if(nGramWorkbook.Sheets.Count < 3)
                                    {
                                        ws.Name = "Results";
                                    }
                                }
                            }
                            
                            //Move this newly added worksheet to the last position.
                            //nGramWorkbook.Sheets.Move(After: nGramWorkbook.Sheets[nGramWorkbook.Sheets.Count]);

                            //create an instance of the currently active sheet.
                            Worksheet activeSheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


                            Excel.Range pastingRange = activeSheet.get_Range("A1");
                            if (pastingRange != null)
                            {
                                pastingRange.Value = nGramColl; 
                                nGramWorkbook.Sheets.FillAcrossSheets(pastingRange);
                            }


                            int masterCounter = 1;

                            foreach (List<string> list in nGramColl)
                            {
                                //do something usefull: you select now an individual cell.
                                //var range = resultsWorksheet.get_Range("A1", "A1");
                                //range.Value2 = "test"; //Value2 is not a typo.
                                int counter = masterCounter;

                                foreach (string str in list)
                                {
                                    var cellName = "A" + counter.ToString();
                                    var range = activeSheet.get_Range(cellName, cellName);
                                    range.Value2 = str;
                                    ++counter;
                                    ++masterCounter;
                                }
                            }
                            
                        }
                        else
                        {
                            MessageBox.Show(@"The selected range does not contain any values for processing.", @"Empty Range", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                    }
                    else
                    {
                        MessageBox.Show(@"Either no range has been selected or the currently selected range does not contains any values for processing.", @"No Range/Empty Range", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                        //var dgs = Globals.ThisAddIn.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogSummaryInfo];
                        //dgs.Show();
                    }
                }
                else
                {
                    MessageBox.Show(@"Range is null.", @"Null Range", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //worksheet.Controls.AddNamedRange(selection, Name);
                }
            }
            else
            {
                worksheet.Controls.Remove(buttonName);
            }
        }

        /// <summary>
        /// Iterates over the words in the provided text (separated by spaces) and returns a collection of sequential word groups that are each "n" words long.
        /// </summary>
        /// <param name="text">The <see cref="string"/> phrase to process.</param>
        /// <param name="nGramSize">A <see cref="byte"/> specifying the number of words each result in the collection should contain.</param>
        /// <returns><see cref="IEnumerable{string}"/></returns>
        private IEnumerable<string> makeNgrams(string text, byte nGramSize)
        {
            StringBuilder nGram = new StringBuilder();
            Queue<int> wordLengths = new Queue<int>();
            int wordCount = 0;
            int lastWordLen = 0;

            //append the first character, if valid.
            //avoids if statement for each for loop to check i==0 for the before and after vars.
            if (text != string.Empty && char.IsLetterOrDigit(text[0]))
            {
                nGram.Append(text[0]);
                lastWordLen++;
            }

            //generate ngrams.
            for (int i = 1; i < text.Length - 1; i++)
            {
                char before = text[i - 1];
                char after = text[i + 1];

                //keep all punctuation that is surrounded by letters or numbers on both sides.
                if (char.IsLetterOrDigit(text[i]) || text[i] != ' ' && (char.IsSeparator(text[i]) || char.IsPunctuation(text[i])) && char.IsLetterOrDigit(before) && char.IsLetterOrDigit(after))
                {
                    nGram.Append(text[i]);
                    lastWordLen++;
                }
                else
                {
                    if (lastWordLen > 0)
                    {
                        wordLengths.Enqueue(lastWordLen);
                        lastWordLen = 0;
                        wordCount++;

                        if (wordCount >= nGramSize)
                        {
                            yield return nGram.ToString();
                            nGram.Remove(0, wordLengths.Dequeue() + 1);
                            wordCount -= 1;
                        }

                        nGram.Append(" ");
                    }
                }
            }
        }

        bool IsAllDigits(string s)
        {
            foreach(char c in s)
            {
                if(!char.IsDigit(c))
                    return false;
            }
            return true;
        }

        public string cboGetItemID(Microsoft.Office.Tools.Ribbon.RibbonComboBox control, int index)
        {
            if(control.Id == "cmbNumOfWords")
            {
                return index.ToString();
            }
            return string.Empty;
        }

        private void chkNamedRange_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Tools.Excel.Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            string name = "chkNamedRange";

            if (((RibbonCheckBox)sender).Checked)
            {
                Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

                if (selection != null)
                {
                    worksheet.Controls.AddNamedRange(selection, name);
                }
            }
            else
            {
                worksheet.Controls.Remove(name);
            }
        }

        private void cmbNumOfWords_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void cmbNumOfWords_ItemsLoading(object sender, RibbonControlEventArgs e)
        {

        }

        private void ExtRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
    }
}
