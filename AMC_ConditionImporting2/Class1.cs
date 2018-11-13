using System;
using System.Windows.Forms;
using EllieMae.Encompass.Automation;
using EllieMae.Encompass.BusinessObjects.Loans;
using EllieMae.Encompass.BusinessObjects.Loans.Logging;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AMC_ConditionImporting2
{
    public class ConditionImporting : EllieMae.Encompass.Forms.Form
    {

        private EllieMae.Encompass.Forms.Button ImportButton = null;
        public override void CreateControls()
        {
            ImportButton = (EllieMae.Encompass.Forms.Button)FindControl("ImportConditionsBtn");
            ImportButton.Click += new EventHandler(ImportButton_Click);
        }
        public void ImportButton_Click(object sender, EventArgs e)
        {
            string fileName = null;

            // allow user to select an excel file for processing
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                }
            }

            if (fileName != null)
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;

                string loanExceptionId;
                string loanID;
                string exceptionStatus;
                int rCnt = 0;
                double startingrow;
                double row;
                int ConditionCount = 0;

                xlWorkBook = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "", false, false, 0, false, 1, 0);
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Exception Standard Report"];
                xlWorkSheet = xlWorkBook.Worksheets["Exception Standard Report"];

                range = xlWorkSheet.UsedRange;
                rCnt = range.Rows.Count;
                
                try
                {
                    if (range.Cells[5, 2].Value.ToString().Contains("Customer"))
                    {
                        //AMC using alternate template, condition data starts on row 6
                        loanID = range.Cells[6, 2].Value.ToString();
                        startingrow = 6;
                    }
                    else
                    {
                        //standard template used
                        loanID = range.Cells[5, 2].Value.ToString();
                        startingrow = 5;
                    }

                    if (loanID != Loan.LoanNumber || loanID == null)
                    {
                        // add check to see if loan number in spreadsheet matches current loan
                        Macro.Alert("Loan number does not match the customer loan ID.  Select a different file.");
                    }
                    else
                    {
                        for (row = startingrow; row <= rCnt; row++)
                        {
                                    exceptionStatus = range.Cells[row, 18].Value.ToString();
                                    loanExceptionId = range.Cells[row, 4].Value.ToString();

                                    // for this row, if the exception status = open, and if "loan exception id" value doesnt already exist in the field "CX.Underwriting.Imported" of this loan,  
                                    // then import the required fields, otherwise ignore this row
                                    string leids = Loan.Fields["CX.UNDERWRITING.IMPORTED"].Value.ToString();
                                    if (exceptionStatus == "open" && !leids.Contains(loanExceptionId))
                                    {
                                        //add underwriting conditions
                                        UnderwritingCondition underwritingCondition = Loan.Log.UnderwritingConditions.Add(Loan.LoanName);
                                        underwritingCondition.Title = "Due Diligence Finding";
                                        underwritingCondition.Source = "Exception Grade " + range.Cells[row, 20].Value.ToString() + " (AMC Condition Importer)";
                                        underwritingCondition.Description = range.Cells[row, 21].Value + " - " + range.Cells[row, 23].Value + " - " + range.Cells[row, 22].Value;
                                        underwritingCondition.ForExternalUse = false;
                                        underwritingCondition.ForInternalUse = true;

                                        //add this loan exception id to the CX.UNDERWRITING.IMPORTED string
                                        Loan.Fields["CX.UNDERWRITING.IMPORTED"].Value += (loanExceptionId + ",");
                                        ConditionCount++;
                                    }
                        }
                        //here send an alert to let the user know the processing is complete
                        Macro.Alert("Condition Importing now complete. " + ConditionCount.ToString() + " Conditions imported");
                    }

                    xlWorkBook.Close(false, null, null);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                }
                catch (Exception ex)
                {
                    Loan.Fields["CX.UNDERWRITING.ERROR"].Value = ex;
                    Macro.Alert("Importing process error, conditions not imported, contact Encompass team for assistance.");
                }
                finally
                {
                    //look into adding the assembly version number to the form for reference
                    //typeof(ConditionImporting).Assembly.GetName().Version;
                }
            }

        } //button1_click
    } //class
} //namespace