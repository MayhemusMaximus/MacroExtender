using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Windows.Forms;
using System.IO;

//using Outlook = Microsoft.Office.Interop.Outlook;

//using Microsoft.VisualBasic;
//using System;
//using System.Collections;
////using System.Collections.Generic;
//using System.Data;
//using System.Diagnostics;
////using Excel = Microsoft.Office.Interop.Excel;
//using Olook = Microsoft.Office.Interop.Outlook;

namespace MacroExtender
{
    class ApplicationEngine : EngineBase
    {
        ////~~> Define your Excel Objects
        //Excel.Application xlApp = new Excel.Application();
        //Excel.Workbook xlWorkBook;
        //Excel.Worksheet xlWorkSheet;

        //Excel.Range xlRange;
        ////~~> Define Outlook Objects
        //Olook.Application olApp = new Olook.Application();

        //Olook.MailItem olMail;



        /// <summary>
        /// This method uses scope to determine how to handle the method to be run.
        /// </summary>
        /// <param name="Scope">The scope of the method to be run.</param>
        /// <param name="BeginningRow">The "Begin" row of the method to be run.</param>
        public void DelegationCheck(string Scope, int BeginningRow)
        {
            if (Scope == "Application")
                Engine(BeginningRow);
            else
                delegateMacro(Scope, BeginningRow);
        }

        /// <summary>
        /// This method forwards non-application scope methods to the workbook engine.
        /// </summary>
        /// <param name="Scope">The scope of the method to be run.</param>
        /// <param name="BeginningRow">The "Begin" row of the method to be run.</param>
        private void delegateMacro(string Scope, int BeginningRow)
        {
            WorkbookEngine workbookEngine = new WorkbookEngine();
            workbookEngine.DelegationCheck(Scope, BeginningRow);
        }

        /// <summary>
        /// Handles application scope methods.
        /// </summary>
        /// <param name="beginningRow">
        /// The line that the "Begin" method is located on.
        /// </param>
        private void Engine(int beginningRow)
        {

            methodCell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(MethodColumn + beginningRow);

            int curRow = beginningRow;

            Boolean end = true;

            if (methodCell.Value2 == "Begin")
            {
                end = false;
            }
            else
            {
                MessageBox.Show("The first method of a macro must be 'Begin'");
                return;
            }

            do
            {

                UpdateRanges(curRow);


                switch (methodName)
                {
                    case "End": // RELEASES THE DO LOOP TO END THE MACRO
                        end = true;
                        break;

                    case "Open Workbook": // USED TO OPEN A WORKBOOK
                        MessageBox.Show("This is you pretending to open a Workbook.");
                        // NASTY ERROR.
                        //    //WBToOpen.Open("C:Users\rmatton\desktop\");
                        //    FileInfo fi = new FileInfo("C:\\Users\\RMatton\\Desktop\\PROCESSING WORK ORDER.xls");
                        //    if (!fi.Exists)
                        //    {
                        //        MessageBox.Show("C:\\Users\\RMatton\\Desktop\\PROCESSING WORK ORDER.xls does not exist.");
                        //    }
                        //    else
                        //    {
                        //        Excel.Workbook WBToOpen = new Excel.Workbook();
                        //        WBToOpen = ExcelObj.Workbooks.Open("C:\\Users\\RMatton\\Desktop\\PROCESSING WORK ORDER.xls");
                        //    }
                        ////= "C:\\Users\\RMatton\\Desktop\\PROCESSING WORK ORDER.xls";

                        break;

                    case "Close Workbook": // USED TO CLOSE AN OPEN WORKBOOK
                        OpenFileDialog OFD = new OpenFileDialog();
                        DialogResult userClickedOk = OFD.ShowDialog();
                        if (userClickedOk == DialogResult.OK)
                        {
                            arg1Cell.Value2 = OFD.FileName;
                        }
                        else
                            end = true;
                        //userClickedOk = true ? arg1Cell.Value2 = userClickedOk : end = true;

                        break;

                    case "Input Box": // ALLOWS THE USER TO ENTER A RUNTIME VARIABLE
                        break;

                    case "Run Macro": // USED TO CALL ANOTHER MACRO
                        RunMacro(arg1);
                        break;

                    //case "Create Email": // USED TO AUTOMATE THE CREATION OF AN E-MAIL
                    //    sendEmailThroughOutlook();
                    //    break;

                } // switch (methodName)

                curRow++;
            } while (end == false); // THE DO LOOP WILL CONTINUE UNTIL THE End METHOD IS PASSED FROM THE MACRO

        }

        #region ENGINE HELPER METHODS

        /// <summary>
        /// This method collects macro information to handle the "Run Macro" user method.
        /// </summary>
        /// <param name="macroName">The name of the macro to be run.</param>
        private void RunMacro(string macroName)
        {
            for (int x = 0; x < MacrosList.Count; x++)
            {
                if (MacrosList[x].MacroName == macroName)
                {
                    WorkbookEngine workbookEngine = new WorkbookEngine();
                    workbookEngine.DelegationCheck(MacrosList[x].Scope, MacrosList[x].Row);
                }
            }
        }

        ///// <summary>
        ///// This method is used to generate Microsoft Outlook Emails.
        ///// </summary>
        //public void sendEmailThroughOutlook()
        //{
        //    try
        //    {
        //        // CREATE THE OUTLOOK APPLICATION.
        //        Outlook.Application oApp = new Outlook.Application();

        //        // CREATE A NEW MAIL ITEM.
        //        Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

        //        // SET HTMLBody.
        //        // ADD THE BODY OF THE EMAIL
        //        Excel.Range sendRange = Sheet.get_Range("A1:E30");
        //        sendRange.Copy();
        //        //sendRange.Select();
        //        oMsg.HTMLBody = "Hello, Processing your message body will go here!!";
        //        //oMsg.HTMLBody += sendRange.PasteSpecial(Excel.XlPasteType.xlPasteAll);
        //        oMsg.HTMLBody += Clipboard.GetData(System.Windows.Forms.DataFormats.Html);
        //        //oMsg.HTMLBody += Clipboard.GetData(System.Windows.Forms.DataFormats.Rtf);
        //        //oMsg.HTMLBody += Clipboard.GetImage();
        //        //oMsg.HTMLBody += sendRange;
        //        //oMsg.HTMLBody += Clipboard.GetText();
        //        //oMsg.HTMLBody += Clipboard.GetDataObject();
        //        //oMsg.HTMLBody += sendRange;
        //        //oMsg.HTMLBody += Clipboard.GetData(System.Windows.Forms.DataFormats.MetafilePict);
        //        //oMsg.HTMLBody += sendRange.PasteSpecial(Excel.XlPasteType.xlPasteAllUsingSourceTheme);


        //        // ADD AN ATTACHMENT.
        //        String sDisplayName = "MyAttachment";
        //        int iPosition = (int)oMsg.Body.Length + 1;
        //        int iAttachType = (int)Outlook.OlAttachmentType.olByValue;

        //        //NOW ATTACH THE FILE
        //        Outlook.Attachment oAttach = oMsg.Attachments.Add(@"C:\\users\\rmatton\\desktop\\material description manual.pdf", iAttachType, iPosition, sDisplayName);
        //        oAttach = oMsg.Attachments.Add(@"C:\\Users\\rmatton\\Desktop\\OFFICE STUFF\\DAILY MACHINE TOLERANCE LOG.pdf", iAttachType, iPosition, sDisplayName);


        //        //SUBJECT LINE
        //        oMsg.Subject = "Automating attachments.";

        //        // ADD A RECIPIENT.
        //        Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
        //        // CHANGE THE RECIPIENT IN THE NEXT LINE IF NECESSARY.
        //        Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("rmatton@supstl.com");
        //        oRecip.Resolve();

        //        //SEND.
        //        oMsg.Save();
        //        oMsg.Display();
        //        //((Outlook._MailItem)oMsg).Send();

        //        // CLEAN UP.
        //        oRecip = null;
        //        oRecips = null;
        //        oMsg = null;
        //        oApp = null;
        //        Clipboard.Clear();
        //    }
        //    catch
        //    {

        //    }
        //}

        //public object RangetoHTML(Excel.Range rng)
        //{
        //    object functionReturnValue = null;
        //    // Changed by Ron de Bruin 28-Oct-2006
        //    // Working in Office 2000-2010
        //    object fso = null;
        //    object ts = null;
        //    string TempFile = null;
        //    Excel.Workbook TempWB = default(Excel.Workbook);

        //    TempFile = Interaction.Environ("temp") + "/" + Strings.Format(DateAndTime.Now, "dd-mm-yy h-mm-ss") + ".htm";

        //    //Copy the range and create a new workbook to past the data in
        //    rng.Copy();

        //    TempWB = xlApp.Workbooks.Add(1);

        //    var _with2 = TempWB.Sheets(1);
        //    _with2.Cells(1).PasteSpecial(Paste: 8);
        //    _with2.Cells(1).PasteSpecial(-4163, , false, false);
        //    _with2.Cells(1).PasteSpecial(-4122, , false, false);
        //    _with2.Cells(1).Select();
        //    xlApp.CutCopyMode = false;
        //     // ERROR: Not supported in C#: OnErrorStatement

        //    _with2.DrawingObjects.Visible = true;
        //    _with2.DrawingObjects.Delete();
        //     // ERROR: Not supported in C#: OnErrorStatement


        //    //Publish the sheet to a htm file
        //    var _with3 = TempWB.PublishObjects.Add(SourceType: 4, Filename: TempFile, Sheet: TempWB.Sheets(1).Name, Source: TempWB.Sheets(1).UsedRange.Address, HtmlType: 0);
        //    _with3.Publish(true);

        //    //Read all data from the htm file into RangetoHTML
        //    fso = Interaction.CreateObject("Scripting.FileSystemObject");
        //    ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2);
        //    functionReturnValue = ts.ReadAll;
        //    ts.Close();
        //    functionReturnValue = Strings.Replace(RangetoHTML(), "align=center x:publishsource=", "align=left x:publishsource=");

        //    //Close TempWB
        //    TempWB.Close(savechanges: false);

        //    //Delete the htm file we used in this function
        //    FileSystem.Kill(TempFile);

        //    ts = null;
        //    fso = null;
        //    TempWB = null;
        //    return functionReturnValue;
        //}
        //private void Button1_Click(System.Object sender, System.EventArgs e)
        //{
        //    //~~> Opens an exisiting Workbook. Change path and filename as applicable
        //    //xlWorkBook = xlApp.Workbooks.Open("C:\\Sample.xlsx");
        //    //~~> Set the relevant sheet that we want to work with
        //    //xlWorkSheet = xlWorkBook.Sheets("Sheet1");

        //    //xlRange = xlWorkSheet.Range("A1:F20");
        //    xlRange = Sheet.get_Range("A1:F20");

        //    olMail = olApp.CreateItem(0);

        //    // ERROR: Not supported in C#: OnErrorStatement

        //    var _with1 = olMail;
        //    _with1.To = "INSERT TO EMAIL HERE";
        //    _with1.CC = "";
        //    _with1.BCC = "";
        //    _with1.Subject = "This is the Subject line";
        //    _with1.HTMLBody = (string)RangetoHTML(xlRange);
        //    _with1.Display();
        //    //or use .Send to send it
        //    // ERROR: Not supported in C#: OnErrorStatement


        //    //~~> Close the File
        //    xlWorkBook.Close(false);

        //    //~~> Quit the Excel Application
        //    xlApp.Quit();

        //    //~~> Clean Up
        //    releaseObject(xlApp);
        //    releaseObject(xlWorkBook);

        //    //~~> Similarly cleanup for outlook. not including as I am using .Display()

        //}

        ////~~> Release the objects
        //private void releaseObject(object obj)
        //{
        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //    }
        //    finally
        //    {
        //        GC.Collect();
        //    }
        //}

        #endregion
    }
}



//public class Form1
//{
//}
