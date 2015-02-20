using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Windows.Forms;

namespace MacroExtender
{
    public class WorkbookEngine : EngineBase
    {
        /// <summary>
        /// This method uses scope to determine how to handle the method to be run.
        /// </summary>
        /// <param name="Scope">The scope of the method to be run.</param>
        /// <param name="BeginningRow">The "Begin" row of the method to be run.</param>
        public void DelegationCheck(string Scope, int BeginningRow)
        {
            if (Scope == "Workbook")
                Engine(BeginningRow);
            else
                delegateMacro(BeginningRow);
        }

        /// <summary>
        /// This method forwards non-application scope methods to the workbook engine.
        /// </summary>
        /// <param name="BeginningRow">The "Begin" row of the method to be run.</param>
        private void delegateMacro(int BeginningRow)
        {
            WorksheetEngine worksheetEngine = new WorksheetEngine();
            worksheetEngine.Engine(BeginningRow);
        }

        /// <summary>
        /// Handles workbook scope methods.
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
                    case "Activate Worksheet": // USED TO BRING A WORKBOOK WORKSHEET TO FOCUS
                        Microsoft.Office.Interop.Excel.Worksheet oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)Sheets.get_Item(arg1);
                        ((Microsoft.Office.Interop.Excel._Worksheet)oWorksheet).Activate();
                        break;

                    case "End": // RELEASES THE DO LOOP TO END THE MACRO
                        end = true;
                        break;

                    case "Hide Worksheet": // USED TO HIDE A WORKBOOK WORKSHEET
                        Microsoft.Office.Interop.Excel.Worksheet hideSheet = (Microsoft.Office.Interop.Excel.Worksheet)Sheets.get_Item(arg1);
                        ((Microsoft.Office.Interop.Excel._Worksheet)hideSheet).Visible = XlSheetVisibility.xlSheetHidden;
                        break;

                    case "Input Box": // ALLOWS THE USER TO ENTER A RUNTIME VARIABLE
                        break;

                    case "Run Macro": // USED TO CALL ANOTHER MACRO
                        RunMacro(arg1);
                        break;

                    case "Show Worksheet": // USED TO SHOW A WORKBOOK WORKSHEET
                        Microsoft.Office.Interop.Excel.Worksheet showSheet = (Microsoft.Office.Interop.Excel.Worksheet)Sheets.get_Item(arg1);
                        ((Microsoft.Office.Interop.Excel._Worksheet)showSheet).Visible = XlSheetVisibility.xlSheetVisible;
                        break;

                    case "Sort Sheets": // USED TO ALPHABETISE THE WORKSHEETS IN A WORKBOOK
                        break;
                } // switch (methodName)

                curRow++;
            } while (end == false); // THE DO LOOP WILL CONTINUE UNTIL THE End METHOD IS PASSED FROM THE MACRO

        } // private void WorkbookEngine(int row)

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
                    WorksheetEngine worksheetEngine = new WorksheetEngine();
                    worksheetEngine.Engine(MacrosList[x].Row);
                }
            }
        }
    }
}
