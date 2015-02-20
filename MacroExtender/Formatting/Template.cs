using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;

namespace MacroExtender
{

    struct templateStruct
    {
        private readonly string text;
        private readonly string comment;
        private readonly int commentHeight;
        private readonly Boolean border;
        private readonly Boolean header;
        private readonly int columnWidth;
        private readonly string cell;

        public templateStruct(string text, string comment, int commentHeight, Boolean border, Boolean header, int size, string cell)
        {
            this.text = text;
            this.border = border;
            this.header = header;
            this.columnWidth = size;
            this.cell = cell;
            this.comment = comment;
            this.commentHeight = commentHeight;
        }

        public string Text { get { return text; } }
        public string Comment { get { return comment; } }
        public int CommentHeight { get { return commentHeight; } }
        public Boolean Border { get { return border; } }
        public Boolean Header { get { return header; } }
        public int ColumnWidth { get { return columnWidth; } }
        public string Cell { get { return cell; } }
    }
    class Template : MacroExtenderRibbon
    {

        public void insertMacrosSheet()
        {
            Microsoft.Office.Interop.Excel.Range curCell;

            // STRUCT/ARRAY FOR STORING STRINGS IN CELLS FOR USER REFERENCE
            IList<templateStruct> templateStructArray = new ReadOnlyCollection<templateStruct>
            (new[] {
                // HEADERS
                new templateStruct ("Scope List","This column contains the constant and user-defined scopes used by the Macro Extender Add-in in this workbook.", 70,true,true,16,ScopeListColumn + "1"),
                new templateStruct ("Method List","This column contains a list of all possible methods that can be used to build macros.", 50,true,true,18,MethodListColumn + "1"),
                new templateStruct ("","This column is meant to break the page; it is purely aesthetic. You can hide it, shrink it, or grow it, but don't delete it.", 70,true,true,2,"C1"),
                new templateStruct ("Scope","Use this column to define the scope of macros on each macros 'Begin' row.", 50,true,true,16,ScopeColumn + "1"),
                new templateStruct ("Worksheet","Use this column to define a sheet associated with 'Worksheet' scope Macros (Begin row only).", 70,true,true,16,WorksheetColumn + "1"),
                new templateStruct ("Macro Name","Use this column to name macros.", 30,true,true,16,MacroNameColumn + "1"),
                new templateStruct ("Method","Use this column to sequence the methods.", 40,true,true,18,MethodColumn + "1"),
                new templateStruct ("Argument 1","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg1Column + "1"),
                new templateStruct ("Argument 2","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg2Column + "1"),
                new templateStruct ("Argument 3","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg3Column + "1"),
                new templateStruct ("Argument 4","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg4Column + "1"),
                new templateStruct ("Argument 5","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg5Column + "1"),
                new templateStruct ("Argument 6","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg6Column + "1"),
                new templateStruct ("Argument 7","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg7Column + "1"),
                new templateStruct ("Argument 8","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg8Column + "1"),
                new templateStruct ("Argument 9","This column is used to supply methods with user defined specifications.", 50,true,true,14,Arg9Column + "1"),
                new templateStruct ("Argument 10","This column is used to supply methods with user defined specifications.", 50,true,true,15,Arg10Column + "1"),
                // "Scope List"
                new templateStruct ("Application","The application scope is used to access methods that interact with the Excel application itself.", 60,false,false,0,ScopeListColumn + "2"),
                new templateStruct ("Workbook","The workbook scope is used to access methods that interact with Excel within a single workbook.", 60,false,false,0,ScopeListColumn + "3"),
                new templateStruct ("Worksheet","The worksheet scope is used to access methods that interact with Excel within a single worksheet.  "
                                    + "It is the only scope that accepts user defined names.", 90,false,false,0,ScopeListColumn + "4"),
                // "Method List"
                new templateStruct ("Activate Worksheet","This method is used to bring a worksheet into focus. The same as clicking the sheet's name tab.", 60,false,false,0,MethodListColumn + "2"),
                new templateStruct ("Close Workbook","This method is used to close a workbook.", 30,false,false,0,MethodListColumn + "3"),
                new templateStruct ("Create Email","This method is used to generate an e-mail.", 30,false,false,0,MethodListColumn + "4"),
                new templateStruct ("Filter Table","This method is used to filter a table using values specified by the user.", 50,false,false,0,MethodListColumn + "5"),
                new templateStruct ("Hide Column(s)","This method is used to hide columns specified by the user.", 40,false,false,0,MethodListColumn + "6"),
                new templateStruct ("Hide Row(s)","This method is used to hide rows specified by the user.", 40,false,false,0,MethodListColumn + "7"),
                new templateStruct ("Hide Worksheet","This method is used to hide a worksheet specified by the user.", 40,false,false,0,MethodListColumn + "8"),
                new templateStruct ("Input Box","This method is used to display an input box which will allow users to supply run-time variables.", 60,false,false,0,MethodListColumn + "9"),
                new templateStruct ("Open Workbook","This method is used to open a workbook.", 30,false,false,0,MethodListColumn + "10"),
                new templateStruct ("Run Macro","This method is used to run another user defined macro."
                                    + " Notes: The application scope version of this method will run all scopes, while the workbook scope version will only run worksheet scope macros.", 130,false,false,0,MethodListColumn + "11"),
                new templateStruct ("Show Column(s)","This method is used to show columns specified by the user.", 40,false,false,0,MethodListColumn + "12"),
                new templateStruct ("Show Row(s)","This method is used to show rows specified by the user.", 40,false,false,0,MethodListColumn + "13"),
                new templateStruct ("Sort Sheets","This method is used to alphabetize the sheets in a workbook.", 40,false,false,0,MethodListColumn + "14"),
                new templateStruct ("Show Worksheet","This method is used to unhide a worksheet specified by the user.", 50,false,false,0,MethodListColumn + "15"),
                new templateStruct ("Unfilter Table","This method removes all filters from a table specified by the user.", 40,false,false,0,MethodListColumn + "16")
            });

            UpdateExcelHandles();

            try
            {
                //DETERMINE IF THE "Macros" WORKSHEET EXISTS
                Microsoft.Office.Interop.Excel.Worksheet oWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)Sheets.get_Item("Macros");
            }
            catch (Exception ex)
            {
                string msg = ex.Message;

                // INSERT "Macros" SHEET
                Worksheet newMacrosSheet = null;
                newMacrosSheet = (Worksheet)Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                newMacrosSheet.Name = "Macros";

                InsertMacrosSheetButtonEnabledState(false);
                ScopeSelectionCBoxEnabledState(true);
                RefreshMacrosButtonEnabledState(true);
                MacroSelectionCBoxEnabledState(true);
            }

            // SETUP THE MACROS "TEMPLATE"
            // GET THE MACROS SHEET
            Sheet = (Microsoft.Office.Interop.Excel.Worksheet)Sheets.get_Item("Macros");
            curCell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range("A1");


            for (int i = 0; i < templateStructArray.Count; i++) // LOOP THROUGH LIST WITH FOR
            {
                curCell = (Microsoft.Office.Interop.Excel.Range)Sheet.get_Range(templateStructArray[i].Cell, System.Reflection.Missing.Value);
                if (templateStructArray[i].Border == true)
                {
                    curCell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    curCell.Font.Size = 14;
                    curCell.ColumnWidth = templateStructArray[i].ColumnWidth;
                }
                else
                {
                    curCell.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    curCell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    curCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                }
                curCell.set_Value(System.Reflection.Missing.Value, templateStructArray[i].Text);
                curCell.AddComment(templateStructArray[i].Comment);
                curCell.Comment.Shape.Height = templateStructArray[i].CommentHeight;

            }

            //EXIT PROCEDURE
            return;
        }
    }
}
