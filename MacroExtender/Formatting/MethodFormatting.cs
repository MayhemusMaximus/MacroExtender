using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace MacroExtender
{
    /// <summary>
    /// MethodStruct contains formatting data pertaining to dealing with each method and its' arguments.
    /// </summary>
    public struct MethodStruct
    {
        private readonly string methodName;
        private readonly bool hideArg1;
        private readonly bool hideArg2;
        private readonly bool hideArg3;
        private readonly bool hideArg4;
        private readonly bool hideArg5;
        private readonly bool hideArg6;
        private readonly bool hideArg7;
        private readonly bool hideArg8;
        private readonly bool hideArg9;
        private readonly bool hideArg10;
        private readonly string arg1Comment;
        private readonly string arg2Comment;
        private readonly string arg3Comment;
        private readonly string arg4Comment;
        private readonly string arg5Comment;
        private readonly string arg6Comment;
        private readonly string arg7Comment;
        private readonly string arg8Comment;
        private readonly string arg9Comment;
        private readonly string arg10Comment;

        /// <summary>
        /// This structure is used to define which arguments to show/hide and contain any comments for the arguments.
        /// </summary>
        /// <param name="methodName">The name of the method the record refers to.</param>
        /// <param name="hideArg1">True to blacken out Arg1, false to show it.</param>
        /// <param name="arg1Comment">Insert the comment associated with Arg1 here.</param>
        /// <param name="hideArg2">True to blacken out Arg2, false to show it.</param>
        /// <param name="arg2Comment">Insert the comment associated with Arg2 here.</param>
        /// <param name="hideArg3">True to blacken out Arg3, false to show it.</param>
        /// <param name="arg3Comment">Insert the comment associated with Arg3 here.</param>
        /// <param name="hideArg4">True to blacken out Arg4, false to show it.</param>
        /// <param name="arg4Comment">Insert the comment associated with Arg4 here.</param>
        /// <param name="hideArg5">True to blacken out Arg5, false to show it.</param>
        /// <param name="arg5Comment">Insert the comment associated with Arg5 here.</param>
        /// <param name="hideArg6">True to blacken out Arg6, false to show it.</param>
        /// <param name="arg6Comment">Insert the comment associated with Arg6 here.</param>
        /// <param name="hideArg7">True to blacken out Arg7, false to show it.</param>
        /// <param name="arg7Comment">Insert the comment associated with Arg7 here.</param>
        /// <param name="hideArg8">True to blacken out Arg8, false to show it.</param>
        /// <param name="arg8Comment">Insert the comment associated with Arg8 here.</param>
        /// <param name="hideArg9">True to blacken out Arg9, false to show it.</param>
        /// <param name="arg9Comment">Insert the comment associated with Arg9 here.</param>
        /// <param name="hideArg10">True to blacken out Arg10, false to show it.</param>
        /// <param name="arg10Comment">Insert the comment associated with Arg10 here.</param>
        public MethodStruct(string methodName,
                            bool hideArg1, string arg1Comment,
                            bool hideArg2, string arg2Comment,
                            bool hideArg3, string arg3Comment,
                            bool hideArg4, string arg4Comment,
                            bool hideArg5, string arg5Comment,
                            bool hideArg6, string arg6Comment,
                            bool hideArg7, string arg7Comment,
                            bool hideArg8, string arg8Comment,
                            bool hideArg9, string arg9Comment,
                            bool hideArg10, string arg10Comment)
        {
            this.methodName = methodName;
            this.hideArg1 = hideArg1;
            this.hideArg2 = hideArg2;
            this.hideArg3 = hideArg3;
            this.hideArg4 = hideArg4;
            this.hideArg5 = hideArg5;
            this.hideArg6 = hideArg6;
            this.hideArg7 = hideArg7;
            this.hideArg8 = hideArg8;
            this.hideArg9 = hideArg9;
            this.hideArg10 = hideArg10;
            this.arg1Comment = arg1Comment;
            this.arg2Comment = arg2Comment;
            this.arg3Comment = arg3Comment;
            this.arg4Comment = arg4Comment;
            this.arg5Comment = arg5Comment;
            this.arg6Comment = arg6Comment;
            this.arg7Comment = arg7Comment;
            this.arg8Comment = arg8Comment;
            this.arg9Comment = arg9Comment;
            this.arg10Comment = arg10Comment;
        }

        public string MethodName { get { return methodName; } }
        public bool HideArg1 { get { return hideArg1; } }
        public bool HideArg2 { get { return hideArg2; } }
        public bool HideArg3 { get { return hideArg3; } }
        public bool HideArg4 { get { return hideArg4; } }
        public bool HideArg5 { get { return hideArg5; } }
        public bool HideArg6 { get { return hideArg6; } }
        public bool HideArg7 { get { return hideArg7; } }
        public bool HideArg8 { get { return hideArg8; } }
        public bool HideArg9 { get { return hideArg9; } }
        public bool HideArg10 { get { return hideArg10; } }
        public string Arg1Comment { get { return arg1Comment; } }
        public string Arg2Comment { get { return arg2Comment; } }
        public string Arg3Comment { get { return arg3Comment; } }
        public string Arg4Comment { get { return arg4Comment; } }
        public string Arg5Comment { get { return arg5Comment; } }
        public string Arg6Comment { get { return arg6Comment; } }
        public string Arg7Comment { get { return arg7Comment; } }
        public string Arg8Comment { get { return arg8Comment; } }
        public string Arg9Comment { get { return arg9Comment; } }
        public string Arg10Comment { get { return arg10Comment; } }
    }
    class MethodFormatting : InteractiveFormatting
    {
        #region Fields and Properties

        // A FEW COLORS TO CALL BY NAME
        System.Drawing.Color Black = System.Drawing.Color.Black;
        System.Drawing.Color Green = System.Drawing.Color.Green;
        System.Drawing.Color Red = System.Drawing.Color.Red;

        private int hideColor = 16;

        /// <summary>
        /// This list defines which arguments to show/hide and contain any comments for the arguments.
        /// </summary>
        List<MethodStruct> methodStruct = new List<MethodStruct>
            (new[]{
                new MethodStruct("Activate Worksheet",
                    false,"Insert the name of the worksheet to be activated here.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("BAD METHOD", 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty),

                new MethodStruct("Begin", 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty),

                new MethodStruct("Close Workbook",
                    false,"Insert the name of the workbook to be closed here.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Create Email",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Filter Table",
                    false,"Insert the name of the table to be filtered here.",
                    false,"Insert the name of the table header to be filtered here.",
                    false,"Insert the/a filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here.",
                    false,"(Optional) Insert another filter to be applied to the table column here."),

                new MethodStruct("Hide Columns", 
                    false, "Insert the name of the first column to hide here.", 
                    false, "Insert the name of the last column to hide here; if there is only one column to hide, insert the name of the first column here, also.", 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty, 
                    true, string.Empty),

                new MethodStruct("Hide Rows",
                    false,"Insert the first row to hide here.",
                    false,"Insert the last row to hide here; if there is only one row to hide, insert the first row here, also.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Hide Worksheet",
                    false,"Insert the name of the worksheet to hide here.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Open Workbook",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Input Box",
                    false,"Insert the name of the user variable that will be assigned the value collected by the input box here.",
                    false,"Insert a prompt for the input box here; the prompt serves as a reminder of the purpose of input box when it is displayed.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Run Macro",
                    false,"Insert the name of the macro to be run here.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Show Columns",
                    false,"Insert the name of the first column to be shown here.",
                    false,"Insert the name of the last column to shown here; if there is only one column to show, insert the name of the first column here, also.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Show Rows",
                    false,"Insert the name of the first row to be shown here.",
                    false,"Insert the name of the last row to shown here; if there is only one row to show, insert the name of the first row here, also.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Show Worksheet",
                    false, "Insert the name of the worksheet to be shown here.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Sort Sheets",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("Unfilter Table",
                    false,"Insert the name of the table to be unfiltered here.",
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty,
                    true,string.Empty),

                new MethodStruct("WHITE SPACE",
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty,
                    false,string.Empty)
            });

        #endregion

        /// <summary>
        /// This method checks for syntax, then calls methods to execute formatting according to the results of the check.
        /// </summary>
        /// <param name="target"></param>
        public void delegateMethodFormatting(Microsoft.Office.Interop.Excel.Range target)
        {
            // THE FOCUS OF THE METHOD COLUMNS WILL BE
            // CORRECT SYNTAX (I.E. SPELLING OF METHOD NAMES), AND
            // THE ARGUMENTS, THEIR COMMENTS, AND
            // WHETHER OR NOT TO HIDE THEM.

            int methodRow = target.Row;
            Excel.Range argCell = (Excel.Range)Sheet.get_Range(Arg1Column + methodRow);

            string cellValue = target.Value2;
            switch (cellValue)
            {
                case "Activate Worksheet":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Begin":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Close Workbook":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Create Email":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "End":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Filter Table":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Hide Columns":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Hide Rows":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Hide Worksheet":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Input Box":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Open Workbook":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Run Macro":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Show Columns":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Show Rows":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Sort Sheets":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Show Worksheet":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case "Unfilter Table":
                    setTargetFontColor(target, Black);
                    methodFormatting(methodRow, cellValue);
                    break;

                case null:
                    methodFormatting(methodRow, "WHITE SPACE");
                    break;

                default:
                    setTargetFontColor(target, Red);
                    methodFormatting(methodRow, "BAD METHOD");
                    break;
            }
        }

        /// <summary>
        /// This method sets formatting conditions according to the definitions provided by the methodStruct[].
        /// </summary>
        /// <param name="row">This is the row to be formatted.</param>
        /// <param name="methodName">The methodStruct[].methodName who's values are to be applied to the given row.</param>
        private void methodFormatting(int row, string methodName)
        {
            int index = -1;
            for (int x = 0; x < methodStruct.Count; x++)
            {
                if (methodStruct[x].MethodName == methodName)
                {
                    index = x;
                    break;
                }
            }

            if (index != -1)
            {
                HighlightCellIf(methodStruct[index].HideArg1, Arg1Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg2, Arg2Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg3, Arg3Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg4, Arg4Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg5, Arg5Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg6, Arg6Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg7, Arg7Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg8, Arg8Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg9, Arg9Column, row, hideColor);
                HighlightCellIf(methodStruct[index].HideArg10, Arg10Column, row, hideColor);

                AddCommentIf(methodStruct[index].Arg1Comment != string.Empty, Arg1Column, row, methodStruct[index].Arg1Comment);
                AddCommentIf(methodStruct[index].Arg2Comment != string.Empty, Arg2Column, row, methodStruct[index].Arg2Comment);
                AddCommentIf(methodStruct[index].Arg3Comment != string.Empty, Arg3Column, row, methodStruct[index].Arg3Comment);
                AddCommentIf(methodStruct[index].Arg4Comment != string.Empty, Arg4Column, row, methodStruct[index].Arg4Comment);
                AddCommentIf(methodStruct[index].Arg5Comment != string.Empty, Arg5Column, row, methodStruct[index].Arg5Comment);
                AddCommentIf(methodStruct[index].Arg6Comment != string.Empty, Arg6Column, row, methodStruct[index].Arg6Comment);
                AddCommentIf(methodStruct[index].Arg7Comment != string.Empty, Arg7Column, row, methodStruct[index].Arg7Comment);
                AddCommentIf(methodStruct[index].Arg8Comment != string.Empty, Arg8Column, row, methodStruct[index].Arg8Comment);
                AddCommentIf(methodStruct[index].Arg9Comment != string.Empty, Arg9Column, row, methodStruct[index].Arg9Comment);
                AddCommentIf(methodStruct[index].Arg10Comment != string.Empty, Arg10Column, row, methodStruct[index].Arg10Comment);
            }
        }

        /// <summary>
        /// This method will highlight a cell if the hide parameter is true; otherwise it will be made white (Excel "No Fill").
        /// </summary>
        /// <param name="highlight">True to highlight, false to "No Fill".</param>
        /// <param name="column">The column that the cell is in.</param>
        /// <param name="row">The row that the cell is in.</param>
        private void HighlightCellIf(bool highlight, string column, int row, int color)
        {
            if (highlight)
                setTargetBackgroundColor((Excel.Range)Sheet.get_Range(column + row), color);
            else
                Sheet.get_Range(column + row).Interior.ColorIndex = 0;
        }

        /// <summary>
        /// This method will remove a comment from the provided cell if one exists, and add a comment if the "add" argument evaluates to true.
        /// </summary>
        /// <param name="add">True to add a comment.</param>
        /// <param name="column">The column that the cell is in.</param>
        /// <param name="row">The row that the cell is in.</param>
        /// <param name="comment">The comment to be added if the "add" parameter evaluates to true.</param>
        private void AddCommentIf(bool add, string column, int row, string comment)
        {
            Sheet.get_Range(column + row).ClearComments();

            if (add)
                Sheet.get_Range(column + row).AddComment(comment);
        }

        /// <summary>
        /// This method will set the font color of the range that is passed to it.
        /// </summary>
        /// <param name="target">The range who's font will be changed.</param>
        /// <param name="color">The color to change the font to.</param>
        private void setTargetFontColor(Excel.Range target, System.Drawing.Color color)
        {
            target.Font.Color = System.Drawing.ColorTranslator.ToOle(color);
        }

        /// <summary>
        /// This method will highlight the cell(s) in the provided range the color provided.
        /// </summary>
        /// <param name="target">The range to be highlighted.</param>
        /// <param name="color">The Excel color index number representing the color to highlight the range.</param>
        private void setTargetBackgroundColor(Excel.Range target, int color)
        {
            target.Interior.ColorIndex = color;
        }
    }
}
