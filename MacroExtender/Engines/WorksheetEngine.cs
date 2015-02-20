using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace MacroExtender
{
    class WorksheetEngine : EngineBase
    {

        /// <summary>
        /// Handles Worksheet scope methods.
        /// </summary>
        /// <param name="beginningRow">The line that the "Begin" method is located on.</param>
        public void Engine(int beginningRow)
        {
            Excel.Worksheet macroSheet = (Excel.Worksheet)Sheets.get_Item(getBeginningRowWorksheet(beginningRow));

            methodCell = (Excel.Range)Sheet.get_Range(MethodColumn + beginningRow);
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
                    case "End": // RELEASES THE DO LOOP TO END A MACRO
                        end = true;
                        break;

                    case "Filter Table": // USED TO FILTER A TABLE
                        string[] stringArray = buildStringArray(arg9, arg10, arg3, arg4, arg5, arg6, arg7, arg8);
                        filterTable(macroSheet, arg1, arg2, stringArray);
                        break;

                    case "Hide Columns": // USED TO HIDE ONE OR MORE WORKSHEET COLUMNS
                        macroSheet.get_Range(arg1 + ":" + arg2, Type.Missing).EntireColumn.Hidden = true;
                        break;

                    case "Hide Rows": // USED TO HIDE ONE OR MORE WORKSHEET ROWS
                        macroSheet.get_Range(arg1 + ":" + arg2, Type.Missing).EntireRow.Hidden = true;
                        break;

                    case "Input Box": // USED TO ALLOW THE USER TO INPUT RUNTIME VARIABLES
                        inputBox(arg2, arg1);
                        break;

                    case "Unfilter Table":
                        unfilterTable(macroSheet, arg1);
                        break;

                    case "Show Columns": // USED TO SHOW ONE OR MORE WORKSHEET COLUMNS
                        macroSheet.get_Range(arg1 + ":" + arg2, Type.Missing).EntireColumn.Hidden = false;
                        break;

                    case "Show Rows": // USED TO SHOW ONE OR MORE WORKSHEET ROWS
                        macroSheet.get_Range(arg1 + ":" + arg2, Type.Missing).EntireRow.Hidden = false;
                        break;

                } // switch (methodName)

                curRow++;

            } while (end == false); // THE DO LOOP WILL CONTINUE UNTIL THE End METHOD IS PASSED FROM THE MACRO

        } // private void WorksheetEngine(int row)

        /// <summary>
        /// This function returns a string that represents the worksheet that is altered by the macro.
        /// </summary>
        /// <param name="beginningRow">The "Begin" line of the macro.</param>
        /// <returns>String</returns>
        private string getBeginningRowWorksheet(int beginningRow)
        {
            Boolean foundScope = false;

            int curIndex = -1;

            do
            {
                curIndex++;
                if (MacrosList[curIndex].Row == beginningRow)
                {
                    foundScope = true;
                }
                //curIndex++;
            } while (foundScope == false);

            return MacrosList[curIndex].Worksheet;
        }

        /// <summary>
        /// This function returns a string array representing the non-null arguments that are passed to it.
        /// </summary>
        /// <param name="arg1">A string to be added to the return array.</param>
        /// <param name="arg2">A string to be added to the return array.</param>
        /// <param name="arg3">A string to be added to the return array.</param>
        /// <param name="arg4">A string to be added to the return array.</param>
        /// <param name="arg5">A string to be added to the return array.</param>
        /// <param name="arg6">A string to be added to the return array.</param>
        /// <param name="arg7">A string to be added to the return array.</param>
        /// <param name="arg8">A string to be added to the return array.</param>
        /// <returns>String[]</returns>
        private String[] buildStringArray(String arg1, String arg2, String arg3, String arg4,
                                           String arg5, String arg6, String arg7, String arg8)
        {

            int arg1Null = argNullStatus(arg1);
            int arg2Null = argNullStatus(arg2);
            int arg3Null = argNullStatus(arg3);
            int arg4Null = argNullStatus(arg4);
            int arg5Null = argNullStatus(arg5);
            int arg6Null = argNullStatus(arg6);
            int arg7Null = argNullStatus(arg7);
            int arg8Null = argNullStatus(arg8);


            int stringListIndex = 0;

            int nonNullCount = arg1Null
                                + arg2Null
                                + arg3Null
                                + arg4Null
                                + arg5Null
                                + arg6Null
                                + arg7Null
                                + arg8Null;

            String[] stringList = new String[nonNullCount];

            if (arg1Null != 0)
            {
                stringList[stringListIndex] = arg1;
                stringListIndex++;
            }

            if (arg2Null != 0)
            {
                stringList[stringListIndex] = arg2;
                stringListIndex++;
            }

            if (arg3Null != 0)
            {
                stringList[stringListIndex] = arg3;
                stringListIndex++;
            }

            if (arg4Null != 0)
            {
                stringList[stringListIndex] = arg4;
                stringListIndex++;
            }

            if (arg5Null != 0)
            {
                stringList[stringListIndex] = arg5;
                stringListIndex++;
            }

            if (arg6Null != 0)
            {
                stringList[stringListIndex] = arg6;
                stringListIndex++;
            }

            if (arg7Null != 0)
            {
                stringList[stringListIndex] = arg7;
                stringListIndex++;
            }

            if (arg8Null != 0)
            {
                stringList[stringListIndex] = arg8;
                stringListIndex++;
            }

            return stringList;
        }

        /// <summary>
        /// This method filters the specified table.
        /// </summary>
        /// <param name="worksheet">The worksheet that the table is on.</param>
        /// <param name="tableName">The name of the table to be filtered.</param>
        /// <param name="header">The name of the table field header to be filtered.</param>
        /// <param name="filterList">A string[] containing the values to filter the table field by.</param>
        private void filterTable(Microsoft.Office.Interop.Excel.Worksheet worksheet, String tableName, string header, String[] filterList)
        {
            int field = 1;
            bool foundHeader = false;
            foreach (Excel.Range cell in worksheet.ListObjects[tableName].HeaderRowRange)
            {
                if (header == cell.Value2)
                {
                    worksheet.ListObjects[tableName].Range.AutoFilter(field, filterList, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlFilterValues);
                    foundHeader = true;
                }
                if (foundHeader == true)
                    break;
                field++;
            }

        } // private void filterTable(Microsoft.Office.Interop.Excel.Worksheet worksheet, String tableName, int field, String[] filterList)

        /// <summary>
        /// This method removes the filters from a specified table.
        /// </summary>
        /// <param name="worksheet">The name of the worksheet that the table is located on.</param>
        /// <param name="table">The name of the table who's filters are to be removed.</param>
        private void unfilterTable(Microsoft.Office.Interop.Excel.Worksheet worksheet, String table)
        {
            int fieldCount = worksheet.ListObjects[table].ListColumns.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                worksheet.ListObjects[table].Range.AutoFilter(i);
            }


        }

        /// <summary>
        ///  This method is displays an input box, which is used collect user defined run-time variables.
        /// </summary>
        /// <param name="prompt">Used to set the prompt text for the input box.</param>
        /// <param name="userVariable">The user variable that the collected run time variable will be assigned to.</param>
        private void inputBox(string prompt, string userVariable)
        {
            Int32 userIndex = Convert.ToInt32(StringExt.Right(userVariable, 1));
            Boolean foundMatch = false;
            Int32 listIndex = 0;

            switch (StringExt.Left(userVariable, 3))
            {
                case "Int":
                    foreach (KeyValuePair<int, int> pair in userIntIndexRef)
                    {
                        if (pair.Key == userIndex)
                        {
                            foundMatch = true;
                            listIndex = pair.Value;
                        }
                    }

                    if (foundMatch == true)
                    {
                        userIntegerList[listIndex] = Convert.ToInt32(ExcelObj.InputBox(prompt, "Input Box", Type.Missing, Type.Missing,
                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing));
                    }
                    else
                    {
                        Int32 item = Convert.ToInt32(ExcelObj.InputBox(prompt, "Input Box", Type.Missing, Type.Missing,
                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                        userIntIndexRef.Add(userIndex, userIntegerList.Count());
                        userIntegerList.Add(item);
                    }

                    break;

                case "Str":
                    foreach (KeyValuePair<int, int> pair in userStrIndexRef)
                    {
                        if (pair.Key == userIndex)
                        {
                            foundMatch = true;
                            listIndex = pair.Value;
                        }
                    }

                    if (foundMatch == true)
                    {
                        userStringList[listIndex] = Convert.ToString(ExcelObj.InputBox(prompt, "Input Box", Type.Missing, Type.Missing,
                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing));
                    }
                    else
                    {
                        String item = Convert.ToString(ExcelObj.InputBox(prompt, "Input Box", Type.Missing, Type.Missing,
                                                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                        userStrIndexRef.Add(userIndex, userIntegerList.Count());
                        userStringList.Add(item);
                    }
                    break;
            }

        }

        /// <summary>
        /// This function determines whether or not an argument contains a a value.
        /// </summary>
        /// <param name="arg">Value to check.</param>
        /// <returns>int (0 - null)(1 - populated)</returns>
        private int argNullStatus(string arg)
        {
            int i = 0;
            if (arg != null)
            {
                i = 1;
            }

            return i;
        }
    }
}
