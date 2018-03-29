using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using _5QDataExtractor.Library.Model;

namespace _5QDataExtractor.Library
{
    public static class Transform
    {
        // user input for each column from frmDEShowExcelTableColumnsDataType
        // Lookup Efficiency Key: O(1), Manipulate Efficiency O(1)
        public static Dictionary<string, ColDTModel> usrInputForAggregationProcess =
            new Dictionary<string, ColDTModel>();

        // mapping between excel column data type and user friendly
        // if all elements are the same data type, the column that data type is the common type
        // if not all are the same default to string data type
        public static Dictionary<string, string> ExcelDataTypeToUserDataTypeDict =
           new Dictionary<string, string>()
           {
                { "string"   , "string" },
                { "double"   , "number" },
                { "boolean"  , "true-false" },
                { "datetime" , "date and time" },
                { "null"     , "empty" }
           };

        // column data type conversion rules
        public static Dictionary<string, List<string>> PossibleColumnDataTypeChangesDict =
            new Dictionary<string, List<string>>()
            {
                // from excel default column data type to user selected column data type
                // see the booleans as strings = concatenate them, or convert them to double
                // boolean => double : TRUE = 1, FALSE = 0
                { "boolean"  , new List<string>() { "string", "double" } },

                // datetime can be converted to string = see the dates as strings and concatenate them
                { "datetime" , new List<string>() { "string" } },

                // double => string, consider the number as a string
                // double => boolean, number > 0 is TURE, number <= 0 is FALSE
                { "double"   , new List<string>() { "string", "boolean" } },

                // see certain strings as booleans = apply logic operations
                // a user input string is considered TRUE, and another input string is considered as FALSE
                 { "string"  , new List<string>() { "boolean" } }
            };

        // default row aggregation operation
        public static string defaultRowAggOperation = "no operation";

        // possible aggregate operations for a column's rows
        //public static List<string> RowAggregateOperationsLst = new List<string> { "no operation", "concatenate", "+", "-", "*", "/", "MIN", "MAX", "logic AND", "logic OR" };
        public static List<string> RowAggregateOperationsLst = new List<string> { "no operation", "concatenate", "minimum", "logic AND" };

        // column data type operations that can be performed on each column's rows
        public static Dictionary<string, List<string>> ColumnRowsOperations =
            new Dictionary<string, List<string>>()
            {
                // "no operation" = user hasn't selected anything
               
                //{ "boolean"  , new List<string>() { "logic AND", "logic OR" } },
                { "boolean"  , new List<string>() { "logic AND" } },
                //{ "datetime" , new List<string>() { "MIN", "MAX" } },
                { "datetime" , new List<string>() { "minimum" } },
                //{ "double"   , new List<string>() { "+", "-", "*", "/", "MIN", "MAX" } },
                { "double"   , new List<string>() { "minimum" } },
                { "null"     , new List<string>() { "no operation" } },
                { "string"   , new List<string>() { "concatenate" } }
            };

        private static bool CurrCellValExistsInUsrMapVals(Interop.Range cellVal, string columnName, out object valToWriteToCSV)
        {
            object val;

            var dict = Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict;

            if (cellVal.Value2 != null)
            {
                val = cellVal.Value2;
            }
            else
            {
                val = String.Empty;
            }

            if (Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict.Keys.ToList().Contains(val))
            {
                // now the value valToWriteToCSV is of the old column data type
                CellBoxUnbox.CellValueToObject(cellVal, out valToWriteToCSV);

                return true;
            }
            else
            {
                valToWriteToCSV = String.Empty;

                return false;
            }
        }

        /// <summary>
        /// Get the user replacement value or cell value converted to the column's old data type implemented as a .Net data structure
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static object GetCellValueOrUsrMappedVal(Interop.Range cell, string columnName)
        {
            object cellVal;
            object valToWriteToCSV = String.Empty;

            // see if the value in the cell is the same with the value in the cell val-user value dict.
            // if it is the value to write is the value in the dict
            // if it's not the value to write is the cell
            try
            {
                if (cell.Value2 != null)
                {
                    cellVal = cell.Value2;
                }
                else
                {
                    cellVal = String.Empty;
                }

                // the user might have entered something as a repalcement value for current cell value
                if (Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict != null &&
                    Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict.Keys.ToList().Contains(cellVal))
                {
                    List<object> dictVal = Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict[cellVal];

                    if ((Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict[cellVal] != null) &&
                        (dictVal.Count > 1))
                    {
                        if (Convert.ToString(Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict[cellVal][1]) == String.Empty)
                        {
                            // identify cell from the dict with cell unique values of the column and get it's string representation
                            if (CurrCellValExistsInUsrMapVals(cell, columnName, out valToWriteToCSV) == false)
                            {
                                throw new Exception($"Could not find Excell cell value \"{cell.Value2}\" from table column {columnName} Excel row {cell.Row} " +
                                                     "in the excel values - user values dictionary as a key to get it' string value. Could not write to CSV further.");
                            }
                        }
                        else
                        {
                            // get vaulue from dict whch will be of  old column data type
                            valToWriteToCSV = Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict[cellVal][1];
                        }
                    }
                    else
                    {
                        throw new Exception($"Implementation of the cell value - user value dictionary did not respect the pattern " +
                                            "(object) cell.Value2: List{ Convert.ToString(cell.Value2), (object) user value }" +
                                            $"when the application tried to reach the user value " +
                                            $"stored in the dictionary to replace the current cell value \"{cell.Value2}\" at table column {columnName} Excel row {cell.Row}. " +
                                            $"Can't continue writing values to the csv file.");
                    }
                }
                // if the user did not map any values
                else
                {
                    try
                    {
                        // now the value valToWriteToCSV is of the old column data type
                        CellBoxUnbox.CellValueToObject(cell, out valToWriteToCSV);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error at cell from table column {columnName} Excel row {cell.Row}  value \"{cell.Value2}\": " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error at table column {columnName} Excel row number {cell.Row} value \"{cell.Value2}\": " + ex.Message);
            }

            return valToWriteToCSV;
        }
    }
}
