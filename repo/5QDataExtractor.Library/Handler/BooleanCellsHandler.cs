using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _5QDataExtractor.Library.Handler
{
    public static class BooleanCellsHandler
    {
        public static void AggregateCells(string columnName, int tableColumnNum, 
                                          Interop.Range row2Cell, ref Dictionary<string, object> resultRow)
        {
            bool RezultCellToBoolean;

            // old column data type
            object currentCellOfResRow = resultRow[columnName];

            // new column data type
            string strCurrentCellOfResRow;

            try
            {
                // null cell value to new column data type
                if (Convert.ToString(currentCellOfResRow) == String.Empty)
                {
                    // dummy value
                    RezultCellToBoolean = false;

                    // new column data type, conversion to new column data type can't be done
                    strCurrentCellOfResRow = String.Empty;
                }
                else
                {
                    // convet to new column data type
                    RezultCellToBoolean = ConvertToBooleanDataType(currentCellOfResRow);

                    // new column data type
                    strCurrentCellOfResRow = Convert.ToString(RezultCellToBoolean);
                }
               
            }
            catch (Exception ex)
            {
                throw new Exception($"Error when converting result row old column data type value, " +
                                    $"on table column {tableColumnNum}  to the new column data type: " +
                                    ex.Message);
            }

            // if key yes - insert value in dict
            if (Transform.usrInputForAggregationProcess[columnName].IsAggregationKey)
            {
                resultRow[columnName] = strCurrentCellOfResRow;
            }
            // if key no - perform user operation on result cell and row2Cell and insert in dict
            else
            {
                bool Row2CellCanBeConvToNewDT;

                // old column data type
                object currenceRow2Cell = Transform.GetCellValueOrUsrMappedVal(row2Cell, columnName);

                bool Row2CellToBoolean;

                try
                {
                    if (Convert.ToString(currenceRow2Cell) == String.Empty)
                    {
                        // dummy value
                        Row2CellToBoolean = false;

                        // convertion to the new column data type can't be done, f = false
                        Row2CellCanBeConvToNewDT = false;
                    }
                    else
                    {
                        // new column data type
                        Row2CellToBoolean = ConvertToBooleanDataType(currenceRow2Cell);

                        // convertion to the new column data type can be done
                        Row2CellCanBeConvToNewDT = true;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error when converting old column data type for row 2 cell value, " +
                                        $"on table column {tableColumnNum} Excel row {row2Cell.Row}  to the new column data type: " +
                                        ex.Message);
                }

                resultRow[columnName] = ComputeResultOfTwoBooleans(Transform.usrInputForAggregationProcess[columnName].RowsOperation,
                                                                   RezultCellToBoolean, strCurrentCellOfResRow,
                                                                   Row2CellToBoolean, Row2CellCanBeConvToNewDT,
                                                                   columnName, row2Cell.Row);
            }
        }

        public static bool ConvertToBooleanDataType(object elem)
        {
            try
            {
                return Convert.ToBoolean(elem);
            }
            catch (FormatException ex)
            {
                throw new FormatException($"FormatException: The {elem.GetType().Name} value {Convert.ToString(elem)} " +
                                          $"is not recognized as a valid boolean value.");
            }
            catch (InvalidCastException ex)
            {
                throw new InvalidCastException($"InvalidCastException: Conversion of the {elem.GetType().Name} value {Convert.ToString(elem)} " +
                                               $"to a boolean value is not supported.");
            }
        }

        public static object ComputeResultOfTwoBooleans(string operation,
                                                        bool val1, string canWorkWithResCell,
                                                        bool val2, bool canWorkWithRow2Cell,
                                                        string columnName, long row2CellExcelIndex)
        {
            string newColDataType = Transform.usrInputForAggregationProcess[columnName].ColumnDataType;

            // see if the values we will operate on are nulls and give back the appropiate result
            if ((canWorkWithResCell == String.Empty) && canWorkWithRow2Cell)
            {
                return val2;
            }

            // can't work with the result row value and can work with row 2 cell value
            if ((canWorkWithResCell != String.Empty) && !canWorkWithRow2Cell)
            {
                return val1;
            }

            if ((canWorkWithResCell == String.Empty) && !canWorkWithRow2Cell)
            {
                return String.Empty;
            }

            // the values we will operate on are not nulls and we can compute a boolean result
            switch (operation)
            {
                case "logic AND":
                    try
                    {
                        return (val1 && val2);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error when computing the logic AND between " +
                                            $"result row value \"{Convert.ToString(val1)}\" and the row2 cell value \"{Convert.ToString(val2)}\" " +
                                            $"at Excel row {row2CellExcelIndex} on the tables's {columnName}: " +
                                            $"{ ex.Message }. Aggregation process stopped.");
                    }

                default:
                    throw new Exception($"Aggregation method unknown on the tables's {columnName} between  " +
                                        $"the result row value {Convert.ToString(val1)} and " +
                                        $"the {newColDataType} row2 cell value\"{Convert.ToString(val2)}\" " +
                                        $"at Excel row {row2CellExcelIndex}. Aggregation process stopped.");
            }
        }
    }
}
