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
    public static class NumberCellsHandler
    {
        public static void AggregateCells(string columnName, int tableColNum, 
                                          Interop.Range  row2Cell, ref Dictionary<string, object> resultRow)
        {
            decimal RezultCellToDecimal;

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
                    RezultCellToDecimal = -666;

                    // new column data type, conversion to new column data type can't be done
                    strCurrentCellOfResRow = String.Empty;
                }
                else
                {
                    // convet non-empty cell to new column data type
                    RezultCellToDecimal = ConvertToDecimalDataType(currentCellOfResRow);

                    // new column data type, conversion to new column data type can be done
                    strCurrentCellOfResRow = Convert.ToString(RezultCellToDecimal);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error when converting result row old column data type value, " +
                                    $"on table column {tableColNum} to the new column data type: " + ex.Message);
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

                decimal Row2CellToDecimal;

                try
                {
                    if (Convert.ToString(currenceRow2Cell) == String.Empty)
                    {
                        // dummy value
                        Row2CellToDecimal = -666;

                        // convertion to the new column data type can't be done
                        Row2CellCanBeConvToNewDT = false;
                    }
                    else
                    {
                        // new column data type
                        Row2CellToDecimal = ConvertToDecimalDataType(currenceRow2Cell);

                        // convertion to the new column data type can be done
                        Row2CellCanBeConvToNewDT = true;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error when converting old column data type for row 2 cell value, " +
                                        $"on table column {tableColNum} Excel row {row2Cell.Row}  to the new column data type: " +
                                        ex.Message);
                }

                resultRow[columnName] = ComputeResultOfTwoDecimals(Transform.usrInputForAggregationProcess[columnName].RowsOperation,
                                                                   RezultCellToDecimal, strCurrentCellOfResRow,
                                                                   Row2CellToDecimal, Row2CellCanBeConvToNewDT,
                                                                   columnName, row2Cell.Row);
            }
        }

        public static decimal ConvertToDecimalDataType(object elem)
        {
            try
            {
                return (Convert.ToDecimal(elem));
            }
            catch (OverflowException ex)
            {
                throw new OverflowException($"OverflowException: The {elem.GetType().Name} value {Convert.ToString(elem)} " +
                                            $"is out of range of the Decimal type.");
            }
            catch (FormatException ex)
            {
                throw new FormatException($"FormatException: The {elem.GetType().Name} value {Convert.ToString(elem)} " +
                                          $"is not recognized as a valid Decimal value.");
            }
            catch (InvalidCastException ex)
            {
                throw new InvalidCastException($"InvalidCastException: Conversion of the {elem.GetType().Name} value {Convert.ToString(elem)} " +
                                               $"to a Decimal is not supported.");
            }
        }

        public static object ComputeResultOfTwoDecimals(string operation, 
                                                        decimal val1, string canWorkWithResCell,
                                                        decimal val2, bool canWorkWithRow2Cell, 
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

            // the values we will operate on are not nulls and we can compute a decimal result
            switch (operation)
            {
                case "minimum":
                    try
                    {
                        // get the result of the result cell and the row 2 cell value
                        return Math.Min(val1, val2);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error when computing the decimal minimum of  " +
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
