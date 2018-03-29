using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace _5QDataExtractor.Library.Handler
{
    public class DateTimeCellsHandler
    {
        public static void AggregateCells(string columnName, int tableColumnNum,
                                          Interop.Range row2Cell, ref Dictionary<string, object> resultRow)
        {
            DateTime RezultCellToDateTime;

            // cell value or user mapped value of old column data type
            object currentCellOfResRow = resultRow[columnName];

            // new column data type
            string strCurrentCellOfResRow;

            try
            {
                // null cell value to new column data type
                if (Convert.ToString(currentCellOfResRow) == String.Empty)
                {
                    // dummy value
                    RezultCellToDateTime = new DateTime(1800, 12, 1);

                    // new column data type, conversion to new column data type can't be done
                    strCurrentCellOfResRow = String.Empty;
                }
                else
                {
                    // convet to new column data type
                    RezultCellToDateTime = CovertCellValuesToDateTime(currentCellOfResRow);

                    // new column data type, conversion to new column data type can be done
                    strCurrentCellOfResRow = Convert.ToString(RezultCellToDateTime);
                }
                    
            }
            catch (Exception ex)
            {
                throw new Exception($"Error when converting result row old column data type value, " +
                                    $"on table column {tableColumnNum} to the new column data type: " + ex.Message);
            }

            // if key yes - insert value in dict
            if (Transform.usrInputForAggregationProcess[columnName].IsAggregationKey)
            {
                // result row will now have new column data type
                resultRow[columnName] = strCurrentCellOfResRow;
            }
            // if key no - perform user operation on result cell and row2Cell and insert in dict
            else
            {
                bool Row2CellCanBeConvToNewDT;

                DateTime row2CellVal;

                // old column data type
                object currenceRow2Cell = Transform.GetCellValueOrUsrMappedVal(row2Cell, columnName);

                try
                {
                    // check if row 2 cell value is empty
                    if (Convert.ToString(currenceRow2Cell) == String.Empty)
                    {
                        // dummy value
                        row2CellVal = new DateTime(1800, 12, 1);

                        // convertion to the new column data type can't be done
                        Row2CellCanBeConvToNewDT = false;
                    }
                    else
                    {
                        // convert to new column data type
                        row2CellVal = CovertCellValuesToDateTime(currenceRow2Cell);

                        // convertion to the new column data type can be done
                        Row2CellCanBeConvToNewDT = true;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error when converting old column data type for row 2 cell value, " +
                                        $"on table column {columnName} Excel row {row2Cell.Row} to the new column data type: " +
                                        ex.Message);
                }

                resultRow[columnName] = ComputeResultOfTwoDates(Transform.usrInputForAggregationProcess[columnName].RowsOperation,
                                                                RezultCellToDateTime, strCurrentCellOfResRow,
                                                                row2CellVal, Row2CellCanBeConvToNewDT,
                                                                columnName, row2Cell.Row);
            }
        }

        public static DateTime CovertCellValuesToDateTime(object elem)
        {
            // if value contains os language specific number decimal separator
            // https://stackoverflow.com/questions/14513468/detect-decimal-separator
            //char decimalSep = Convert.ToChar(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);

            //if (Decimal.TryParse(Convert.ToString(elem), out decimal decValue))
            if (Double.TryParse(Convert.ToString(elem), out double decValue))
            {
                try
                {
                    //Excel behaves as if the date 1900-Feb-29 existed, it did not. So we must subtract 1
                    //return new DateTime(1899, 12, 31).AddDays(castedV - 1);
                    return new DateTime(1899, 12, 31).AddDays(decValue - 1);
                }
                catch (Exception ex)
                {
                    throw new Exception($"The {elem.GetType().Name} value {Convert.ToString(elem)}." +
                                        $"Could not add {decValue} days from 31.12.1899.");
                }
            }
            else
            {
                try
                {
                    return Convert.ToDateTime(elem);
                }
                catch (Exception ex)
                {
                    throw new Exception($"The {elem.GetType().Name} value {Convert.ToString(elem)} " +
                                        $"could not be converted to DateTime .Net data structure.");
                }
            }
        }

        public static object ComputeResultOfTwoDates(string operation,
                                                       DateTime val1, string canWorkWithResCell,
                                                       DateTime val2, bool canWorkWithRow2Cell,
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

            switch (operation)
            {
                case "minimum":
                    try
                    {
                        // DateTime structure also contains a Kind property, that is not retained in the new value. 
                        // This is normally not a problem; if you compare DateTime values of different kinds the comparison doesn't make sense anyway.
                        return (new DateTime(Math.Min(val1.Ticks, val2.Ticks)));
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Could not create a DateTime object from the earliest date/time between " +
                                            $"{val1.ToString()} and {val2.ToString()}. " +
                                            ex.Message + $". Aggregation process stopped.");
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
