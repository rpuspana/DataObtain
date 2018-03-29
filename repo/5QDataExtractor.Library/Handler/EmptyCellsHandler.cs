using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _5QDataExtractor.Library.Handler
{
    public static class EmptyCellsHandler
    {
        // at 14 March 2018 empty cells can't be mapped so we are using the old column data type
        public static void AggregateCells(string columnName, int tableColumnNum,
                                          Interop.Range row2Cell, ref Dictionary<string, object> resultRow)
        {
            // old column data type
            object currentCellOfResRow = resultRow[columnName];
            string strCurrentCellOfResRow = Convert.ToString(currentCellOfResRow);

            // old column data type
            object currenceRow2Cell = row2Cell.Value2;

            if (Transform.usrInputForAggregationProcess[columnName].IsAggregationKey)
            {
                resultRow[columnName] = strCurrentCellOfResRow;
            }
            // if key no - perform user operation on dict cell and  + insert in dict
            else
            {
                switch (Transform.usrInputForAggregationProcess[columnName].RowsOperation)
                {
                    case "no operation":
                        try
                        {
                            resultRow[columnName] = strCurrentCellOfResRow + Convert.ToString(currenceRow2Cell);
                        }
                        catch (Exception ex)
                        {
                            string newColDT = Transform.usrInputForAggregationProcess[columnName].ColumnDataType;
                            throw new Exception($"Could not concatenate empty cells on the tables's {columnName} between {newColDT} cell " +
                                                $"at row {row2Cell.Row} and {newColDT} result cell {strCurrentCellOfResRow}. {ex.Message}." +
                                                $" Aggregation process stopped.");
                        }

                        break;

                    default:
                        string newColDataType = Transform.usrInputForAggregationProcess[columnName].ColumnDataType;
                        throw new Exception($"Aggregation method unknown on the tables's {columnName} between {newColDataType} cell at row {row2Cell.Row} and" +
                                            $" {newColDataType} result cell {strCurrentCellOfResRow}. " +
                                            $"Aggregation process stopped.");
                }
            }
        }
    }
}
