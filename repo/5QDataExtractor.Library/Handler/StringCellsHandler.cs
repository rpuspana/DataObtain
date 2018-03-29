using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _5QDataExtractor.Library.Handler
{
    public static class StringCellsHandler
    {
        public static void AggregateCells(string columnName, int tableColumnCount, 
                                          Interop.Range row2Cell, ref Dictionary<string, object> resultRow)
        {
            object currentCellOfResRow = resultRow[columnName];

            // convert to the new column data type
            string strCurrentCellOfResRow = Convert.ToString(currentCellOfResRow);

            object currRow2Cell = String.Empty;

            // if key yes - insert value in dict
            if (Transform.usrInputForAggregationProcess[columnName].IsAggregationKey)
            {
                // insert in result row the new column data type
                resultRow[columnName] = strCurrentCellOfResRow;
            }
            // if key no - perform user operation on dict cell and  + insert in dict
            else
            {
                // get the cell value in old column data type
                try
                {
                    currRow2Cell = Transform.GetCellValueOrUsrMappedVal(row2Cell, columnName);
                }
                catch (Exception ex)
                {
                    string oldColDT = Transform.usrInputForAggregationProcess[columnName].OldDataType;
                    string newColDT = Transform.usrInputForAggregationProcess[columnName].ColumnDataType;

                    throw new Exception($"Aggregation method on the tables's {columnName} Excel row {row2Cell.Row} value {row2Cell.Value2}: " +
                                        $"Could not convert value to {newColDT}. Agregation process stopped.");
                }

                // convert to new column data type - string
                string strRow2ColIvalue = Convert.ToString(currRow2Cell);

                switch (Transform.usrInputForAggregationProcess[columnName].RowsOperation)
                {
                    case "concatenate":
                        try
                        {
                            resultRow[columnName] = strCurrentCellOfResRow + strRow2ColIvalue;
                        }
                        catch (Exception ex)
                        {
                            string newColDT = Transform.usrInputForAggregationProcess[columnName].ColumnDataType;
                            throw new Exception($"Aggregation method on the tables's {columnName} between {newColDT} cell at row {row2Cell.Row} and {newColDT} previous aggregations result cell cound not be done." +
                                                $" Aggregation process stopped.");
                        }
                        break;

                    default:
                        string newColDataType = Transform.usrInputForAggregationProcess[columnName].ColumnDataType;
                        throw new Exception($"Aggregation method unknown on the tables's {columnName} between {newColDataType} cell at row {row2Cell.Row} and {newColDataType} previous aggregations result cell." +
                                            $" Aggregation process stopped.");
                }
            }
        }
    }
}
