using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.CodeDom.Compiler;
using System.CodeDom;
using System.Text.RegularExpressions;
using _5QDataExtractor.Library.Model;
using _5QDataExtractor.Library.Parser;
using System.Runtime.InteropServices;
using _5QDataExtractor.Library.Handler;

namespace _5QDataExtractor.Library.DataAccess
{
    public static class TextProcessor
    {
        private static string AppDirectoryName = "5QDataExtractor";
        private static string CSVoutFolderName = "out";

        private static string windowsListSeparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;

        // path to My Documents
        private static string MyDocumentsPath = String.Empty;

        // path to app out folder
        private static string PathToAppOutFolder = String.Empty;

        private static string filePath = String.Empty;

        public static void SaveExcelTableToCSVFile(Interop.ListObject excelTable, bool boolVarTableHeaderInExport, char csvDelimiter,
                                                   bool aggregateRows, Workbook wb, Worksheet activeWS)
        {
            string CSVfileName = String.Empty;

            if (aggregateRows == false)
            {
                CSVfileName = excelTable.Name + "_to_csv.csv";
            }
            else
            {
                CSVfileName = excelTable.Name + "_aggregate_cells_to_csv.csv";
            }
            
            try
            {
                CreateAppFolderOutFolderAndCSVFile(excelTable, CSVfileName);

                if (boolVarTableHeaderInExport)
                {
                    WriteTableHeaderRangeToCSV(excelTable, csvDelimiter);
                }

                if (aggregateRows == false)
                {
                    WriteTableBodyRangeToCSV(excelTable, csvDelimiter);
                }
                else
                {
                    AggregateTableRowsAndWriteToCSVfile(excelTable, csvDelimiter, wb, activeWS);
                }

                MessageBox.Show($"Table {excelTable.Name} exported to {filePath} successfully!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.Diagnostics.Process.Start(filePath);
            }
            catch(Exception ex)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                }
                catch (Exception fileEx)
                {
                    throw new Exception($"Error when checking the existance of the csv file path {filePath} and trying to delete it " +
                                        $"because in the process copying to csv file/ aggregating and copying to csv file en error came through: " +
                                        $"{fileEx.Message}");
                }

                throw new Exception($"Error: {ex.Message}. Table {excelTable.Name} was not exported to csv.");
            }
        }

        private static void CreateAppFolderOutFolderAndCSVFile(Interop.ListObject excelTable, string csvFileName)
        {
            // path to My Documents
            MyDocumentsPath = Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents");

            // path to app out folder
            PathToAppOutFolder = Path.Combine(MyDocumentsPath, AppDirectoryName, CSVoutFolderName);

            // create directory to app out folder
            Directory.CreateDirectory(Path.Combine(MyDocumentsPath, AppDirectoryName, CSVoutFolderName));
            //Directory.CreateDirectory(Path.Combine("Z:", AppDirectoryName, CSVoutFolderName));

            // path to the actual csv file where the contents of the Excel table will be dumped
            filePath = Path.Combine(PathToAppOutFolder, csvFileName);
        }

        private static void WriteTableHeaderRangeToCSV(Interop.ListObject excelTable, char csvDelimiter)
        {
            Interop.Range tableHeaderRange = excelTable.HeaderRowRange;

            string Headerline = String.Empty;
            int columnInTableCount = 1;

            foreach (Interop.Range cell in tableHeaderRange)
            {
                if (columnInTableCount == 1)
                {
                    // assume the header is not null or empty
                    //Headerline = Escape(Convert.ToString(cell.Value2), csvDelimiter);
                    Headerline = EscapeCSVstyle(Convert.ToString(cell.Text), csvDelimiter);
                }
                else
                {
                    // assume the header is not null or empty
                    //Headerline = String.Join(Convert.ToString(csvDelimiter), Headerline, Escape(Convert.ToString(cell.Value2), csvDelimiter));
                    Headerline = String.Join(Convert.ToString(csvDelimiter), Headerline, EscapeCSVstyle(Convert.ToString(cell.Text), csvDelimiter));
                }

                columnInTableCount++;
            }

            Headerline += Environment.NewLine;

            File.WriteAllText(filePath, Headerline);

            if (tableHeaderRange != null) Marshal.ReleaseComObject(tableHeaderRange);
        }

        private static void WriteTableBodyRangeToCSV(Interop.ListObject excelTable, char csvDelimiter)
        {
            object oldColDTVal;
            object newColDTVal;
            string columnName;

            Interop.Range tableBodyRange = excelTable.DataBodyRange;

            foreach (Interop.Range row in tableBodyRange.Rows)
            {
                string line = String.Empty;
                int cellInRowCount = 1;

                foreach (Interop.Range cell in row.Columns)
                {
                    columnName = "Column " + cellInRowCount;

                    // old column data type
                    oldColDTVal = Transform.GetCellValueOrUsrMappedVal(cell, columnName);

                    if (oldColDTVal.ToString() != String.Empty)
                    {
                        // if the user selected changed the old column data type
                        if (Transform.usrInputForAggregationProcess[columnName].OldDataType != Transform.usrInputForAggregationProcess[columnName].ColumnDataType)
                        {
                            // the cell might be empty and can not convert to the new column data type
                            // in this case the newColDTVal = String.Empty
                            CellBoxUnbox.CastToCellType(oldColDTVal,
                                                   Transform.usrInputForAggregationProcess[columnName].ColumnDataType,
                                                   true,
                                                   out newColDTVal);
                        }
                        else
                        {
                            newColDTVal = oldColDTVal;
                        }
                    }
                    else
                    {
                        newColDTVal = String.Empty;
                    }

                    if (cellInRowCount == 1)
                    {
                        line = EscapeCSVstyle(Convert.ToString(newColDTVal), csvDelimiter);
                    }
                    // column 2 onwards per string
                    else
                    {
                        line = String.Join(Convert.ToString(csvDelimiter), line, EscapeCSVstyle(Convert.ToString(newColDTVal), csvDelimiter));
                    }

                    cellInRowCount++;
                }

                line += Environment.NewLine;

                File.AppendAllText(filePath, line);
            }

            if (tableBodyRange != null) Marshal.ReleaseComObject(tableBodyRange);
        }

        private static object GetCellValueOrUsrMappedVal(Interop.Range cell, string columnName )
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
                            if (CurrCellValExistsInUsrMapVals(cellVal, columnName, out valToWriteToCSV) == false)
                            {
                                throw new Exception($"Could not find Excell cell value \"{cell.Value2}\" from table column {columnName} Excel row {cell.Row} " +
                                                     "in the excel values - user values dictionary as a key to get it' string value. Could not write to CSV further.");
                            }
                        }
                        else
                        {
                            // get vaulue from dict
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
                        // identify cell from the dict with cell unique values of the column and get it's string representation
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

        private static bool CurrCellValExistsInUsrMapVals(object cellVal, string columnName, out object valToWriteToCSV)
        {
            var dict = Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict;

            if (Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict.Keys.ToList().Contains(cellVal))
            {
                valToWriteToCSV = Transform.usrInputForAggregationProcess[columnName].ExCellValToUserValDict[cellVal][0];
                return true;
            }
            else
            {
                valToWriteToCSV = String.Empty;

                return false;
            }

            //foreach (object key in dict.Keys)
            //{
            //    if (cellVal == key)
            //    {
            //        valToWriteToCSV = dict[key][0];
            //        return true;
            //    }
            //}

            //valToWriteToCSV = String.Empty;

            //return false;
        }

        private static void AggregateTableRowsAndWriteToCSVfile(Interop.ListObject excelTable, char csvDelimiter, Workbook wb, Worksheet activeWS)
        {
            Interop.Range tableBodyRange = excelTable.DataBodyRange;

            // table "row" with a result from applying a two row operation
            string excelTableRow1Str = String.Empty;

            // line to write to csv file
            string csvLine = String.Empty;

            //// result row (aggregated rows)
            //Dictionary<string, object> resultRow = new Dictionary<string, object>();

            // excel row or result row (aggregated rows) as string
            string resultStr = String.Empty;

            int lastColumnInTableExcelIndex = 0;
            int lastRowInTableExcelIndex = 0;

            GetLastExcelRowAndLastColumn(tableBodyRange, ref lastColumnInTableExcelIndex, ref lastRowInTableExcelIndex);

            int firstColumnInTableExcelIndex = lastColumnInTableExcelIndex - (tableBodyRange.Columns.Count - 1);
            int firstRowInTableExcelIndex = lastRowInTableExcelIndex - (tableBodyRange.Rows.Count - 1);

            List<string> columnKeysValuesVisited = new List<string>();

            // values from each cell corresponding to a key column united by "~"
            string row1KeyColumnValuesStr = String.Empty;

            foreach (Interop.Range row1 in tableBodyRange.Rows)
            {
                // result row (aggregated rows)
                Dictionary<string, object> resultRow = new Dictionary<string, object>();

                // transform row1 into a string and get the value from each key column in a string separated with ~
                excelTableRow1Str = GetCellsInRowInOneStr(row1, ref row1KeyColumnValuesStr, csvDelimiter);

                // if the key column values are not in this list
                if (!columnKeysValuesVisited.Contains(row1KeyColumnValuesStr))
                {
                    foreach (Interop.Range row2 in tableBodyRange.Rows)
                    {
                        //good if (row1.Row != row2.Row)
                        if (row1.Row != row2.Row && row2.Row > row1.Row)
                        {
                            //if tableBodyRange.Cells[row1,] &&  &&
                            if (AllCellValuesFromKeyColumnsAreEqual(excelTableRow1Str, row2,
                                                                    ref columnKeysValuesVisited,
                                                                    csvDelimiter))
                            {
                                if (resultRow.Keys.Count == 0)
                                {
                                    // copy row1 in resultRow. all values are the old column data type
                                    CopyRowToDictionary(row1, ref resultRow);
                                }

                                // column aggregating keys found, 
                                // aggregating result row and row2 and dump result to result row
                                // escape to CSV " to "" for strings
                                AggregateRows(row2, ref resultRow,
                                            firstColumnInTableExcelIndex,
                                            lastColumnInTableExcelIndex,
                                            firstRowInTableExcelIndex,
                                            lastRowInTableExcelIndex,
                                            wb, activeWS);

                            }
                        }
                    }

                    // we didn't aggregate rows
                    if (resultRow.Keys.Count == 0)
                    {
                        // escape the value to CSV: " is escaped as ""
                        ExcelRowToString(row1, csvDelimiter, ref csvLine);
                    }
                    else
                    {
                        // convert result list to string
                        // enclose in "..."
                        JoinDictElemsInStr(row1, ref resultRow, csvDelimiter, ref csvLine);

                        // we don't need the result row anymore as we will need a new result row
                        resultRow = null;
                    }

                    csvLine += Environment.NewLine;

                    // write the string version of result row / row1 to file
                    File.AppendAllText(filePath, csvLine);

                    csvLine = String.Empty;
                }

                row1KeyColumnValuesStr = null;
            }

            if (tableBodyRange != null) Marshal.ReleaseComObject(tableBodyRange);
        }

        // we didn't aggregate rows
        private static void ExcelRowToString(Interop.Range row, char csvDelimiter, ref string csvLine)
        {
            string strToAppend;
            int tableColumNum = 1;
            string currentColName;
            object cellValObj;
            object newColDTVal;
            string cellValStr;

            StringBuilder resultRowStrBuilder = new StringBuilder();

            try
            {
                foreach (Interop.Range cell in row.Columns)
                {
                    currentColName = "Column " + tableColumNum;

                    // old val column data type
                    cellValObj = Transform.GetCellValueOrUsrMappedVal(cell, currentColName);

                    // if the user selected changed the old column data type
                    if (Transform.usrInputForAggregationProcess[currentColName].OldDataType != 
                        Transform.usrInputForAggregationProcess[currentColName].ColumnDataType)
                    {
                        // new column data type
                        // default to "" when value can't be converted to the new column data type
                        CellBoxUnbox.CastToCellType(cellValObj,
                                               Transform.usrInputForAggregationProcess[currentColName].ColumnDataType,
                                               true,
                                               out newColDTVal);
                    }
                    else
                    {
                        // old column data type because user didn't change
                        newColDTVal = cellValObj;
                    }

                    cellValStr = Convert.ToString(newColDTVal);

                    // row was not escaped to CSV before
                    // CSV: replace " with "" and enclose string in "..."
                    if (cellValStr != String.Empty)
                    {
                        strToAppend = "\"" + Convert.ToString(newColDTVal).Replace("\"", "\"\"") + "\"";
                        resultRowStrBuilder.Append(csvDelimiter + strToAppend);

                        //switch (Transform.usrInputForAggregationProcess[currentColName].ColumnDataType)
                        //{
                        //    case "datetime":
                        //        DateTime dateTimeVal;
                        //        try
                        //        {
                        //            dateTimeVal = Convert.ToDateTime(cellValObj);
                                    
                        //        }
                        //        catch (Exception ex)
                        //        {
                        //            throw new Exception($"Error when converting Excel value/user mapped value {Convert.ToString(cellValObj)} corresponding to cell " +
                        //                                $"from table column {tableColumNum} Excel row {row.Row}: " + ex.Message);
                        //        }

                        //        strToAppend = "\"" + Convert.ToString(dateTimeVal).Replace("\"", "\"\"") + "\"";
                        //        resultRowStrBuilder.Append(csvDelimiter + strToAppend);

                        //        break;

                        //    default:
                        //        strToAppend = "\"" + Convert.ToString(cellValObj).Replace("\"", "\"\"") + "\"";
                        //        resultRowStrBuilder.Append(csvDelimiter + strToAppend);
                        //        break;
                        //}
                    }
                    else
                    {
                        resultRowStrBuilder.Append(csvDelimiter + String.Empty);
                    }

                    tableColumNum++;
                }

                resultRowStrBuilder.Remove(0, 1);
                    
                csvLine = resultRowStrBuilder.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error in the process of converting cells value/mapped values of table column {tableColumNum} row {row.Row} " +
                                $"to their string representation: {ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw ex;
            }
        }

        private static void JoinDictElemsInStr(Interop.Range row1, ref Dictionary<string, object> resultRow, char csvDelimiter, ref string csvLine)
        {
            if (resultRow.Keys.Count > 0)
            {
                StringBuilder resultRowStrBuilder = new StringBuilder();
                string csvEscapedStr;

                try
                {
                    foreach (string key in resultRow.Keys)
                    {
                        // check for empty cells which were stored as null
                        if (resultRow[key] != null)
                        {
                            // dict elemetns already had " replaced with ""
                            csvEscapedStr = "\"" + Convert.ToString(resultRow[key]).Replace("\"","\"\"") + "\"";
                            resultRowStrBuilder.Append(csvDelimiter + csvEscapedStr);
                        }
                        else
                        {
                            resultRowStrBuilder.Append(csvDelimiter + String.Empty);
                        }
                            
                    }

                    resultRowStrBuilder.Remove(0, 1);

                    csvLine = resultRowStrBuilder.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error when converting resultRow elements to string after comparing row {row1.Row} " +
                                    $"with the table rows below it. Application will now exit.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw ex;
                }
            }
        }

        private static void CopyRowToDictionary(Interop.Range row, ref Dictionary<string, object> dict)
        {
            object cellValObj;
            string columnName;

            if (row != null && dict != null)
            {
                int tableColumnNum = 1;

                foreach (Interop.Range cell in row.Columns)
                {
                    columnName = "Column " + tableColumnNum;

                    // getting the old column data type
                    cellValObj = GetCellValueOrUsrMappedVal(cell, columnName);

                    dict.Add(columnName, cellValObj);

                    tableColumnNum++;
                }
            }
        }

        private static void AggregateRows(Interop.Range row2, ref Dictionary<string, object> resultRow,
                                          int firstColumnInTableExcelIndex, int lastColumnInTableExcelIndex,
                                          int firstRowInTableExcelIndex, int lastRowInTableExcelIndex, Workbook wb, Worksheet activeWS)
        {
            try
            {
                int tableColumnNum = 1;

                string columnName;

                foreach (Interop.Range row2Cell in row2.Columns)
                {
                    columnName = "Column " + tableColumnNum;

                    if (Transform.usrInputForAggregationProcess.Keys.Contains(columnName))
                    {
                        // do the operations on the same column for both rows based on the new column type
                        // we know that each cell of the column can be converted to this type because Validator.AllTblColsCanBeParsedToTheNewType()
                        // passed

                        // convert cell from result and cell from row2 to the new column data type
                        switch (Transform.usrInputForAggregationProcess[columnName].ColumnDataType)
                        {
                            case "string":
                                // escape to CSV " to ""
                                StringCellsHandler.AggregateCells(columnName, tableColumnNum, row2Cell, ref resultRow);
                                break;

                            case "double":
                                NumberCellsHandler.AggregateCells(columnName, tableColumnNum, row2Cell, ref resultRow);
                                break;

                            case "boolean":
                                 BooleanCellsHandler.AggregateCells(columnName, tableColumnNum, row2Cell, ref resultRow);
                                break;

                            case "datetime":
                                DateTimeCellsHandler.AggregateCells(columnName, tableColumnNum, row2Cell, ref resultRow);
                                break;

                            // 16 March 2018 - cells can't be converted to a "null" column data type
                            //case "null":
                            //    EmptyCellsHandler.AggregateCells(columnName, tableColumnNum, row2Cell, ref resultRow);
                            //    break;

                            default:
                                throw new Exception($"Can't handle aggregation operation for column data type {Transform.usrInputForAggregationProcess[columnName].ColumnDataType}.");
                        }
                    }

                    tableColumnNum++;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static bool AllCellValuesFromKeyColumnsAreEqual(string row1, Interop.Range row2, 
                                                                ref List<string> columnKeysValuesVisited,
                                                                char valSeparator)
        {
            List<string> temp = new List<string>();

            object row2CellObject;
            string row2Val = String.Empty;


            string[] row1Values = row1.Split('~');

            if (row1Values.Length == row2.Columns.Count)
            {
                string dictKey = String.Empty;
                long idxOfColumnInTable = 1;

                foreach (Interop.Range row2Cell in row2.Columns)
                {
                    dictKey = $"Column { idxOfColumnInTable }";

                    // look for the keys
                     if (Transform.usrInputForAggregationProcess.Keys.Contains(dictKey))
                     {
                        if (Transform.usrInputForAggregationProcess[dictKey].IsAggregationKey)
                        {
                            row2CellObject = GetCellValueOrUsrMappedVal(row2Cell, "Column " + idxOfColumnInTable);
                            row2Val = Convert.ToString(row2CellObject);

                            if (row2Val != String.Empty)
                            {
                                bool mustQuote = (row2Val.Contains(valSeparator) || row2Val.Contains("\"") || row2Val.Contains("\'") ||
                                                  row2Val.Contains("\\") || row2Val.Contains("/") || row2Val.Contains("\0") ||
                                                  row2Val.Contains("\a") || row2Val.Contains("\b") || row2Val.Contains("\f") ||
                                                  row2Val.Contains("\n") || row2Val.Contains("\r") || row2Val.Contains("\t") ||
                                                  row2Val.Contains("\v"));
                                if (mustQuote)
                                {
                                    row2Val = Parser.Parser.ToLiteral(row2Val);
                                }
                                    
                            }
                            else
                            {
                                row2Val = "`";
                            }

                            if (Transform.usrInputForAggregationProcess[dictKey].ColumnDataType == "boolean")
                            {
                                // each value has been C# style escaped and because we convert to boolean
                                // the case of strings doesn't matter
                                if (row1Values[idxOfColumnInTable - 1].ToLower() != row2Val.ToLower())
                                {
                                    temp = null;

                                    // not all key values are equal
                                    return false;
                                }

                                // add TRUE/FALSE visited key value  as true/false visited key value because we want to aggregate all boolean values
                                temp.Add(row1Values[idxOfColumnInTable - 1].ToLower());
                            }
                            else
                            {
                                if (row1Values[idxOfColumnInTable - 1] != row2Val)
                                {
                                    temp = null;

                                    // not all key values are equal
                                    return false;
                                }

                                // add TRUE/FALSE visited key value as true/false visited key value because the case of the keys matter
                                temp.Add(row1Values[idxOfColumnInTable - 1]);
                            }

                            //temp.Add(row1Values[idxOfColumnInTable - 1]);
                        }
                     }

                    idxOfColumnInTable++;
                }

                // add the current visited column key val 1, column key val 2, ..., column key value N
                columnKeysValuesVisited.Add(String.Join("~", temp));

                // all key values are equal
                return true;
            }

            // row1 and row 2 don't have the same number of columns
            return false;
        }

        private static string GetCellsInRowInOneStr(Interop.Range excelTableRow, ref string row1KeyColumnValuesStr, char csvDelimiter)
        {
            string excelRowAsStr = String.Empty;
            string cellValue;
            int columnInTable = 1;
            object cellValOrUsrMappedVal;

            foreach (Interop.Range cell in excelTableRow.Columns)
            {
                cellValOrUsrMappedVal = GetCellValueOrUsrMappedVal(cell, "Column " + columnInTable);
                cellValue = Convert.ToString(cellValOrUsrMappedVal);

                // cell value is empty or there is no user replacement for the cell value
                if (cellValue != String.Empty)
                {
                    bool mustQuote = (cellValue.Contains(csvDelimiter) || cellValue.Contains("\"") || cellValue.Contains("\'") ||
                                      cellValue.Contains("\\") || cellValue.Contains("/") || cellValue.Contains("\0") ||
                                      cellValue.Contains("\a") || cellValue.Contains("\b") || cellValue.Contains("\f") ||
                                      cellValue.Contains("\n") || cellValue.Contains("\r") || cellValue.Contains("\t") ||
                                      cellValue.Contains("\v"));
                    if (mustQuote)
                    {
                        cellValue = Parser.Parser.ToLiteral(cellValue);
                    }
                }
                else
                {
                    cellValue = "`";
                }

                excelRowAsStr = excelRowAsStr + "~" + cellValue;

                // name of a column is Column 1..NrColumns
                if (Transform.usrInputForAggregationProcess.Keys.Contains("Column " + columnInTable) && 
                    Transform.usrInputForAggregationProcess["Column " + columnInTable].IsAggregationKey)
                {
                    // because we want TRUE or true to be the same when taken into consideration boolean keys
                    if (Transform.usrInputForAggregationProcess["Column " + columnInTable].ColumnDataType == "boolean")
                    { 
                        row1KeyColumnValuesStr = row1KeyColumnValuesStr + "~" + cellValue.ToLower();
                    }
                    else
                    {
                        row1KeyColumnValuesStr = row1KeyColumnValuesStr + "~" + cellValue;
                    }
                        //row1KeyColumnValuesStr = row1KeyColumnValuesStr + "~" + cellValue;
                }

                columnInTable++;
            }

            row1KeyColumnValuesStr = row1KeyColumnValuesStr.Remove(0, 1);

            return (excelRowAsStr.Remove(0, 1));
        }
        
        private static void GetLastExcelRowAndLastColumn(Interop.Range excelTableDataBodyRange, ref int lastColumnInTableExcelIndex, ref int lastRowInTableExcelIndex)
        {
            foreach (Interop.Range row in excelTableDataBodyRange.Rows)
            {
                foreach (Interop.Range cell in excelTableDataBodyRange.Columns)
                {
                    lastColumnInTableExcelIndex = cell.Column + (excelTableDataBodyRange.Columns.Count - 1);
                    break;
                }
                break;
            }

            foreach (Interop.Range row in excelTableDataBodyRange.Rows)
            {
                foreach (Interop.Range cell in excelTableDataBodyRange.Columns)
                {
                    lastRowInTableExcelIndex = cell.Row + (excelTableDataBodyRange.Rows.Count - 1);
                    break;
                }
                break;
            }
        }

        // turn cell string into valid csv output cell
        private static string EscapeCSVstyle(string str, char csvDel)
        {
            bool mustQuote = (str.Contains(csvDel) || str.Contains("\"") || str.Contains("\'")  ||
                              str.Contains("\\") || str.Contains("/") || str.Contains("\0") ||
                              str.Contains("\a")   || str.Contains("\b") || str.Contains("\f") ||
                              str.Contains("\n")   || str.Contains("\r") || str.Contains("\t") ||
                              str.Contains("\v"));
            if (mustQuote)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("\"");
                foreach (char nextChar in str)
                {
                    sb.Append(nextChar);
                    if (nextChar == '"')
                        sb.Append("\"");
                }
                sb.Append("\"");
                return sb.ToString();
            }

            return str;
        }

        public static string ReplaceFirstOccurrence(string Source, string Find, string Replace)
        {
            int Place = Source.IndexOf(Find);
            string result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
            return result;
        }

        public static string ReplaceLastOccurrence(string Source, string Find, string Replace)
        {
            int Place = Source.LastIndexOf(Find);
            string result = Source.Remove(Place, Find.Length).Insert(Place, Replace);
            return result;
        }
    }
}
