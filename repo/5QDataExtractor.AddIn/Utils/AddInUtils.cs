using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace _5QDataExtractor.AddIn.Utils
{
    public static class AddInUtils
    {
        public static Win32Window GetExcelWindowHandle()
        {
            var nw = new Win32Window(Globals.ThisAddIn.Application.Hwnd);
            return nw;
        }

        public static Workbook GetActiveWorkbook_IfExists()
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook == null)
            {
                return null;
            }

            return Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
        }

        public static Worksheet GetActiveWorksheet_IfExists()
        {
            var res = Globals.ThisAddIn.Application.ActiveSheet;
            if (res == null)
            {
                return null;
            }
            return Globals.Factory.GetVstoObject((Interop.Worksheet)res);
        }

        public static dynamic currentWorkbookTableCount(Workbook wb)
        {
            if (wb == null)
            {
                return null;
            }

            int workbookTableCount = 0;

            foreach (Interop.Worksheet nativeWS in wb.Worksheets)
            {
                int wsTableCount = nativeWS.ListObjects.Count;

                workbookTableCount = workbookTableCount + wsTableCount;
            }

           return workbookTableCount;
        }

        public static Interop.ListObject ReturnInputTableFromWorkbook(Workbook wb, string inputTableName)
        {
            if (wb == null)
            {
                return null;
            }

            foreach (Interop.Worksheet nativeWS in wb.Worksheets)
            {
                int wsTableCount = nativeWS.ListObjects.Count;

                for (int i = 1; i <= wsTableCount; ++i)
                {
                    if (nativeWS.ListObjects[i].Name == inputTableName)
                    {
                        return nativeWS.ListObjects[i];
                    }
                }
            }

            return null;
        }

        // get the column data type for each column in the table
        public static void GetExcelTableColumnDataTypes (Win32Window excelHandle, 
                                                        Microsoft.Office.Tools.Excel.Workbook wb, 
                                                        Interop.ListObject excelTable, 
                                                        out List<string> excelTableColumnsDataTypes)
        {
            int rowIndex = 1;
            excelTableColumnsDataTypes = new List<string>();
            string excelColumnDataType = String.Empty;

            Interop.Worksheet activeSheet = ((Interop.Worksheet)wb.ActiveSheet);
            Interop.Range tableBodyRange = excelTable.DataBodyRange;

            try
            {
                foreach (Interop.Range column in tableBodyRange.Columns)
                {
                    List<string> cellsOfAcolumnDataTypes = new List<string>();

                    foreach (Interop.Range cell in column.Rows)
                    {
                        object value = cell.Value;
                        string typeName;

                        if (value == null)
                        {
                            typeName = "null";
                        }
                        else
                        {
                            //// v2 to find out time  -- PUT AN OPTION ON the "Get the table name" window to look for time in column
                            //typeName = value.GetType().ToString().Replace("System.", "").ToLower();

                            //if (typeName == "double")
                            //{
                            //    // validte time in cell.Text with regex 
                            //    //  if the typeName includes AM/PM  ^(((0?[0-9])|(1[0-2])):[0-5]?\d(:[0-5]?\d)?\s(AM|PM))$  for AM/PM time
                            //    //  if typeName DOESN'T include  AM/PM   ^(((0?\d)|(1\d)|2[0-3]):[0-5]?\d(:[0-5]?\d)?)$   for 00:00:00 to 23:00:00

                            //    // if it's not valid, it's data type will be string

                            //    // when using functions use cell.value2
                            //    // for testing regex http://regexstorm.net/tester
                            //}
                            //else
                            //{
                            //    typeName = value.GetType().ToString().Replace("System.", "").ToLower();
                            //}
                            //// v2 END to find out time

                            // v1 that works !
                            typeName = value.GetType().ToString().Replace("System.", "").ToLower();
                        }

                        // write cell data type to excell cell
                        // ((Interop.Range)activeSheet.Cells[cell.Row, cell.Column - 6]).Value = typeName;

                        cellsOfAcolumnDataTypes.Add(typeName);

                        rowIndex = cell.Row;
                    }

                    // aproximate the table's current column data type based on it's cells' data type
                    excelColumnDataType = GetExcelColumnDataTypeBasedOnCellsDataTypes(cellsOfAcolumnDataTypes, column);

                    excelTableColumnsDataTypes.Add(excelColumnDataType);

                    // write to cell the column's aproximated type
                    // ((Interop.Range)activeSheet.Cells[rowIndex + 2, column.Column - 6]).Value = excelColumnDataType;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(excelHandle, $"{ex.Message}","", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
                if (tableBodyRange != null) Marshal.ReleaseComObject(tableBodyRange);
            }
        }

        public static string GetExcelColumnDataTypeBasedOnCellsDataTypes(List<string> cellDataTypes, Interop.Range column)
        {
            //// if all elements are the same data type, or cells are of type null and at least one cell is of a type different than null 
            ////   return that data type
            //// else default to string
            //if (cellDataTypes.Any(o => o != cellDataTypes[0] && o != "null" ))
            //{
            //    return "string";
            //}
            //else
            //{
            //    return cellDataTypes[0];
            //}

            string emptyCellDataType = "null";

            string defaultMultipleCellDataTypesInCol = "string";

            List<string> uniqueCellDT;

            try
            {
                // get the distinct cell data types
                uniqueCellDT = cellDataTypes.Distinct().ToList();

                switch (uniqueCellDT.Count)
                {
                    // one data type per column
                    case 1:
                        return uniqueCellDT[0];

                    // two data types per column
                    case 2:
                        // if one of them is null
                        if (uniqueCellDT.Contains(emptyCellDataType))
                        {
                            if (uniqueCellDT[0] == emptyCellDataType)
                            {
                                return uniqueCellDT[1];
                            }

                            if (uniqueCellDT[1] == emptyCellDataType)
                            {
                                return uniqueCellDT[0];
                            }
                        }

                        // if one of them is not null
                        return defaultMultipleCellDataTypesInCol;

                    // three or more data types per column
                    default:
                        return defaultMultipleCellDataTypesInCol;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error when trying to figure out what the column data type is for Excel column {column.Column}: " +
                                    $"{ex.Message}");
            }
        }
    }
}
