using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Globalization;
using _5QDataExtractor.Library.Handler;

namespace _5QDataExtractor.Library
{
    public static class Validator
    {
        public static bool TblColCanBeParsedToTheNewType(Interop.ListObject inputTable, int tableColToInspect, string oldColDT, string newColDT)
        {
            Interop.Range tableBodyRange = null;

            try
            {
                tableBodyRange = inputTable.DataBodyRange;

                string columnName = String.Empty;
                string cellVal;

                int tableColumnNum = 1;

                // v2 improvement - validate only unique column values
                foreach (Interop.Range column in tableBodyRange.Columns)
                {
                    columnName = "Column " + tableColumnNum;

                    //if (Transform.usrInputForAggregationProcess.Keys.Contains(columnName))
                    //{
                        if (tableColumnNum == tableColToInspect)
                        {
                            switch (oldColDT)
                            {
                                case "string":
                                    // new data type that the user provided
                                    switch (newColDT)
                                    {
                                        case "string":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToString(Transform.GetCellValueOrUsrMappedVal(cell, columnName));
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value \"{cell.Value2}\" " +
                                                                        $"from table column {columnName} Excel row {cell.Row} or it's mapped value to string. " + ex.Message);
                                                }
                                            }
                                            break;

                                        case "boolean":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToBoolean(Transform.GetCellValueOrUsrMappedVal(cell, columnName));
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value \"{cell.Value2}\" from table column {columnName} " +
                                                                        $"Excel row {cell.Row} or it's mapped value to string to boolean." +
                                                                        $"Table column {tableColumnNum} should only contain these values: TRUE, FALSE. " +
                                                                        ex.Message);
                                                }
                                            }
                                            break;

                                        default:
                                            throw new Exception($"The application doesn't have a validation process when trying to convert cells from " +
                                                                $"{Transform.usrInputForAggregationProcess[columnName].OldDataType} to {Transform.usrInputForAggregationProcess[columnName].ColumnDataType} " +
                                                                $"on column {tableColumnNum}. Validation process stopped.");
                                    }
                                    break;

                                case "double":
                                    // new data type that the user provided
                                    switch (newColDT)
                                    {
                                        case "double":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToDecimal(Transform.GetCellValueOrUsrMappedVal(cell, columnName));
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value \"{cell.Value2}\" from table column {columnName} " +
                                                                        $"Excel row {cell.Row} or it's mapped vallue to decimal. " + ex.Message);
                                                }
                                            }
                                            break;

                                        case "string":
                                            Decimal dummy;

                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    dummy = Convert.ToDecimal(Transform.GetCellValueOrUsrMappedVal(cell, columnName));
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value \"{cell.Value2}\" from table column {columnName} " +
                                                                        $"Excel row {cell.Row} or it's mapped value to decimal. " + ex.Message);
                                                }

                                                try
                                                {
                                                    Convert.ToString(dummy);
                                                }
                                                catch (Exception ex)
                                                {

                                                    throw new Exception($"Error when converting cell value \"{cell.Value2}\" from  table column {columnName} " +
                                                                        $"Excel row {cell.Row} or it's mapped value to string after being successfully converted to decimal. " +
                                                                        ex.Message);
                                                }
                                            }
                                            break;

                                        case "boolean":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToBoolean(Transform.GetCellValueOrUsrMappedVal(cell, columnName));
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value \"{cell.Value2}\" from  table column {columnName} " +
                                                                        $"Excel row {cell.Row} or it's mapped value to boolean. " +
                                                                        ex.Message);
                                                }
                                            }
                                            break;

                                        default:
                                            throw new Exception($"The application doesn't have a validation process when trying to convert cells from " +
                                                                $"{Transform.usrInputForAggregationProcess[columnName].OldDataType} to {Transform.usrInputForAggregationProcess[columnName].ColumnDataType} " +
                                                                $"on column {tableColumnNum}. Validation process stopped.");
                                    }
                                    break; 

                                case "boolean":
                                    // new data type that the user provided
                                    switch (newColDT)
                                    {
                                        case "boolean":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToBoolean(cell.Value2);
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value {cell.Value2} from  table column {columnName} " +
                                                                        $"Excel row {cell.Row} to boolean. " +
                                                                        ex.Message);
                                                }
                                            }
                                            break;

                                        case "string":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToString(cell.Value2);
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value {cell.Value2} from  table column {columnName} " +
                                                                        $"Excel row {cell.Row} to string. " +
                                                                        ex.Message);
                                                }
                                            }
                                            break;

                                        case "double":
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                try
                                                {
                                                    Convert.ToDecimal(cell.Value2);
                                                }
                                                catch (Exception ex)
                                                {
                                                    throw new Exception($"Error when converting cell value {cell.Value2} from  table column {columnName} " +
                                                                        $"Excel row {cell.Row} to decimal. " +
                                                                        ex.Message);
                                                }
                                            }
                                            break;

                                        default:
                                            throw new Exception($"The application doesn't have a validation process when trying to convert cells from " +
                                                                $"{Transform.usrInputForAggregationProcess[columnName].OldDataType} to {Transform.usrInputForAggregationProcess[columnName].ColumnDataType} " +
                                                                $"on column {tableColumnNum}. Validation process stopped.");
                                    }
                                    break;

                                case "datetime":
                                    switch (newColDT)
                                    {
                                        case "datetime":
                                            // Excel datetime http://www.cpearson.com/excel/datetime.htm
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                string str = Convert.ToString(Transform.GetCellValueOrUsrMappedVal(cell, columnName));

                                                if (Decimal.TryParse(str, out decimal decValue))
                                                {
                                                    int castedV = Decimal.ToInt32(Math.Round(decValue));

                                                    //Excel behaves as if the date 1900-Feb-29 existed, it did not. So we must subtract 1
                                                    DateTime validCell = new DateTime(1899, 12, 31).AddDays(castedV - 1);
                                                }
                                                else
                                                {
                                                    throw new Exception($"Cell value {cell.Value2} at table column {columnName} Excel row {cell.Row} or it's mapped value " +
                                                                        $"could not be converted to a valid DateTime structure. +");
                                                }
                                            }
                                            break;

                                        case "string":
                                            // Excel datetime http://www.cpearson.com/excel/datetime.htm
                                            foreach (Interop.Range cell in column.Rows)
                                            {
                                                string str = Convert.ToString(Transform.GetCellValueOrUsrMappedVal(cell, columnName));

                                                if (Decimal.TryParse(str, out decimal decValue))
                                                {
                                                    int castedV = Decimal.ToInt32(Math.Round(decValue));

                                                    //Excel behaves as if the date 1900-Feb-29 existed, it did not. So we must subtract 1
                                                    DateTime validCell = new DateTime(1899, 12, 31).AddDays(castedV - 1);

                                                    Convert.ToString(validCell);
                                                }
                                                else
                                                {
                                                    throw new Exception($"Cell value {cell.Value2} at table column {columnName} Excel row {cell.Row} or it's mapped value " +
                                                                        $"could not be converted to a valid DateTime structure. +");
                                                }
                                            }
                                            break;

                                        default:
                                            throw new Exception($"The application doesn't have a validation process when trying to convert cells from " +
                                                                $"{Transform.usrInputForAggregationProcess[columnName].OldDataType} to {Transform.usrInputForAggregationProcess[columnName].ColumnDataType} " +
                                                                $"on column {tableColumnNum}. Validation process stopped. Application will now exit.");
                                    }
                                    break;

                                case "null":
                                    // doesn't need checking because Excel validated the cells as null when I got the column's type previously
                                    //foreach (Interop.Range cell in column.Rows)
                                    //{
                                    //    if (Convert.ToString(cell.Value2) != String.Empty)
                                    //    {
                                    //        throw new Exception($"Column {tableColumnNum} is not empty. Column data type validation stopped. Application will now exit.");
                                    //    }
                                    //}
                                    break;

                                default:
                                    throw new Exception($"Table column {tableColumnNum} has an unsupported old column data type ({Transform.usrInputForAggregationProcess[columnName].OldDataType}). " +
                                                        $"Column data type validation stopped. Application will now exit.");
                            }

                            return true;
                        }
                        else
                        {
                            tableColumnNum++;
                        }
                    //}
                    //else
                    //{
                    //    throw new Exception($"No key in column data type dictionary with the name {columnName}. Column data type validation stopped.");
                    //}
                }

                // the column specified could not be found
                //throw new Exception($"Could not find column index {tableColToInspect} in the table");
                return false;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (tableBodyRange != null) Marshal.ReleaseComObject(tableBodyRange);
            }
        }

        private static void ObjValIsTrueOrFalseStr(object value, string columName, object key)
        {
            try
            {
                if (!(value.ToString().ToLower() == "true") &&
                    !(value.ToString().ToLower() == "false"))
                {
                    throw new Exception($"User replacement value {Convert.ToString(value)} was not a match with either true or false strings.");
                }
            }
            catch (Exception ex)
            {
                string exceptionMsg = $"Error found inside the user value replacement dictionary correponding to " +
                                      $"{columName} for cell value \"{key.ToString()}\" (form column 1): " +
                                      ex.Message;

                throw new Exception(exceptionMsg);
            }
        }

        private static void ObjValToDecimal(object value)
        {
            try
            {
                Convert.ToDecimal(value);
            }
            catch (OverflowException ex)
            {
                throw new OverflowException($"he value {Convert.ToString(value)} is out of range of the Decimal type.");
            }
            catch (FormatException ex)
            {
                throw new FormatException($"The value {Convert.ToString(value)} is not recognized as a valid Decimal value.");
            }
            catch (InvalidCastException ex)
            {
                throw new InvalidCastException($"Conversion of the value {Convert.ToString(value)} to a Decimal data type is not supported.");
            }
        }

        public static void UsrMappedValsCanConvToColNewDT(string columName, string newColDataTypeKey, string oldColDT,
                                                          Dictionary<object, List<object>> ExTblColCellValToUserValDict,
                                                          Interop.ListObject inputTable, int tableColToInspect)
        {
            string exceptionMsg = String.Empty;

            switch (newColDataTypeKey)
            {
                // DONE
                case "boolean":
                    if (ExTblColCellValToUserValDict == null)
                    {
                        //if (oldColDT == "string")
                        //{
                        //    throw new Exception($"Please use the \"Map column values\" feature to map the unique strings in your Excel table {columName} " +
                        //                        $"to one of these values: true, false.");
                        //}

                        Interop.Range tableBodyRange = inputTable.DataBodyRange;

                        bool columnSearchStatus = false;

                        try
                        {
                            int tableColumnNum = 1;

                            foreach (Interop.Range column in tableBodyRange.Columns)
                            {
                                if (tableColumnNum == tableColToInspect)
                                {
                                    columnSearchStatus = true;

                                    foreach (Interop.Range cell in column.Rows)
                                    {
                                        try
                                        {
                                            // accept null values as valid Booleans
                                            if (cell.Value2 != null) { Convert.ToBoolean(cell.Value2); }
                                                
                                        }
                                        catch (Exception ex)
                                        {
                                            throw new Exception($"Error when converting Excel cell value {cell.Value2} at table {tableColumnNum} and " +
                                                                $"Excel row {cell.Row} to boolean: " + ex.Message);
                                        }
                                    }
                                }
                                else
                                {
                                    tableColumnNum++;
                                }
                            }

                            if (columnSearchStatus == false)
                            {
                                throw new Exception($"Could not find table column index {tableColToInspect}");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            if (tableBodyRange != null) { Marshal.ReleaseComObject(tableBodyRange); }
                        }

                        return;
                    }

                    foreach (object key in ExTblColCellValToUserValDict.Keys)
                    {
                        object value;

                        // if user didn't want to replace the current value of inside the cell value - user value dictionary, don't check user input
                        if (ExTblColCellValToUserValDict[key][1].ToString() != String.Empty)
                        {
                            value = ExTblColCellValToUserValDict[key][1];

                            ObjValIsTrueOrFalseStr(value, columName, key);
                        }
                        // check cell table which is the key
                        else
                        {
                            // accept empty cells as valid boolean values
                            // if old column data type is not boolean and the value in column 1 is not a null cell
                            if (oldColDT != "boolean" && key.ToString() != String.Empty)
                            {
                                value = key;

                                ObjValIsTrueOrFalseStr(value, columName, key);
                            }
                        }
                    }
                    break;

                case "datetime":
                    if (ExTblColCellValToUserValDict == null)
                    {
                        Interop.Range tableBodyRange = inputTable.DataBodyRange;

                        bool columnSearchStatus = false;

                        try
                        {
                            int tableColumnNum = 1;

                            foreach (Interop.Range column in tableBodyRange.Columns)
                            {
                                if (tableColumnNum == tableColToInspect)
                                {
                                    columnSearchStatus = true;

                                    foreach (Interop.Range cell in column.Rows)
                                    {
                                        try
                                        {
                                            // accept empty cells as valid DateTime
                                            if (cell.Value2 != null) { Handler.DateTimeCellsHandler.CovertCellValuesToDateTime(cell.Value2); }
                                        }
                                        catch (Exception ex)
                                        {
                                            throw new Exception($"Error when converting Excel cell value {cell.Value2} at table {tableColumnNum} and " +
                                                                $"Excel row {cell.Row} to a DateTime structure: " + ex.Message);
                                        }
                                    }
                                }
                                else
                                {
                                    tableColumnNum++;
                                }
                            }

                            if (columnSearchStatus == false)
                            {
                                throw new Exception($"Could not find table column index {tableColToInspect}");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            if (tableBodyRange != null) { Marshal.ReleaseComObject(tableBodyRange); }
                        }

                        return;
                    }

                    foreach (object key in ExTblColCellValToUserValDict.Keys)
                    {
                        // don't check empty values in column 2 and don't check nulls in column 1
                        // if user didn't want to replace the current value of inside the cell value - user value dictionary, don't check user input
                        if (ExTblColCellValToUserValDict[key][1].ToString() != String.Empty)
                        {
                            try
                            {
                                DateTime.ParseExact(ExTblColCellValToUserValDict[key][1].ToString(), "M/d/yyyy h:mm tt", CultureInfo.InvariantCulture);
                            }
                            catch (ArgumentNullException ex)
                            {
                                throw new ArgumentNullException($"Error found inside the user value replacement dictionary correponding to {columName} for cell value \"{key.ToString()}\" (form column 1): " +
                                                                $"User did not provide a string to be converted to Datetime.");
                            }
                            catch (FormatException ex)
                            {
                                try
                                {
                                    DateTime keyToDateTme = Handler.DateTimeCellsHandler.CovertCellValuesToDateTime(key);

                                    throw new FormatException($"Inside the user value replacement dictionary correponding to {columName} for cell value \"{keyToDateTme.ToString()}\" (form column 1) " +
                                                              $"the user did not provide a string to be converted to Datetime or " +
                                                              $"the user value {Convert.ToString(ExTblColCellValToUserValDict[key][1])} is not in the DateTime format M/d/yyyy h:mm AM/PM or " +
                                                              $"the hour component and the AM/PM designator in the user value do not agree.");
                                }
                                catch (Exception ex2)
                                {
                                    throw new Exception(ex2.Message);
                                }

                            }
                        }
                    }
                    break;

                case "double":
                    if (ExTblColCellValToUserValDict == null)
                    {
                        Interop.Range tableBodyRange = inputTable.DataBodyRange;

                        bool columnSearchStatus = false;

                        try
                        {
                            int tableColumnNum = 1;

                            foreach (Interop.Range column in tableBodyRange.Columns)
                            {
                                if (tableColumnNum == tableColToInspect)
                                {
                                    columnSearchStatus = true;

                                    foreach (Interop.Range cell in column.Rows)
                                    {
                                        try
                                        {
                                            // accept null cells as valid decimals
                                            if (cell.Value2 != null) { Convert.ToDecimal(cell.Value2); }
                                        }
                                        catch (Exception ex)
                                        {
                                            throw new Exception($"Error when converting Excel cell value {cell.Value2} at Column {tableColumnNum} and " +
                                                                $"Excel row {cell.Row} to decimal: " + ex.Message);
                                        }
                                    }
                                }
                                else
                                {
                                    tableColumnNum++;
                                }
                            }

                            if (columnSearchStatus == false)
                            {
                                throw new Exception($"Could not find table column index {tableColToInspect}");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            if (tableBodyRange != null) { Marshal.ReleaseComObject(tableBodyRange); }
                        }

                        return;
                    }

                    object val;
                    foreach (object key in ExTblColCellValToUserValDict.Keys)
                    {
                        if (ExTblColCellValToUserValDict[key][1].ToString() != String.Empty)
                        {
                            val = ExTblColCellValToUserValDict[key][1];

                            try
                            {
                                ObjValToDecimal(val);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Error found at column 2 and key \"{key.ToString()}\" inside the user value replacement dictionary " +
                                                    $"corresponding to {columName}: " + ex.Message);
                            }
                        }
                        // no replacement value
                        else
                        {
                            if (key.ToString() != String.Empty)
                            { 
                                val = key;

                                try
                                {
                                    ObjValToDecimal(val);
                                }
                                catch (OverflowException ex)
                                {
                                    throw new Exception($"Error found at column 1 inside the user value replacement dictionary " +
                                                        $"corresponding to {columName}:" + ex.Message);
                                }
                            }
                        }
                    }
                    break;

                case "string":
                    if (ExTblColCellValToUserValDict == null)
                    {
                        Interop.Range tableBodyRange = inputTable.DataBodyRange;

                        bool columnSearchStatus = false;

                        try
                        {
                            int tableColumnNum = 1;

                            foreach (Interop.Range column in tableBodyRange.Columns)
                            {
                                if (tableColumnNum == tableColToInspect)
                                {
                                    columnSearchStatus = true;

                                    foreach (Interop.Range cell in column.Rows)
                                    {
                                        try
                                        {
                                            // accept null cells as valid strings
                                            if (cell.Value2 != null) { Convert.ToString(cell.Value2); }
                                        }
                                        catch (Exception ex)
                                        {
                                            throw new Exception($"Error when converting Excel cell value {cell.Value2} at Column {tableColumnNum} and " +
                                                                $"Excel row {cell.Row} to string: " + ex.Message);
                                        }
                                    }
                                }
                                else
                                {
                                    tableColumnNum++;
                                }
                            }

                            if (columnSearchStatus == false)
                            {
                                throw new Exception($"Could not find table column index {tableColToInspect}");
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            if (tableBodyRange != null) { Marshal.ReleaseComObject(tableBodyRange); }
                        }

                        return;
                    }
                    
                    foreach (object key in ExTblColCellValToUserValDict.Keys)
                    {
                        // if user didn't want to replace the current value of inside the cell value - user value dictionary, don't check user input
                        if (ExTblColCellValToUserValDict[key][1].ToString() != String.Empty)
                        {
                            try
                            {
                                Convert.ToString(ExTblColCellValToUserValDict[key][1]);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Error found at column 2 key \"{key.ToString()}\" inside the user value replacement dictionary " +
                                                    $"corresponding to {columName} when trying to convert value {Convert.ToString(ExTblColCellValToUserValDict[key][1])}" +
                                                    $" to string: " + ex.Message);
                            }
                        }
                        else
                        {
                            try
                            {
                                Convert.ToString(key);
                            }
                            catch (Exception ex)
                            {
                                throw new Exception($"Error found at column 1 key \"{key.ToString()}\" inside the user value replacement dictionary " +
                                                    $"corresponding to {columName} when trying to convert this key to string: " +
                                                    ex.Message);
                            }
                        }
                    }
                    break;

                // 15 March 2018 - for the moment no validation is to be done on null
                case "null":
                    break;
                    
               
                default:
                    throw new Exception($"Error when validating user mapped values on the table's {columName.ToLower()} to the new column " +
                                $"data type {newColDataTypeKey}: Could not convert values to this new column data type.");
            }
        }
    }
}
