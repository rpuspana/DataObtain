using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _5QDataExtractor.Library.Handler;

namespace _5QDataExtractor.Library
{
    public static class CellBoxUnbox
    {
        public static void CellValueToStr(Interop.Range cell, int tblColIdx, out string cellValToObj)
        {
            string typeName;

            if (cell.Value == null)
            {
                typeName = "null";
            }
            else
            {
                typeName = cell.Value.GetType().ToString().Replace("System.", "").ToLower();
            }

            switch (typeName)
            {
                case "string":
                    cellValToObj = Convert.ToString(cell.Value2);
                    break;

                case "double":
                    cellValToObj = Convert.ToString(
                                        NumberCellsHandler.ConvertToDecimalDataType(cell.Value2)
                                   );
                    break;

                case "boolean":
                    try
                    {
                        cellValToObj = Convert.ToString(Convert.ToBoolean(cell.Value2)).ToUpper();
                    }
                    catch (FormatException ex)
                    {
                        throw new FormatException(
                            $"The {cell.Value2.GetType().Name} value {Convert.ToString(cell.Value2)} is not recognized as a valid boolean value."
                        );
                    }
                    catch (InvalidCastException ex)
                    {
                        throw new InvalidCastException(
                            $"Conversion of the {cell.Value.GetType().Name} value {Convert.ToString(cell.Value)} to a boolean value is not supported."
                        );
                    }

                    break;

                case "datetime":
                    cellValToObj = Convert.ToString(DateTimeCellsHandler.CovertCellValuesToDateTime(cell.Value2));
                    break;

                case "null":
                    cellValToObj = String.Empty;
                    break;

                default:
                    throw new Exception($"Application can't convert {typeName} table cells to their string representation.");
            }
        }

        public static void CellValueToObject(Interop.Range cell, out object cellValToObj)
        {
            string typeName;

            if (cell.Value == null)
            {
                typeName = "null";
            }
            else
            {
                typeName = cell.Value.GetType().ToString().Replace("System.", "").ToLower();
            }

            switch (typeName)
            {
                case "string":
                    cellValToObj = cell.Value2;
                    break;

                case "double":
                    try
                    {
                        cellValToObj = (Convert.ToDecimal(cell.Value2));
                    }
                    catch (OverflowException ex)
                    {
                        throw new OverflowException($"The {cell.Value2.GetType().Name} value {Convert.ToString(cell.Value2)} is out of range of the Decimal type.");
                    }
                    catch (FormatException ex)
                    {
                        throw new FormatException($"The {cell.Value2.GetType().Name} value {Convert.ToString(cell.Value2)} is not recognized as a valid Decimal value.");
                    }
                    catch (InvalidCastException ex)
                    {
                        throw new InvalidCastException($"Conversion of the {cell.Value2.GetType().Name} value {Convert.ToString(cell.Value2)} to a Decimal is not supported.");
                    }
                    break;

                case "boolean":
                    try
                    {
                        cellValToObj = Convert.ToBoolean(cell.Value2);
                    }
                    catch (FormatException ex)
                    {
                        throw new FormatException(
                            $"The {cell.Value2.GetType().Name} value {Convert.ToString(cell.Value2)} is not recognized as a valid boolean value."
                        );
                    }
                    catch (InvalidCastException ex)
                    {
                        throw new InvalidCastException(
                            $"Conversion of the {cell.Value.GetType().Name} value {Convert.ToString(cell.Value)} to a boolean value is not supported."
                        );
                    }
                    break;

                case "datetime":
                    cellValToObj = DateTimeCellsHandler.CovertCellValuesToDateTime(cell.Value2);
                    break;

                case "null":
                    cellValToObj = String.Empty;
                    break;

                default:
                    throw new Exception($"Application can't convert {typeName} table cells to {typeName} data structure.");
            }
        }

        public static void CastToCellType(object value, string dataType, bool fallToDefault, out object cellValToObj)
        {
            if (fallToDefault)
            {
                if (Convert.ToString(value) == String.Empty)
                {
                    cellValToObj = String.Empty;
                    return;
                }
            }

            switch (dataType)
            {
                case "string":
                    cellValToObj = Convert.ToString(value);
                    break;

                case "double":
                    try
                    {
                        cellValToObj = (Convert.ToDecimal(value));
                    }
                    catch (OverflowException ex)
                    {
                        throw new OverflowException($"The value \"{Convert.ToString(value)}\" is out of range of the Decimal type.");
                    }
                    catch (FormatException ex)
                    {
                        throw new FormatException($"The value \"{Convert.ToString(value)}\" is not recognized as a valid Decimal value.");
                    }
                    catch (InvalidCastException ex)
                    {
                        throw new InvalidCastException($"Conversion of the value \"{Convert.ToString(value)}\" to a Decimal is not supported.");
                    }
                   
                    break;

                case "boolean":
                    try
                    {
                        cellValToObj = Convert.ToBoolean(value);
                    }
                    catch (FormatException ex)
                    {
                        throw new FormatException(
                            $"The value \"{Convert.ToString(value)}\" is not recognized as a valid boolean value."
                        );
                    }
                    catch (InvalidCastException ex)
                    {
                        throw new InvalidCastException(
                            $"Conversion of the value \"{Convert.ToString(value)}\" to a boolean value is not supported."
                        );
                    }
                    break;

                case "datetime":
                    if (Double.TryParse(Convert.ToString(value), out double decValue))
                    {
                        try
                        {
                            //Excel behaves as if the date 1900-Feb-29 existed, it did not. So we must subtract 1
                            //return new DateTime(1899, 12, 31).AddDays(castedV - 1);
                            cellValToObj = new DateTime(1899, 12, 31).AddDays(decValue - 1);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"The {value.GetType().Name} value \"{Convert.ToString(value)}\"." +
                                                $"Could not add {decValue} days from 31.12.1899.");
                        }
                    }
                    else
                    {
                        try
                        {
                            cellValToObj = Convert.ToDateTime(value);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"The {value.GetType().Name} value \"{Convert.ToString(value)}\" could not be conveted to DateTime .Net data structure.");
                        }
                            
                    }
                    break;

                case "null":
                    cellValToObj = String.Empty;
                    break;

                default:
                    throw new Exception($"Application can't convert {dataType} table cells to their string representation.");
            }
        }
    }
}
