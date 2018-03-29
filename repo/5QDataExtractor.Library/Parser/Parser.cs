using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using _5QDataExtractor.Library.Handler;

namespace _5QDataExtractor.Library.Parser
{
    public static class Parser
    {
        private static Dictionary<string, char> csvChrListSeparator = new Dictionary<string, char>
        {
            {"[comma] ,"        , ','},
            {"[tab character]"  , '\t'},
            {"[colon] :"        , ':'},
            {"[semicolon] ;"    , ';'},
            {"[vertical bar] |" , '|'}
        };

        public static char GetListSeparatorFromUserInput(string userCSVdelimiterInput)
        {
            // user did't select anything from the list separator combo list box and didn't write anything
            if (userCSVdelimiterInput.Length == 0)
            {
                // the csv delimiter will be the OS' list separator for the current selected language
                return Convert.ToChar(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator);
            }

            // if the user selected a value from the list separator combo list box
            foreach (string cmbCSVdelimiter in csvChrListSeparator.Keys)
            {
                if (userCSVdelimiterInput == cmbCSVdelimiter)
                {
                    return csvChrListSeparator[cmbCSVdelimiter];
                }
            }

            // the user entered his own delimitator and we take the first character from it
            return userCSVdelimiterInput[0];
        }

        /// <summary>
        /// Convert to escaped string literal
        /// </summary>
        /// <param name="input"></param>
        /// <returns>Escaped string literal</returns>
        public static string ToLiteral(string input)
        {
            using (var writer = new StringWriter())
            {
                using (var provider = CodeDomProvider.CreateProvider("CSharp"))
                {
                    provider.GenerateCodeFromExpression(new CodePrimitiveExpression(input), writer, null);
                    return writer.ToString();
                }
            }
        }

        public static Dictionary<object, string> GetUniqueExTblColVals(Interop.ListObject inExTable, string colNameToInspect, 
                                                              string colToInspectNewDTKey, string colToInspOldDTKey)
        {
            Interop.Range tableBodyRange = null;

            Dictionary<object, string> ExTblColUniqueVals = new Dictionary<object, string>();

            bool breakOutOfSecondFor = false;

            tableBodyRange = inExTable.DataBodyRange;

            int idxOfSpaceInStr = colNameToInspect.IndexOf(' ');
            if (idxOfSpaceInStr > 0)
            {
                string ExcelTableColIndex = colNameToInspect.Substring(idxOfSpaceInStr + 1);
                if (ExcelTableColIndex != String.Empty)
                {
                    if (Int32.TryParse(ExcelTableColIndex, out int tableColNum))
                    {
                        int tableColCount = 1;
                       
                        foreach (Interop.Range column in tableBodyRange.Columns)
                        {
                            if (tableColCount == tableColNum)
                            {
                                long tableRowsCount = 1;
                                string cellValToObj;

                                // get Excel column's unique values
                                foreach (Interop.Range cell in column.Rows)
                                {
                                    if (tableRowsCount == 1)
                                    {
                                        try
                                        {
                                            CellBoxUnbox.CellValueToStr(cell, tableColNum, out cellValToObj);
                                        }
                                        catch (Exception ex)
                                        {
                                            throw new Exception($"Error at cell value {cell.Value2} on table column {tableColCount} row {tableRowsCount} " +
                                                                $"when trying to display cell value on the frmDEMapColumnValues form: " + ex.Message);
                                        }

                                        if (cell.Value2 != null)
                                            ExTblColUniqueVals.Add(cell.Value2, cellValToObj);
                                        else
                                            ExTblColUniqueVals.Add("", cellValToObj);

                                        tableRowsCount++;

                                        continue;
                                    }

                                    try
                                    {
                                        CellBoxUnbox.CellValueToStr(cell, tableColNum, out cellValToObj);
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception($"Error at cell value {cell.Value2} on table column {tableColCount} row {tableRowsCount} " +
                                                            $"when trying to display cell value on the frmDEMapColumnValues form: " + ex.Message);
                                    }
                                   
                                    if (cell.Value2 != null)
                                    {
                                        if (!ExTblColUniqueVals.Keys.ToList().Contains(cell.Value2))
                                            ExTblColUniqueVals.Add(cell.Value2, cellValToObj);
                                    }
                                    else
                                    {
                                        if (!ExTblColUniqueVals.Keys.ToList().Contains(""))
                                            ExTblColUniqueVals.Add("", cellValToObj);
                                    }

                                    tableRowsCount++;
                                }

                                breakOutOfSecondFor = true;

                                break;
                            }

                            if (breakOutOfSecondFor) break;

                            tableColCount++;
                        }
                    }
                }
                else
                {
                    // no number in column name
                    throw new Exception($"Could not get the distinct values from the excel table column passed " +
                                        $"to the frmDEMapColumnValues form for the Shown event, because the column name does not contain an integer.");
                }
            }
            else
            {
                // no space in column name
                throw new Exception($"Could not get the distinct values from the excel table column passed " +
                                    $"to the frmDEMapColumnValues form for the Shown event, because the column " +
                                    $"name does not contain a space between Column and the table column index.");
            }

            if (tableBodyRange != null) Marshal.ReleaseComObject(tableBodyRange);

            return ExTblColUniqueVals;
        }
    }
}
