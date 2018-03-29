using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using _5QDataExtractor.AddIn.Utils;
using System.Windows.Forms;
using _5QDataExtractor.UI.Forms;
using _5QDataExtractor.UI.Forms.DynamicExtract;
using System.Threading.Tasks;
using _5QDataExtractor.Library.DataAccess;
using _5QDataExtractor.Library.Parser;
using _5QDataExtractor.Library;
using _5QDataExtractor.Library.Model;
using System.Runtime.InteropServices;

namespace _5QDataExtractor.AddIn
{
    public partial class Ribbon
    {
        private int userInputOption;

        private bool _secondWindowClosed = false;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// Extract the data from :
        /// - Excel table
        /// and create a csv file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDynamicExtract_Click(object sender, RibbonControlEventArgs e)
        {
            var excelHandle = AddInUtils.GetExcelWindowHandle();

            //userInputOption = GetUserInputOption(excelHandle);
            bool secondWindowResult = false; 
            // ask for data input type
            using (var frmDESelectInput = new frmDESelectInput())
            {
                do
                {
                    if (frmDESelectInput.ShowDialog(excelHandle) == DialogResult.Cancel)
                    {
                        frmDESelectInput.Close();
                        return;
                    }

                    //frmDESelectInput.Hide();

                    // get what the user selected from the frmDESelectInput
                    // 1 = Tabular Form radio button
                    userInputOption = frmDESelectInput.UserSelection;

                    // user didn't press "Cancel" on the data input type form
                    if (userInputOption > 0)
                    {
                        switch (userInputOption)
                        {
                            // user selected the Tabular Form radio button
                            case 1:
                                // the user chose "Yes"
                                processTabularFormInput(excelHandle, frmDESelectInput);

                                break;

                            // user selected option 2 radio button
                            case 2:
                                processOption2Input(excelHandle, frmDESelectInput);
                                break;
                        }
                    }
                }
                while (_secondWindowClosed == true);

                frmDESelectInput.Close();
            }
        }

        //private int GetUserInputOption(Win32Window excelHandle)
        //{
        //    // ask for data input type
        //    using (var frmDESelectInput = new frmDESelectInput())
        //    {
        //        if (frmDESelectInput.ShowDialog(excelHandle) == DialogResult.Cancel)
        //        {
        //            frmDESelectInput.Close();

        //            // user pressed "Cancel"
        //            return 0;
        //        }

        //        // get what the user selected from the frmDESelectInput
        //        // 1 = Tabular Form radio button
        //        return frmDESelectInput.UserSelection;
        //    }
        //}

        //private void decideHowToProcessUserInput(Win32Window excelHandle)
        //{
        //    switch (userInputOption)
        //    {
        //        // user selected the Tabular Form radio button
        //        case 1:
        //            // the user chose "Yes"
        //            processTabularFormInput(excelHandle);

        //            break;

        //        // user selected option 2 radio button
        //        case 2:
        //            processOption2Input(excelHandle);
        //            break;
        //    }
        //}

        /// <summary>
        /// Take data from the current sheet of the Excel table, process it and dump the results to a csv file
        /// </summary>
        private void processTabularFormInput(Win32Window excelHandle, frmDESelectInput fDESelectInput)
        {
            //Interop.Workbook wb = AddInUtils.GetActiveWorkbook_IfExists();
            //Interop.Worksheet activeWS = AddInUtils.GetActiveWorksheet_IfExists();

            bool keepWindow2 = false;

            try
            {
                //MessageBox.Show(excelHandle, $"You have selected radio button {userInputOption}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                string inputTableName;
                Interop.ListObject inputTable = null;
                char InputCSVlistSeparatorParsed = '!';
                bool boolVarTableHeaderInExport;
                string csvUsrIinputListSeparator = String.Empty;
                bool AggregateTableRowsSwitch;
                List<string> excelTableColumnsDataTypes = new List<string>();

                var wb = AddInUtils.GetActiveWorkbook_IfExists();
                var activeWS = AddInUtils.GetActiveWorksheet_IfExists();

                // ask for table name
                using (wb)
                using (activeWS)
                using (var frmDETInputTableName = new frmDETInputTableName())
                {
                    do
                    {
                        if (frmDETInputTableName.ShowDialog(fDESelectInput) == DialogResult.Cancel)
                        {
                            frmDETInputTableName.Close();

                            _secondWindowClosed = true;

                            return;
                        }

                        inputTableName = frmDETInputTableName.GetInputTableName();
                        boolVarTableHeaderInExport = frmDETInputTableName.IncludeTableHeaderValuesInExportFile();
                        csvUsrIinputListSeparator = frmDETInputTableName.GetCSVlistSeparator();
                        AggregateTableRowsSwitch = frmDETInputTableName.GetAggregationSwitchValue();

                        // check the table name so it's not empty, whitespace or null
                        if (String.IsNullOrEmpty(inputTableName) || String.IsNullOrWhiteSpace(inputTableName))
                        {
                            MessageBox.Show(excelHandle, "Please enter a table name of at least 1 visible character.",
                                            "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            keepWindow2 = true;
                        }
                        else
                        {
                            keepWindow2 = false;
                        }

                        if (keepWindow2 == false)
                        {
                            ////MessageBox.Show(excelHandle, $"table name {inputTableName}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            InputCSVlistSeparatorParsed = Parser.GetListSeparatorFromUserInput(csvUsrIinputListSeparator);

                            var numOfTablesInCurrentWorkbook = AddInUtils.currentWorkbookTableCount(wb);

                            //MessageBox.Show(excelHandle, $"tables in workbook {numOfTablesInCurrentWorkbook}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            if (numOfTablesInCurrentWorkbook == 0 || numOfTablesInCurrentWorkbook == null)
                            {
                                MessageBox.Show(excelHandle, $"There are no tables in your current workbook. Please create at least one table in order to use this form of data input.",
                                                "", MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);

                                keepWindow2 = true;
                            }
                            else
                            {
                                keepWindow2 = false;
                            }

                            if (keepWindow2 == false)
                            {
                                inputTable = AddInUtils.ReturnInputTableFromWorkbook(wb, inputTableName);

                                // if the input table was not found
                                if (inputTable == null)
                                {
                                    MessageBox.Show(excelHandle, $"The input table could not be found. Please create at least one table with the specified name in order to use this form of data input.",
                                                    "", MessageBoxButtons.OK,
                                                    MessageBoxIcon.Error);

                                    keepWindow2 = true;
                                }
                                else
                                {
                                    keepWindow2 = false;
                                }
                            }

                            if (keepWindow2 == false)
                            {
                                // the column data types are in the excelTableColumnsDataTypes var
                                AddInUtils.GetExcelTableColumnDataTypes(excelHandle, wb, inputTable, out excelTableColumnsDataTypes);

                                // change from default type to new type
                                using (var frmDEShowExcelTableColumnsDataType = new frmDEShowExcelTableColumnsDataType(excelTableColumnsDataTypes,
                                                                                                                       AggregateTableRowsSwitch,
                                                                                                                       inputTable))
                                {
                                    if (frmDEShowExcelTableColumnsDataType.ShowDialog(frmDETInputTableName) != DialogResult.Cancel)
                                    {
                                        keepWindow2 = false;

                                        // close form 2
                                        frmDETInputTableName.Close();

                                        // user input for each column store in Transform.usrInputForAggregationProcess

                                        // based on user input aggregate rows or not
                                        TextProcessor.SaveExcelTableToCSVFile(inputTable, boolVarTableHeaderInExport, InputCSVlistSeparatorParsed,
                                                                              AggregateTableRowsSwitch, wb, activeWS);

                                        // erase all the user info about each column.
                                        // From my brief search System.Collections.Generic.Dictionary does not implement IDisposable interface
                                        Transform.usrInputForAggregationProcess.Clear();
                                    }
                                    else
                                    {
                                        keepWindow2 = true;

                                        // close form 3
                                        frmDEShowExcelTableColumnsDataType.Close();
                                    }
                                }
                            }
                        }
                    }
                    while (keepWindow2 == true);

                    // don't come back to window 1
                    _secondWindowClosed = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(excelHandle, "Error occured: " + ex.Message, "Oops!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // NOT WORKING
                //if (wb != null) Marshal.ReleaseComObject(wb);
                //if (activeWS != null) Marshal.ReleaseComObject(activeWS);
            }
        }

        // dummy code
        private void processOption2Input(Win32Window excelHandle, frmDESelectInput fDESelectInput)
        {
            MessageBox.Show(excelHandle, $"You have selected radio button {userInputOption}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}
