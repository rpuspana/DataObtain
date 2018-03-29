using Microsoft.Office.Tools.Excel;
using Interop = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using _5QDataExtractor.Library;
using _5QDataExtractor.Library.Parser;
using _5QDataExtractor.Library.Model;

namespace _5QDataExtractor.UI.Forms.DynamicExtract
{
    public partial class frmDEShowExcelTableColumnsDataType : Form
    {
        private List<string> _excelTableColumnsDataTypesData;
        private bool _aggregateRows;
        private Interop.ListObject _inputTable;

        // Excell cell val object: List{Excel cell value to string, user input as object}
        private Dictionary<string, Dictionary<object, List<object>>> _ExCellValToUserValDict;

        public frmDEShowExcelTableColumnsDataType(List<string> excelTableColumnsDataTypes,
                                                  bool aggOpperatinsColumnPresent,
                                                  Interop.ListObject excelTable)
        {
            InitializeComponent();
            _excelTableColumnsDataTypesData = excelTableColumnsDataTypes;
            _aggregateRows = aggOpperatinsColumnPresent;
            _inputTable = excelTable;
            _ExCellValToUserValDict = new Dictionary<string, Dictionary<object, List<object>>>();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void frmViewExcelColumnDataTypes_Shown(object sender, EventArgs e)
        {
            DataGridViewCheckBoxColumn columnIsAggregationKey = new DataGridViewCheckBoxColumn();
            DataGridViewComboBoxColumn columnRowsOperation = new DataGridViewComboBoxColumn();

            this.Enabled = false;
            try
            {
                this.dgvData.AutoGenerateColumns = false;

                DataTable dt = new DataTable();
                dt.Columns.Add("Table column name", typeof(String));
                dt.Columns.Add("Table column data type", typeof(String));
                if (_aggregateRows)
                {
                    dt.Columns.Add("Table column aggregation key", typeof(Boolean));
                }

                if (_aggregateRows)
                {
                    dt.Columns.Add("Rows operation", typeof(String));
                }

                for (int i = 0; i < _excelTableColumnsDataTypesData.Count; i++)
                {
                    if (!Transform.ExcelDataTypeToUserDataTypeDict.Keys.Contains(_excelTableColumnsDataTypesData[i]))
                    {
                        throw new Exception($"Application can not handle {_excelTableColumnsDataTypesData[i]} column data types." +
                                            $"The application will not exit.");
                    }

                    if (_aggregateRows)
                    {
                        dt.Rows.Add(new object[] { "Column " + (i + 1),
                                                   Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[i]],
                                                   false,
                                                   Transform.defaultRowAggOperation });
                    }
                    else
                    {
                        dt.Rows.Add(new object[] { "Column " + (i + 1),
                                                   Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[i]] });
                    }
                    
                }

                DataGridViewTextBoxColumn columnName = new DataGridViewTextBoxColumn();
                columnName.HeaderText = "Table column name";
                columnName.DataPropertyName = "Table column name";
                columnName.ReadOnly = true;
                //columnName.Name = "colName";

                DataGridViewComboBoxColumn columnDataType = new DataGridViewComboBoxColumn();
                var usrFriendlyColumnDataTypes = Transform.ExcelDataTypeToUserDataTypeDict.Values.ToList();
                columnDataType.DataSource = usrFriendlyColumnDataTypes;
                columnDataType.HeaderText = "Table column data type";
                columnDataType.DataPropertyName = "Table column data type";
                //columnName.Name = "colDataType";
                ////dropdown with white color but blue border :money.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;

                if (_aggregateRows)
                {
                    columnIsAggregationKey.Selected = false;
                    columnIsAggregationKey.HeaderText = "Table column aggregation key";
                    columnIsAggregationKey.DataPropertyName = "Table column aggregation key";
                    columnIsAggregationKey.TrueValue = true;
                    columnIsAggregationKey.FalseValue = false;
                    //columnIsAggregationKey.Name = "colAggKey";
                }

                if (_aggregateRows)
                {
                    columnRowsOperation.DataSource = Transform.RowAggregateOperationsLst;
                    columnRowsOperation.HeaderText = "Rows operation";
                    columnRowsOperation.DataPropertyName = "Rows operation";
                    //columnRowsOperation.Name = "colInput";
                }

                this.dgvData.DataSource = dt;

                if (_aggregateRows)
                {
                    this.dgvData.Columns.AddRange(columnName, columnDataType, columnIsAggregationKey,
                                                    columnRowsOperation);
                }
                else
                {
                    this.dgvData.Columns.AddRange(columnName, columnDataType);
                }

                // mapping column is present if you aggregate rows or not
                var mappingCol = new DataGridViewButtonColumn()
                {
                    UseColumnTextForButtonValue = true,
                    Text = "Map column values",
                    Name = "Map column values",
                    Width = 120,
                };

                this.dgvData.Columns.Add(mappingCol);

                // show info about the data type conversions and rows operations of each column data type
                ShowHelpTextColsConvAndRowsOperation(_aggregateRows);

                // we need this call because we are disabeling the form (all children affected) 
                // the scroll is not refreshing properly if DataGridView is disabled
                // other stuff might not work properly regarding drawing as well
                this.dgvData.PerformLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                DialogResult = DialogResult.Cancel;
            }
            finally
            {
                this.Enabled = true;
                //MessageBox.Show(this, "Form enabled again", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Show help text on the form regarding allowed column data type conversions and rows operations
        /// </summary>
        private void ShowHelpTextColsConvAndRowsOperation(bool aggOpperatinsColumnPresent)
        {
            string startMsgPart;
            string colDTconversionsForCurrColDT;
            string colRowOperationsForCurrColDT;
            string helpMsg = String.Empty;
            string msg2 = String.Empty;

            foreach (string key in Transform.PossibleColumnDataTypeChangesDict.Keys)
            {
                msg2 = String.Empty;

                startMsgPart = $"For the \"{Transform.ExcelDataTypeToUserDataTypeDict[key]}\" data type column you can:\n";

                if (Transform.PossibleColumnDataTypeChangesDict[key].Count == 1)
                {
                    msg2 = $"\"{Transform.ExcelDataTypeToUserDataTypeDict[Transform.PossibleColumnDataTypeChangesDict[key][0]]}\"";

                    colDTconversionsForCurrColDT = $" - change it's data type to this/one of these: {msg2}\n";
                }
                else
                {
                    foreach (string colConvDTKey in Transform.PossibleColumnDataTypeChangesDict[key])
                    {
                        msg2 = msg2 + $"\"{Transform.ExcelDataTypeToUserDataTypeDict[colConvDTKey]}\"" + "\n";
                    }

                    msg2 = msg2.Remove(msg2.Length - 1, 1);

                    colDTconversionsForCurrColDT = $" - change it's data type to this/one of these:\n{msg2}\n";
                }

                if (aggOpperatinsColumnPresent)
                {
                    colRowOperationsForCurrColDT = $" - do this/these row operation(s): {String.Join(" , ", Transform.ColumnRowsOperations[key].ToArray())}\n\n";

                    helpMsg = helpMsg + startMsgPart + colDTconversionsForCurrColDT + colRowOperationsForCurrColDT;
                }
                else
                {
                    colRowOperationsForCurrColDT = "\n\n";

                    helpMsg = helpMsg + startMsgPart + colDTconversionsForCurrColDT + colRowOperationsForCurrColDT;
                }
            }

            if (aggOpperatinsColumnPresent)
            {
                helpMsg = helpMsg + $"For the \"empty\" data type column you can't:\n" +
                     $" - change it's data type\n" +
                     $" - select other row operation rather than \"{Transform.defaultRowAggOperation}\"";
            }
            else
            {
                helpMsg = helpMsg + $"For the \"empty\" data type column you can't:\n" +
                     $" - change it's data type\n" ;
            }

            // display header of help msg
            lblHelpOnDTandRowOps.Visible = true;

            // display help msg
            lblInfoDTColConvOperations.Text = helpMsg;
            lblInfoDTColConvOperations.Visible = true;
        }

        // get each column data type selectd by the user in this form
        private void btnOK_Click(object sender, EventArgs e)
        {
            string[] ALL_SETTINGS_ACCEPTED_MSG = { "All settings accepted. Do you want to continue aggregating the rows and export the result to a csv file ?",
                                                   "All settings accepted. Do you want to continue exporting the table to a csv file ?" };

            lblErrUserColumnDataTypeConv.Visible = false;
            lblErrRowsOperationDenied.Visible = false;
            bool columnIsAggregationKey = false;
            bool atLeastOneAggKey = false;

            Dictionary<string, string> IntermediateStateTableColumnsDataType = new Dictionary<string, string>();

            string DefaultExcelTableColumnTypeUsrFriendly = String.Empty;
            string UsrSelColumnTypeKey = String.Empty;
            string newColDataTypeKey = String.Empty;
            string strForBooleanTRUE = String.Empty;
            string strForBooleanFALSE = String.Empty;
            string rowAggregationOperation = String.Empty;
            string allOKMsg = String.Empty;


            try
            {
                for (int i = 0; i < dgvData.Rows.Count; i++)
                {
                    //DefaultExcelTableColumnTypeUsrFriendly = UserColumnTypeToExcelType[_excelTableColumnsDataTypesData[i]];
                    DefaultExcelTableColumnTypeUsrFriendly = Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[i]];

                    // user changed from the default type to another type
                    if (DefaultExcelTableColumnTypeUsrFriendly != dgvData.Rows[i].Cells[1].Value.ToString())
                    {
                        // column "new column data type" hardcoded as column 1
                        UsrSelColumnTypeKey = Transform.ExcelDataTypeToUserDataTypeDict.
                                                FirstOrDefault(x => x.Value == dgvData.Rows[i].Cells[1].Value.ToString()).
                                                Key;

                        // validate change of old column data type to new column data type
                        if (valiDataColumnDataTypeConversion(i, DefaultExcelTableColumnTypeUsrFriendly, UsrSelColumnTypeKey))
                        {
                            newColDataTypeKey = UsrSelColumnTypeKey;
                        }
                        else
                        {
                            return;
                        }
                    }
                    // user didn't change column data type dropdown
                    else
                    {
                        // data type is _excelColDataType[i]
                        newColDataTypeKey = _excelTableColumnsDataTypesData[i];
                    }

                    // column is aggregation key
                    // column is aggregation key hardcoded as column 2
                    if (_aggregateRows)
                    {
                        columnIsAggregationKey = Convert.ToBoolean(dgvData.Rows[i].Cells[2].Value);
                        if (columnIsAggregationKey)
                        {
                            atLeastOneAggKey = true;
                        }
                    }

                    if (_aggregateRows)
                    {
                        if ((i == (dgvData.Rows.Count - 1)) && (!atLeastOneAggKey))
                        {
                            lblErrRowsOperationDenied.Text =
                                $"Please select at least one column as aggregation key";

                            lblErrRowsOperationDenied.Visible = true;

                            return;
                        }
                    }

                    if (_aggregateRows)
                    {
                        if (!columnIsAggregationKey)
                        {
                            string usrRowsAggregationMethod = dgvData.Rows[i].Cells[3].Value.ToString();

                            // validate if an operation is permited to aggregate the column's rows
                            if (validateColRowsOperationForUrColDataType(i, newColDataTypeKey, usrRowsAggregationMethod))
                            {
                                rowAggregationOperation = usrRowsAggregationMethod;
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            // no rows operation if column is not an aggregation key
                            rowAggregationOperation = Transform.defaultRowAggOperation;
                        }
                    }

                    // column name has not been escaped because initially it was named Column N
                    string colName = this.dgvData.Rows[i].Cells[0].Value.ToString();

                    // validate dictionary value and if empty cell value
                    try
                    {
                        string columName = "Column " + (i + 1);

                        if (_ExCellValToUserValDict.Keys.Contains(columName))
                        {
                            Validator.UsrMappedValsCanConvToColNewDT(("Column " + (i + 1)), newColDataTypeKey, _excelTableColumnsDataTypesData[i],
                                                                      _ExCellValToUserValDict[("Column " + (i + 1))], _inputTable, (i + 1));
                        }
                        else
                        {
                            // 15 March 2018 - for the moment no validation is to be done on null
                            if (newColDataTypeKey != "null")
                            {
                                // validate excel column
                                Validator.UsrMappedValsCanConvToColNewDT(("Column " + (i + 1)), newColDataTypeKey, _excelTableColumnsDataTypesData[i],
                                                                          null, _inputTable, (i + 1));
                            }
                        }
                        
                        //Validator.UsrMappedValsCanConvToColNewDT(_inputTable, (i + 1), _excelTableColumnsDataTypesData[i], newColDataTypeKey, _ExCellValToUserValDict[("Column " + (i + 1))]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"{ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        return;
                    }

                    // mapped values
                    // SEE IF MAPPED VALUES CAN BE CONVERTED TO the new data type
                    //  out bool usrMappedAllCellsToEmptyStr
                    //if (_ExCellValToUserValDict.Keys.Contains(colName) && (_ExCellValToUserValDict[colName] != null))
                    //{
                    //    try
                    //    {
                    //        Validator.UsrMappedValsCanConvToColNewDT(colName, newColDataTypeKey,
                    //                                             _ExCellValToUserValDict[colName]);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        MessageBox.Show(this, $"{ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    //        return;
                    //    }
                    //}
                    //else
                    //{
                    //    if (_excelTableColumnsDataTypesData[i] == "string" && newColDataTypeKey == "boolean")
                    //    {
                    //        MessageBox.Show(this, $"Please map the new boolean values on Column { i + 1 }",
                    //                        "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    //        return;
                    //    }
                    //}

                    if (Transform.usrInputForAggregationProcess.Keys.Count == 0 ||
                        !Transform.usrInputForAggregationProcess.Keys.Contains(colName))
                    {
                        // new column info
                        // create list with custom objects because all input on a dgvData row  is now valid

                        ColDTModel currColInputDataAggProcess;

                        if (_ExCellValToUserValDict.Keys.Contains(colName) && (_ExCellValToUserValDict[colName] != null))
                        {
                            currColInputDataAggProcess = new ColDTModel(_excelTableColumnsDataTypesData[i], newColDataTypeKey,
                                                                        columnIsAggregationKey, rowAggregationOperation,
                                                                        _ExCellValToUserValDict[colName]);
                        }
                        else
                        {
                            currColInputDataAggProcess = new ColDTModel(_excelTableColumnsDataTypesData[i], newColDataTypeKey,
                                                                        columnIsAggregationKey, rowAggregationOperation,
                                                                        null);
                        }

                        Transform.usrInputForAggregationProcess.Add(colName, currColInputDataAggProcess);
                    }
                    else
                    {
                        // edit column info because it exists in the dictionary
                        // a row has been updated, so update correspondng object from the dictionary of user input about 
                        // column data type and row aggregation method
                        Transform.usrInputForAggregationProcess[colName].ColumnDataType = newColDataTypeKey;
                        Transform.usrInputForAggregationProcess[colName].IsAggregationKey = columnIsAggregationKey;
                        Transform.usrInputForAggregationProcess[colName].RowsOperation = rowAggregationOperation;
                        if (_ExCellValToUserValDict.Keys.Contains(colName) && (_ExCellValToUserValDict[colName] != null))
                        {
                            Transform.usrInputForAggregationProcess[colName].ExCellValToUserValDict = _ExCellValToUserValDict[colName];
                        }
                    }
                }

                if (_aggregateRows)
                {
                    allOKMsg = ALL_SETTINGS_ACCEPTED_MSG[0];
                }
                else
                {
                    allOKMsg = ALL_SETTINGS_ACCEPTED_MSG[1];
                }

                if (MessageBox.Show(this, allOKMsg, "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);

                DialogResult = DialogResult.Cancel;
            }
        }

        // presumes and int number of columns
        private bool ValidteColDTconversion(int dgvRowIdx)
        {
            lblErrUserColumnDataTypeConv.Text = String.Empty;
            lblErrUserColumnDataTypeConv.Visible = false;

            string DefaultExcelTableColumnTypeUsrFriendly = Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[dgvRowIdx]];
            string UsrSelColumnTypeKey;

            // user changed from the default type to another type
            if (DefaultExcelTableColumnTypeUsrFriendly != dgvData.Rows[dgvRowIdx].Cells[1].Value.ToString())
            {
                // column "new column data type" hardcoded as column 1
                UsrSelColumnTypeKey = Transform.ExcelDataTypeToUserDataTypeDict.
                                             FirstOrDefault(x => x.Value == dgvData.Rows[dgvRowIdx].Cells[1].Value.ToString()).
                                             Key;

                // validate change of old column data type to new column data type
                if (valiDataColumnDataTypeConversion(dgvRowIdx, DefaultExcelTableColumnTypeUsrFriendly, UsrSelColumnTypeKey))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            // user didn't change column data type dropdown
            else
            {
                // data type is _excelColDataType[i]
                return true;
            }
        }

        private bool validateColRowsOperationForUrColDataType(int dgvRowIndex, string colDataTypeKey, string usrRowsAggregationMethod)
        {
            // if a column type does not exist
            if (!Transform.ColumnRowsOperations.Keys.Contains(colDataTypeKey))
            {
                lblErrRowsOperationDenied.Text =
                    $"Error for Column {(dgvRowIndex + 1)} : \"{Transform.ExcelDataTypeToUserDataTypeDict[colDataTypeKey]}\" column data type does not have operations";

                //dgvData.Rows[dgvRowIndex].Cells[5].Value = Transform.RowAggregateOperationsLst;

                lblErrRowsOperationDenied.Visible = true;

                return false;
            }

            // column rows operation does not exist for a column data type
            if (Transform.ColumnRowsOperations[colDataTypeKey].Count == 0)
            {
                lblErrRowsOperationDenied.Text =
                    $"Error for Column {(dgvRowIndex + 1)} : An operation for a(n) \"{Transform.ExcelDataTypeToUserDataTypeDict[colDataTypeKey]}\" column data type does not exist";

                lblErrRowsOperationDenied.Visible = true;

                return false;
            }

            // operation selected can't be used for a column data type
            if (!Transform.ColumnRowsOperations[colDataTypeKey].Contains(usrRowsAggregationMethod))
            {
                string usrFriendlyColDataTYpe = Transform.ExcelDataTypeToUserDataTypeDict[colDataTypeKey];
                string aggregationOpForColDataType = String.Join(" , ", Transform.ColumnRowsOperations[colDataTypeKey].ToArray());
                string msgPart = String.Empty;

                if (usrRowsAggregationMethod != Transform.defaultRowAggOperation)
                {
                    msgPart = $"Can't apply rows aggregation method \"{usrRowsAggregationMethod}\" " +
                              $"on a(n) \"{usrFriendlyColDataTYpe}\" column data type.";
                }
                else
                {
                    msgPart = $"Please choose an aggregation operation for the rows of \"Column {(dgvRowIndex + 1)}\"";
                }

                string msgPart2 = colDataTypeKey != "null" ? $"\nAllowed aggregation operations: {aggregationOpForColDataType}" : $"\nPlease select again \"{Transform.defaultRowAggOperation}\" from the column's dropdown list.";

                lblErrRowsOperationDenied.Text = $"Error for Column {(dgvRowIndex + 1)} : " + msgPart + msgPart2;

                lblErrRowsOperationDenied.Visible = true;

                return false;
            }

            //MessageBox.Show(this, $"operation acccepted for col data type {Transform.ExcelDataTypeToUserDataTypeDict[colDataTypeKey]}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            lblErrUserColumnDataTypeConv.Visible = false;

            return true;
        }

        private void dgvData_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            DataRowView item = this.dgvData.Rows[e.RowIndex].DataBoundItem as DataRowView;

            if (item != null)
            {
                string colName = this.dgvData.Columns[e.ColumnIndex].HeaderText;

                switch (colName)
                {
                    case "Table column aggregation key":
                        //MessageBox.Show(this, $"checkbox form cell column {e.ColumnIndex} row {e.RowIndex}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        DataGridViewCheckBoxCell checkBoxClicked = (DataGridViewCheckBoxCell)this.dgvData.Rows[e.RowIndex].Cells[e.ColumnIndex];

                        // select the corresponding "Column rows operation" cell that is 1 columns after the currently selected cell
                        DataGridViewComboBoxCell cmbBox = (DataGridViewComboBoxCell)this.dgvData.Rows[e.RowIndex].Cells[e.ColumnIndex + 1];

                        // check the "Table column aggregation key" checkbox cell state
                        if (Convert.ToBoolean(checkBoxClicked.Value) == Convert.ToBoolean(checkBoxClicked.FalseValue) || checkBoxClicked.Value == null)
                        {
                            checkBoxClicked.Value = checkBoxClicked.TrueValue;

                            // insert String.Empty in the cell
                            // TODO

                            // make the cell ReadOnly
                            cmbBox.ReadOnly = true;

                            // explain why the combobox on the same row is disabled
                            if (!lblInputExplain.Visible)
                            {
                                lblInputExplain.Visible = true;
                                lblInputExplain.Text = "Info: The rows of a column key will not support an operation applied to them. " +
                                                       "\nThe corresponding \"Rows operation\" dropdown has been blocked and any value selected will not be taken into account";
                            }
                        }
                        // disable the "Table column aggregation key" checkbox cell
                        else
                        {
                            checkBoxClicked.Value = checkBoxClicked.FalseValue;

                            // insert String.Empty
                            //cmbBox.DataSource = Transform.RowAggregateOperationsLst;

                            // make the cell ReadOnly
                            cmbBox.ReadOnly = false;

                            lblInputExplain.Visible = false;
                        }

                        this.dgvData.EndEdit();

                        break;

                    case "Map column values":
                        // MessageBox.Show(this, $"button pressed from cell column {e.ColumnIndex} row {e.RowIndex}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        // see if conversion from the old column data type to the new column data type is possible
                        if (ValidteColDTconversion(e.RowIndex))
                        {
                            Dictionary<object, List<object>> obj;

                            // first cell of the datagridview row corresponding to the clicked button
                            string columnName = dgvData.Rows[e.RowIndex].Cells[0].Value.ToString();

                            string newColumnDTkey = Transform.ExcelDataTypeToUserDataTypeDict.
                                            FirstOrDefault(x => x.Value == dgvData.Rows[e.RowIndex].Cells[1].Value.ToString()).
                                            Key;

                            int columnIndex;

                            // can't map values of an empty column
                            if (newColumnDTkey != "null")
                            {
                                try
                                {
                                    columnIndex = Convert.ToInt32(columnName.Substring(columnName.IndexOf(" ")));
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception($"Could not get the number that suppose to be inside this column name: {columnName}");
                                }

                                string columnOldDTkey = _excelTableColumnsDataTypesData[columnIndex - 1];

                                if (!_ExCellValToUserValDict.Keys.Contains(columnName)) { obj = new Dictionary<object, List<object>>(); }
                                else { obj = _ExCellValToUserValDict[columnName]; }


                                using (var frmMapTableColValToUsrVal = new frmDEMapColumnValues(_inputTable, columnName,
                                                                                                newColumnDTkey, columnOldDTkey,
                                                                                                obj))
                                {
                                    if (frmMapTableColValToUsrVal.ShowDialog(this) != DialogResult.OK)
                                    {
                                        return;
                                    }

                                    if (!_ExCellValToUserValDict.Keys.Contains(columnName))
                                    {
                                        _ExCellValToUserValDict.Add(columnName, frmMapTableColValToUsrVal.Get_ExCellValToUserValDict());
                                    }
                                    else
                                    {
                                        _ExCellValToUserValDict[columnName] = frmMapTableColValToUsrVal.Get_ExCellValToUserValDict();
                                    }
                                }
                            }
                            else
                            {
                                lblErrRowsOperationDenied.Text =
                                $"An empty column does not support mapping of values.";

                                lblErrRowsOperationDenied.Visible = true;
                            }
                        }
                        break;
                }
            }
        }

        private bool valiDataColumnDataTypeConversion(int _0BsedTblColumnIndex, string DefaultColTypeUsrFriendly, string UserSelectedColTypeKey)
        {
            // Excel table columns that can't change from the default column data type
            if (!Transform.PossibleColumnDataTypeChangesDict.Keys.Contains(_excelTableColumnsDataTypesData[_0BsedTblColumnIndex]))
            {
                lblErrUserColumnDataTypeConv.Text =
                    $"Can't convert the column's data type from \"{DefaultColTypeUsrFriendly}\" to any other data type.\n" +
                    $"Column {_0BsedTblColumnIndex + 1}'s data type reverted back to \"{Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[_0BsedTblColumnIndex]]}\"";

                dgvData.Rows[_0BsedTblColumnIndex].Cells[1].Value = Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[_0BsedTblColumnIndex]];

                lblErrUserColumnDataTypeConv.Visible = true;

                return false;
            }

            // Excel table columns that can change from the default type to another type
            if (!Transform.PossibleColumnDataTypeChangesDict[_excelTableColumnsDataTypesData[_0BsedTblColumnIndex]].Contains(UserSelectedColTypeKey))
            {
                lblErrUserColumnDataTypeConv.Text =
                    $"Can't convert the column's data type from \"{DefaultColTypeUsrFriendly}\" to \"{Transform.ExcelDataTypeToUserDataTypeDict[UserSelectedColTypeKey]}\"." +
                    $"Column {_0BsedTblColumnIndex + 1}'s data type reverted back to \"{Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[_0BsedTblColumnIndex]]}\"";

                dgvData.Rows[_0BsedTblColumnIndex].Cells[1].Value = Transform.ExcelDataTypeToUserDataTypeDict[_excelTableColumnsDataTypesData[_0BsedTblColumnIndex]];

                lblErrUserColumnDataTypeConv.Visible = true;

                return false;
            }

            //MessageBox.Show(this, "column data type conversion acccepted", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            lblErrUserColumnDataTypeConv.Visible = false;

            return true;
        }
    }
}
