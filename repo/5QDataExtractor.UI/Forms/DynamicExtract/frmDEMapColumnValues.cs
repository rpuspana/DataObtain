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

namespace _5QDataExtractor.UI.Forms.DynamicExtract
{
    public partial class frmDEMapColumnValues : Form
    {
        private Interop.ListObject _intputExTable;
        private string _colToInspect;
        private string _colToInspectNewDTKey;
        private string _colToInspectOldDTKey;
        private Dictionary<object, List<object>> _ExCellValToUserValDict;
        private bool _ExCellValUsrValueDictHasValues;

        // Excell cell val object: Excel cell value to string
        private Dictionary<object, string> _mappingExlColValsUsrVals;

        public frmDEMapColumnValues(Interop.ListObject inExcelTable, string columnName,
                                    string newColDTusrKey, string oldColDTkey,
                                    Dictionary<object, List<object>> ExCellValToUsrVal)
        {
            InitializeComponent();

            _intputExTable = inExcelTable;
            _colToInspect = columnName;

            // store the new data type of the column
            _colToInspectNewDTKey = newColDTusrKey;

            // store the aproximated column value
            _colToInspectOldDTKey = oldColDTkey;

            // Excel cell value to be mapped to a user value dict
            if (ExCellValToUsrVal != null)
            {
               _ExCellValToUserValDict = ExCellValToUsrVal;
            }
            else
            {
                _ExCellValToUserValDict = new Dictionary<object, List<object>>();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void frmMapExlTblValsToUsrVals_Shown(object sender, EventArgs e)
        {
            this.Enabled = false;
            try
            {
                string colDTUsrFriendly = Transform.ExcelDataTypeToUserDataTypeDict[_colToInspectNewDTKey];

                // place the data type in which the user wants to convert the column data type in the first label of the form
                lblMapColUniqueValsToUsrVals.Text = lblMapColUniqueValsToUsrVals.Text.Replace("NEW_TYPE", colDTUsrFriendly);

                // get unique values from the selected table column including empty cells
                //List<object> mappingExlColValsUsrVals = Parser.GetUniqueExTblColumnValues(_intputExTable, _colToInspect,
                //                                                                          _colToInspectNewDTKey, _colToInspectOldDTKey);

                this.dgvTblColValsMapUsrVals.AutoGenerateColumns = false;

                DataTable dt = new DataTable();
                dt.Columns.Add("Excel table column distinct values", typeof(String));
                dt.Columns.Add("New value", typeof(String));

                _mappingExlColValsUsrVals = Parser.GetUniqueExTblColVals(_intputExTable, _colToInspect,
                                                                         _colToInspectNewDTKey, _colToInspectOldDTKey);

                string col2Str;

                foreach (object key in _mappingExlColValsUsrVals.Keys)
                {
                    if (_ExCellValToUserValDict.Keys.Count == 0)
                    {
                        dt.Rows.Add(new object[] { _mappingExlColValsUsrVals[key], String.Empty });
                    }
                    else
                    {
                        col2Str = String.Empty;

                        if (_ExCellValToUserValDict.Keys.ToList().Contains(key))
                        {
                            if (_ExCellValToUserValDict[key][1].ToString() != String.Empty)
                            {
                                col2Str = _ExCellValToUserValDict[key][1].ToString();
                            }
                        }

                        // list of : excell value to string, user value to obj
                        dt.Rows.Add(new object[] { _mappingExlColValsUsrVals[key], col2Str });
                    }
                }

                DataGridViewTextBoxColumn colUsrMappedVal = new DataGridViewTextBoxColumn();
                colUsrMappedVal.HeaderText = "Excel table column distinct values";
                colUsrMappedVal.DataPropertyName = "Excel table column distinct values";
                //columnName.Name = "colName";
                colUsrMappedVal.ReadOnly = true;

                DataGridViewTextBoxColumn colUsrNewValue = new DataGridViewTextBoxColumn();
                colUsrNewValue.HeaderText = "New value";
                colUsrNewValue.DataPropertyName = "New value";
                //columnName.Name = "colName";

                this.dgvTblColValsMapUsrVals.DataSource = dt;

                this.dgvTblColValsMapUsrVals.Columns.AddRange(colUsrMappedVal, colUsrNewValue);

                // we need this call because we are disabeling the form (all children affected) 
                // the scroll is not refreshing properly if DataGridView is disabled
                // other stuff might not work properly regarding drawing as well
                this.dgvTblColValsMapUsrVals.PerformLayout();
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

        private void btnOK_Click(object sender, EventArgs e)
        {
            string col1CurrentValue;
            object col1CurrentKey;

            object col2CurrentKey;

            try
            {
                // collect the input from the user
                for (int i = 0; i < dgvTblColValsMapUsrVals.Rows.Count; i++)
                {
                    col1CurrentValue = dgvTblColValsMapUsrVals.Rows[i].Cells[0].Value.ToString();
                    col1CurrentKey = _ExCellValToUserValDict.FirstOrDefault(x => x.Value[0].ToString() == col1CurrentValue).Key;

                    // you didn't map any values in the past
                    if (col1CurrentKey == null)
                    {
                        col1CurrentKey = _mappingExlColValsUsrVals.FirstOrDefault(x => x.Value == col1CurrentValue).Key;
                    }

                    col2CurrentKey = dgvTblColValsMapUsrVals.Rows[i].Cells[1].Value;

                    //if (col2CurrentKey.ToString() == String.Empty)
                    //{
                    //    col2CurrentKey = "null";
                    //}

                    // GOOD
                    if (!_ExCellValToUserValDict.Keys.ToList().Contains(col1CurrentKey))
                    {
                        _ExCellValToUserValDict.Add(col1CurrentKey, new List<object>() { col1CurrentValue, col2CurrentKey });
                    }
                    else
                    {
                        _ExCellValToUserValDict[col1CurrentKey][1] = col2CurrentKey;
                    }

                    //if (col2CurrentKey.ToString() != String.Empty)
                    //{
                    //    if (!_ExCellValToUserValDict.Keys.ToList().Contains(col1CurrentKey))
                    //    {
                    //        _ExCellValToUserValDict.Add(col1CurrentKey, new List<object>() { col1CurrentValue, col2CurrentKey });
                    //    }
                    //    else
                    //    {
                    //        _ExCellValToUserValDict[col1CurrentKey][1] = col2CurrentKey;
                    //    }
                    //}
                    //else
                    //{
                    //    if (col1CurrentValue == String.Empty)
                    //    {
                    //        if (!_ExCellValToUserValDict.Keys.ToList().Contains(col1CurrentKey))
                    //        {
                    //            _ExCellValToUserValDict.Add(col1CurrentKey, new List<object>() { col1CurrentValue, col2CurrentKey });
                    //        }
                    //        else
                    //        {
                    //            _ExCellValToUserValDict[col1CurrentKey][1] = col2CurrentKey;
                    //        }
                    //    }
                    //}
                }

                DialogResult = DialogResult.OK;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public Dictionary<object, List<object>> Get_ExCellValToUserValDict()
        {
            return _ExCellValToUserValDict;
        }

        //private void ConvertObjToUsrFriendlyVal(object cellValBoxed, out string usrFriendlyValue, string currColName)
        //{
        //    // unbox object based on it's C# type
        //    switch (cellValBoxed.GetType().Name.ToLower())
        //    {
        //        case "string":
        //            // check for empty cell
        //            if (Convert.ToString(cellValBoxed) == "null")
        //            {
        //                // in the form it will appear as a blank cell
        //                usrFriendlyValue = "";
        //                return;
        //            }
        //            else
        //            {
        //                usrFriendlyValue = Convert.ToString(cellValBoxed);
        //            }
        //            break;

        //        case "decimal":
        //        case "datetime":
        //            usrFriendlyValue = Convert.ToString(cellValBoxed);
        //            break;

        //        case "boolean":
        //            usrFriendlyValue = Convert.ToString(cellValBoxed).ToUpper();
        //            break;

        //        default:
        //                throw new Exception($"Error on column {currColName}: {cellValBoxed.GetType().FullName} not a valid type to be converted to a user friendly data type when showing boxed cell value {Convert.ToString(cellValBoxed)} on frmDEMapColumnValues form.");
        //    }
        //}
    }
}
