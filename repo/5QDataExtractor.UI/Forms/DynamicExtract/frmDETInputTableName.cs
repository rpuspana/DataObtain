using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _5QDataExtractor.UI.Forms.DynamicExtract
{
    public partial class frmDETInputTableName : Form
    {
        public bool IncludeTableHeaderValuesInExportFile()
        {
            return checkbxInclTbleHeader.Checked;
        }

        public bool GetAggregationSwitchValue()
        {
            return checkbxAggregationSwitch.Checked;
        }

        public string GetCSVlistSeparator()
        {
            //MessageBox.Show($"user entered text: {cmbListSeparator.SelectedText} and user selected text: {cmbListSeparator.Text}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // user has not touched the List separator combobox
            if (cmbListSeparator.SelectedText.Length == 0 && cmbListSeparator.Text.Length == 0)
            {
                // the csv file's list separator will be the os' list separator
                return String.Empty;
            }

            if (cmbListSeparator.SelectedText.Length == 0 && cmbListSeparator.Text.Length > 0)
            {
                // the csv file's list separator will be the os' list separator
                return cmbListSeparator.Text;
            }

            // in case of unknown error
            return "`";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        public frmDETInputTableName()
        {
            InitializeComponent();
        }

        public string GetInputTableName()
        {
            return txtInputTableName.Text;
        }

        public void DisplayErrorLabelOnInputTableNameForm(string msgError)
        {
            lblErrTableName.Text = msgError;
            lblErrTableName.Visible = true;
        }

        public void HideErrorLabelOnInputTableNameForm()
        {
            lblErrTableName.Visible = false;
        }
    }
}
