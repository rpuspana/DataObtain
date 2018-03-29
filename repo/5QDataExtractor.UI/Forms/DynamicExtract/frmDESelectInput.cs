using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _5QDataExtractor.UI.Forms
{
    public partial class frmDESelectInput : Form
    {
        public int UserSelection { get; private set;}

        public frmDESelectInput()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rbtnTabularForm.Checked == true)
            {
                // user selected Tabular Form radio button
                UserSelection = 1;

                DialogResult = DialogResult.OK;
            }
            // dummy code
            else if (rbtnInputOption2.Checked == true)
            {
                // user selected data input 2 radio button
                UserSelection = 2;

                DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show(this, "Please click on a radio button.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
