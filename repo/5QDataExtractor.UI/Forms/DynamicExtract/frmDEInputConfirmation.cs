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
    public partial class frmDEInputConfirmation : Form
    {
        public frmDEInputConfirmation(string confirmationMsgChunck)
        {
            InitializeComponent();

            // customize the confirmation message based on user input selection
            lblInputTypeConfirmation.Text = lblInputTypeConfirmation.Text.Replace("[INPUT_TYPE]", confirmationMsgChunck);

        }

        private void btnNo_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.No;
        }

        private void btnYes_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
        }
    }
}
