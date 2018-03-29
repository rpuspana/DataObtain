namespace _5QDataExtractor.UI.Forms.DynamicExtract
{
    partial class frmDEMapColumnValues
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblMapColUniqueValsToUsrVals = new System.Windows.Forms.Label();
            this.lblErrUsrInputNotValid = new System.Windows.Forms.Label();
            this.dgvTblColValsMapUsrVals = new System.Windows.Forms.DataGridView();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTblColValsMapUsrVals)).BeginInit();
            this.SuspendLayout();
            // 
            // lblMapColUniqueValsToUsrVals
            // 
            this.lblMapColUniqueValsToUsrVals.AutoSize = true;
            this.lblMapColUniqueValsToUsrVals.Location = new System.Drawing.Point(13, 13);
            this.lblMapColUniqueValsToUsrVals.Name = "lblMapColUniqueValsToUsrVals";
            this.lblMapColUniqueValsToUsrVals.Size = new System.Drawing.Size(391, 13);
            this.lblMapColUniqueValsToUsrVals.TabIndex = 0;
            this.lblMapColUniqueValsToUsrVals.Text = "Map your column\'s values to values of your choice (but of NEW_TYPE data type)";
            // 
            // lblErrUsrInputNotValid
            // 
            this.lblErrUsrInputNotValid.AutoSize = true;
            this.lblErrUsrInputNotValid.ForeColor = System.Drawing.Color.Red;
            this.lblErrUsrInputNotValid.Location = new System.Drawing.Point(13, 45);
            this.lblErrUsrInputNotValid.Name = "lblErrUsrInputNotValid";
            this.lblErrUsrInputNotValid.Size = new System.Drawing.Size(164, 13);
            this.lblErrUsrInputNotValid.TabIndex = 5;
            this.lblErrUsrInputNotValid.Text = "ERR_USR_INPUT_NOT_VALID";
            this.lblErrUsrInputNotValid.Visible = false;
            // 
            // dgvTblColValsMapUsrVals
            // 
            this.dgvTblColValsMapUsrVals.AllowUserToAddRows = false;
            this.dgvTblColValsMapUsrVals.AllowUserToDeleteRows = false;
            this.dgvTblColValsMapUsrVals.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvTblColValsMapUsrVals.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTblColValsMapUsrVals.Location = new System.Drawing.Point(13, 77);
            this.dgvTblColValsMapUsrVals.Name = "dgvTblColValsMapUsrVals";
            this.dgvTblColValsMapUsrVals.RowHeadersVisible = false;
            this.dgvTblColValsMapUsrVals.Size = new System.Drawing.Size(594, 372);
            this.dgvTblColValsMapUsrVals.TabIndex = 6;
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(531, 467);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 31);
            this.btnClose.TabIndex = 7;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(450, 467);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 31);
            this.btnOK.TabIndex = 8;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // frmDEMapColumnValues
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 510);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.dgvTblColValsMapUsrVals);
            this.Controls.Add(this.lblErrUsrInputNotValid);
            this.Controls.Add(this.lblMapColUniqueValsToUsrVals);
            this.Name = "frmDEMapColumnValues";
            this.Text = "Map table column values";
            this.Shown += new System.EventHandler(this.frmMapExlTblValsToUsrVals_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTblColValsMapUsrVals)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblMapColUniqueValsToUsrVals;
        private System.Windows.Forms.Label lblErrUsrInputNotValid;
        private System.Windows.Forms.DataGridView dgvTblColValsMapUsrVals;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOK;
    }
}