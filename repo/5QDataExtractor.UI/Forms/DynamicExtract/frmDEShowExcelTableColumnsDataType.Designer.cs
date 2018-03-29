namespace _5QDataExtractor.UI.Forms.DynamicExtract
{
    partial class frmDEShowExcelTableColumnsDataType
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
            this.lblFormTitle = new System.Windows.Forms.Label();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblErrUserColumnDataTypeConv = new System.Windows.Forms.Label();
            this.lblInputExplain = new System.Windows.Forms.Label();
            this.lblErrRowsOperationDenied = new System.Windows.Forms.Label();
            this.lblInfoDTColConvOperations = new System.Windows.Forms.Label();
            this.lblHelpOnDTandRowOps = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.SuspendLayout();
            // 
            // lblFormTitle
            // 
            this.lblFormTitle.AutoSize = true;
            this.lblFormTitle.Location = new System.Drawing.Point(13, 9);
            this.lblFormTitle.Name = "lblFormTitle";
            this.lblFormTitle.Size = new System.Drawing.Size(159, 13);
            this.lblFormTitle.TabIndex = 0;
            this.lblFormTitle.Text = "Default table column data types.";
            // 
            // dgvData
            // 
            this.dgvData.AllowUserToAddRows = false;
            this.dgvData.AllowUserToDeleteRows = false;
            this.dgvData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvData.Location = new System.Drawing.Point(387, 158);
            this.dgvData.Name = "dgvData";
            this.dgvData.RowHeadersVisible = false;
            this.dgvData.Size = new System.Drawing.Size(648, 289);
            this.dgvData.TabIndex = 1;
            this.dgvData.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvData_CellClick);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(960, 468);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 31);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(879, 468);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 31);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblErrUserColumnDataTypeConv
            // 
            this.lblErrUserColumnDataTypeConv.AutoSize = true;
            this.lblErrUserColumnDataTypeConv.ForeColor = System.Drawing.Color.Red;
            this.lblErrUserColumnDataTypeConv.Location = new System.Drawing.Point(13, 39);
            this.lblErrUserColumnDataTypeConv.Name = "lblErrUserColumnDataTypeConv";
            this.lblErrUserColumnDataTypeConv.Size = new System.Drawing.Size(216, 13);
            this.lblErrUserColumnDataTypeConv.TabIndex = 4;
            this.lblErrUserColumnDataTypeConv.Text = "ERR_USR_COLUMN_DATA_TYPE_CONV";
            this.lblErrUserColumnDataTypeConv.Visible = false;
            // 
            // lblInputExplain
            // 
            this.lblInputExplain.AutoSize = true;
            this.lblInputExplain.ForeColor = System.Drawing.Color.Black;
            this.lblInputExplain.Location = new System.Drawing.Point(13, 113);
            this.lblInputExplain.Name = "lblInputExplain";
            this.lblInputExplain.Size = new System.Drawing.Size(71, 13);
            this.lblInputExplain.TabIndex = 5;
            this.lblInputExplain.Text = "INPUT_INFO";
            this.lblInputExplain.Visible = false;
            // 
            // lblErrRowsOperationDenied
            // 
            this.lblErrRowsOperationDenied.AutoSize = true;
            this.lblErrRowsOperationDenied.ForeColor = System.Drawing.Color.Red;
            this.lblErrRowsOperationDenied.Location = new System.Drawing.Point(13, 75);
            this.lblErrRowsOperationDenied.Name = "lblErrRowsOperationDenied";
            this.lblErrRowsOperationDenied.Size = new System.Drawing.Size(215, 13);
            this.lblErrRowsOperationDenied.TabIndex = 6;
            this.lblErrRowsOperationDenied.Text = "ERR_USR_ROWS_OPERATION_DENIED";
            this.lblErrRowsOperationDenied.Visible = false;
            // 
            // lblInfoDTColConvOperations
            // 
            this.lblInfoDTColConvOperations.AutoEllipsis = true;
            this.lblInfoDTColConvOperations.AutoSize = true;
            this.lblInfoDTColConvOperations.ForeColor = System.Drawing.Color.Black;
            this.lblInfoDTColConvOperations.Location = new System.Drawing.Point(12, 190);
            this.lblInfoDTColConvOperations.Name = "lblInfoDTColConvOperations";
            this.lblInfoDTColConvOperations.Size = new System.Drawing.Size(141, 13);
            this.lblInfoDTColConvOperations.TabIndex = 7;
            this.lblInfoDTColConvOperations.Text = "lblInfoDTColConvOperations";
            this.lblInfoDTColConvOperations.Visible = false;
            // 
            // lblHelpOnDTandRowOps
            // 
            this.lblHelpOnDTandRowOps.AutoSize = true;
            this.lblHelpOnDTandRowOps.ForeColor = System.Drawing.Color.Black;
            this.lblHelpOnDTandRowOps.Location = new System.Drawing.Point(13, 158);
            this.lblHelpOnDTandRowOps.Name = "lblHelpOnDTandRowOps";
            this.lblHelpOnDTandRowOps.Size = new System.Drawing.Size(29, 13);
            this.lblHelpOnDTandRowOps.TabIndex = 8;
            this.lblHelpOnDTandRowOps.Text = "Help";
            this.lblHelpOnDTandRowOps.Visible = false;
            // 
            // frmDEShowExcelTableColumnsDataType
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1047, 511);
            this.Controls.Add(this.lblHelpOnDTandRowOps);
            this.Controls.Add(this.lblInfoDTColConvOperations);
            this.Controls.Add(this.lblErrRowsOperationDenied);
            this.Controls.Add(this.lblInputExplain);
            this.Controls.Add(this.lblErrUserColumnDataTypeConv);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.dgvData);
            this.Controls.Add(this.lblFormTitle);
            this.Name = "frmDEShowExcelTableColumnsDataType";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Table columns aggregation input";
            this.Shown += new System.EventHandler(this.frmViewExcelColumnDataTypes_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFormTitle;
        private System.Windows.Forms.DataGridView dgvData;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblErrUserColumnDataTypeConv;
        private System.Windows.Forms.Label lblInputExplain;
        private System.Windows.Forms.Label lblErrRowsOperationDenied;
        private System.Windows.Forms.Label lblInfoDTColConvOperations;
        private System.Windows.Forms.Label lblHelpOnDTandRowOps;
    }
}