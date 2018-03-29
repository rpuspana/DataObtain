namespace _5QDataExtractor.UI.Forms.DynamicExtract
{
    partial class frmDETInputTableName
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
            this.lblTableNameInput = new System.Windows.Forms.Label();
            this.txtInputTableName = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblErrTableName = new System.Windows.Forms.Label();
            this.checkbxInclTbleHeader = new System.Windows.Forms.CheckBox();
            this.lblValueSeparator = new System.Windows.Forms.Label();
            this.cmbListSeparator = new System.Windows.Forms.ComboBox();
            this.checkbxAggregationSwitch = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // lblTableNameInput
            // 
            this.lblTableNameInput.AutoSize = true;
            this.lblTableNameInput.Location = new System.Drawing.Point(9, 18);
            this.lblTableNameInput.Name = "lblTableNameInput";
            this.lblTableNameInput.Size = new System.Drawing.Size(295, 13);
            this.lblTableNameInput.TabIndex = 0;
            this.lblTableNameInput.Text = "Please enter a table name that exists in the current workbook";
            // 
            // txtInputTableName
            // 
            this.txtInputTableName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtInputTableName.Location = new System.Drawing.Point(12, 87);
            this.txtInputTableName.Name = "txtInputTableName";
            this.txtInputTableName.Size = new System.Drawing.Size(330, 20);
            this.txtInputTableName.TabIndex = 1;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(369, 250);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(92, 33);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(271, 250);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(92, 33);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblErrTableName
            // 
            this.lblErrTableName.AutoSize = true;
            this.lblErrTableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblErrTableName.ForeColor = System.Drawing.Color.Red;
            this.lblErrTableName.Location = new System.Drawing.Point(9, 52);
            this.lblErrTableName.Name = "lblErrTableName";
            this.lblErrTableName.Size = new System.Drawing.Size(85, 13);
            this.lblErrTableName.TabIndex = 4;
            this.lblErrTableName.Text = "lblErrTableName";
            this.lblErrTableName.Visible = false;
            // 
            // checkbxInclTbleHeader
            // 
            this.checkbxInclTbleHeader.AutoSize = true;
            this.checkbxInclTbleHeader.Checked = true;
            this.checkbxInclTbleHeader.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkbxInclTbleHeader.Location = new System.Drawing.Point(12, 124);
            this.checkbxInclTbleHeader.Name = "checkbxInclTbleHeader";
            this.checkbxInclTbleHeader.Size = new System.Drawing.Size(187, 17);
            this.checkbxInclTbleHeader.TabIndex = 5;
            this.checkbxInclTbleHeader.Text = "Include table headers in export file";
            this.checkbxInclTbleHeader.UseVisualStyleBackColor = true;
            // 
            // lblValueSeparator
            // 
            this.lblValueSeparator.AutoSize = true;
            this.lblValueSeparator.Location = new System.Drawing.Point(12, 162);
            this.lblValueSeparator.Name = "lblValueSeparator";
            this.lblValueSeparator.Size = new System.Drawing.Size(72, 13);
            this.lblValueSeparator.TabIndex = 6;
            this.lblValueSeparator.Text = "List Separator";
            // 
            // cmbListSeparator
            // 
            this.cmbListSeparator.FormattingEnabled = true;
            this.cmbListSeparator.Items.AddRange(new object[] {
            "[comma] ,",
            "[tab character]",
            "[colon] :",
            "[semicolon] ;",
            "[vertical bar] |"});
            this.cmbListSeparator.Location = new System.Drawing.Point(90, 159);
            this.cmbListSeparator.Name = "cmbListSeparator";
            this.cmbListSeparator.Size = new System.Drawing.Size(121, 21);
            this.cmbListSeparator.TabIndex = 7;
            // 
            // checkbxAggregationSwitch
            // 
            this.checkbxAggregationSwitch.AutoSize = true;
            this.checkbxAggregationSwitch.Checked = true;
            this.checkbxAggregationSwitch.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkbxAggregationSwitch.Location = new System.Drawing.Point(12, 199);
            this.checkbxAggregationSwitch.Name = "checkbxAggregationSwitch";
            this.checkbxAggregationSwitch.Size = new System.Drawing.Size(434, 17);
            this.checkbxAggregationSwitch.TabIndex = 8;
            this.checkbxAggregationSwitch.Text = "Aggregate rows based on columns (This process takes a bit more time to export to " +
    "csv)";
            this.checkbxAggregationSwitch.UseVisualStyleBackColor = true;
            // 
            // frmDETInputTableName
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(466, 299);
            this.Controls.Add(this.checkbxAggregationSwitch);
            this.Controls.Add(this.cmbListSeparator);
            this.Controls.Add(this.lblValueSeparator);
            this.Controls.Add(this.checkbxInclTbleHeader);
            this.Controls.Add(this.lblErrTableName);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.txtInputTableName);
            this.Controls.Add(this.lblTableNameInput);
            this.MaximizeBox = false;
            this.Name = "frmDETInputTableName";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Table Name Input";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTableNameInput;
        private System.Windows.Forms.TextBox txtInputTableName;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblErrTableName;
        private System.Windows.Forms.CheckBox checkbxInclTbleHeader;
        private System.Windows.Forms.Label lblValueSeparator;
        private System.Windows.Forms.ComboBox cmbListSeparator;
        private System.Windows.Forms.CheckBox checkbxAggregationSwitch;
    }
}