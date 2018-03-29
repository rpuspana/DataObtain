namespace _5QDataExtractor.UI.Forms
{
    partial class frmDESelectInput
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
            this.lblDataInput = new System.Windows.Forms.Label();
            this.rbtnTabularForm = new System.Windows.Forms.RadioButton();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.rbtnInputOption2 = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // lblDataInput
            // 
            this.lblDataInput.AutoSize = true;
            this.lblDataInput.Location = new System.Drawing.Point(13, 13);
            this.lblDataInput.Name = "lblDataInput";
            this.lblDataInput.Size = new System.Drawing.Size(87, 13);
            this.lblDataInput.TabIndex = 0;
            this.lblDataInput.Text = "Select data input";
            // 
            // rbtnTabularForm
            // 
            this.rbtnTabularForm.AutoSize = true;
            this.rbtnTabularForm.Location = new System.Drawing.Point(16, 46);
            this.rbtnTabularForm.Name = "rbtnTabularForm";
            this.rbtnTabularForm.Size = new System.Drawing.Size(87, 17);
            this.rbtnTabularForm.TabIndex = 1;
            this.rbtnTabularForm.TabStop = true;
            this.rbtnTabularForm.Text = "Tabular Form";
            this.rbtnTabularForm.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(180, 216);
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
            this.btnOK.Location = new System.Drawing.Point(82, 216);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(92, 33);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // rbtnInputOption2
            // 
            this.rbtnInputOption2.AutoSize = true;
            this.rbtnInputOption2.Location = new System.Drawing.Point(16, 69);
            this.rbtnInputOption2.Name = "rbtnInputOption2";
            this.rbtnInputOption2.Size = new System.Drawing.Size(92, 17);
            this.rbtnInputOption2.TabIndex = 4;
            this.rbtnInputOption2.TabStop = true;
            this.rbtnInputOption2.Text = "Input Option 2";
            this.rbtnInputOption2.UseVisualStyleBackColor = true;
            // 
            // frmDESelectInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.ControlBox = false;
            this.Controls.Add(this.rbtnInputOption2);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.rbtnTabularForm);
            this.Controls.Add(this.lblDataInput);
            this.Name = "frmDESelectInput";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Data Extractor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblDataInput;
        private System.Windows.Forms.RadioButton rbtnTabularForm;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.RadioButton rbtnInputOption2;
    }
}