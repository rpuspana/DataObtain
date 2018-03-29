namespace _5QDataExtractor.UI.Forms
{
    partial class frmProgress
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
            this.lblProgressBarMessage = new System.Windows.Forms.Label();
            this.prbarProgressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // lblProgressBarMessage
            // 
            this.lblProgressBarMessage.AutoSize = true;
            this.lblProgressBarMessage.Location = new System.Drawing.Point(13, 13);
            this.lblProgressBarMessage.Name = "lblProgressBarMessage";
            this.lblProgressBarMessage.Size = new System.Drawing.Size(16, 13);
            this.lblProgressBarMessage.TabIndex = 0;
            this.lblProgressBarMessage.Text = "...";
            // 
            // prbarProgressBar
            // 
            this.prbarProgressBar.Location = new System.Drawing.Point(12, 96);
            this.prbarProgressBar.Name = "prbarProgressBar";
            this.prbarProgressBar.Size = new System.Drawing.Size(698, 23);
            this.prbarProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.prbarProgressBar.TabIndex = 1;
            // 
            // frmProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(723, 131);
            this.Controls.Add(this.prbarProgressBar);
            this.Controls.Add(this.lblProgressBarMessage);
            this.MaximizeBox = false;
            this.Name = "frmProgress";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "frmProgress";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblProgressBarMessage;
        private System.Windows.Forms.ProgressBar prbarProgressBar;
    }
}