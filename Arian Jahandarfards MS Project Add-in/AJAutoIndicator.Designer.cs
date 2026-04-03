namespace Arian_Jahandarfards_MS_Project_Add_in
{
    partial class AJAutoIndicator
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AJAutoIndicator));
            this.picSpinner = new System.Windows.Forms.PictureBox();
            this.lblStatus = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.picSpinner)).BeginInit();
            this.SuspendLayout();
            // 
            // picSpinner
            // 
            this.picSpinner.Image = ((System.Drawing.Image)(resources.GetObject("picSpinner.Image")));
            this.picSpinner.Location = new System.Drawing.Point(4, 6);
            this.picSpinner.Name = "picSpinner";
            this.picSpinner.Size = new System.Drawing.Size(42, 37);
            this.picSpinner.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picSpinner.TabIndex = 0;
            this.picSpinner.TabStop = false;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Trebuchet MS", 8.5F, System.Drawing.FontStyle.Bold);
            this.lblStatus.ForeColor = System.Drawing.Color.White;
            this.lblStatus.Location = new System.Drawing.Point(46, 15);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(89, 18);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "Auto Track ON";
            // 
            // AJAutoIndicator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(160, 48);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.picSpinner);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AJAutoIndicator";
            this.Text = "Tracker Status";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.picSpinner)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox picSpinner;
        private System.Windows.Forms.Label lblStatus;
    }
}