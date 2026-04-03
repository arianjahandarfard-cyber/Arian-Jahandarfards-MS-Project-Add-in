namespace Arian_Jahandarfards_MS_Project_Add_in
{
    partial class AJProgress
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AJProgress));
            this.lblStatus = new System.Windows.Forms.Label();
            this.pnlBarBg = new System.Windows.Forms.Panel();
            this.pnlBarFill = new System.Windows.Forms.Panel();
            this.lblPercent = new System.Windows.Forms.Label();
            this.pnlSeparator = new System.Windows.Forms.Panel();
            this.pnlShimmer = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.lblStatus.Location = new System.Drawing.Point(12, 20);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 18);
            this.lblStatus.TabIndex = 0;
            // 
            // pnlBarBg
            // 
            this.pnlBarBg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(230)))), ((int)(((byte)(245)))));
            this.pnlBarBg.Location = new System.Drawing.Point(12, 43);
            this.pnlBarBg.Name = "pnlBarBg";
            this.pnlBarBg.Size = new System.Drawing.Size(350, 22);
            this.pnlBarBg.TabIndex = 1;
            // 
            // pnlBarFill
            // 
            this.pnlBarFill.BackColor = System.Drawing.Color.MediumBlue;
            this.pnlBarFill.Location = new System.Drawing.Point(12, 43);
            this.pnlBarFill.Name = "pnlBarFill";
            this.pnlBarFill.Size = new System.Drawing.Size(0, 22);
            this.pnlBarFill.TabIndex = 2;
            // 
            // lblPercent
            // 
            this.lblPercent.AutoSize = true;
            this.lblPercent.BackColor = System.Drawing.Color.Transparent;
            this.lblPercent.Font = new System.Drawing.Font("Trebuchet MS", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPercent.ForeColor = System.Drawing.Color.Navy;
            this.lblPercent.Location = new System.Drawing.Point(163, 63);
            this.lblPercent.Name = "lblPercent";
            this.lblPercent.Size = new System.Drawing.Size(30, 22);
            this.lblPercent.TabIndex = 3;
            this.lblPercent.Text = "0%";
            this.lblPercent.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // pnlSeparator
            // 
            this.pnlSeparator.BackColor = System.Drawing.Color.Transparent;
            this.pnlSeparator.Location = new System.Drawing.Point(12, 43);
            this.pnlSeparator.Name = "pnlSeparator";
            this.pnlSeparator.Size = new System.Drawing.Size(350, 22);
            this.pnlSeparator.TabIndex = 5;
            // 
            // pnlShimmer
            // 
            this.pnlShimmer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(1)))), ((int)(((byte)(44)))), ((int)(((byte)(100)))));
            this.pnlShimmer.Location = new System.Drawing.Point(-1, 105);
            this.pnlShimmer.Name = "pnlShimmer";
            this.pnlShimmer.Size = new System.Drawing.Size(377, 3);
            this.pnlShimmer.TabIndex = 6;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(330, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(30, 20);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(235)))), ((int)(((byte)(255)))));
            this.pictureBox2.Location = new System.Drawing.Point(-16, 107);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(434, 48);
            this.pictureBox2.TabIndex = 7;
            this.pictureBox2.TabStop = false;
            // 
            // AJProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(372, 99);
            this.Controls.Add(this.pnlShimmer);
            this.Controls.Add(this.pnlBarFill);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.pnlBarBg);
            this.Controls.Add(this.lblPercent);
            this.Controls.Add(this.pnlSeparator);
            this.Controls.Add(this.pictureBox2);
            this.Name = "AJProgress";
            this.Text = "Milestone Impact Tracker";
            this.Load += new System.EventHandler(this.AJProgress_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Panel pnlBarBg;
        private System.Windows.Forms.Panel pnlBarFill;
        private System.Windows.Forms.Label lblPercent;
        private System.Windows.Forms.Panel pnlSeparator;
        private System.Windows.Forms.Panel pnlShimmer;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
    }
}