namespace ArianJahandarfardsAddIn
{
    partial class frmGoToUID
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGoToUID));
            this.lblTitle = new System.Windows.Forms.Label();
            this.txtUID = new System.Windows.Forms.TextBox();
            this.chkSearchAllOpenProjects = new System.Windows.Forms.CheckBox();
            this.pnlSeparator = new System.Windows.Forms.Panel();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblError = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.lblTitle.ForeColor = System.Drawing.Color.White;
            this.lblTitle.Location = new System.Drawing.Point(4, 13);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(48, 24);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "UID:";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtUID
            // 
            this.txtUID.BackColor = System.Drawing.Color.White;
            this.txtUID.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtUID.Location = new System.Drawing.Point(8, 37);
            this.txtUID.MaxLength = 20;
            this.txtUID.Name = "txtUID";
            this.txtUID.Size = new System.Drawing.Size(288, 23);
            this.txtUID.TabIndex = 2;
            this.txtUID.TextChanged += new System.EventHandler(this.txtUID_TextChanged);
            this.txtUID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUID_KeyDown);
            // 
            // chkSearchAllOpenProjects
            // 
            this.chkSearchAllOpenProjects.AutoSize = true;
            this.chkSearchAllOpenProjects.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.chkSearchAllOpenProjects.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.chkSearchAllOpenProjects.ForeColor = System.Drawing.Color.White;
            this.chkSearchAllOpenProjects.Location = new System.Drawing.Point(15, 63);
            this.chkSearchAllOpenProjects.Name = "chkSearchAllOpenProjects";
            this.chkSearchAllOpenProjects.Size = new System.Drawing.Size(164, 19);
            this.chkSearchAllOpenProjects.TabIndex = 3;
            this.chkSearchAllOpenProjects.Text = "Search in all open projects";
            this.chkSearchAllOpenProjects.UseVisualStyleBackColor = false;
            this.chkSearchAllOpenProjects.CheckedChanged += new System.EventHandler(this.chkSearchAllOpenProjects_CheckedChanged);
            // 
            // pnlSeparator
            // 
            this.pnlSeparator.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.pnlSeparator.Location = new System.Drawing.Point(-7, 88);
            this.pnlSeparator.Name = "pnlSeparator";
            this.pnlSeparator.Size = new System.Drawing.Size(320, 3);
            this.pnlSeparator.TabIndex = 4;
            // 
            // btnOK
            // 
            this.btnOK.BackColor = System.Drawing.Color.White;
            this.btnOK.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.btnOK.Location = new System.Drawing.Point(123, 97);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(67, 24);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "Go";
            this.btnOK.UseVisualStyleBackColor = false;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblError
            // 
            this.lblError.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblError.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.lblError.Location = new System.Drawing.Point(39, 14);
            this.lblError.Name = "lblError";
            this.lblError.Size = new System.Drawing.Size(213, 23);
            this.lblError.TabIndex = 5;
            this.lblError.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(254, 4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(42, 31);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // frmGoToUID
            // 
            this.AcceptButton = this.btnOK;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(305, 125);
            this.Controls.Add(this.chkSearchAllOpenProjects);
            this.Controls.Add(this.pnlSeparator);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.txtUID);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblError);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmGoToUID";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Go To UID";
            this.Load += new System.EventHandler(this.frmGoToUID_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        // ── Control declarations ───────────────────────────────────
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.TextBox txtUID;
        private System.Windows.Forms.CheckBox chkSearchAllOpenProjects;
        private System.Windows.Forms.Panel pnlSeparator;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblError;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}
