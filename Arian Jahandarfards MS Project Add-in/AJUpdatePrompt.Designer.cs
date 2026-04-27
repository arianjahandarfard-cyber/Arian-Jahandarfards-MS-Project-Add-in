namespace ArianJahandarfardsAddIn
{
    partial class AJUpdatePrompt
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.PictureBox pictureBoxLogo;
        private System.Windows.Forms.Label labelSubtitle;
        private System.Windows.Forms.Label labelVersion;
        private System.Windows.Forms.Panel panelAccent;
        private System.Windows.Forms.Panel panelBody;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Label labelBody;
        private System.Windows.Forms.Panel panelDivider;
        private AJShimmerBar shimmerBar;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Panel panelBottomBorder;
        private System.Windows.Forms.Button buttonContinue;
        private System.Windows.Forms.Button buttonCancel;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.panelTop = new System.Windows.Forms.Panel();
            this.labelVersion = new System.Windows.Forms.Label();
            this.labelSubtitle = new System.Windows.Forms.Label();
            this.pictureBoxLogo = new System.Windows.Forms.PictureBox();
            this.panelAccent = new System.Windows.Forms.Panel();
            this.panelBody = new System.Windows.Forms.Panel();
            this.labelStatus = new System.Windows.Forms.Label();
            this.shimmerBar = new ArianJahandarfardsAddIn.AJUpdatePrompt.AJShimmerBar();
            this.panelDivider = new System.Windows.Forms.Panel();
            this.labelBody = new System.Windows.Forms.Label();
            this.labelTitle = new System.Windows.Forms.Label();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.buttonContinue = new System.Windows.Forms.Button();
            this.panelBottomBorder = new System.Windows.Forms.Panel();
            this.panelTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLogo)).BeginInit();
            this.panelBody.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.labelVersion);
            this.panelTop.Controls.Add(this.labelSubtitle);
            this.panelTop.Controls.Add(this.pictureBoxLogo);
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(520, 150);
            this.panelTop.TabIndex = 0;
            // 
            // labelVersion
            // 
            this.labelVersion.AutoSize = true;
            this.labelVersion.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.labelVersion.Location = new System.Drawing.Point(22, 122);
            this.labelVersion.Name = "labelVersion";
            this.labelVersion.Size = new System.Drawing.Size(37, 15);
            this.labelVersion.TabIndex = 2;
            this.labelVersion.Text = "v0.0.0";
            // 
            // labelSubtitle
            // 
            this.labelSubtitle.AutoSize = true;
            this.labelSubtitle.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.labelSubtitle.Location = new System.Drawing.Point(146, 122);
            this.labelSubtitle.Name = "labelSubtitle";
            this.labelSubtitle.Size = new System.Drawing.Size(104, 15);
            this.labelSubtitle.TabIndex = 1;
            this.labelSubtitle.Text = "MS Project Add-in";
            // 
            // pictureBoxLogo
            // 
            this.pictureBoxLogo.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxLogo.Location = new System.Drawing.Point(20, 25);
            this.pictureBoxLogo.Name = "pictureBoxLogo";
            this.pictureBoxLogo.Size = new System.Drawing.Size(230, 90);
            this.pictureBoxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxLogo.TabIndex = 0;
            this.pictureBoxLogo.TabStop = false;
            // 
            // panelAccent
            // 
            this.panelAccent.Location = new System.Drawing.Point(0, 150);
            this.panelAccent.Name = "panelAccent";
            this.panelAccent.Size = new System.Drawing.Size(520, 3);
            this.panelAccent.TabIndex = 1;
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.labelStatus);
            this.panelBody.Controls.Add(this.shimmerBar);
            this.panelBody.Controls.Add(this.panelDivider);
            this.panelBody.Controls.Add(this.labelBody);
            this.panelBody.Controls.Add(this.labelTitle);
            this.panelBody.Location = new System.Drawing.Point(0, 153);
            this.panelBody.Name = "panelBody";
            this.panelBody.Size = new System.Drawing.Size(520, 185);
            this.panelBody.TabIndex = 2;
            // 
            // labelStatus
            // 
            this.labelStatus.AutoSize = true;
            this.labelStatus.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.labelStatus.Location = new System.Drawing.Point(20, 152);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 15);
            this.labelStatus.TabIndex = 4;
            // 
            // shimmerBar
            // 
            this.shimmerBar.AccentColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.shimmerBar.Location = new System.Drawing.Point(20, 132);
            this.shimmerBar.Name = "shimmerBar";
            this.shimmerBar.NavyColor = System.Drawing.Color.FromArgb(((int)(((byte)(1)))), ((int)(((byte)(44)))), ((int)(((byte)(100)))));
            this.shimmerBar.Size = new System.Drawing.Size(474, 12);
            this.shimmerBar.TabIndex = 3;
            this.shimmerBar.Visible = false;
            // 
            // panelDivider
            // 
            this.panelDivider.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(230)))), ((int)(((byte)(230)))));
            this.panelDivider.Location = new System.Drawing.Point(20, 118);
            this.panelDivider.Name = "panelDivider";
            this.panelDivider.Size = new System.Drawing.Size(474, 1);
            this.panelDivider.TabIndex = 2;
            // 
            // labelBody
            // 
            this.labelBody.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.labelBody.Location = new System.Drawing.Point(20, 50);
            this.labelBody.Name = "labelBody";
            this.labelBody.Size = new System.Drawing.Size(474, 80);
            this.labelBody.TabIndex = 1;
            this.labelBody.Text = "Body";
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold);
            this.labelTitle.Location = new System.Drawing.Point(20, 16);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(83, 25);
            this.labelTitle.TabIndex = 0;
            this.labelTitle.Text = "AJ Tools";
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.buttonCancel);
            this.panelBottom.Controls.Add(this.buttonContinue);
            this.panelBottom.Controls.Add(this.panelBottomBorder);
            this.panelBottom.Location = new System.Drawing.Point(0, 338);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(520, 58);
            this.panelBottom.TabIndex = 3;
            // 
            // buttonCancel
            // 
            this.buttonCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCancel.Font = new System.Drawing.Font("Segoe UI", 9.5F);
            this.buttonCancel.Location = new System.Drawing.Point(293, 11);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(85, 36);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // buttonContinue
            // 
            this.buttonContinue.FlatAppearance.BorderSize = 0;
            this.buttonContinue.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonContinue.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold);
            this.buttonContinue.Location = new System.Drawing.Point(388, 11);
            this.buttonContinue.Name = "buttonContinue";
            this.buttonContinue.Size = new System.Drawing.Size(110, 36);
            this.buttonContinue.TabIndex = 1;
            this.buttonContinue.Text = "Update";
            this.buttonContinue.UseVisualStyleBackColor = true;
            this.buttonContinue.Click += new System.EventHandler(this.buttonContinue_Click);
            // 
            // panelBottomBorder
            // 
            this.panelBottomBorder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(215)))), ((int)(((byte)(215)))), ((int)(((byte)(215)))));
            this.panelBottomBorder.Location = new System.Drawing.Point(0, 0);
            this.panelBottomBorder.Name = "panelBottomBorder";
            this.panelBottomBorder.Size = new System.Drawing.Size(520, 1);
            this.panelBottomBorder.TabIndex = 0;
            // 
            // AJUpdatePrompt
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(520, 400);
            this.Controls.Add(this.panelBottom);
            this.Controls.Add(this.panelBody);
            this.Controls.Add(this.panelAccent);
            this.Controls.Add(this.panelTop);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AJUpdatePrompt";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AJ Tools";
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLogo)).EndInit();
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }
    }
}
