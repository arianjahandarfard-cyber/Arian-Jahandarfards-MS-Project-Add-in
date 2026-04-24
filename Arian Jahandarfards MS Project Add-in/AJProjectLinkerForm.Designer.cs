namespace Arian_Jahandarfards_MS_Project_Add_in
{
    partial class AJProjectLinkerForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelShell;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.PictureBox pictureBoxLogo;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Panel panelAccent;
        private System.Windows.Forms.Panel panelBody;
        private System.Windows.Forms.Label labelMode;
        private System.Windows.Forms.Label labelModeValue;
        private System.Windows.Forms.Panel panelPowerDot;
        private System.Windows.Forms.Label labelPower;
        private System.Windows.Forms.Label labelPowerValue;
        private System.Windows.Forms.Panel panelDivider;
        private System.Windows.Forms.Label labelStatus;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.panelShell = new System.Windows.Forms.Panel();
            this.panelBody = new System.Windows.Forms.Panel();
            this.labelStatus = new System.Windows.Forms.Label();
            this.panelDivider = new System.Windows.Forms.Panel();
            this.labelPowerValue = new System.Windows.Forms.Label();
            this.labelPower = new System.Windows.Forms.Label();
            this.panelPowerDot = new System.Windows.Forms.Panel();
            this.labelModeValue = new System.Windows.Forms.Label();
            this.labelMode = new System.Windows.Forms.Label();
            this.panelAccent = new System.Windows.Forms.Panel();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelTitle = new System.Windows.Forms.Label();
            this.pictureBoxLogo = new System.Windows.Forms.PictureBox();
            this.panelShell.SuspendLayout();
            this.panelBody.SuspendLayout();
            this.panelHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // panelShell
            // 
            this.panelShell.Controls.Add(this.panelBody);
            this.panelShell.Controls.Add(this.panelAccent);
            this.panelShell.Controls.Add(this.panelHeader);
            this.panelShell.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelShell.Location = new System.Drawing.Point(1, 1);
            this.panelShell.Name = "panelShell";
            this.panelShell.Size = new System.Drawing.Size(198, 66);
            this.panelShell.TabIndex = 0;
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.labelStatus);
            this.panelBody.Controls.Add(this.panelDivider);
            this.panelBody.Controls.Add(this.labelPowerValue);
            this.panelBody.Controls.Add(this.labelPower);
            this.panelBody.Controls.Add(this.panelPowerDot);
            this.panelBody.Controls.Add(this.labelModeValue);
            this.panelBody.Controls.Add(this.labelMode);
            this.panelBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBody.Location = new System.Drawing.Point(0, 25);
            this.panelBody.Name = "panelBody";
            this.panelBody.Size = new System.Drawing.Size(198, 41);
            this.panelBody.TabIndex = 2;
            // 
            // labelStatus
            // 
            this.labelStatus.Font = new System.Drawing.Font("Segoe UI", 6.8F, System.Drawing.FontStyle.Bold);
            this.labelStatus.Location = new System.Drawing.Point(10, 21);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(180, 16);
            this.labelStatus.TabIndex = 6;
            this.labelStatus.Text = "Project Linker is on.";
            // 
            // panelDivider
            // 
            this.panelDivider.Location = new System.Drawing.Point(8, 17);
            this.panelDivider.Name = "panelDivider";
            this.panelDivider.Size = new System.Drawing.Size(184, 1);
            this.panelDivider.TabIndex = 5;
            // 
            // labelPowerValue
            // 
            this.labelPowerValue.AutoSize = true;
            this.labelPowerValue.Font = new System.Drawing.Font("Segoe UI", 6.7F, System.Drawing.FontStyle.Bold);
            this.labelPowerValue.Location = new System.Drawing.Point(88, 0);
            this.labelPowerValue.Name = "labelPowerValue";
            this.labelPowerValue.Size = new System.Drawing.Size(18, 12);
            this.labelPowerValue.TabIndex = 4;
            this.labelPowerValue.Text = "On";
            // 
            // labelPower
            // 
            this.labelPower.AutoSize = true;
            this.labelPower.Font = new System.Drawing.Font("Segoe UI", 6.7F);
            this.labelPower.Location = new System.Drawing.Point(50, 0);
            this.labelPower.Name = "labelPower";
            this.labelPower.Size = new System.Drawing.Size(30, 12);
            this.labelPower.TabIndex = 3;
            this.labelPower.Text = "Status";
            // 
            // panelPowerDot
            // 
            this.panelPowerDot.Location = new System.Drawing.Point(36, 3);
            this.panelPowerDot.Name = "panelPowerDot";
            this.panelPowerDot.Size = new System.Drawing.Size(8, 8);
            this.panelPowerDot.TabIndex = 2;
            // 
            // labelModeValue
            // 
            this.labelModeValue.AutoSize = true;
            this.labelModeValue.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold);
            this.labelModeValue.Location = new System.Drawing.Point(52, 0);
            this.labelModeValue.Name = "labelModeValue";
            this.labelModeValue.Size = new System.Drawing.Size(31, 13);
            this.labelModeValue.TabIndex = 1;
            this.labelModeValue.Text = "Excel";
            // 
            // labelMode
            // 
            this.labelMode.AutoSize = true;
            this.labelMode.Font = new System.Drawing.Font("Segoe UI", 6.7F);
            this.labelMode.Location = new System.Drawing.Point(10, 1);
            this.labelMode.Name = "labelMode";
            this.labelMode.Size = new System.Drawing.Size(28, 12);
            this.labelMode.TabIndex = 0;
            this.labelMode.Text = "Mode";
            // 
            // panelAccent
            // 
            this.panelAccent.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelAccent.Location = new System.Drawing.Point(0, 24);
            this.panelAccent.Name = "panelAccent";
            this.panelAccent.Size = new System.Drawing.Size(198, 1);
            this.panelAccent.TabIndex = 1;
            // 
            // panelHeader
            // 
            this.panelHeader.Controls.Add(this.buttonClose);
            this.panelHeader.Controls.Add(this.labelTitle);
            this.panelHeader.Controls.Add(this.pictureBoxLogo);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(198, 24);
            this.panelHeader.TabIndex = 0;
            this.panelHeader.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // buttonClose
            // 
            this.buttonClose.FlatAppearance.BorderSize = 0;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.Font = new System.Drawing.Font("Segoe UI", 6.1F, System.Drawing.FontStyle.Bold);
            this.buttonClose.Location = new System.Drawing.Point(176, 3);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(16, 14);
            this.buttonClose.TabIndex = 2;
            this.buttonClose.Text = "X";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 8.3F, System.Drawing.FontStyle.Bold);
            this.labelTitle.Location = new System.Drawing.Point(34, 5);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(75, 15);
            this.labelTitle.TabIndex = 1;
            this.labelTitle.Text = "Project Linker";
            this.labelTitle.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // pictureBoxLogo
            // 
            this.pictureBoxLogo.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxLogo.Location = new System.Drawing.Point(8, 2);
            this.pictureBoxLogo.Name = "pictureBoxLogo";
            this.pictureBoxLogo.Size = new System.Drawing.Size(20, 20);
            this.pictureBoxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxLogo.TabIndex = 0;
            this.pictureBoxLogo.TabStop = false;
            this.pictureBoxLogo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // AJProjectLinkerForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.ClientSize = new System.Drawing.Size(200, 68);
            this.Controls.Add(this.panelShell);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AJProjectLinkerForm";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Project Linker";
            this.TopMost = true;
            this.panelShell.ResumeLayout(false);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelHeader.ResumeLayout(false);
            this.panelHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLogo)).EndInit();
            this.ResumeLayout(false);

        }
    }
}
