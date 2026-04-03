namespace Arian_Jahandarfards_MS_Project_Add_in
{
    partial class AJSettings
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.Label lblFlag;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.Label lblDate;
        private System.Windows.Forms.Label lblStartDate;
        private System.Windows.Forms.Label lblNumber;
        private System.Windows.Forms.ComboBox cboFlag;
        private System.Windows.Forms.ComboBox cboText;
        private System.Windows.Forms.ComboBox cboDate;
        private System.Windows.Forms.ComboBox cboStartDate;
        private System.Windows.Forms.ComboBox cboNumber;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel pnlSeparator;
        private System.Windows.Forms.CheckBox chkLoadFields;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AJSettings));
            this.lblFlag = new System.Windows.Forms.Label();
            this.lblText = new System.Windows.Forms.Label();
            this.lblDate = new System.Windows.Forms.Label();
            this.lblStartDate = new System.Windows.Forms.Label();
            this.lblNumber = new System.Windows.Forms.Label();
            this.cboFlag = new System.Windows.Forms.ComboBox();
            this.cboText = new System.Windows.Forms.ComboBox();
            this.cboDate = new System.Windows.Forms.ComboBox();
            this.cboStartDate = new System.Windows.Forms.ComboBox();
            this.cboNumber = new System.Windows.Forms.ComboBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pnlSeparator = new System.Windows.Forms.Panel();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.chkLoadFields = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // lblFlag
            // 
            this.lblFlag.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.lblFlag.Location = new System.Drawing.Point(12, 20);
            this.lblFlag.Name = "lblFlag";
            this.lblFlag.Size = new System.Drawing.Size(240, 20);
            this.lblFlag.TabIndex = 0;
            this.lblFlag.Text = "Flag Field - Marks Milestones to Track:";
            // 
            // lblText
            // 
            this.lblText.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.lblText.Location = new System.Drawing.Point(12, 52);
            this.lblText.Name = "lblText";
            this.lblText.Size = new System.Drawing.Size(240, 20);
            this.lblText.TabIndex = 2;
            this.lblText.Text = "Text Field - Milestones Affected:";
            // 
            // lblDate
            // 
            this.lblDate.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.lblDate.Location = new System.Drawing.Point(12, 84);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(240, 32);
            this.lblDate.TabIndex = 4;
            this.lblDate.Text = "Finish Date Field - Stores Snapshot Finish Dates:";
            this.lblDate.Click += new System.EventHandler(this.lblDate_Click);
            // 
            // lblStartDate
            // 
            this.lblStartDate.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.lblStartDate.Location = new System.Drawing.Point(12, 116);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(240, 20);
            this.lblStartDate.TabIndex = 6;
            this.lblStartDate.Text = "Start Date Field - Stores Snapshot Start:";
            // 
            // lblNumber
            // 
            this.lblNumber.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.lblNumber.Location = new System.Drawing.Point(12, 148);
            this.lblNumber.Name = "lblNumber";
            this.lblNumber.Size = new System.Drawing.Size(240, 20);
            this.lblNumber.TabIndex = 8;
            this.lblNumber.Text = "Duration Field - Stores Snapshot Duration:";
            // 
            // cboFlag
            // 
            this.cboFlag.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFlag.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.cboFlag.Location = new System.Drawing.Point(260, 17);
            this.cboFlag.Name = "cboFlag";
            this.cboFlag.Size = new System.Drawing.Size(140, 26);
            this.cboFlag.TabIndex = 1;
            // 
            // cboText
            // 
            this.cboText.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboText.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.cboText.Location = new System.Drawing.Point(260, 49);
            this.cboText.Name = "cboText";
            this.cboText.Size = new System.Drawing.Size(140, 26);
            this.cboText.TabIndex = 3;
            // 
            // cboDate
            // 
            this.cboDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDate.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.cboDate.Location = new System.Drawing.Point(260, 81);
            this.cboDate.Name = "cboDate";
            this.cboDate.Size = new System.Drawing.Size(140, 26);
            this.cboDate.TabIndex = 5;
            // 
            // cboStartDate
            // 
            this.cboStartDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboStartDate.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.cboStartDate.Location = new System.Drawing.Point(260, 113);
            this.cboStartDate.Name = "cboStartDate";
            this.cboStartDate.Size = new System.Drawing.Size(140, 26);
            this.cboStartDate.TabIndex = 7;
            // 
            // cboNumber
            // 
            this.cboNumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboNumber.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.cboNumber.Location = new System.Drawing.Point(260, 145);
            this.cboNumber.Name = "cboNumber";
            this.cboNumber.Size = new System.Drawing.Size(140, 26);
            this.cboNumber.TabIndex = 9;
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.Lavender;
            this.btnSave.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.btnSave.Location = new System.Drawing.Point(337, 233);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(65, 26);
            this.btnSave.TabIndex = 10;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.Lavender;
            this.btnCancel.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.btnCancel.Location = new System.Drawing.Point(262, 233);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(65, 26);
            this.btnCancel.TabIndex = 11;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(235)))), ((int)(((byte)(255)))));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(14, 219);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(187, 46);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // pnlSeparator
            // 
            this.pnlSeparator.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.pnlSeparator.Location = new System.Drawing.Point(-2, 209);
            this.pnlSeparator.Name = "pnlSeparator";
            this.pnlSeparator.Size = new System.Drawing.Size(450, 3);
            this.pnlSeparator.TabIndex = 20;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(235)))), ((int)(((byte)(235)))), ((int)(((byte)(255)))));
            this.pictureBox2.Location = new System.Drawing.Point(-15, 212);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(454, 71);
            this.pictureBox2.TabIndex = 21;
            this.pictureBox2.TabStop = false;
            // 
            // chkLoadFields
            // 
            this.chkLoadFields.Font = new System.Drawing.Font("Trebuchet MS", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLoadFields.Location = new System.Drawing.Point(14, 179);
            this.chkLoadFields.Name = "chkLoadFields";
            this.chkLoadFields.Size = new System.Drawing.Size(265, 20);
            this.chkLoadFields.TabIndex = 12;
            this.chkLoadFields.Text = "Load these fields into current view";
            // 
            // AJSettings
            // 
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(428, 281);
            this.Controls.Add(this.pnlSeparator);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.lblFlag);
            this.Controls.Add(this.cboFlag);
            this.Controls.Add(this.lblText);
            this.Controls.Add(this.cboText);
            this.Controls.Add(this.lblDate);
            this.Controls.Add(this.cboDate);
            this.Controls.Add(this.chkLoadFields);
            this.Controls.Add(this.lblStartDate);
            this.Controls.Add(this.cboStartDate);
            this.Controls.Add(this.lblNumber);
            this.Controls.Add(this.cboNumber);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.pictureBox2);
            this.Font = new System.Drawing.Font("Trebuchet MS", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AJSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Milestone Tracker Settings";
            this.Load += new System.EventHandler(this.AJSettings_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.PictureBox pictureBox2;
    }
}