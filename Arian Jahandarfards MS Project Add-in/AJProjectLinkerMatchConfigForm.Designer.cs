namespace Arian_Jahandarfards_MS_Project_Add_in
{
    partial class AJProjectLinkerMatchConfigForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel panelShell;
        private System.Windows.Forms.Panel panelHeader;
        private System.Windows.Forms.PictureBox pictureBoxLogo;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Label labelSubtitle;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Panel panelAccent;
        private System.Windows.Forms.Panel panelBody;
        private System.Windows.Forms.Label labelIntro;
        private System.Windows.Forms.CheckBox checkBoxUniqueId;
        private System.Windows.Forms.ComboBox comboBoxUniqueId;
        private System.Windows.Forms.CheckBox checkBoxTaskName;
        private System.Windows.Forms.ComboBox comboBoxTaskName;
        private System.Windows.Forms.Label labelNote;
        private System.Windows.Forms.Panel panelSeparator;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonSave;

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
            this.buttonSave = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.panelSeparator = new System.Windows.Forms.Panel();
            this.labelNote = new System.Windows.Forms.Label();
            this.comboBoxTaskName = new System.Windows.Forms.ComboBox();
            this.checkBoxTaskName = new System.Windows.Forms.CheckBox();
            this.comboBoxUniqueId = new System.Windows.Forms.ComboBox();
            this.checkBoxUniqueId = new System.Windows.Forms.CheckBox();
            this.labelIntro = new System.Windows.Forms.Label();
            this.panelAccent = new System.Windows.Forms.Panel();
            this.panelHeader = new System.Windows.Forms.Panel();
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelSubtitle = new System.Windows.Forms.Label();
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
            this.panelShell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(13)))), ((int)(((byte)(31)))));
            this.panelShell.Controls.Add(this.panelBody);
            this.panelShell.Controls.Add(this.panelAccent);
            this.panelShell.Controls.Add(this.panelHeader);
            this.panelShell.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelShell.Location = new System.Drawing.Point(1, 1);
            this.panelShell.Name = "panelShell";
            this.panelShell.Size = new System.Drawing.Size(386, 370);
            this.panelShell.TabIndex = 0;
            // 
            // panelBody
            // 
            this.panelBody.BackColor = System.Drawing.Color.White;
            this.panelBody.Controls.Add(this.buttonSave);
            this.panelBody.Controls.Add(this.buttonCancel);
            this.panelBody.Controls.Add(this.panelSeparator);
            this.panelBody.Controls.Add(this.labelNote);
            this.panelBody.Controls.Add(this.comboBoxTaskName);
            this.panelBody.Controls.Add(this.checkBoxTaskName);
            this.panelBody.Controls.Add(this.comboBoxUniqueId);
            this.panelBody.Controls.Add(this.checkBoxUniqueId);
            this.panelBody.Controls.Add(this.labelIntro);
            this.panelBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBody.Location = new System.Drawing.Point(0, 84);
            this.panelBody.Name = "panelBody";
            this.panelBody.Size = new System.Drawing.Size(386, 286);
            this.panelBody.TabIndex = 2;
            // 
            // buttonSave
            // 
            this.buttonSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.buttonSave.FlatAppearance.BorderSize = 0;
            this.buttonSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonSave.ForeColor = System.Drawing.Color.White;
            this.buttonSave.Location = new System.Drawing.Point(256, 244);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(96, 30);
            this.buttonSave.TabIndex = 8;
            this.buttonSave.Text = "Save";
            this.buttonSave.UseVisualStyleBackColor = false;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.BackColor = System.Drawing.Color.WhiteSmoke;
            this.buttonCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonCancel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(58)))), ((int)(((byte)(74)))));
            this.buttonCancel.Location = new System.Drawing.Point(152, 244);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(96, 30);
            this.buttonCancel.TabIndex = 7;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = false;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // panelSeparator
            // 
            this.panelSeparator.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.panelSeparator.Location = new System.Drawing.Point(0, 232);
            this.panelSeparator.Name = "panelSeparator";
            this.panelSeparator.Size = new System.Drawing.Size(386, 3);
            this.panelSeparator.TabIndex = 6;
            // 
            // labelNote
            // 
            this.labelNote.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.labelNote.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(88)))), ((int)(((byte)(97)))), ((int)(((byte)(109)))));
            this.labelNote.Location = new System.Drawing.Point(24, 184);
            this.labelNote.Name = "labelNote";
            this.labelNote.Size = new System.Drawing.Size(336, 48);
            this.labelNote.TabIndex = 5;
            this.labelNote.Text = "Click anywhere in the Excel sheet row to find the task. If both options are chec" +
    "ked, Unique ID is used first and Task Name is the fallback.";
            // 
            // comboBoxTaskName
            // 
            this.comboBoxTaskName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxTaskName.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.comboBoxTaskName.FormattingEnabled = true;
            this.comboBoxTaskName.Location = new System.Drawing.Point(42, 148);
            this.comboBoxTaskName.Name = "comboBoxTaskName";
            this.comboBoxTaskName.Size = new System.Drawing.Size(310, 23);
            this.comboBoxTaskName.TabIndex = 4;
            // 
            // checkBoxTaskName
            // 
            this.checkBoxTaskName.AutoSize = true;
            this.checkBoxTaskName.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.checkBoxTaskName.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(13)))), ((int)(((byte)(31)))));
            this.checkBoxTaskName.Location = new System.Drawing.Point(24, 122);
            this.checkBoxTaskName.Name = "checkBoxTaskName";
            this.checkBoxTaskName.Size = new System.Drawing.Size(87, 19);
            this.checkBoxTaskName.TabIndex = 3;
            this.checkBoxTaskName.Text = "Task Name";
            this.checkBoxTaskName.UseVisualStyleBackColor = true;
            this.checkBoxTaskName.CheckedChanged += new System.EventHandler(this.checkBoxTaskName_CheckedChanged);
            // 
            // comboBoxUniqueId
            // 
            this.comboBoxUniqueId.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxUniqueId.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.comboBoxUniqueId.FormattingEnabled = true;
            this.comboBoxUniqueId.Location = new System.Drawing.Point(42, 84);
            this.comboBoxUniqueId.Name = "comboBoxUniqueId";
            this.comboBoxUniqueId.Size = new System.Drawing.Size(310, 23);
            this.comboBoxUniqueId.TabIndex = 2;
            // 
            // checkBoxUniqueId
            // 
            this.checkBoxUniqueId.AutoSize = true;
            this.checkBoxUniqueId.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.checkBoxUniqueId.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(13)))), ((int)(((byte)(31)))));
            this.checkBoxUniqueId.Location = new System.Drawing.Point(24, 58);
            this.checkBoxUniqueId.Name = "checkBoxUniqueId";
            this.checkBoxUniqueId.Size = new System.Drawing.Size(79, 19);
            this.checkBoxUniqueId.TabIndex = 1;
            this.checkBoxUniqueId.Text = "Unique ID";
            this.checkBoxUniqueId.UseVisualStyleBackColor = true;
            this.checkBoxUniqueId.CheckedChanged += new System.EventHandler(this.checkBoxUniqueId_CheckedChanged);
            // 
            // labelIntro
            // 
            this.labelIntro.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.labelIntro.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(47)))), ((int)(((byte)(58)))), ((int)(((byte)(74)))));
            this.labelIntro.Location = new System.Drawing.Point(22, 14);
            this.labelIntro.Name = "labelIntro";
            this.labelIntro.Size = new System.Drawing.Size(342, 38);
            this.labelIntro.TabIndex = 0;
            this.labelIntro.Text = "Choose which Excel column Project Linker should read when you click anywhere on " +
    "a row.";
            // 
            // panelAccent
            // 
            this.panelAccent.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.panelAccent.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelAccent.Location = new System.Drawing.Point(0, 82);
            this.panelAccent.Name = "panelAccent";
            this.panelAccent.Size = new System.Drawing.Size(386, 2);
            this.panelAccent.TabIndex = 1;
            // 
            // panelHeader
            // 
            this.panelHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(13)))), ((int)(((byte)(31)))));
            this.panelHeader.Controls.Add(this.buttonClose);
            this.panelHeader.Controls.Add(this.labelSubtitle);
            this.panelHeader.Controls.Add(this.labelTitle);
            this.panelHeader.Controls.Add(this.pictureBoxLogo);
            this.panelHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelHeader.Location = new System.Drawing.Point(0, 0);
            this.panelHeader.Name = "panelHeader";
            this.panelHeader.Size = new System.Drawing.Size(386, 82);
            this.panelHeader.TabIndex = 0;
            this.panelHeader.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // buttonClose
            // 
            this.buttonClose.FlatAppearance.BorderSize = 0;
            this.buttonClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonClose.ForeColor = System.Drawing.Color.White;
            this.buttonClose.Location = new System.Drawing.Point(358, 10);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(24, 22);
            this.buttonClose.TabIndex = 3;
            this.buttonClose.Text = "X";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelSubtitle
            // 
            this.labelSubtitle.AutoSize = true;
            this.labelSubtitle.BackColor = System.Drawing.Color.Transparent;
            this.labelSubtitle.Font = new System.Drawing.Font("Segoe UI", 8.7F);
            this.labelSubtitle.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(234)))), ((int)(((byte)(250)))));
            this.labelSubtitle.Location = new System.Drawing.Point(66, 46);
            this.labelSubtitle.Name = "labelSubtitle";
            this.labelSubtitle.Size = new System.Drawing.Size(101, 15);
            this.labelSubtitle.TabIndex = 2;
            this.labelSubtitle.Text = "Excel Match Setup";
            this.labelSubtitle.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.BackColor = System.Drawing.Color.Transparent;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.labelTitle.ForeColor = System.Drawing.Color.White;
            this.labelTitle.Location = new System.Drawing.Point(64, 20);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(104, 21);
            this.labelTitle.TabIndex = 1;
            this.labelTitle.Text = "Project Linker";
            this.labelTitle.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // pictureBoxLogo
            // 
            this.pictureBoxLogo.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxLogo.Location = new System.Drawing.Point(18, 16);
            this.pictureBoxLogo.Name = "pictureBoxLogo";
            this.pictureBoxLogo.Size = new System.Drawing.Size(36, 36);
            this.pictureBoxLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxLogo.TabIndex = 0;
            this.pictureBoxLogo.TabStop = false;
            this.pictureBoxLogo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Header_MouseDown);
            // 
            // AJProjectLinkerMatchConfigForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(146)))), ((int)(((byte)(231)))));
            this.ClientSize = new System.Drawing.Size(388, 372);
            this.Controls.Add(this.panelShell);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "AJProjectLinkerMatchConfigForm";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Project Linker Setup";
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
