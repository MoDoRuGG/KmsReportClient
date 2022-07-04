namespace KmsReportClient.Forms
{
    partial class NotificationForm
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
            this.CmbReport = new System.Windows.Forms.ComboBox();
            this.CmbMonth = new System.Windows.Forms.ComboBox();
            this.NudYear = new System.Windows.Forms.NumericUpDown();
            this.RbtnLater = new System.Windows.Forms.RadioButton();
            this.RbtnRefuse = new System.Windows.Forms.RadioButton();
            this.RbtnSelect = new System.Windows.Forms.RadioButton();
            this.panel = new System.Windows.Forms.Panel();
            this.TxtbListFilial = new System.Windows.Forms.TextBox();
            this.BtnMinus = new System.Windows.Forms.Button();
            this.BtnPlus = new System.Windows.Forms.Button();
            this.CmbFilials = new System.Windows.Forms.ComboBox();
            this.TxtbText = new System.Windows.Forms.TextBox();
            this.BtnSaveText = new System.Windows.Forms.Button();
            this.BtnSend = new System.Windows.Forms.Button();
            this.BtnClear = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.txtbTheme = new System.Windows.Forms.TextBox();
            this.chkbDefault = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.NudYear)).BeginInit();
            this.panel.SuspendLayout();
            this.SuspendLayout();
            // 
            // CmbReport
            // 
            this.CmbReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CmbReport.FormattingEnabled = true;
            this.CmbReport.Location = new System.Drawing.Point(12, 12);
            this.CmbReport.Name = "CmbReport";
            this.CmbReport.Size = new System.Drawing.Size(185, 24);
            this.CmbReport.TabIndex = 0;
            this.CmbReport.SelectedIndexChanged += new System.EventHandler(this.CmbReport_SelectedIndexChanged);
            // 
            // CmbMonth
            // 
            this.CmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbMonth.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CmbMonth.FormattingEnabled = true;
            this.CmbMonth.Location = new System.Drawing.Point(213, 12);
            this.CmbMonth.Name = "CmbMonth";
            this.CmbMonth.Size = new System.Drawing.Size(148, 24);
            this.CmbMonth.TabIndex = 1;
            // 
            // NudYear
            // 
            this.NudYear.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.NudYear.Location = new System.Drawing.Point(377, 13);
            this.NudYear.Maximum = new decimal(new int[] {
            2100,
            0,
            0,
            0});
            this.NudYear.Minimum = new decimal(new int[] {
            2010,
            0,
            0,
            0});
            this.NudYear.Name = "NudYear";
            this.NudYear.Size = new System.Drawing.Size(77, 23);
            this.NudYear.TabIndex = 2;
            this.NudYear.Value = new decimal(new int[] {
            2010,
            0,
            0,
            0});
            // 
            // RbtnLater
            // 
            this.RbtnLater.AutoSize = true;
            this.RbtnLater.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.RbtnLater.Location = new System.Drawing.Point(12, 42);
            this.RbtnLater.Name = "RbtnLater";
            this.RbtnLater.Size = new System.Drawing.Size(129, 21);
            this.RbtnLater.TabIndex = 3;
            this.RbtnLater.TabStop = true;
            this.RbtnLater.Text = "Не сдали отчет";
            this.RbtnLater.UseVisualStyleBackColor = true;
            this.RbtnLater.CheckedChanged += new System.EventHandler(this.RbtnLater_CheckedChanged);
            // 
            // RbtnRefuse
            // 
            this.RbtnRefuse.AutoSize = true;
            this.RbtnRefuse.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.RbtnRefuse.Location = new System.Drawing.Point(147, 42);
            this.RbtnRefuse.Name = "RbtnRefuse";
            this.RbtnRefuse.Size = new System.Drawing.Size(161, 21);
            this.RbtnRefuse.TabIndex = 4;
            this.RbtnRefuse.TabStop = true;
            this.RbtnRefuse.Text = "Отчет на доработке";
            this.RbtnRefuse.UseVisualStyleBackColor = true;
            this.RbtnRefuse.CheckedChanged += new System.EventHandler(this.RbtnRefuse_CheckedChanged);
            // 
            // RbtnSelect
            // 
            this.RbtnSelect.AutoSize = true;
            this.RbtnSelect.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.RbtnSelect.Location = new System.Drawing.Point(314, 42);
            this.RbtnSelect.Name = "RbtnSelect";
            this.RbtnSelect.Size = new System.Drawing.Size(151, 21);
            this.RbtnSelect.TabIndex = 5;
            this.RbtnSelect.TabStop = true;
            this.RbtnSelect.Text = "Выбрать из списка";
            this.RbtnSelect.UseVisualStyleBackColor = true;
            this.RbtnSelect.CheckedChanged += new System.EventHandler(this.RbtnSelect_CheckedChanged);
            // 
            // panel
            // 
            this.panel.Controls.Add(this.TxtbListFilial);
            this.panel.Controls.Add(this.BtnMinus);
            this.panel.Controls.Add(this.BtnPlus);
            this.panel.Controls.Add(this.CmbFilials);
            this.panel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.panel.Location = new System.Drawing.Point(4, 69);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(461, 111);
            this.panel.TabIndex = 6;
            // 
            // TxtbListFilial
            // 
            this.TxtbListFilial.Location = new System.Drawing.Point(8, 33);
            this.TxtbListFilial.Multiline = true;
            this.TxtbListFilial.Name = "TxtbListFilial";
            this.TxtbListFilial.ReadOnly = true;
            this.TxtbListFilial.Size = new System.Drawing.Size(442, 73);
            this.TxtbListFilial.TabIndex = 3;
            // 
            // BtnMinus
            // 
            this.BtnMinus.Location = new System.Drawing.Point(427, 4);
            this.BtnMinus.Name = "BtnMinus";
            this.BtnMinus.Size = new System.Drawing.Size(23, 23);
            this.BtnMinus.TabIndex = 2;
            this.BtnMinus.Text = "-";
            this.BtnMinus.UseVisualStyleBackColor = true;
            this.BtnMinus.Click += new System.EventHandler(this.BtnMinus_Click);
            // 
            // BtnPlus
            // 
            this.BtnPlus.Location = new System.Drawing.Point(398, 4);
            this.BtnPlus.Name = "BtnPlus";
            this.BtnPlus.Size = new System.Drawing.Size(23, 23);
            this.BtnPlus.TabIndex = 1;
            this.BtnPlus.Text = "+";
            this.BtnPlus.UseVisualStyleBackColor = true;
            this.BtnPlus.Click += new System.EventHandler(this.BtnPlus_Click);
            // 
            // CmbFilials
            // 
            this.CmbFilials.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbFilials.FormattingEnabled = true;
            this.CmbFilials.Location = new System.Drawing.Point(8, 4);
            this.CmbFilials.Name = "CmbFilials";
            this.CmbFilials.Size = new System.Drawing.Size(385, 24);
            this.CmbFilials.TabIndex = 0;
            // 
            // TxtbText
            // 
            this.TxtbText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TxtbText.Location = new System.Drawing.Point(10, 256);
            this.TxtbText.Multiline = true;
            this.TxtbText.Name = "TxtbText";
            this.TxtbText.Size = new System.Drawing.Size(442, 111);
            this.TxtbText.TabIndex = 7;
            // 
            // BtnSaveText
            // 
            this.BtnSaveText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnSaveText.Location = new System.Drawing.Point(10, 373);
            this.BtnSaveText.Name = "BtnSaveText";
            this.BtnSaveText.Size = new System.Drawing.Size(140, 43);
            this.BtnSaveText.TabIndex = 8;
            this.BtnSaveText.Text = "Сохранить текст шаблона";
            this.BtnSaveText.UseVisualStyleBackColor = true;
            this.BtnSaveText.Click += new System.EventHandler(this.BtnSaveText_Click);
            // 
            // BtnSend
            // 
            this.BtnSend.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnSend.Location = new System.Drawing.Point(162, 373);
            this.BtnSend.Name = "BtnSend";
            this.BtnSend.Size = new System.Drawing.Size(140, 43);
            this.BtnSend.TabIndex = 9;
            this.BtnSend.Text = "Отправить";
            this.BtnSend.UseVisualStyleBackColor = true;
            this.BtnSend.Click += new System.EventHandler(this.BtnSend_Click);
            // 
            // BtnClear
            // 
            this.BtnClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnClear.Location = new System.Drawing.Point(312, 373);
            this.BtnClear.Name = "BtnClear";
            this.BtnClear.Size = new System.Drawing.Size(140, 43);
            this.BtnClear.TabIndex = 10;
            this.BtnClear.Text = "Очистить";
            this.BtnClear.UseVisualStyleBackColor = true;
            this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(10, 237);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 16);
            this.label1.TabIndex = 11;
            this.label1.Text = "Текст письма";
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(0, 423);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(469, 10);
            this.progressBar1.TabIndex = 12;
            // 
            // txtbTheme
            // 
            this.txtbTheme.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbTheme.Location = new System.Drawing.Point(12, 211);
            this.txtbTheme.Name = "txtbTheme";
            this.txtbTheme.Size = new System.Drawing.Size(442, 23);
            this.txtbTheme.TabIndex = 13;
            // 
            // chkbDefault
            // 
            this.chkbDefault.AutoSize = true;
            this.chkbDefault.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chkbDefault.Location = new System.Drawing.Point(13, 188);
            this.chkbDefault.Name = "chkbDefault";
            this.chkbDefault.Size = new System.Drawing.Size(159, 21);
            this.chkbDefault.TabIndex = 14;
            this.chkbDefault.Text = "Тема по умолчанию";
            this.chkbDefault.UseVisualStyleBackColor = true;
            this.chkbDefault.CheckedChanged += new System.EventHandler(this.ChkbDefault_CheckedChanged);
            // 
            // NotificationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(469, 433);
            this.Controls.Add(this.chkbDefault);
            this.Controls.Add(this.txtbTheme);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.BtnClear);
            this.Controls.Add(this.BtnSend);
            this.Controls.Add(this.BtnSaveText);
            this.Controls.Add(this.TxtbText);
            this.Controls.Add(this.panel);
            this.Controls.Add(this.RbtnSelect);
            this.Controls.Add(this.RbtnRefuse);
            this.Controls.Add(this.RbtnLater);
            this.Controls.Add(this.NudYear);
            this.Controls.Add(this.CmbMonth);
            this.Controls.Add(this.CmbReport);
            this.Name = "NotificationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Форма рассылки уведоимлений";
            ((System.ComponentModel.ISupportInitialize)(this.NudYear)).EndInit();
            this.panel.ResumeLayout(false);
            this.panel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox CmbReport;
        private System.Windows.Forms.ComboBox CmbMonth;
        private System.Windows.Forms.NumericUpDown NudYear;
        private System.Windows.Forms.RadioButton RbtnLater;
        private System.Windows.Forms.RadioButton RbtnRefuse;
        private System.Windows.Forms.RadioButton RbtnSelect;
        private System.Windows.Forms.Panel panel;
        private System.Windows.Forms.TextBox TxtbListFilial;
        private System.Windows.Forms.Button BtnMinus;
        private System.Windows.Forms.Button BtnPlus;
        private System.Windows.Forms.ComboBox CmbFilials;
        private System.Windows.Forms.TextBox TxtbText;
        private System.Windows.Forms.Button BtnSaveText;
        private System.Windows.Forms.Button BtnSend;
        private System.Windows.Forms.Button BtnClear;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox txtbTheme;
        private System.Windows.Forms.CheckBox chkbDefault;
    }
}