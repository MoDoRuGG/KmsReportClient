namespace KmsReportClient.Forms
{
    partial class AuthorizationForm
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
            if (disposing && (components != null)) {
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
            this.CmbLogin = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.BtnEnter = new System.Windows.Forms.Button();
            this.BtnExit = new System.Windows.Forms.Button();
            this.TxtbPassword = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.CmbFilial = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // CmbLogin
            // 
            this.CmbLogin.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.CmbLogin.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CmbLogin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CmbLogin.FormattingEnabled = true;
            this.CmbLogin.Location = new System.Drawing.Point(12, 76);
            this.CmbLogin.Name = "CmbLogin";
            this.CmbLogin.Size = new System.Drawing.Size(289, 24);
            this.CmbLogin.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(96, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "Имя пользователя";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(131, 103);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "Пароль";
            // 
            // BtnEnter
            // 
            this.BtnEnter.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnEnter.Location = new System.Drawing.Point(12, 149);
            this.BtnEnter.Name = "BtnEnter";
            this.BtnEnter.Size = new System.Drawing.Size(117, 29);
            this.BtnEnter.TabIndex = 4;
            this.BtnEnter.Text = "Вход <Enter>";
            this.BtnEnter.UseVisualStyleBackColor = true;
            this.BtnEnter.Click += new System.EventHandler(this.Btn_enter_Click);
            // 
            // BtnExit
            // 
            this.BtnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnExit.Location = new System.Drawing.Point(184, 149);
            this.BtnExit.Name = "BtnExit";
            this.BtnExit.Size = new System.Drawing.Size(117, 29);
            this.BtnExit.TabIndex = 5;
            this.BtnExit.Text = "Выход <Esc>";
            this.BtnExit.UseVisualStyleBackColor = true;
            this.BtnExit.Click += new System.EventHandler(this.Btn_exit_Click);
            // 
            // TxtbPassword
            // 
            this.TxtbPassword.Location = new System.Drawing.Point(12, 123);
            this.TxtbPassword.Name = "TxtbPassword";
            this.TxtbPassword.PasswordChar = '*';
            this.TxtbPassword.Size = new System.Drawing.Size(289, 20);
            this.TxtbPassword.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(131, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Регион";
            // 
            // CmbFilial
            // 
            this.CmbFilial.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.CmbFilial.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CmbFilial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbFilial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.CmbFilial.FormattingEnabled = true;
            this.CmbFilial.Location = new System.Drawing.Point(12, 29);
            this.CmbFilial.Name = "CmbFilial";
            this.CmbFilial.Size = new System.Drawing.Size(289, 24);
            this.CmbFilial.TabIndex = 7;
            this.CmbFilial.SelectedIndexChanged += new System.EventHandler(this.CmbFilial_SelectedIndexChanged);
            // 
            // AuthorizationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(313, 190);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.CmbFilial);
            this.Controls.Add(this.TxtbPassword);
            this.Controls.Add(this.BtnExit);
            this.Controls.Add(this.BtnEnter);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CmbLogin);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.KeyPreview = true;
            this.Name = "AuthorizationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Вход в электронный журнал";
            this.Load += new System.EventHandler(this.Autorization_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Autorization_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox CmbLogin;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button BtnEnter;
        private System.Windows.Forms.Button BtnExit;
        private System.Windows.Forms.TextBox TxtbPassword;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox CmbFilial;
    }
}