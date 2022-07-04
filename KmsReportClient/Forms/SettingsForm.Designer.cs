namespace KmsReportClient.Forms
{
    partial class SettingsForm
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
            this.treeSettings = new System.Windows.Forms.TreeView();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabFilial = new System.Windows.Forms.TabPage();
            this.panelFilial = new System.Windows.Forms.Panel();
            this.txtbFilial = new System.Windows.Forms.TextBox();
            this.txtbPhoneHead = new System.Windows.Forms.MaskedTextBox();
            this.txtbFioHead = new System.Windows.Forms.TextBox();
            this.txtbPlaceHead = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabUser = new System.Windows.Forms.TabPage();
            this.btnAdd = new System.Windows.Forms.Button();
            this.panelUser = new System.Windows.Forms.Panel();
            this.txtbFio = new System.Windows.Forms.TextBox();
            this.txtbPhone = new System.Windows.Forms.MaskedTextBox();
            this.txtbEmail = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.BtnSave = new System.Windows.Forms.Button();
            this.BtnExit = new System.Windows.Forms.Button();
            this.btnEditFilial = new System.Windows.Forms.Button();
            this.btnEditUser = new System.Windows.Forms.Button();
            this.tabControl.SuspendLayout();
            this.tabFilial.SuspendLayout();
            this.panelFilial.SuspendLayout();
            this.tabUser.SuspendLayout();
            this.panelUser.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeSettings
            // 
            this.treeSettings.BackColor = System.Drawing.Color.LightSteelBlue;
            this.treeSettings.Dock = System.Windows.Forms.DockStyle.Left;
            this.treeSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.treeSettings.Location = new System.Drawing.Point(0, 0);
            this.treeSettings.Name = "treeSettings";
            this.treeSettings.Size = new System.Drawing.Size(186, 401);
            this.treeSettings.TabIndex = 0;
            this.treeSettings.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.TreeSettings_AfterSelect);
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabFilial);
            this.tabControl.Controls.Add(this.tabUser);
            this.tabControl.Location = new System.Drawing.Point(192, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(618, 346);
            this.tabControl.TabIndex = 1;
            // 
            // tabFilial
            // 
            this.tabFilial.BackColor = System.Drawing.Color.AntiqueWhite;
            this.tabFilial.Controls.Add(this.btnEditFilial);
            this.tabFilial.Controls.Add(this.panelFilial);
            this.tabFilial.Controls.Add(this.label4);
            this.tabFilial.Controls.Add(this.label3);
            this.tabFilial.Controls.Add(this.label2);
            this.tabFilial.Controls.Add(this.label1);
            this.tabFilial.Location = new System.Drawing.Point(4, 22);
            this.tabFilial.Name = "tabFilial";
            this.tabFilial.Padding = new System.Windows.Forms.Padding(3);
            this.tabFilial.Size = new System.Drawing.Size(610, 320);
            this.tabFilial.TabIndex = 0;
            this.tabFilial.Text = "Филиал";
            // 
            // panelFilial
            // 
            this.panelFilial.Controls.Add(this.txtbFilial);
            this.panelFilial.Controls.Add(this.txtbPhoneHead);
            this.panelFilial.Controls.Add(this.txtbFioHead);
            this.panelFilial.Controls.Add(this.txtbPlaceHead);
            this.panelFilial.Location = new System.Drawing.Point(211, 6);
            this.panelFilial.Name = "panelFilial";
            this.panelFilial.Size = new System.Drawing.Size(393, 136);
            this.panelFilial.TabIndex = 16;
            // 
            // txtbFilial
            // 
            this.txtbFilial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbFilial.Location = new System.Drawing.Point(3, 11);
            this.txtbFilial.Name = "txtbFilial";
            this.txtbFilial.Size = new System.Drawing.Size(379, 23);
            this.txtbFilial.TabIndex = 1;
            // 
            // txtbPhoneHead
            // 
            this.txtbPhoneHead.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbPhoneHead.Location = new System.Drawing.Point(3, 98);
            this.txtbPhoneHead.Mask = "(000) 000-00-00";
            this.txtbPhoneHead.Name = "txtbPhoneHead";
            this.txtbPhoneHead.Size = new System.Drawing.Size(379, 23);
            this.txtbPhoneHead.TabIndex = 15;
            // 
            // txtbFioHead
            // 
            this.txtbFioHead.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbFioHead.Location = new System.Drawing.Point(3, 40);
            this.txtbFioHead.Name = "txtbFioHead";
            this.txtbFioHead.Size = new System.Drawing.Size(379, 23);
            this.txtbFioHead.TabIndex = 3;
            // 
            // txtbPlaceHead
            // 
            this.txtbPlaceHead.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbPlaceHead.Location = new System.Drawing.Point(3, 69);
            this.txtbPlaceHead.Name = "txtbPlaceHead";
            this.txtbPlaceHead.Size = new System.Drawing.Size(379, 23);
            this.txtbPlaceHead.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(33, 107);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(164, 17);
            this.label4.TabIndex = 6;
            this.label4.Text = "Телефон руководителя";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(20, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(177, 17);
            this.label3.TabIndex = 4;
            this.label3.Text = "Должность руководителя";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(59, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(138, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "ФИО руководителя";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(28, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(169, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Наименование филиала";
            // 
            // tabUser
            // 
            this.tabUser.BackColor = System.Drawing.Color.AntiqueWhite;
            this.tabUser.Controls.Add(this.btnEditUser);
            this.tabUser.Controls.Add(this.btnAdd);
            this.tabUser.Controls.Add(this.panelUser);
            this.tabUser.Controls.Add(this.label6);
            this.tabUser.Controls.Add(this.label7);
            this.tabUser.Controls.Add(this.label8);
            this.tabUser.Location = new System.Drawing.Point(4, 22);
            this.tabUser.Name = "tabUser";
            this.tabUser.Padding = new System.Windows.Forms.Padding(3);
            this.tabUser.Size = new System.Drawing.Size(610, 320);
            this.tabUser.TabIndex = 1;
            this.tabUser.Text = "Пользователь";
            // 
            // btnAdd
            // 
            this.btnAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnAdd.Location = new System.Drawing.Point(366, 256);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(238, 26);
            this.btnAdd.TabIndex = 4;
            this.btnAdd.Text = "Добавить нового пользователя";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // panelUser
            // 
            this.panelUser.Controls.Add(this.txtbFio);
            this.panelUser.Controls.Add(this.txtbPhone);
            this.panelUser.Controls.Add(this.txtbEmail);
            this.panelUser.Location = new System.Drawing.Point(80, 6);
            this.panelUser.Name = "panelUser";
            this.panelUser.Size = new System.Drawing.Size(448, 236);
            this.panelUser.TabIndex = 15;
            // 
            // txtbFio
            // 
            this.txtbFio.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbFio.Location = new System.Drawing.Point(3, 9);
            this.txtbFio.Name = "txtbFio";
            this.txtbFio.Size = new System.Drawing.Size(379, 23);
            this.txtbFio.TabIndex = 9;
            // 
            // txtbPhone
            // 
            this.txtbPhone.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbPhone.Location = new System.Drawing.Point(3, 38);
            this.txtbPhone.Mask = "(000) 000-00-00";
            this.txtbPhone.Name = "txtbPhone";
            this.txtbPhone.Size = new System.Drawing.Size(379, 23);
            this.txtbPhone.TabIndex = 14;
            // 
            // txtbEmail
            // 
            this.txtbEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.txtbEmail.Location = new System.Drawing.Point(3, 67);
            this.txtbEmail.Name = "txtbEmail";
            this.txtbEmail.Size = new System.Drawing.Size(379, 23);
            this.txtbEmail.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(27, 76);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(47, 17);
            this.label6.TabIndex = 12;
            this.label6.Text = "E-mail";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.Location = new System.Drawing.Point(9, 47);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 17);
            this.label7.TabIndex = 10;
            this.label7.Text = "Телефон";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label8.Location = new System.Drawing.Point(7, 18);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(70, 17);
            this.label8.TabIndex = 8;
            this.label8.Text = "Фамилия";
            // 
            // BtnSave
            // 
            this.BtnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnSave.Location = new System.Drawing.Point(526, 352);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(138, 43);
            this.BtnSave.TabIndex = 2;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // BtnExit
            // 
            this.BtnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.BtnExit.Location = new System.Drawing.Point(670, 352);
            this.BtnExit.Name = "BtnExit";
            this.BtnExit.Size = new System.Drawing.Size(130, 43);
            this.BtnExit.TabIndex = 3;
            this.BtnExit.Text = "Закрыть";
            this.BtnExit.UseVisualStyleBackColor = true;
            this.BtnExit.Click += new System.EventHandler(this.BtnExit_Click);
            // 
            // btnEditFilial
            // 
            this.btnEditFilial.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnEditFilial.Location = new System.Drawing.Point(330, 287);
            this.btnEditFilial.Name = "btnEditFilial";
            this.btnEditFilial.Size = new System.Drawing.Size(274, 27);
            this.btnEditFilial.TabIndex = 17;
            this.btnEditFilial.Text = "Редактировать данные о филиале";
            this.btnEditFilial.UseVisualStyleBackColor = true;
            this.btnEditFilial.Click += new System.EventHandler(this.btnEditFilial_Click);
            // 
            // btnEditUser
            // 
            this.btnEditUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnEditUser.Location = new System.Drawing.Point(366, 288);
            this.btnEditUser.Name = "btnEditUser";
            this.btnEditUser.Size = new System.Drawing.Size(238, 26);
            this.btnEditUser.TabIndex = 16;
            this.btnEditUser.Text = "Редактировать пользователя";
            this.btnEditUser.UseVisualStyleBackColor = true;
            this.btnEditUser.Click += new System.EventHandler(this.btnEditUser_Click);
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.ClientSize = new System.Drawing.Size(812, 401);
            this.Controls.Add(this.BtnExit);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.treeSettings);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.tabControl.ResumeLayout(false);
            this.tabFilial.ResumeLayout(false);
            this.tabFilial.PerformLayout();
            this.panelFilial.ResumeLayout(false);
            this.panelFilial.PerformLayout();
            this.tabUser.ResumeLayout(false);
            this.tabUser.PerformLayout();
            this.panelUser.ResumeLayout(false);
            this.panelUser.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeSettings;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabFilial;
        private System.Windows.Forms.TabPage tabUser;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.Button BtnExit;
        private System.Windows.Forms.MaskedTextBox txtbPhoneHead;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtbPlaceHead;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtbFioHead;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtbFilial;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.MaskedTextBox txtbPhone;
        private System.Windows.Forms.TextBox txtbEmail;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtbFio;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Panel panelFilial;
        private System.Windows.Forms.Panel panelUser;
        private System.Windows.Forms.Button btnEditFilial;
        private System.Windows.Forms.Button btnEditUser;
    }
}