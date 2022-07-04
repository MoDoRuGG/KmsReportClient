namespace KmsReportClient.Forms
{
    partial class CreateElementForm
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
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnAdd = new System.Windows.Forms.Button();
            this.TbxDesc = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TbxName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BtnCancel
            // 
            this.BtnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.BtnCancel.Location = new System.Drawing.Point(400, 132);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 29);
            this.BtnCancel.TabIndex = 11;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click_1);
            // 
            // BtnAdd
            // 
            this.BtnAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.BtnAdd.Location = new System.Drawing.Point(299, 132);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(95, 29);
            this.BtnAdd.TabIndex = 10;
            this.BtnAdd.Text = "Добавить";
            this.BtnAdd.UseVisualStyleBackColor = true;
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // TbxDesc
            // 
            this.TbxDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxDesc.Location = new System.Drawing.Point(96, 44);
            this.TbxDesc.Multiline = true;
            this.TbxDesc.Name = "TbxDesc";
            this.TbxDesc.Size = new System.Drawing.Size(379, 82);
            this.TbxDesc.TabIndex = 9;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label2.Location = new System.Drawing.Point(12, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(78, 17);
            this.label2.TabIndex = 8;
            this.label2.Text = "Описание:";
            // 
            // TbxName
            // 
            this.TbxName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxName.Location = new System.Drawing.Point(96, 15);
            this.TbxName.Name = "TbxName";
            this.TbxName.Size = new System.Drawing.Size(379, 23);
            this.TbxName.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label1.Location = new System.Drawing.Point(51, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 17);
            this.label1.TabIndex = 6;
            this.label1.Text = "Имя:";
            // 
            // CreateElementForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AntiqueWhite;
            this.ClientSize = new System.Drawing.Size(492, 178);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnAdd);
            this.Controls.Add(this.TbxDesc);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TbxName);
            this.Controls.Add(this.label1);
            this.Name = "CreateElementForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CreateElementForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnAdd;
        private System.Windows.Forms.TextBox TbxDesc;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TbxName;
        private System.Windows.Forms.Label label1;
    }
}