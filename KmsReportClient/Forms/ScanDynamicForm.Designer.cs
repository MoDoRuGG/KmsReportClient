
namespace KmsReportClient.Forms
{
    partial class ScanDynamicForm
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
            this.BtnDelete = new System.Windows.Forms.Button();
            this.BtnLoad = new System.Windows.Forms.Button();
            this.BtnOpen = new System.Windows.Forms.Button();
            this.LbScan = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // BtnDelete
            // 
            this.BtnDelete.Location = new System.Drawing.Point(155, 168);
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.Size = new System.Drawing.Size(125, 23);
            this.BtnDelete.TabIndex = 7;
            this.BtnDelete.Text = "Удалить";
            this.BtnDelete.UseVisualStyleBackColor = true;
            this.BtnDelete.Click += new System.EventHandler(this.BtnDelete_Click);
            // 
            // BtnLoad
            // 
            this.BtnLoad.Location = new System.Drawing.Point(286, 168);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(144, 23);
            this.BtnLoad.TabIndex = 6;
            this.BtnLoad.Text = "Загрузить";
            this.BtnLoad.UseVisualStyleBackColor = true;
            this.BtnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // BtnOpen
            // 
            this.BtnOpen.Location = new System.Drawing.Point(2, 168);
            this.BtnOpen.Name = "BtnOpen";
            this.BtnOpen.Size = new System.Drawing.Size(147, 23);
            this.BtnOpen.TabIndex = 5;
            this.BtnOpen.Text = "Открыть";
            this.BtnOpen.UseVisualStyleBackColor = true;
            this.BtnOpen.Click += new System.EventHandler(this.BtnOpen_Click);
            // 
            // LbScan
            // 
            this.LbScan.FormattingEnabled = true;
            this.LbScan.Location = new System.Drawing.Point(2, 2);
            this.LbScan.Name = "LbScan";
            this.LbScan.Size = new System.Drawing.Size(640, 160);
            this.LbScan.TabIndex = 4;
            // 
            // ScanDynamicForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(644, 226);
            this.Controls.Add(this.BtnDelete);
            this.Controls.Add(this.BtnLoad);
            this.Controls.Add(this.BtnOpen);
            this.Controls.Add(this.LbScan);
            this.Name = "ScanDynamicForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Сканы отчёта";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnDelete;
        private System.Windows.Forms.Button BtnLoad;
        private System.Windows.Forms.Button BtnOpen;
        private System.Windows.Forms.ListBox LbScan;
    }
}