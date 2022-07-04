namespace KmsReportClient.Forms
{
    partial class ConsolidateForm
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
            this.cmbStart = new System.Windows.Forms.ComboBox();
            this.nudStart = new System.Windows.Forms.NumericUpDown();
            this.labelStart = new System.Windows.Forms.Label();
            this.labelEnd = new System.Windows.Forms.Label();
            this.nudEnd = new System.Windows.Forms.NumericUpDown();
            this.cmbEnd = new System.Windows.Forms.ComboBox();
            this.btnDo = new System.Windows.Forms.Button();
            this.nudSingle = new System.Windows.Forms.NumericUpDown();
            this.panelSt = new System.Windows.Forms.Panel();
            this.panelEnd = new System.Windows.Forms.Panel();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.panelRegion = new System.Windows.Forms.Panel();
            this.cmbRegion = new System.Windows.Forms.ComboBox();
            this.lblRegion = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            ((System.ComponentModel.ISupportInitialize)(this.nudStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudEnd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudSingle)).BeginInit();
            this.panelSt.SuspendLayout();
            this.panelEnd.SuspendLayout();
            this.panelRegion.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbStart
            // 
            this.cmbStart.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cmbStart.FormattingEnabled = true;
            this.cmbStart.Location = new System.Drawing.Point(7, 4);
            this.cmbStart.Name = "cmbStart";
            this.cmbStart.Size = new System.Drawing.Size(148, 24);
            this.cmbStart.TabIndex = 0;
            // 
            // nudStart
            // 
            this.nudStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nudStart.Location = new System.Drawing.Point(197, 5);
            this.nudStart.Maximum = new decimal(new int[] {
            2100,
            0,
            0,
            0});
            this.nudStart.Minimum = new decimal(new int[] {
            2018,
            0,
            0,
            0});
            this.nudStart.Name = "nudStart";
            this.nudStart.Size = new System.Drawing.Size(64, 22);
            this.nudStart.TabIndex = 1;
            this.nudStart.Value = new decimal(new int[] {
            2019,
            0,
            0,
            0});
            // 
            // labelStart
            // 
            this.labelStart.AutoSize = true;
            this.labelStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelStart.Location = new System.Drawing.Point(12, 9);
            this.labelStart.Name = "labelStart";
            this.labelStart.Size = new System.Drawing.Size(109, 16);
            this.labelStart.TabIndex = 2;
            this.labelStart.Text = "Период начала";
            // 
            // labelEnd
            // 
            this.labelEnd.AutoSize = true;
            this.labelEnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelEnd.Location = new System.Drawing.Point(7, 9);
            this.labelEnd.Name = "labelEnd";
            this.labelEnd.Size = new System.Drawing.Size(131, 16);
            this.labelEnd.TabIndex = 3;
            this.labelEnd.Text = "Период окончания";
            // 
            // nudEnd
            // 
            this.nudEnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nudEnd.Location = new System.Drawing.Point(197, 28);
            this.nudEnd.Maximum = new decimal(new int[] {
            2100,
            0,
            0,
            0});
            this.nudEnd.Minimum = new decimal(new int[] {
            2018,
            0,
            0,
            0});
            this.nudEnd.Name = "nudEnd";
            this.nudEnd.Size = new System.Drawing.Size(64, 22);
            this.nudEnd.TabIndex = 5;
            this.nudEnd.Value = new decimal(new int[] {
            2019,
            0,
            0,
            0});
            // 
            // cmbEnd
            // 
            this.cmbEnd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbEnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cmbEnd.FormattingEnabled = true;
            this.cmbEnd.Location = new System.Drawing.Point(7, 27);
            this.cmbEnd.Name = "cmbEnd";
            this.cmbEnd.Size = new System.Drawing.Size(148, 24);
            this.cmbEnd.TabIndex = 4;
            // 
            // btnDo
            // 
            this.btnDo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnDo.Location = new System.Drawing.Point(12, 209);
            this.btnDo.Name = "btnDo";
            this.btnDo.Size = new System.Drawing.Size(254, 49);
            this.btnDo.TabIndex = 6;
            this.btnDo.Text = "button1";
            this.btnDo.UseVisualStyleBackColor = true;
            this.btnDo.Click += new System.EventHandler(this.BtnDo_Click);
            // 
            // nudSingle
            // 
            this.nudSingle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.nudSingle.Location = new System.Drawing.Point(127, 7);
            this.nudSingle.Maximum = new decimal(new int[] {
            2100,
            0,
            0,
            0});
            this.nudSingle.Minimum = new decimal(new int[] {
            2018,
            0,
            0,
            0});
            this.nudSingle.Name = "nudSingle";
            this.nudSingle.Size = new System.Drawing.Size(64, 22);
            this.nudSingle.TabIndex = 7;
            this.nudSingle.Value = new decimal(new int[] {
            2019,
            0,
            0,
            0});
            // 
            // panelSt
            // 
            this.panelSt.Controls.Add(this.cmbStart);
            this.panelSt.Controls.Add(this.nudStart);
            this.panelSt.Location = new System.Drawing.Point(5, 35);
            this.panelSt.Name = "panelSt";
            this.panelSt.Size = new System.Drawing.Size(275, 37);
            this.panelSt.TabIndex = 8;
            // 
            // panelEnd
            // 
            this.panelEnd.Controls.Add(this.labelEnd);
            this.panelEnd.Controls.Add(this.cmbEnd);
            this.panelEnd.Controls.Add(this.nudEnd);
            this.panelEnd.Location = new System.Drawing.Point(5, 78);
            this.panelEnd.Name = "panelEnd";
            this.panelEnd.Size = new System.Drawing.Size(275, 62);
            this.panelEnd.TabIndex = 9;
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "Excel|.xlsx";
            // 
            // panelRegion
            // 
            this.panelRegion.Controls.Add(this.cmbRegion);
            this.panelRegion.Controls.Add(this.lblRegion);
            this.panelRegion.Location = new System.Drawing.Point(5, 146);
            this.panelRegion.Name = "panelRegion";
            this.panelRegion.Size = new System.Drawing.Size(275, 57);
            this.panelRegion.TabIndex = 10;
            // 
            // cmbRegion
            // 
            this.cmbRegion.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbRegion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cmbRegion.FormattingEnabled = true;
            this.cmbRegion.Location = new System.Drawing.Point(7, 28);
            this.cmbRegion.Name = "cmbRegion";
            this.cmbRegion.Size = new System.Drawing.Size(254, 24);
            this.cmbRegion.TabIndex = 4;
            // 
            // lblRegion
            // 
            this.lblRegion.AutoSize = true;
            this.lblRegion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblRegion.Location = new System.Drawing.Point(12, 9);
            this.lblRegion.Name = "lblRegion";
            this.lblRegion.Size = new System.Drawing.Size(55, 16);
            this.lblRegion.TabIndex = 3;
            this.lblRegion.Text = "Регион";
            // 
            // ConsolidateForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 266);
            this.Controls.Add(this.panelRegion);
            this.Controls.Add(this.panelEnd);
            this.Controls.Add(this.panelSt);
            this.Controls.Add(this.nudSingle);
            this.Controls.Add(this.btnDo);
            this.Controls.Add(this.labelStart);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ConsolidateForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Формирование сводных отчетов";
            this.Load += new System.EventHandler(this.ConsolidateForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.nudStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudEnd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudSingle)).EndInit();
            this.panelSt.ResumeLayout(false);
            this.panelEnd.ResumeLayout(false);
            this.panelEnd.PerformLayout();
            this.panelRegion.ResumeLayout(false);
            this.panelRegion.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbStart;
        private System.Windows.Forms.NumericUpDown nudStart;
        private System.Windows.Forms.Label labelStart;
        private System.Windows.Forms.Label labelEnd;
        private System.Windows.Forms.NumericUpDown nudEnd;
        private System.Windows.Forms.ComboBox cmbEnd;
        private System.Windows.Forms.Button btnDo;
        private System.Windows.Forms.NumericUpDown nudSingle;
        private System.Windows.Forms.Panel panelSt;
        private System.Windows.Forms.Panel panelEnd;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Panel panelRegion;
        private System.Windows.Forms.ComboBox cmbRegion;
        private System.Windows.Forms.Label lblRegion;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}