namespace KmsReportClient.Forms
{
    partial class CopyDataForm
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
            this.components = new System.ComponentModel.Container();
            this.BtnSave = new System.Windows.Forms.Button();
            this.CbxShow = new System.Windows.Forms.ComboBox();
            this.CbxPage = new System.Windows.Forms.ComboBox();
            this.BtnNextOne = new System.Windows.Forms.Button();
            this.BtnAllNext = new System.Windows.Forms.Button();
            this.BtnBackOne = new System.Windows.Forms.Button();
            this.BtnBackAll = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.TreeListNew = new BrightIdeasSoftware.TreeListView();
            this.TreeListOld = new BrightIdeasSoftware.TreeListView();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TreeListNew)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TreeListOld)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(449, 402);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(88, 33);
            this.BtnSave.TabIndex = 4;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // CbxShow
            // 
            this.CbxShow.FormattingEnabled = true;
            this.CbxShow.Location = new System.Drawing.Point(50, 12);
            this.CbxShow.Name = "CbxShow";
            this.CbxShow.Size = new System.Drawing.Size(409, 21);
            this.CbxShow.TabIndex = 8;
            this.CbxShow.SelectedIndexChanged += new System.EventHandler(this.CbxShow_SelectedIndexChanged);
            // 
            // CbxPage
            // 
            this.CbxPage.FormattingEnabled = true;
            this.CbxPage.Location = new System.Drawing.Point(38, 15);
            this.CbxPage.Name = "CbxPage";
            this.CbxPage.Size = new System.Drawing.Size(409, 21);
            this.CbxPage.TabIndex = 7;
            this.CbxPage.SelectedIndexChanged += new System.EventHandler(this.CbxPage_SelectedIndexChanged);
            // 
            // BtnNextOne
            // 
            this.BtnNextOne.Location = new System.Drawing.Point(453, 96);
            this.BtnNextOne.Name = "BtnNextOne";
            this.BtnNextOne.Size = new System.Drawing.Size(43, 32);
            this.BtnNextOne.TabIndex = 31;
            this.BtnNextOne.Text = ">";
            this.BtnNextOne.UseVisualStyleBackColor = true;
            this.BtnNextOne.Click += new System.EventHandler(this.BtnNextOne_Click);
            // 
            // BtnAllNext
            // 
            this.BtnAllNext.Location = new System.Drawing.Point(453, 134);
            this.BtnAllNext.Name = "BtnAllNext";
            this.BtnAllNext.Size = new System.Drawing.Size(43, 32);
            this.BtnAllNext.TabIndex = 32;
            this.BtnAllNext.Text = ">>";
            this.BtnAllNext.UseVisualStyleBackColor = true;
            this.BtnAllNext.Click += new System.EventHandler(this.BtnAllNext_Click);
            // 
            // BtnBackOne
            // 
            this.BtnBackOne.Location = new System.Drawing.Point(453, 172);
            this.BtnBackOne.Name = "BtnBackOne";
            this.BtnBackOne.Size = new System.Drawing.Size(43, 32);
            this.BtnBackOne.TabIndex = 33;
            this.BtnBackOne.Text = "<";
            this.BtnBackOne.UseVisualStyleBackColor = true;
            this.BtnBackOne.Click += new System.EventHandler(this.BtnBackOne_Click);
            // 
            // BtnBackAll
            // 
            this.BtnBackAll.Location = new System.Drawing.Point(453, 210);
            this.BtnBackAll.Name = "BtnBackAll";
            this.BtnBackAll.Size = new System.Drawing.Size(43, 32);
            this.BtnBackAll.TabIndex = 34;
            this.BtnBackAll.Text = "<<";
            this.BtnBackAll.UseVisualStyleBackColor = true;
            this.BtnBackAll.Click += new System.EventHandler(this.BtnBackAll_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.TreeListNew);
            this.groupBox1.Controls.Add(this.TreeListOld);
            this.groupBox1.Controls.Add(this.BtnBackAll);
            this.groupBox1.Controls.Add(this.BtnBackOne);
            this.groupBox1.Controls.Add(this.BtnAllNext);
            this.groupBox1.Controls.Add(this.CbxPage);
            this.groupBox1.Controls.Add(this.BtnNextOne);
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox1.Location = new System.Drawing.Point(12, 42);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(944, 353);
            this.groupBox1.TabIndex = 35;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // TreeListNew
            // 
            this.TreeListNew.BackColor = System.Drawing.Color.LightSteelBlue;
            this.TreeListNew.CellEditUseWholeCell = false;
            this.TreeListNew.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TreeListNew.FullRowSelect = true;
            this.TreeListNew.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.TreeListNew.HeaderUsesThemes = true;
            this.TreeListNew.HeaderWordWrap = true;
            this.TreeListNew.HideSelection = false;
            this.TreeListNew.Location = new System.Drawing.Point(502, 42);
            this.TreeListNew.Name = "TreeListNew";
            this.TreeListNew.OwnerDrawnHeader = true;
            this.TreeListNew.RevealAfterExpand = false;
            this.TreeListNew.SelectedBackColor = System.Drawing.Color.Teal;
            this.TreeListNew.SelectedForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.TreeListNew.ShowGroups = false;
            this.TreeListNew.ShowHeaderInAllViews = false;
            this.TreeListNew.ShowItemToolTips = true;
            this.TreeListNew.Size = new System.Drawing.Size(421, 280);
            this.TreeListNew.TabIndex = 36;
            this.TreeListNew.UseCompatibleStateImageBehavior = false;
            this.TreeListNew.View = System.Windows.Forms.View.Details;
            this.TreeListNew.VirtualMode = true;
            this.TreeListNew.FormatRow += new System.EventHandler<BrightIdeasSoftware.FormatRowEventArgs>(this.TreeListNew_FormatRow);
            // 
            // TreeListOld
            // 
            this.TreeListOld.BackColor = System.Drawing.Color.LightSteelBlue;
            this.TreeListOld.CellEditUseWholeCell = false;
            this.TreeListOld.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TreeListOld.FullRowSelect = true;
            this.TreeListOld.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.TreeListOld.HeaderUsesThemes = true;
            this.TreeListOld.HeaderWordWrap = true;
            this.TreeListOld.HideSelection = false;
            this.TreeListOld.Location = new System.Drawing.Point(38, 42);
            this.TreeListOld.Name = "TreeListOld";
            this.TreeListOld.OwnerDrawnHeader = true;
            this.TreeListOld.RevealAfterExpand = false;
            this.TreeListOld.SelectedBackColor = System.Drawing.Color.Teal;
            this.TreeListOld.SelectedForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.TreeListOld.ShowGroups = false;
            this.TreeListOld.ShowHeaderInAllViews = false;
            this.TreeListOld.ShowItemToolTips = true;
            this.TreeListOld.Size = new System.Drawing.Size(409, 280);
            this.TreeListOld.TabIndex = 35;
            this.TreeListOld.UseCompatibleStateImageBehavior = false;
            this.TreeListOld.View = System.Windows.Forms.View.Details;
            this.TreeListOld.VirtualMode = true;
            this.TreeListOld.FormatRow += new System.EventHandler<BrightIdeasSoftware.FormatRowEventArgs>(this.TreeListOld_FormatRow);
            // 
            // CopyDataForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(968, 447);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.CbxShow);
            this.Controls.Add(this.BtnSave);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "CopyDataForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Скопировать данные из других вкладок";
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TreeListNew)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TreeListOld)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.ComboBox CbxShow;
        private System.Windows.Forms.ComboBox CbxPage;
        private System.Windows.Forms.Button BtnNextOne;
        private System.Windows.Forms.Button BtnAllNext;
        private System.Windows.Forms.Button BtnBackOne;
        private System.Windows.Forms.Button BtnBackAll;
        private System.Windows.Forms.GroupBox groupBox1;
        private BrightIdeasSoftware.TreeListView TreeListNew;
        private BrightIdeasSoftware.TreeListView TreeListOld;
    }
}