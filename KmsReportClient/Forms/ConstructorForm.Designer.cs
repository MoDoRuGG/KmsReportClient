namespace KmsReportClient.Forms
{
    partial class ConstuctorForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConstuctorForm));
            this.CbxShow = new System.Windows.Forms.ComboBox();
            this.Menu = new System.Windows.Forms.ToolStrip();
            this.BtnCreatePage = new System.Windows.Forms.ToolStripButton();
            this.BtnCopyPage = new System.Windows.Forms.ToolStripButton();
            this.BtnAdd = new System.Windows.Forms.ToolStripButton();
            this.Группа = new System.Windows.Forms.ToolStripButton();
            this.BtnSaveReport = new System.Windows.Forms.ToolStripButton();
            this.TbxTabDesc = new System.Windows.Forms.TextBox();
            this.CbxPage = new System.Windows.Forms.ComboBox();
            this.btnAddContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.столбецToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.группаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.BtnDelete = new System.Windows.Forms.Button();
            this.treeListView1 = new BrightIdeasSoftware.TreeListView();
            this.treeViewContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.снятьВыделениеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.BtnEditElement = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.TbxDescElement = new System.Windows.Forms.TextBox();
            this.TbxNameElement = new System.Windows.Forms.TextBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.LbEmail = new System.Windows.Forms.CheckedListBox();
            this.CbxUserRow = new System.Windows.Forms.CheckBox();
            this.TbxDescReport = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.TbxName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.DtmDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.Menu.SuspendLayout();
            this.btnAddContextMenuStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.treeListView1)).BeginInit();
            this.treeViewContextMenuStrip.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // CbxShow
            // 
            this.CbxShow.BackColor = System.Drawing.SystemColors.Control;
            this.CbxShow.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbxShow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.CbxShow.FormattingEnabled = true;
            this.CbxShow.Location = new System.Drawing.Point(4, 87);
            this.CbxShow.Name = "CbxShow";
            this.CbxShow.Size = new System.Drawing.Size(594, 24);
            this.CbxShow.TabIndex = 18;
            this.CbxShow.SelectedIndexChanged += new System.EventHandler(this.CbxShow_SelectedIndexChanged);
            // 
            // Menu
            // 
            this.Menu.BackColor = System.Drawing.SystemColors.Control;
            this.Menu.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.Menu.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.Menu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.BtnCreatePage,
            this.BtnCopyPage,
            this.BtnAdd,
            this.Группа,
            this.BtnSaveReport});
            this.Menu.Location = new System.Drawing.Point(0, 0);
            this.Menu.Name = "Menu";
            this.Menu.Size = new System.Drawing.Size(1085, 31);
            this.Menu.TabIndex = 29;
            this.Menu.Text = " ";
            // 
            // BtnCreatePage
            // 
            this.BtnCreatePage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.BtnCreatePage.Image = ((System.Drawing.Image)(resources.GetObject("BtnCreatePage.Image")));
            this.BtnCreatePage.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BtnCreatePage.Name = "BtnCreatePage";
            this.BtnCreatePage.Size = new System.Drawing.Size(28, 28);
            this.BtnCreatePage.Text = "Сохранить в Базу Данных";
            this.BtnCreatePage.ToolTipText = "Добавить вкладку";
            this.BtnCreatePage.Click += new System.EventHandler(this.BtnCreatePage_Click);
            // 
            // BtnCopyPage
            // 
            this.BtnCopyPage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.BtnCopyPage.Image = ((System.Drawing.Image)(resources.GetObject("BtnCopyPage.Image")));
            this.BtnCopyPage.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BtnCopyPage.Name = "BtnCopyPage";
            this.BtnCopyPage.Size = new System.Drawing.Size(28, 28);
            this.BtnCopyPage.Text = "toolStripButton1";
            this.BtnCopyPage.ToolTipText = "Скопировать строки столбцы с другой вкладки";
            this.BtnCopyPage.Click += new System.EventHandler(this.BtnCopyPage_Click);
            // 
            // BtnAdd
            // 
            this.BtnAdd.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.BtnAdd.Image = ((System.Drawing.Image)(resources.GetObject("BtnAdd.Image")));
            this.BtnAdd.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(28, 28);
            this.BtnAdd.Text = "toolStripButton1";
            this.BtnAdd.ToolTipText = "Столбец";
            this.BtnAdd.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // Группа
            // 
            this.Группа.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.Группа.Image = ((System.Drawing.Image)(resources.GetObject("Группа.Image")));
            this.Группа.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Группа.Name = "Группа";
            this.Группа.Size = new System.Drawing.Size(28, 28);
            this.Группа.Text = "Группа";
            this.Группа.Click += new System.EventHandler(this.toolStripButton2_Click);
            // 
            // BtnSaveReport
            // 
            this.BtnSaveReport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.BtnSaveReport.Image = ((System.Drawing.Image)(resources.GetObject("BtnSaveReport.Image")));
            this.BtnSaveReport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BtnSaveReport.Name = "BtnSaveReport";
            this.BtnSaveReport.Size = new System.Drawing.Size(101, 28);
            this.BtnSaveReport.Text = "Сохранить отчёт";
            this.BtnSaveReport.Click += new System.EventHandler(this.BtnSaveReport_Click);
            // 
            // TbxTabDesc
            // 
            this.TbxTabDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TbxTabDesc.BackColor = System.Drawing.Color.White;
            this.TbxTabDesc.Enabled = false;
            this.TbxTabDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxTabDesc.Location = new System.Drawing.Point(4, 32);
            this.TbxTabDesc.Multiline = true;
            this.TbxTabDesc.Name = "TbxTabDesc";
            this.TbxTabDesc.Size = new System.Drawing.Size(594, 47);
            this.TbxTabDesc.TabIndex = 29;
            // 
            // CbxPage
            // 
            this.CbxPage.BackColor = System.Drawing.SystemColors.Control;
            this.CbxPage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CbxPage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.CbxPage.FormattingEnabled = true;
            this.CbxPage.Location = new System.Drawing.Point(3, 5);
            this.CbxPage.Name = "CbxPage";
            this.CbxPage.Size = new System.Drawing.Size(595, 24);
            this.CbxPage.TabIndex = 27;
            this.CbxPage.SelectedIndexChanged += new System.EventHandler(this.CbxPage_SelectedIndexChanged);
            // 
            // btnAddContextMenuStrip
            // 
            this.btnAddContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.столбецToolStripMenuItem,
            this.группаToolStripMenuItem});
            this.btnAddContextMenuStrip.Name = "contextMenuStrip1";
            this.btnAddContextMenuStrip.Size = new System.Drawing.Size(122, 48);
            this.btnAddContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(this.btnAddContextMenuStrip_Opening);
            // 
            // столбецToolStripMenuItem
            // 
            this.столбецToolStripMenuItem.Name = "столбецToolStripMenuItem";
            this.столбецToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.столбецToolStripMenuItem.Text = "Столбец";
            this.столбецToolStripMenuItem.Click += new System.EventHandler(this.столбецToolStripMenuItem_Click);
            // 
            // группаToolStripMenuItem
            // 
            this.группаToolStripMenuItem.Name = "группаToolStripMenuItem";
            this.группаToolStripMenuItem.Size = new System.Drawing.Size(121, 22);
            this.группаToolStripMenuItem.Text = "Группа";
            this.группаToolStripMenuItem.Click += new System.EventHandler(this.группаToolStripMenuItem_Click);
            // 
            // BtnDelete
            // 
            this.BtnDelete.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.BtnDelete.Location = new System.Drawing.Point(466, 564);
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.Size = new System.Drawing.Size(132, 29);
            this.BtnDelete.TabIndex = 6;
            this.BtnDelete.Text = "Удалить";
            this.BtnDelete.UseVisualStyleBackColor = true;
            this.BtnDelete.Click += new System.EventHandler(this.BtnDelete_Click);
            // 
            // treeListView1
            // 
            this.treeListView1.AlternateRowBackColor = System.Drawing.SystemColors.ActiveCaption;
            this.treeListView1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.treeListView1.CellEditUseWholeCell = false;
            this.treeListView1.CellVerticalAlignment = System.Drawing.StringAlignment.Far;
            this.treeListView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.treeListView1.FullRowSelect = true;
            this.treeListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.treeListView1.HeaderUsesThemes = true;
            this.treeListView1.HeaderWordWrap = true;
            this.treeListView1.HideSelection = false;
            this.treeListView1.Location = new System.Drawing.Point(4, 114);
            this.treeListView1.Name = "treeListView1";
            this.treeListView1.OwnerDrawnHeader = true;
            this.treeListView1.RevealAfterExpand = false;
            this.treeListView1.SelectedBackColor = System.Drawing.Color.DarkCyan;
            this.treeListView1.SelectedForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.treeListView1.ShowGroups = false;
            this.treeListView1.ShowHeaderInAllViews = false;
            this.treeListView1.ShowItemToolTips = true;
            this.treeListView1.Size = new System.Drawing.Size(594, 444);
            this.treeListView1.TabIndex = 30;
            this.treeListView1.UseCompatibleStateImageBehavior = false;
            this.treeListView1.View = System.Windows.Forms.View.Details;
            this.treeListView1.VirtualMode = true;
            this.treeListView1.FormatRow += new System.EventHandler<BrightIdeasSoftware.FormatRowEventArgs>(this.treeListView1_FormatRow);
            this.treeListView1.SelectionChanged += new System.EventHandler(this.treeListView1_SelectionChanged);
            this.treeListView1.SelectedIndexChanged += new System.EventHandler(this.treeListView1_SelectedIndexChanged);
            this.treeListView1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.treeListView1_MouseClick);
            // 
            // treeViewContextMenuStrip
            // 
            this.treeViewContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.снятьВыделениеToolStripMenuItem});
            this.treeViewContextMenuStrip.Name = "treeViewContextMenuStrip";
            this.treeViewContextMenuStrip.Size = new System.Drawing.Size(170, 26);
            // 
            // снятьВыделениеToolStripMenuItem
            // 
            this.снятьВыделениеToolStripMenuItem.Name = "снятьВыделениеToolStripMenuItem";
            this.снятьВыделениеToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.снятьВыделениеToolStripMenuItem.Text = "Снять выделение";
            this.снятьВыделениеToolStripMenuItem.Click += new System.EventHandler(this.снятьВыделениеToolStripMenuItem_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.BtnEditElement);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.TbxDescElement);
            this.groupBox2.Controls.Add(this.TbxNameElement);
            this.groupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox2.Location = new System.Drawing.Point(8, 422);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(409, 186);
            this.groupBox2.TabIndex = 33;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Элемент";
            // 
            // BtnEditElement
            // 
            this.BtnEditElement.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.BtnEditElement.Location = new System.Drawing.Point(247, 142);
            this.BtnEditElement.Name = "BtnEditElement";
            this.BtnEditElement.Size = new System.Drawing.Size(151, 38);
            this.BtnEditElement.TabIndex = 38;
            this.BtnEditElement.Text = "Редактировать";
            this.BtnEditElement.UseVisualStyleBackColor = true;
            this.BtnEditElement.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(10, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 17);
            this.label3.TabIndex = 37;
            this.label3.Text = "Описание:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(10, 24);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(110, 17);
            this.label6.TabIndex = 35;
            this.label6.Text = "Наименование:";
            // 
            // TbxDescElement
            // 
            this.TbxDescElement.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxDescElement.Location = new System.Drawing.Point(13, 93);
            this.TbxDescElement.Multiline = true;
            this.TbxDescElement.Name = "TbxDescElement";
            this.TbxDescElement.Size = new System.Drawing.Size(385, 43);
            this.TbxDescElement.TabIndex = 33;
            this.TbxDescElement.TextChanged += new System.EventHandler(this.TbxDescElement_TextChanged);
            // 
            // TbxNameElement
            // 
            this.TbxNameElement.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxNameElement.Location = new System.Drawing.Point(12, 44);
            this.TbxNameElement.Name = "TbxNameElement";
            this.TbxNameElement.Size = new System.Drawing.Size(386, 23);
            this.TbxNameElement.TabIndex = 31;
            this.TbxNameElement.TextChanged += new System.EventHandler(this.TbxNameElement_TextChanged);
            // 
            // splitContainer1
            // 
            this.splitContainer1.BackColor = System.Drawing.Color.Honeydew;
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer1.Location = new System.Drawing.Point(12, 43);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.BackColor = System.Drawing.SystemColors.Control;
            this.splitContainer1.Panel1.Controls.Add(this.treeListView1);
            this.splitContainer1.Panel1.Controls.Add(this.BtnDelete);
            this.splitContainer1.Panel1.Controls.Add(this.CbxShow);
            this.splitContainer1.Panel1.Controls.Add(this.TbxTabDesc);
            this.splitContainer1.Panel1.Controls.Add(this.CbxPage);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.BackColor = System.Drawing.SystemColors.Control;
            this.splitContainer1.Panel2.Controls.Add(this.groupBox1);
            this.splitContainer1.Panel2.Controls.Add(this.groupBox2);
            this.splitContainer1.Size = new System.Drawing.Size(1052, 628);
            this.splitContainer1.SplitterDistance = 610;
            this.splitContainer1.SplitterWidth = 1;
            this.splitContainer1.TabIndex = 34;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.panel2);
            this.groupBox1.Controls.Add(this.CbxUserRow);
            this.groupBox1.Controls.Add(this.TbxDescReport);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.TbxName);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.DtmDate);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox1.Location = new System.Drawing.Point(8, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(409, 411);
            this.groupBox1.TabIndex = 31;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Отчёт";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(9, 281);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 17);
            this.label5.TabIndex = 70;
            this.label5.Text = "Исполнители:";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.LbEmail);
            this.panel2.Location = new System.Drawing.Point(8, 301);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(390, 101);
            this.panel2.TabIndex = 69;
            // 
            // LbEmail
            // 
            this.LbEmail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.LbEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.LbEmail.FormattingEnabled = true;
            this.LbEmail.Location = new System.Drawing.Point(0, 0);
            this.LbEmail.Name = "LbEmail";
            this.LbEmail.Size = new System.Drawing.Size(390, 101);
            this.LbEmail.TabIndex = 39;
            this.toolTip1.SetToolTip(this.LbEmail, "Содержит информацию о исполнителях. Им придёт письмо на почту");
            // 
            // CbxUserRow
            // 
            this.CbxUserRow.AutoSize = true;
            this.CbxUserRow.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.CbxUserRow.Location = new System.Drawing.Point(12, 261);
            this.CbxUserRow.Name = "CbxUserRow";
            this.CbxUserRow.Size = new System.Drawing.Size(199, 21);
            this.CbxUserRow.TabIndex = 68;
            this.CbxUserRow.Text = "Пользовательские строки";
            this.toolTip1.SetToolTip(this.CbxUserRow, "Определяет могут ли пользователи добавлять новые строки");
            this.CbxUserRow.UseVisualStyleBackColor = true;
            // 
            // TbxDescReport
            // 
            this.TbxDescReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxDescReport.Location = new System.Drawing.Point(12, 136);
            this.TbxDescReport.Multiline = true;
            this.TbxDescReport.Name = "TbxDescReport";
            this.TbxDescReport.Size = new System.Drawing.Size(386, 119);
            this.TbxDescReport.TabIndex = 64;
            this.toolTip1.SetToolTip(this.TbxDescReport, "Описание отчёта");
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(10, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(78, 17);
            this.label4.TabIndex = 67;
            this.label4.Text = "Описание:";
            // 
            // TbxName
            // 
            this.TbxName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.TbxName.Location = new System.Drawing.Point(12, 39);
            this.TbxName.Name = "TbxName";
            this.TbxName.Size = new System.Drawing.Size(385, 23);
            this.TbxName.TabIndex = 62;
            this.toolTip1.SetToolTip(this.TbxName, "Содержит наименование отчёта");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(10, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 17);
            this.label2.TabIndex = 66;
            this.label2.Text = "Дата:";
            // 
            // DtmDate
            // 
            this.DtmDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.DtmDate.Location = new System.Drawing.Point(12, 88);
            this.DtmDate.Name = "DtmDate";
            this.DtmDate.Size = new System.Drawing.Size(386, 23);
            this.DtmDate.TabIndex = 63;
            this.toolTip1.SetToolTip(this.DtmDate, "Дата отчёта");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(10, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 17);
            this.label1.TabIndex = 65;
            this.label1.Text = "Наименование:";
            // 
            // ConstuctorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1085, 679);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.Menu);
            this.Name = "ConstuctorForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Создание отчётной формы";
            this.Load += new System.EventHandler(this.ConstructorForm_Load);
            this.Menu.ResumeLayout(false);
            this.Menu.PerformLayout();
            this.btnAddContextMenuStrip.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.treeListView1)).EndInit();
            this.treeViewContextMenuStrip.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox CbxShow;
        private System.Windows.Forms.ToolStrip Menu;
        private System.Windows.Forms.ToolStripButton BtnCreatePage;
        private System.Windows.Forms.ToolStripButton BtnCopyPage;
        private System.Windows.Forms.TextBox TbxTabDesc;
        private System.Windows.Forms.ComboBox CbxPage;
        private System.Windows.Forms.Button BtnDelete;
        private System.Windows.Forms.ContextMenuStrip btnAddContextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem столбецToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem группаToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip treeViewContextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem снятьВыделениеToolStripMenuItem;
        private BrightIdeasSoftware.TreeListView treeListView1;
        private System.Windows.Forms.ToolStripButton BtnAdd;
        private System.Windows.Forms.ToolStripButton Группа;
        private System.Windows.Forms.ToolStripButton BtnSaveReport;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox TbxDescElement;
        private System.Windows.Forms.TextBox TbxNameElement;
        private System.Windows.Forms.Button BtnEditElement;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckBox CbxUserRow;
        private System.Windows.Forms.TextBox TbxDescReport;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox TbxName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker DtmDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox LbEmail;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}