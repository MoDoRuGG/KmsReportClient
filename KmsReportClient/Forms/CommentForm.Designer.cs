namespace KmsReportClient.Forms
{
    partial class CommentForm
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
            this.BtnDo = new System.Windows.Forms.Button();
            this.BtnClear = new System.Windows.Forms.Button();
            this.BtnClose = new System.Windows.Forms.Button();
            this.TxtbComment = new System.Windows.Forms.TextBox();
            this.TxtbChat = new System.Windows.Forms.TextBox();
            this.BtnRefresh = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // BtnDo
            // 
            this.BtnDo.Location = new System.Drawing.Point(13, 392);
            this.BtnDo.Name = "BtnDo";
            this.BtnDo.Size = new System.Drawing.Size(75, 23);
            this.BtnDo.TabIndex = 0;
            this.BtnDo.Text = "Отправить";
            this.BtnDo.UseVisualStyleBackColor = true;
            this.BtnDo.Click += new System.EventHandler(this.BtnDo_Click);
            // 
            // BtnClear
            // 
            this.BtnClear.Location = new System.Drawing.Point(94, 392);
            this.BtnClear.Name = "BtnClear";
            this.BtnClear.Size = new System.Drawing.Size(75, 23);
            this.BtnClear.TabIndex = 1;
            this.BtnClear.Text = "Очистить";
            this.BtnClear.UseVisualStyleBackColor = true;
            this.BtnClear.Click += new System.EventHandler(this.BtnClear_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(466, 392);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(75, 23);
            this.BtnClose.TabIndex = 2;
            this.BtnClose.Text = "Закрыть";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // TxtbComment
            // 
            this.TxtbComment.Location = new System.Drawing.Point(12, 328);
            this.TxtbComment.Multiline = true;
            this.TxtbComment.Name = "TxtbComment";
            this.TxtbComment.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TxtbComment.Size = new System.Drawing.Size(529, 58);
            this.TxtbComment.TabIndex = 3;
            // 
            // TxtbChat
            // 
            this.TxtbChat.Location = new System.Drawing.Point(13, 13);
            this.TxtbChat.Multiline = true;
            this.TxtbChat.Name = "TxtbChat";
            this.TxtbChat.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TxtbChat.Size = new System.Drawing.Size(528, 309);
            this.TxtbChat.TabIndex = 4;
            // 
            // BtnRefresh
            // 
            this.BtnRefresh.Location = new System.Drawing.Point(336, 392);
            this.BtnRefresh.Name = "BtnRefresh";
            this.BtnRefresh.Size = new System.Drawing.Size(124, 23);
            this.BtnRefresh.TabIndex = 5;
            this.BtnRefresh.Text = "Обновить окно чата";
            this.BtnRefresh.UseVisualStyleBackColor = true;
            this.BtnRefresh.Click += new System.EventHandler(this.BtnRefresh_Click);
            // 
            // CommentForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(553, 429);
            this.Controls.Add(this.BtnRefresh);
            this.Controls.Add(this.TxtbChat);
            this.Controls.Add(this.TxtbComment);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.BtnClear);
            this.Controls.Add(this.BtnDo);
            this.Name = "CommentForm";
            this.Text = "Комментарии к отчету";
            this.Load += new System.EventHandler(this.CommentForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnDo;
        private System.Windows.Forms.Button BtnClear;
        private System.Windows.Forms.Button BtnClose;
        private System.Windows.Forms.TextBox TxtbComment;
        private System.Windows.Forms.TextBox TxtbChat;
        private System.Windows.Forms.Button BtnRefresh;
    }
}