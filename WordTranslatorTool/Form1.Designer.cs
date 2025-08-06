namespace WordTranslatorTool
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            lblInputFile = new Label();
            btnOpenFile = new Button();
            btnSaveFile = new Button();
            cmbSourceLang = new ComboBox();
            cmbTargetLang = new ComboBox();
            txtApiKey = new TextBox();
            cmbModel = new ComboBox();
            progressBar = new ProgressBar();
            lblSourceLang = new Label();
            lblTargetLang = new Label();
            lblApiKey = new Label();
            lblModel = new Label();
            btnCancel = new Button();
            tableProgress = new TableLayoutPanel();
            txtSearchLog = new TextBox();
            btnNextMatch = new Button();
            txtLog = new RichTextBox();
            tableLayoutPanel1 = new TableLayoutPanel();
            lblTranslateMode = new Label();
            cmbTranslateMode = new ComboBox();
            tableLayoutPanel2 = new TableLayoutPanel();
            txtFileName = new TextBox();
            tableProgress.SuspendLayout();
            tableLayoutPanel1.SuspendLayout();
            tableLayoutPanel2.SuspendLayout();
            SuspendLayout();
            // 
            // lblInputFile
            // 
            lblInputFile.AutoSize = true;
            lblInputFile.Location = new Point(12, 149);
            lblInputFile.Name = "lblInputFile";
            lblInputFile.Size = new Size(100, 23);
            lblInputFile.TabIndex = 0;
            lblInputFile.Text = "選擇檔案：";
            // 
            // btnOpenFile
            // 
            btnOpenFile.Location = new Point(968, 3);
            btnOpenFile.Name = "btnOpenFile";
            btnOpenFile.Size = new Size(112, 29);
            btnOpenFile.TabIndex = 1;
            btnOpenFile.Text = "開啟檔案";
            btnOpenFile.UseVisualStyleBackColor = true;
            btnOpenFile.Click += btnOpenFile_Click;
            // 
            // btnSaveFile
            // 
            btnSaveFile.Location = new Point(18, 351);
            btnSaveFile.Name = "btnSaveFile";
            btnSaveFile.Size = new Size(232, 34);
            btnSaveFile.TabIndex = 2;
            btnSaveFile.Text = "開始翻譯並儲存檔案";
            btnSaveFile.UseVisualStyleBackColor = true;
            btnSaveFile.Click += btnSaveFile_Click;
            // 
            // cmbSourceLang
            // 
            cmbSourceLang.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbSourceLang.FormattingEnabled = true;
            cmbSourceLang.Location = new Point(183, 201);
            cmbSourceLang.Name = "cmbSourceLang";
            cmbSourceLang.Size = new Size(230, 31);
            cmbSourceLang.TabIndex = 3;
            cmbSourceLang.SelectedIndexChanged += cmbSourceLang_SelectedIndexChanged;
            // 
            // cmbTargetLang
            // 
            cmbTargetLang.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbTargetLang.FormattingEnabled = true;
            cmbTargetLang.Location = new Point(674, 204);
            cmbTargetLang.Name = "cmbTargetLang";
            cmbTargetLang.Size = new Size(230, 31);
            cmbTargetLang.TabIndex = 4;
            // 
            // txtApiKey
            // 
            txtApiKey.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtApiKey.Location = new Point(183, 33);
            txtApiKey.Name = "txtApiKey";
            txtApiKey.Size = new Size(1083, 30);
            txtApiKey.TabIndex = 5;
            // 
            // cmbModel
            // 
            cmbModel.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbModel.FormattingEnabled = true;
            cmbModel.Location = new Point(183, 89);
            cmbModel.Name = "cmbModel";
            cmbModel.Size = new Size(230, 31);
            cmbModel.TabIndex = 6;
            // 
            // progressBar
            // 
            progressBar.Dock = DockStyle.Fill;
            progressBar.Location = new Point(3, 3);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(1127, 33);
            progressBar.TabIndex = 9;
            // 
            // lblSourceLang
            // 
            lblSourceLang.AutoSize = true;
            lblSourceLang.Location = new Point(13, 205);
            lblSourceLang.Name = "lblSourceLang";
            lblSourceLang.Size = new Size(100, 23);
            lblSourceLang.TabIndex = 11;
            lblSourceLang.Text = "原始語言：";
            // 
            // lblTargetLang
            // 
            lblTargetLang.AutoSize = true;
            lblTargetLang.Location = new Point(510, 208);
            lblTargetLang.Name = "lblTargetLang";
            lblTargetLang.Size = new Size(100, 23);
            lblTargetLang.TabIndex = 12;
            lblTargetLang.Text = "目標語言：";
            // 
            // lblApiKey
            // 
            lblApiKey.AutoSize = true;
            lblApiKey.Location = new Point(12, 37);
            lblApiKey.Name = "lblApiKey";
            lblApiKey.Size = new Size(170, 23);
            lblApiKey.TabIndex = 13;
            lblApiKey.Text = "ChatGPT API Key：";
            // 
            // lblModel
            // 
            lblModel.AutoSize = true;
            lblModel.Location = new Point(12, 93);
            lblModel.Name = "lblModel";
            lblModel.Size = new Size(100, 23);
            lblModel.TabIndex = 14;
            lblModel.Text = "選擇模型：";
            // 
            // btnCancel
            // 
            btnCancel.Enabled = false;
            btnCancel.Location = new Point(1136, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(112, 33);
            btnCancel.TabIndex = 15;
            btnCancel.Text = "停止翻譯";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // tableProgress
            // 
            tableProgress.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tableProgress.ColumnCount = 2;
            tableProgress.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableProgress.ColumnStyles.Add(new ColumnStyle());
            tableProgress.Controls.Add(progressBar, 0, 0);
            tableProgress.Controls.Add(btnCancel, 1, 0);
            tableProgress.Location = new Point(15, 415);
            tableProgress.Name = "tableProgress";
            tableProgress.RowCount = 1;
            tableProgress.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableProgress.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tableProgress.Size = new Size(1251, 39);
            tableProgress.TabIndex = 16;
            tableProgress.Paint += tableProgress_Paint;
            // 
            // txtSearchLog
            // 
            txtSearchLog.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtSearchLog.Location = new Point(3, 3);
            txtSearchLog.Name = "txtSearchLog";
            txtSearchLog.PlaceholderText = "輸入關鍵字…";
            txtSearchLog.Size = new Size(1127, 30);
            txtSearchLog.TabIndex = 17;
            txtSearchLog.KeyDown += txtSearchLog_KeyDown;
            // 
            // btnNextMatch
            // 
            btnNextMatch.Location = new Point(1136, 3);
            btnNextMatch.Name = "btnNextMatch";
            btnNextMatch.Size = new Size(112, 34);
            btnNextMatch.TabIndex = 18;
            btnNextMatch.Text = "下一筆";
            btnNextMatch.UseVisualStyleBackColor = true;
            btnNextMatch.Click += btnNextMatch_Click;
            // 
            // txtLog
            // 
            txtLog.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            txtLog.Location = new Point(15, 522);
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.Size = new Size(1251, 681);
            txtLog.TabIndex = 19;
            txtLog.Text = "";
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle());
            tableLayoutPanel1.Controls.Add(txtSearchLog, 0, 0);
            tableLayoutPanel1.Controls.Add(btnNextMatch, 1, 0);
            tableLayoutPanel1.Location = new Point(15, 474);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 1;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Size = new Size(1251, 42);
            tableLayoutPanel1.TabIndex = 16;
            // 
            // lblTranslateMode
            // 
            lblTranslateMode.AutoSize = true;
            lblTranslateMode.Location = new Point(16, 261);
            lblTranslateMode.Name = "lblTranslateMode";
            lblTranslateMode.Size = new Size(100, 23);
            lblTranslateMode.TabIndex = 20;
            lblTranslateMode.Text = "翻譯模式：";
            lblTranslateMode.Click += label1_Click;
            // 
            // cmbTranslateMode
            // 
            cmbTranslateMode.FormattingEnabled = true;
            cmbTranslateMode.Location = new Point(183, 261);
            cmbTranslateMode.Name = "cmbTranslateMode";
            cmbTranslateMode.Size = new Size(230, 31);
            cmbTranslateMode.TabIndex = 21;
            // 
            // tableLayoutPanel2
            // 
            tableLayoutPanel2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tableLayoutPanel2.ColumnCount = 2;
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.ColumnStyles.Add(new ColumnStyle());
            tableLayoutPanel2.Controls.Add(btnOpenFile, 1, 0);
            tableLayoutPanel2.Controls.Add(txtFileName, 0, 0);
            tableLayoutPanel2.Location = new Point(183, 143);
            tableLayoutPanel2.Name = "tableLayoutPanel2";
            tableLayoutPanel2.RowCount = 1;
            tableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel2.Size = new Size(1083, 35);
            tableLayoutPanel2.TabIndex = 22;
            // 
            // txtFileName
            // 
            txtFileName.Dock = DockStyle.Fill;
            txtFileName.Location = new Point(3, 3);
            txtFileName.Name = "txtFileName";
            txtFileName.ReadOnly = true;
            txtFileName.Size = new Size(959, 30);
            txtFileName.TabIndex = 2;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(11F, 23F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1281, 1215);
            Controls.Add(tableLayoutPanel2);
            Controls.Add(cmbTranslateMode);
            Controls.Add(lblTranslateMode);
            Controls.Add(txtLog);
            Controls.Add(tableLayoutPanel1);
            Controls.Add(tableProgress);
            Controls.Add(lblModel);
            Controls.Add(lblApiKey);
            Controls.Add(lblTargetLang);
            Controls.Add(lblSourceLang);
            Controls.Add(cmbModel);
            Controls.Add(txtApiKey);
            Controls.Add(cmbTargetLang);
            Controls.Add(cmbSourceLang);
            Controls.Add(btnSaveFile);
            Controls.Add(lblInputFile);
            Name = "Form1";
            Text = "Form1";
            tableProgress.ResumeLayout(false);
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            tableLayoutPanel2.ResumeLayout(false);
            tableLayoutPanel2.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lblInputFile;
        private Button btnOpenFile;
        private Button btnSaveFile;
        private ComboBox cmbSourceLang;
        private ComboBox cmbTargetLang;
        private TextBox txtApiKey;
        private ComboBox cmbModel;
        private ProgressBar progressBar;
        private Label lblSourceLang;
        private Label lblTargetLang;
        private Label lblApiKey;
        private Label lblModel;
        private Button btnCancel;
        private TableLayoutPanel tableProgress;
        private TextBox txtSearchLog;
        private Button btnNextMatch;
        private RichTextBox txtLog;
        private TableLayoutPanel tableLayoutPanel1;
        private Label lblTranslateMode;
        private ComboBox cmbTranslateMode;
        private TableLayoutPanel tableLayoutPanel2;
        private TextBox txtFileName;
    }
}
