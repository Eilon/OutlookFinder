namespace OutlookFinderApp
{
    partial class OutlookFinderAppForm
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
            this.label1 = new System.Windows.Forms.Label();
            this._totalEmailsValueLabel = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this._folderValueLabel = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this._taggedEmailsValueLabel = new System.Windows.Forms.Label();
            this._logOutputTextBox = new System.Windows.Forms.TextBox();
            this._runNowButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this._tagResultsListView = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label4 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(134, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Total emails:";
            // 
            // _totalEmailsValueLabel
            // 
            this._totalEmailsValueLabel.AutoSize = true;
            this._totalEmailsValueLabel.Location = new System.Drawing.Point(168, 25);
            this._totalEmailsValueLabel.Name = "_totalEmailsValueLabel";
            this._totalEmailsValueLabel.Size = new System.Drawing.Size(0, 25);
            this._totalEmailsValueLabel.TabIndex = 1;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.label5, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this._folderValueLabel, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this._totalEmailsValueLabel, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this._taggedEmailsValueLabel, 1, 2);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(18, 12);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(417, 96);
            this.tableLayoutPanel1.TabIndex = 2;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(79, 25);
            this.label5.TabIndex = 4;
            this.label5.Text = "Folder:";
            // 
            // _folderValueLabel
            // 
            this._folderValueLabel.AutoSize = true;
            this._folderValueLabel.Location = new System.Drawing.Point(168, 0);
            this._folderValueLabel.Name = "_folderValueLabel";
            this._folderValueLabel.Size = new System.Drawing.Size(0, 25);
            this._folderValueLabel.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(159, 25);
            this.label3.TabIndex = 2;
            this.label3.Text = "Tagged emails:";
            // 
            // _taggedEmailsValueLabel
            // 
            this._taggedEmailsValueLabel.AutoSize = true;
            this._taggedEmailsValueLabel.Location = new System.Drawing.Point(168, 50);
            this._taggedEmailsValueLabel.Name = "_taggedEmailsValueLabel";
            this._taggedEmailsValueLabel.Size = new System.Drawing.Size(0, 25);
            this._taggedEmailsValueLabel.TabIndex = 3;
            // 
            // _logOutputTextBox
            // 
            this._logOutputTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._logOutputTextBox.Location = new System.Drawing.Point(18, 484);
            this._logOutputTextBox.Multiline = true;
            this._logOutputTextBox.Name = "_logOutputTextBox";
            this._logOutputTextBox.ReadOnly = true;
            this._logOutputTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this._logOutputTextBox.Size = new System.Drawing.Size(1029, 261);
            this._logOutputTextBox.TabIndex = 3;
            this._logOutputTextBox.WordWrap = false;
            // 
            // _runNowButton
            // 
            this._runNowButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._runNowButton.Location = new System.Drawing.Point(485, 30);
            this._runNowButton.Name = "_runNowButton";
            this._runNowButton.Size = new System.Drawing.Size(562, 78);
            this._runNowButton.TabIndex = 4;
            this._runNowButton.Text = "&Run now";
            this._runNowButton.UseVisualStyleBackColor = true;
            this._runNowButton.Click += new System.EventHandler(this.OnRunNowButton_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 456);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 25);
            this.label2.TabIndex = 5;
            this.label2.Text = "&Log output:";
            // 
            // _tagResultsListView
            // 
            this._tagResultsListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._tagResultsListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this._tagResultsListView.FullRowSelect = true;
            this._tagResultsListView.GridLines = true;
            this._tagResultsListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this._tagResultsListView.Location = new System.Drawing.Point(18, 156);
            this._tagResultsListView.MultiSelect = false;
            this._tagResultsListView.Name = "_tagResultsListView";
            this._tagResultsListView.ShowGroups = false;
            this._tagResultsListView.Size = new System.Drawing.Size(1029, 266);
            this._tagResultsListView.TabIndex = 6;
            this._tagResultsListView.UseCompatibleStateImageBehavior = false;
            this._tagResultsListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Tag";
            this.columnHeader1.Width = 200;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Count";
            this.columnHeader2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.columnHeader2.Width = 80;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(18, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(125, 25);
            this.label4.TabIndex = 7;
            this.label4.Text = "&Tag results:";
            // 
            // OutlookFinderAppForm
            // 
            this.AcceptButton = this._runNowButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1059, 757);
            this.Controls.Add(this.label4);
            this.Controls.Add(this._tagResultsListView);
            this.Controls.Add(this.label2);
            this.Controls.Add(this._runNowButton);
            this.Controls.Add(this._logOutputTextBox);
            this.Controls.Add(this.tableLayoutPanel1);
            this.MinimumSize = new System.Drawing.Size(1085, 828);
            this.Name = "OutlookFinderAppForm";
            this.Text = "Outlook Finder";
            this.Load += new System.EventHandler(this.OutlookFinderAppForm_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label _totalEmailsValueLabel;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label _taggedEmailsValueLabel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox _logOutputTextBox;
        private System.Windows.Forms.Button _runNowButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListView _tagResultsListView;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label _folderValueLabel;
    }
}

