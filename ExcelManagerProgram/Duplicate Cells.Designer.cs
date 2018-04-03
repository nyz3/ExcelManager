namespace ExcelManager
{
    partial class DuplicateCells
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
            this.label4 = new System.Windows.Forms.Label();
            this.fileSelection1 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tableSelection1 = new System.Windows.Forms.ComboBox();
            this.selectAll1 = new System.Windows.Forms.Button();
            this.unselectAll1 = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 13);
            this.label4.TabIndex = 16;
            this.label4.Text = "Choose which file";
            // 
            // fileSelection1
            // 
            this.fileSelection1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.fileSelection1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.fileSelection1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fileSelection1.FormattingEnabled = true;
            this.fileSelection1.Location = new System.Drawing.Point(15, 25);
            this.fileSelection1.Name = "fileSelection1";
            this.fileSelection1.Size = new System.Drawing.Size(200, 21);
            this.fileSelection1.TabIndex = 15;
            this.fileSelection1.SelectedIndexChanged += new System.EventHandler(this.fileSelection1_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 60);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Choose which sheets to merge";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 106);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(163, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Choose which columns to search";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // tableSelection1
            // 
            this.tableSelection1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.tableSelection1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.tableSelection1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tableSelection1.FormattingEnabled = true;
            this.tableSelection1.Location = new System.Drawing.Point(15, 76);
            this.tableSelection1.Name = "tableSelection1";
            this.tableSelection1.Size = new System.Drawing.Size(200, 21);
            this.tableSelection1.TabIndex = 12;
            this.tableSelection1.SelectedIndexChanged += new System.EventHandler(this.tableSelection1_SelectedIndexChanged);
            // 
            // selectAll1
            // 
            this.selectAll1.Location = new System.Drawing.Point(15, 131);
            this.selectAll1.Name = "selectAll1";
            this.selectAll1.Size = new System.Drawing.Size(75, 23);
            this.selectAll1.TabIndex = 19;
            this.selectAll1.Text = "Select All";
            this.selectAll1.UseVisualStyleBackColor = true;
            this.selectAll1.Click += new System.EventHandler(this.selectAll1_Click);
            // 
            // unselectAll1
            // 
            this.unselectAll1.Location = new System.Drawing.Point(15, 153);
            this.unselectAll1.Name = "unselectAll1";
            this.unselectAll1.Size = new System.Drawing.Size(75, 23);
            this.unselectAll1.TabIndex = 18;
            this.unselectAll1.Text = "Unselect All";
            this.unselectAll1.UseVisualStyleBackColor = true;
            this.unselectAll1.Click += new System.EventHandler(this.unselectAll1_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(15, 191);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(200, 244);
            this.checkedListBox1.TabIndex = 17;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(244, 25);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(569, 448);
            this.dataGridView1.TabIndex = 20;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(46, 449);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(142, 24);
            this.button1.TabIndex = 21;
            this.button1.Text = "Find Duplicates";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // DuplicateCells
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(825, 485);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.selectAll1);
            this.Controls.Add(this.unselectAll1);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.fileSelection1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableSelection1);
            this.Name = "DuplicateCells";
            this.Text = "Duplicate_Cells";
            this.Load += new System.EventHandler(this.Duplicate_Cells_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox fileSelection1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox tableSelection1;
        private System.Windows.Forms.Button selectAll1;
        private System.Windows.Forms.Button unselectAll1;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
    }
}