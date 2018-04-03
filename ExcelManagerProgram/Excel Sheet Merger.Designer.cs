namespace ExcelManager
{
    partial class MergeCells
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
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.tableSelection1 = new System.Windows.Forms.ComboBox();
            this.tableSelection2 = new System.Windows.Forms.ComboBox();
            this.mergeAndSave = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.mergeColumnSelection = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.fileSelection1 = new System.Windows.Forms.ComboBox();
            this.fileSelection2 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.unselectAll1 = new System.Windows.Forms.Button();
            this.selectAll1 = new System.Windows.Forms.Button();
            this.selectAll2 = new System.Windows.Forms.Button();
            this.unselectAll2 = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.sameNameSelection = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.newName = new System.Windows.Forms.TextBox();
            this.changeName = new System.Windows.Forms.Button();
            this.keyColumnList = new System.Windows.Forms.ListBox();
            this.addColumn = new System.Windows.Forms.Button();
            this.removeColumn = new System.Windows.Forms.Button();
            this.Instructions = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(10, 183);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(201, 319);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.CheckOnClick = true;
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(268, 183);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(198, 319);
            this.checkedListBox2.TabIndex = 1;
            // 
            // tableSelection1
            // 
            this.tableSelection1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.tableSelection1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.tableSelection1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tableSelection1.FormattingEnabled = true;
            this.tableSelection1.Location = new System.Drawing.Point(11, 77);
            this.tableSelection1.Name = "tableSelection1";
            this.tableSelection1.Size = new System.Drawing.Size(200, 21);
            this.tableSelection1.TabIndex = 2;
            this.tableSelection1.SelectedIndexChanged += new System.EventHandler(this.tableSelection1_SelectedIndexChanged);
            // 
            // tableSelection2
            // 
            this.tableSelection2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.tableSelection2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.tableSelection2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tableSelection2.FormattingEnabled = true;
            this.tableSelection2.Location = new System.Drawing.Point(268, 77);
            this.tableSelection2.Name = "tableSelection2";
            this.tableSelection2.Size = new System.Drawing.Size(198, 21);
            this.tableSelection2.TabIndex = 3;
            this.tableSelection2.SelectedIndexChanged += new System.EventHandler(this.tableSelection2_SelectedIndexChanged);
            // 
            // mergeAndSave
            // 
            this.mergeAndSave.Location = new System.Drawing.Point(184, 508);
            this.mergeAndSave.Name = "mergeAndSave";
            this.mergeAndSave.Size = new System.Drawing.Size(108, 28);
            this.mergeAndSave.TabIndex = 4;
            this.mergeAndSave.Text = "Merge and Save";
            this.mergeAndSave.UseVisualStyleBackColor = true;
            this.mergeAndSave.Click += new System.EventHandler(this.mergeAndSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 107);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(155, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Choose which columns to keep";
            // 
            // mergeColumnSelection
            // 
            this.mergeColumnSelection.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.mergeColumnSelection.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.mergeColumnSelection.FormattingEnabled = true;
            this.mergeColumnSelection.Location = new System.Drawing.Point(484, 45);
            this.mergeColumnSelection.Name = "mergeColumnSelection";
            this.mergeColumnSelection.Size = new System.Drawing.Size(121, 21);
            this.mergeColumnSelection.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(481, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(269, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Choose which columns to be keys (only from first Sheet)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 61);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Choose which sheets to merge";
            // 
            // fileSelection1
            // 
            this.fileSelection1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.fileSelection1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.fileSelection1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fileSelection1.FormattingEnabled = true;
            this.fileSelection1.Location = new System.Drawing.Point(11, 26);
            this.fileSelection1.Name = "fileSelection1";
            this.fileSelection1.Size = new System.Drawing.Size(200, 21);
            this.fileSelection1.TabIndex = 9;
            this.fileSelection1.SelectedIndexChanged += new System.EventHandler(this.fileSelection1_SelectedIndexChanged_1);
            // 
            // fileSelection2
            // 
            this.fileSelection2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.fileSelection2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.fileSelection2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.fileSelection2.FormattingEnabled = true;
            this.fileSelection2.Location = new System.Drawing.Point(268, 26);
            this.fileSelection2.Name = "fileSelection2";
            this.fileSelection2.Size = new System.Drawing.Size(198, 21);
            this.fileSelection2.TabIndex = 10;
            this.fileSelection2.SelectedIndexChanged += new System.EventHandler(this.fileSelection2_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 10);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(90, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Choose which file";
            // 
            // unselectAll1
            // 
            this.unselectAll1.Location = new System.Drawing.Point(10, 145);
            this.unselectAll1.Name = "unselectAll1";
            this.unselectAll1.Size = new System.Drawing.Size(75, 23);
            this.unselectAll1.TabIndex = 12;
            this.unselectAll1.Text = "Unselect All";
            this.unselectAll1.UseVisualStyleBackColor = true;
            this.unselectAll1.Click += new System.EventHandler(this.unselectAll1_Click);
            // 
            // selectAll1
            // 
            this.selectAll1.Location = new System.Drawing.Point(10, 123);
            this.selectAll1.Name = "selectAll1";
            this.selectAll1.Size = new System.Drawing.Size(75, 23);
            this.selectAll1.TabIndex = 13;
            this.selectAll1.Text = "Select All";
            this.selectAll1.UseVisualStyleBackColor = true;
            this.selectAll1.Click += new System.EventHandler(this.selectAll1_Click);
            // 
            // selectAll2
            // 
            this.selectAll2.Location = new System.Drawing.Point(268, 123);
            this.selectAll2.Name = "selectAll2";
            this.selectAll2.Size = new System.Drawing.Size(75, 23);
            this.selectAll2.TabIndex = 14;
            this.selectAll2.Text = "Select All";
            this.selectAll2.UseVisualStyleBackColor = true;
            this.selectAll2.Click += new System.EventHandler(this.selectAll2_Click);
            // 
            // unselectAll2
            // 
            this.unselectAll2.Location = new System.Drawing.Point(268, 145);
            this.unselectAll2.Name = "unselectAll2";
            this.unselectAll2.Size = new System.Drawing.Size(75, 23);
            this.unselectAll2.TabIndex = 15;
            this.unselectAll2.Text = "Unselect All";
            this.unselectAll2.UseVisualStyleBackColor = true;
            this.unselectAll2.Click += new System.EventHandler(this.unselectAll2_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.SystemColors.Control;
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox1.Location = new System.Drawing.Point(484, 366);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(250, 72);
            this.richTextBox1.TabIndex = 16;
            this.richTextBox1.Text = "These Columns have the same name, if you would like to keep them seperate, please" +
    " change their name. This list contains only the columns from the 2nd file.";
            // 
            // sameNameSelection
            // 
            this.sameNameSelection.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.sameNameSelection.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.sameNameSelection.FormattingEnabled = true;
            this.sameNameSelection.Location = new System.Drawing.Point(484, 426);
            this.sameNameSelection.Name = "sameNameSelection";
            this.sameNameSelection.Size = new System.Drawing.Size(250, 21);
            this.sameNameSelection.TabIndex = 17;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(481, 466);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(60, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "New Name";
            // 
            // newName
            // 
            this.newName.Location = new System.Drawing.Point(484, 482);
            this.newName.Name = "newName";
            this.newName.Size = new System.Drawing.Size(170, 20);
            this.newName.TabIndex = 19;
            // 
            // changeName
            // 
            this.changeName.Location = new System.Drawing.Point(660, 479);
            this.changeName.Name = "changeName";
            this.changeName.Size = new System.Drawing.Size(75, 23);
            this.changeName.TabIndex = 20;
            this.changeName.Text = "Change";
            this.changeName.UseVisualStyleBackColor = true;
            this.changeName.Click += new System.EventHandler(this.changeName_Click);
            // 
            // keyColumnList
            // 
            this.keyColumnList.FormattingEnabled = true;
            this.keyColumnList.Location = new System.Drawing.Point(484, 110);
            this.keyColumnList.Name = "keyColumnList";
            this.keyColumnList.Size = new System.Drawing.Size(254, 238);
            this.keyColumnList.TabIndex = 21;
            this.keyColumnList.SelectedIndexChanged += new System.EventHandler(this.keyColumnList_SelectedIndexChanged);
            // 
            // addColumn
            // 
            this.addColumn.Location = new System.Drawing.Point(484, 72);
            this.addColumn.Name = "addColumn";
            this.addColumn.Size = new System.Drawing.Size(75, 23);
            this.addColumn.TabIndex = 22;
            this.addColumn.Text = "Add";
            this.addColumn.UseVisualStyleBackColor = true;
            this.addColumn.Click += new System.EventHandler(this.addColumn_Click);
            // 
            // removeColumn
            // 
            this.removeColumn.Location = new System.Drawing.Point(566, 73);
            this.removeColumn.Name = "removeColumn";
            this.removeColumn.Size = new System.Drawing.Size(75, 23);
            this.removeColumn.TabIndex = 23;
            this.removeColumn.Text = "Remove";
            this.removeColumn.UseVisualStyleBackColor = true;
            this.removeColumn.Click += new System.EventHandler(this.removeColumn_Click);
            // 
            // Instructions
            // 
            this.Instructions.AutoSize = true;
            this.Instructions.Location = new System.Drawing.Point(481, 29);
            this.Instructions.Name = "Instructions";
            this.Instructions.Size = new System.Drawing.Size(259, 13);
            this.Instructions.TabIndex = 24;
            this.Instructions.Text = "The key columns selected must not have empty cells.";
            // 
            // MergeCells
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(774, 557);
            this.Controls.Add(this.Instructions);
            this.Controls.Add(this.removeColumn);
            this.Controls.Add(this.addColumn);
            this.Controls.Add(this.keyColumnList);
            this.Controls.Add(this.changeName);
            this.Controls.Add(this.newName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.sameNameSelection);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.unselectAll2);
            this.Controls.Add(this.selectAll2);
            this.Controls.Add(this.selectAll1);
            this.Controls.Add(this.unselectAll1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.fileSelection2);
            this.Controls.Add(this.fileSelection1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.mergeColumnSelection);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.mergeAndSave);
            this.Controls.Add(this.tableSelection2);
            this.Controls.Add(this.tableSelection1);
            this.Controls.Add(this.checkedListBox2);
            this.Controls.Add(this.checkedListBox1);
            this.Name = "MergeCells";
            this.Text = "Excel Merger";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.CheckedListBox checkedListBox2;
        private System.Windows.Forms.ComboBox tableSelection1;
        private System.Windows.Forms.ComboBox tableSelection2;
        private System.Windows.Forms.Button mergeAndSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox mergeColumnSelection;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox fileSelection1;
        private System.Windows.Forms.ComboBox fileSelection2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button unselectAll1;
        private System.Windows.Forms.Button selectAll1;
        private System.Windows.Forms.Button selectAll2;
        private System.Windows.Forms.Button unselectAll2;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ComboBox sameNameSelection;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox newName;
        private System.Windows.Forms.Button changeName;
        private System.Windows.Forms.ListBox keyColumnList;
        private System.Windows.Forms.Button addColumn;
        private System.Windows.Forms.Button removeColumn;
        private System.Windows.Forms.Label Instructions;
    }
}