using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ExportToExcel;
using System.Text;
using System.Collections;
using System.IO;

namespace ExcelManager
{
    public partial class MergeCells : Form
    {
        private static List<DataSet> _dataSheets1 = new List<DataSet>();
        private static List<DataSet> _dataSheets2 = new List<DataSet>();
        private static List<DataTable> _listTables1 = new List<DataTable>();
        private static List<DataTable> _listTables2 = new List<DataTable>();
        private static readonly BindingList<string> _keyColumns = new BindingList<string>();
        private static BindingList<DataColumn> _duplicateColumns = new BindingList<DataColumn>();
        private bool showWindow = false;

        public MergeCells(List<DataSet> data)
        {
            InitializeComponent();
            _dataSheets1 = new List<DataSet>();
            _dataSheets2 = new List<DataSet>();
            foreach (var dataSheet in data)
            {
                _dataSheets1.Add(dataSheet);
                _dataSheets2.Add(dataSheet);
            }
            DisplayDataSetName1();
            DisplayDataSetName2();

            if (Enumerable.Range(0, _dataSheets1.Count).Contains(fileSelection1.SelectedIndex) && Enumerable.Range(0, _dataSheets2.Count).Contains(fileSelection2.SelectedIndex))
            {
                DisplayTableName1(_dataSheets1[fileSelection1.SelectedIndex]);
                DisplayTableName2(_dataSheets2[fileSelection2.SelectedIndex]);
                showWindow = true;
            }
            else
            {
                showWindow = false;
            }
        }

        private void DisplayDataSetName1()
        {
            var bindingSource1 = new BindingSource();
            bindingSource1.DataSource = _dataSheets1;
            fileSelection1.DataSource = bindingSource1.DataSource;
            fileSelection1.DisplayMember = "DataSetName";
        }

        private void DisplayDataSetName2()
        {
            var bindingSource1 = new BindingSource();
            bindingSource1.DataSource = _dataSheets2;
            fileSelection2.DataSource = bindingSource1.DataSource;
            fileSelection2.DisplayMember = "DataSetName";
        }

        private void DisplayTableName1(DataSet dataTables)
        {
            var bindingSource1 = new BindingSource();
            var tables = dataTables.Tables;
            _listTables1 = new List<DataTable>();
            foreach (DataTable table in tables)
                _listTables1.Add(table);
            bindingSource1.DataSource = _listTables1;
            tableSelection1.DataSource = bindingSource1.DataSource;
            tableSelection1.DisplayMember = "TableName";
        }

        private void DisplayTableName2(DataSet dataTables)
        {
            var bindingSource1 = new BindingSource();
            var tables = dataTables.Tables;
            _listTables2 = new List<DataTable>();
            foreach (DataTable table in tables)
                _listTables2.Add(table);
            bindingSource1.DataSource = _listTables2;
            tableSelection2.DataSource = bindingSource1.DataSource;
            tableSelection2.DisplayMember = "TableName";
        }

        private void DisplayMergeColumn(int index)          //box where you select the keys to merge(displays columns from first sheet, because first prereq)
        {
            var bindingSource1 = new BindingSource();
            var columns = _listTables1[index].Columns;
            var listColumns = new List<DataColumn>();
            foreach (DataColumn column in columns)
                listColumns.Add(column);
            bindingSource1.DataSource = listColumns;
            mergeColumnSelection.DataSource = bindingSource1.DataSource;
            mergeColumnSelection.DisplayMember = "ColumnName";
        }

        private void DisplayColumnName1(int index)
        {
            var bindingSource1 = new BindingSource();
            var columns = _listTables1[index].Columns;
            var listColumns = new List<DataColumn>();
            foreach (DataColumn column in columns)
                listColumns.Add(column);
            bindingSource1.DataSource = listColumns;
            checkedListBox1.DataSource = bindingSource1.DataSource;
            checkedListBox1.DisplayMember = "ColumnName";
            CheckColumnNames();                             
        }

        private void DisplayColumnName2(int index)
        {
            try
            {
                var bindingSource1 = new BindingSource();
                var columns = _listTables2[index].Columns;
                var listColumns = new List<DataColumn>();
                foreach (DataColumn column in columns)
                    listColumns.Add(column);
                bindingSource1.DataSource = listColumns;
                checkedListBox2.DataSource = bindingSource1.DataSource;
                checkedListBox2.DisplayMember = "ColumnName";
                CheckColumnNames();
            }
            catch (NullReferenceException)
            {
            }
        }

        private void DisplaySameColumn()
        {
            var bindingSource1 = new BindingSource();
            bindingSource1.DataSource = _duplicateColumns;
            sameNameSelection.DataSource = bindingSource1.DataSource;
            sameNameSelection.DisplayMember = "ColumnName";
        }

        private void DisplayKeyColumns()
        {
            var bindingSource1 = new BindingSource();
            bindingSource1.DataSource = _keyColumns;
            keyColumnList.DataSource = bindingSource1.DataSource;
        }

        private static DataTable MergeTables(DataTable table1, DataTable table2)
        {
            List<string> keyColumns = new List<string>();
            foreach(string column in _keyColumns)
            {
                keyColumns.Add(column);
            }
            DataTable copyDataTable1 = table1.Copy();           //check the contents of table1
            DataTable copyDataTable2 = table2.Copy();           //TABLE 2 should not be filled with all options in second file

            var key1 = new DataColumn[keyColumns.Count];
            var key2 = new DataColumn[keyColumns.Count];
            for (var i = 0; i < keyColumns.Count; i++)
            {
                key1[i] = copyDataTable1.Columns[keyColumns[i]];
                key2[i] = copyDataTable2.Columns[keyColumns[i]];
            }
            copyDataTable1.PrimaryKey = key1;
            copyDataTable2.PrimaryKey = key2;

    
            copyDataTable1.Merge(copyDataTable2);  //check if merge works correctly

            return copyDataTable1;
        }

        public static DataTable RemoveDuplicateRows(DataTable table, string DistinctColumn)
        {
            try
            {
                ArrayList UniqueRecords = new ArrayList();
                ArrayList DuplicateRecords = new ArrayList();
                foreach (DataRow dRow in table.Rows)
                {
                    if (UniqueRecords.Contains(dRow[DistinctColumn]))
                        DuplicateRecords.Add(dRow);
                    else
                        UniqueRecords.Add(dRow[DistinctColumn]);
                }

                foreach (DataRow dRow in DuplicateRecords)
                {
                    table.Rows.Remove(dRow);
                }

                return table;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public bool windowStatus()
        {
            return showWindow;
        }

        /// <summary>
        /// checks for matching column names between the first table and the second table. Displays this in drop down menu allowing for name changes
        /// for columns with the same name in the second table.
        /// </summary>
        private void CheckColumnNames()
        {
            try
            {
                if (tableSelection1.SelectedIndex != -1 && tableSelection2.SelectedIndex != -1)
                {
                    _duplicateColumns = new BindingList<DataColumn>();
                    var table1 = _listTables1[tableSelection1.SelectedIndex];
                    var table2 = _listTables2[tableSelection2.SelectedIndex];
                    foreach (DataColumn column in table2.Columns)
                    {
                        if (table1.Columns.Contains(column.ColumnName))
                        {
                                _duplicateColumns.Add(column);
                        }
                    }
                    DisplaySameColumn();            
                }
            }
            catch (NullReferenceException)
            {
            }
        }

        private void ChangeColumnName()
        {
            try
            {
                var oldName = _duplicateColumns[sameNameSelection.SelectedIndex].ColumnName;
                var table2 = _listTables2[tableSelection2.SelectedIndex];
                _listTables2[tableSelection2.SelectedIndex].Columns[oldName].ColumnName = newName.Text;
            }
            catch (IndexOutOfRangeException)
            {
                var confirmationDialog = MessageBox.Show("There is no column with this name.", "Error",
                    MessageBoxButtons.OK);
                if (confirmationDialog == DialogResult.OK)
                    return;
            }
        }

        private void fileSelection1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DisplayTableName1(_dataSheets1[fileSelection1.SelectedIndex]);
            ClearKeyColumns();
            UncheckKeepColumns();
        }

        private void fileSelection2_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisplayTableName2(_dataSheets1[fileSelection2.SelectedIndex]);
            ClearKeyColumns();
            UncheckKeepColumns();
        }

        private void selectAll1_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
                checkedListBox1.SetItemChecked(i, true);
        }

        private void selectAll2_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < checkedListBox2.Items.Count; i++)
                checkedListBox2.SetItemChecked(i, true);
        }

        private void unselectAll2_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < checkedListBox2.Items.Count; i++)
                if (!_keyColumns.Any(k => k.Equals(checkedListBox1.Items[i].ToString())))
                    checkedListBox2.SetItemChecked(i, false);
        }

        private void unselectAll1_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
                if (!_keyColumns.Any(k => k.Equals(checkedListBox1.Items[i].ToString())))
                    checkedListBox1.SetItemChecked(i, false);
        }

        private void tableSelection2_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisplayColumnName2(tableSelection2.SelectedIndex);
            ClearKeyColumns();
            UncheckKeepColumns();
        }

        private void tableSelection1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisplayColumnName1(tableSelection1.SelectedIndex);
            DisplayMergeColumn(tableSelection1.SelectedIndex);
            DisplayKeyColumns();
            ClearKeyColumns();
            UncheckKeepColumns();
        }

        private void changeName_Click(object sender, EventArgs e)
        {
            if (newName.Equals(null))
            {
                var dialog = MessageBox.Show("You must enter a valid column name", "Error");
                if (dialog == DialogResult.OK)
                    return;
            }
            ChangeColumnName();
            DisplayTableName1(_dataSheets1[fileSelection1.SelectedIndex]);
            DisplayTableName2(_dataSheets2[fileSelection2.SelectedIndex]);
            CheckColumnNames();
        }

        private void mergeAndSave_Click(object sender, EventArgs e)
        {
            if (_keyColumns.Count == 0)
            {
                var errorDialog =
                    MessageBox.Show("There are no key columns selected. Please select at least one column.", "Error");
                if (errorDialog == DialogResult.OK)
                    return;
            }
            try
            {
                var mergedTable = new DataTable();
                var table1 = _listTables1[tableSelection1.SelectedIndex];
                var table2 = _listTables2[tableSelection2.SelectedIndex];
                var table1Copy = _listTables1[tableSelection1.SelectedIndex]; //used for removal of nonselected columns
                var table2Copy = _listTables2[tableSelection2.SelectedIndex]; //to ensure correct indices for removal

                if (table1.Equals(table2))
                {
                    var errorDialog = MessageBox.Show("You cannot merge the same sheet.", "Error");
                    if (errorDialog == DialogResult.OK)
                        return;
                }

                if (_duplicateColumns.Count() > 0)
                {
                    var confirmationDialog =
                        MessageBox.Show(
                            "Key columns will be merged into one. When merging two rows with common keys, cells within the same column that are not key columns will only display data overriden by the second file. Rows with key column cells that do not have absolute matches will be created as independent rows. Would you like to continue?",
                            "Confirmation", MessageBoxButtons.YesNo);
                    if (confirmationDialog == DialogResult.No)
                        return;
                    else
                    {
                        for (var index = 0; index < table1Copy.Columns.Count; index++)
                            if (checkedListBox1.CheckedIndices.IndexOf(index) == -1)
                                table1.Columns.Remove(table1Copy.Columns[index].ColumnName);
                        for (var index = 0; index < table2Copy.Columns.Count; index++)
                            if (checkedListBox2.CheckedIndices.IndexOf(index) == -1)
                                table2.Columns.Remove(table2Copy.Columns[index].ColumnName); 
                        
                        mergedTable = MergeTables(table1, table2);                                 
                       
                        var dialog = new SaveFileDialog();
                        dialog.InitialDirectory = "C:\\Users";
                        dialog.Filter = "Excel|*.xlsx"; //"Excel Files|(*.xlsx)|*.xlsx"; //changed from xlsx to xls
                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            try
                            {
                                FileStream checkOpen = File.Open(dialog.FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                                checkOpen.Close();
                            }
                            catch(IOException)
                            {
                                var dialog1 = MessageBox.Show("A file with the given name is currently running. Please close the open file and try again to override.", "Error");
                                if (dialog1 == DialogResult.OK)
                                    return;
                            }
                            CreateExcelFile.CreateExcelDocument(mergedTable, dialog.FileName);
                            var confirmation = new Confirmation();
                            if (confirmation.ShowDialog(this) == DialogResult.OK)
                            {
                                Close();
                            }
                        }
                    }
                }
            }
            catch (ArgumentNullException)
            {
                var dialog = MessageBox.Show("The column you are merging must be a column in both sheets.",
                    "Error");
                if (dialog == DialogResult.OK)
                    return;
            }
            catch (DataException)
            {
                var dialog = MessageBox.Show("There are some cells that are empty in the key columns.", "Error");
                if (dialog == DialogResult.OK)
                    return;
            }
            catch (ArgumentException)
            {
                var dialog = MessageBox.Show("There is duplicate data in the key columns selected. Check to make sure there are no empty rows in the datasheet.", "Error");
                if (dialog == DialogResult.OK)
                    return;
            }
        }

        private void ClearKeyColumns()
        {
            _keyColumns.Clear();
        }

        private void UncheckKeepColumns()
        {
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
                if (!_keyColumns.Any(k => k.Equals(checkedListBox1.Items[i].ToString())))
                    checkedListBox1.SetItemChecked(i, false);
            for (var i = 0; i < checkedListBox2.Items.Count; i++)
                if (!_keyColumns.Any(k => k.Equals(checkedListBox1.Items[i].ToString())))
                    checkedListBox2.SetItemChecked(i, false);
        }

        private void addColumn_Click(object sender, EventArgs e)
        {
            var table1 = _listTables1[tableSelection1.SelectedIndex];
            var table2 = _listTables2[tableSelection2.SelectedIndex];
            if (_keyColumns.Any(i => i.Equals(table1.Columns[mergeColumnSelection.SelectedIndex].ColumnName)))
            {
                var dialog = MessageBox.Show("You cannot add the same column twice.", "Error");
                if (dialog == DialogResult.OK)
                    return;
            }
            else if(table2.Columns.IndexOf(table1.Columns[mergeColumnSelection.SelectedIndex].ColumnName) == -1)
            {
                var dialog = MessageBox.Show("This column does not exist in the 2nd table.", "Error");
                if (dialog == DialogResult.OK)
                    return;
            }
            else
            {
                _keyColumns.Add(table1.Columns[mergeColumnSelection.SelectedIndex].ColumnName);
                for (var i = 0; i < checkedListBox1.Items.Count; i++)
                    if (checkedListBox1.Items[i].ToString().Equals(table1.Columns[mergeColumnSelection.SelectedIndex].ColumnName))
                        checkedListBox1.SetItemChecked(i, true);
                for (var i = 0; i < checkedListBox2.Items.Count; i++)
                    if (checkedListBox2.Items[i].ToString().Equals(table1.Columns[mergeColumnSelection.SelectedIndex].ColumnName))
                        checkedListBox2.SetItemChecked(i, true);
            }
        }


        private void removeColumn_Click(object sender, EventArgs e)
        {
            if (keyColumnList.SelectedItem != null)
            {        
                for (var i = 0; i < checkedListBox1.Items.Count; i++)
                    if (checkedListBox1.Items[i].ToString().Equals(keyColumnList.SelectedItem.ToString()))          
                        checkedListBox1.SetItemChecked(i, false);
                for (var i = 0; i < checkedListBox2.Items.Count; i++)
                    if (checkedListBox2.Items[i].ToString().Equals(keyColumnList.SelectedItem.ToString()))
                        checkedListBox2.SetItemChecked(i, false);
                _keyColumns.Remove(keyColumnList.SelectedItem.ToString());  
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void keyColumnList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}