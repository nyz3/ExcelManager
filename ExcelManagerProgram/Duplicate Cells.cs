using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelManager
{

    public partial class DuplicateCells : Form
    {
        private static List<DataSet> _dataSheets1 = new List<DataSet>();
        private static List<DataTable> _listTables1 = new List<DataTable>();
        private static List<DataColumn> _keyColumns = new List<DataColumn>();
        private static DataTable duplicateRows = new DataTable();
        private static bool showWindow = false;             //boolean used to aid in error processing, closes window that is defunct due to error

        public DuplicateCells(List<DataSet> data)
        {
            InitializeComponent();
            _dataSheets1 = new List<DataSet>();
            foreach (var dataSheet in data)
            {
                _dataSheets1.Add(dataSheet);
            }
            DisplayDataSetName1();

            if (fileSelection1.SelectedIndex >= 0 && fileSelection1.SelectedIndex < _dataSheets1.Count)     //prevents exceptions when find duplicates is 
            {                                                                                               //called when no file is selected        
                DisplayTableName1(_dataSheets1[fileSelection1.SelectedIndex]);
                showWindow = true;                                                                          //window has valid files to dupe, keep window open
            }
            else
            {
                showWindow = false;                                                                         //window does not have valid files to dupe, do not open window
            }
        }

        private void DisplayDataSetName1()
        {
            var bindingSource1 = new BindingSource();
            bindingSource1.DataSource = _dataSheets1;
            fileSelection1.DataSource = bindingSource1.DataSource;
            fileSelection1.DisplayMember = "DataSetName";
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
        }

        private void bindData(DataTable dataTables)
        {
            BindingSource bindingSource1 = new BindingSource();
            bindingSource1.DataSource = dataTables;
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = bindingSource1;
            dataGridView1.AutoResizeColumns();
        }

        private bool FindDuplicates()         //true if duplicates exist, else false, used for error handling/user friendly GUI
        {
            DataTable dt = _listTables1[tableSelection1.SelectedIndex]; 
            foreach (int i in checkedListBox1.SelectedIndices)
            {
                var duplicates = dt.Copy().AsEnumerable().GroupBy(r => r[i]).Where(gr => gr.Count() > 1).Select(g => g.Key).ToList();
                if (duplicates.Count != 0)
                {
                    foreach (double data in duplicates)
                    {
                        duplicateRows.Rows.Add(dt.Rows[(int)data]);
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        public bool windowStatus()
        {
            return showWindow;
        }

        private void Duplicate_Cells_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var table1 = _listTables1[tableSelection1.SelectedIndex];
            if(e.CurrentValue == CheckState.Unchecked)
            {
                _keyColumns.Add(table1.Columns[e.Index]); 
            }
            else
            {
                _keyColumns.Remove(table1.Columns[e.Index]); 
            }
        }

        private void selectAll1_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
                checkedListBox1.SetItemChecked(i, true);
        }

        private void unselectAll1_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
                    checkedListBox1.SetItemChecked(i, false);
        }


        private void UncheckKeepColumns()
        {
            for (var i = 0; i < checkedListBox1.Items.Count; i++)
                    checkedListBox1.SetItemChecked(i, false);

        }

        private void tableSelection1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisplayColumnName1(tableSelection1.SelectedIndex);
            UncheckKeepColumns();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!FindDuplicates())
            {
                MessageBox.Show("No duplicates were found.");
            }
            bindData(duplicateRows);
        }

        private void fileSelection1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisplayTableName1(_dataSheets1[fileSelection1.SelectedIndex]);
            UncheckKeepColumns();
        }
    }
}
