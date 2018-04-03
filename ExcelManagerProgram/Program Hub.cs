using CSOpenXmlExcelToXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using ExportToExcel;

namespace ExcelManager
{
    public partial class Form1 : Form
    {
        //Generates a new list<DataSet> to store all files for function use
        static List<DataSet> dataSheets = new List<DataSet>();
        //Generates a new list<DataSet> to store all files for function use, but no formatting of first row into columns
        static List<DataSet> dataSheetsNoFormat = new List<DataSet>();
        //bindingData is used for displaying information in the comboboxs, important to create a bindingSource in order for the 
        //comboboxes to update accordingly
        static BindingList<DataSet> bindingData = new BindingList<DataSet>();
        //static DataSet dataTables = new DataSet();
        public Form1()
        {
            InitializeComponent();
            this.SendToBack();
        }

        //used to display data on the gridview, the int index that should be entered is the index of the table selected in the sheet
        private void bindData(DataSet dataTables, int index)
        {
            BindingSource bindingSource1 = new BindingSource();
            bindingSource1.DataSource = dataTables.Tables[index];
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = bindingSource1;
            dataGridView1.AutoResizeColumns();
        }

        //binds the data so that the grid updates whenever the sheet selected is changed

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            bindData(dataSheets[comboBox2.SelectedIndex], comboBox1.SelectedIndex);
        }

        /// <summary>
        /// This is the "Open..." button press
        /// Checks to make sure that the file has opened before displaying the data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openExcelFile(false) == true)
            {
                displaySheet();
                displayTable(comboBox2.SelectedIndex);
                bindData(dataSheets[comboBox2.SelectedIndex], comboBox1.SelectedIndex);
            }

        }

        /// <summary>
        /// this displays the table in the comboBox for table selection 
        /// </summary>
        /// <param name="index"></param>
        public void displayTable(int index)
        {
            BindingSource bindingSource1 = new BindingSource();
            DataTableCollection tables = dataSheets[index].Tables;
            List<DataTable> listTables = new List<DataTable>();
            foreach (DataTable table in tables)
            {
                listTables.Add(table);
            }
            bindingSource1.DataSource = listTables;
            comboBox1.DataSource = bindingSource1.DataSource;
            comboBox1.DisplayMember = "TableName";
        }

        /// <summary>
        /// Method to display the sheet(file) on the comboBox for sheet selection
        /// </summary>
        public void displaySheet()
        {
            bindingData.ResetBindings();
            bindingData = new BindingList<DataSet>();
            foreach (DataSet data in dataSheets)
            {
                bindingData.Add(data);
            }
            BindingSource bindingSource1 = new BindingSource();
            bindingSource1.DataSource = bindingData;
            comboBox2.DataSource = bindingSource1.DataSource;
            comboBox2.DisplayMember = "DataSetName";
        }

        /// <summary>
        /// Boolean parameter is used to reset the list of datasets
        /// Creates a new DataSet in and adds it to the List<DataSet> 
        /// </summary>
        /// <param name="reset"></param>
        /// <returns></returns>
        private Boolean openExcelFile(Boolean reset)
        {
            if (reset == true)
            {
                dataSheets = new List<DataSet>();
                dataSheetsNoFormat = new List<DataSet>();
            }
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.InitialDirectory = "C:\\Users\\QingshengQuinn\\Desktop\\Nathan's CS Projects\\Projects\\WinterUTHealth\\TestFiles";
                dialog.Filter = "Excel|*.xlsx; *.xls"; //"Excel Files|(*.xlsx, *.xls)|*.xlsx;*.xls";
                dialog.Multiselect = true;
                bool errorDisplayed = false; //prevents errordialog from opening up multiple times if error occurs with several files

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (String FileName in dialog.FileNames)
                    {
                        if (errorDisplayed == true)
                        {
                            break;
                        }
                        using (var stream = File.Open(FileName, FileMode.Open, FileAccess.Read))
                        {
                            var matches = dataSheets.Where(p => p.DataSetName == Path.GetFileNameWithoutExtension(FileName));
                            if (matches.Any())
                            {
                                DialogResult errorDialog = MessageBox.Show("You cannot open any of the same files twice. Try again with unique files only.", "Error");
                                if (errorDialog == DialogResult.OK)
                                {
                                    errorDisplayed = true;
                                    return false;
                                }
                            }
                            // Auto-detect format, supports:
                            //  - Binary Excel files (2.0-2003 format; *.xls)
                            //  - OpenXml Excel files (2007 format; *.xlsx)
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet dataTables = new DataSet();
                                DataSet dataTablesNoFormat = new DataSet();
                                dataTables.DataSetName = Path.GetFileNameWithoutExtension(FileName);
                                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    //Changes the tableReader to make the first row of the excel sheet as column names 
                                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true,
                                        ReadHeaderRow = (rowReader) =>
                                        {
                                        }
                                    }
                                });

                                //read DataSet in without additional formatting
                                var resultNoFormat = reader.AsDataSet();

                                //copy the table to add to the DataSet otherwise only a reference will be copied. 
                                foreach (DataTable table in result.Tables)
                                {
                                    for (int i = 0; i < table.Rows.Count; i++)
                                    {
                                        DataRow row = table.Rows[i];
                                        if (row.ItemArray.All(o => string.IsNullOrEmpty(o.ToString())))
                                        {
                                            table.Rows.Remove(table.Rows[i]);
                                            i--;
                                        }
                                    }
                                    DataTable datatable = table.Copy();
                                    dataTables.Tables.Add(datatable);
                                }

                                foreach (DataTable table in resultNoFormat.Tables)
                                {
                                    for (int i = 0; i < table.Rows.Count; i++)
                                    {
                                        DataRow row = table.Rows[i];
                                        if (row.ItemArray.All(o => string.IsNullOrEmpty(o.ToString())))
                                        {
                                            table.Rows.Remove(table.Rows[i]);
                                            i--;
                                        }
                                    }
                                    DataTable datatable = table.Copy();
                                    dataTablesNoFormat.Tables.Add(datatable);
                                }

                                dataSheets.Add(dataTables);
                                dataSheetsNoFormat.Add(dataTablesNoFormat);
                            }
                        }
                    }
                    return true;
                }
            }
            catch (IOException)
            {
                var dialog = MessageBox.Show("The file you are trying to open is being used by another program. Please close the program and try again.", "Error");
                if (dialog == DialogResult.Yes)
                    return false;
            }
            return false;
        }



        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.Filter = "Excel|*.xlsx"; //"Excel Files|(*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (comboBox2.SelectedIndex >= 0 && comboBox2.SelectedIndex < dataSheets.Count)         //checks if there is content to save
                {
                    CreateExcelFile.CreateExcelDocument(dataSheets[comboBox2.SelectedIndex], dialog.FileName);
                }
                else
                {
                    MessageBox.Show("There is no content to save. Please provide content.", "Error");
                }

            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void mergeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<DataSet> oldSheets = dataSheets;
            MergeCells frm = new MergeCells(dataSheets);
            if (frm.windowStatus())
            {
                frm.ShowDialog(this); //locks parent window when child window is open
            }
            else
            {
                MessageBox.Show("There are no files to merge. Please return and select files.", "Error");
            }
            dataSheets = oldSheets;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            displayTable(comboBox2.SelectedIndex);
        }

        private void tesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void checkListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataSet dataSheet in dataSheets)
            {
                Console.WriteLine(dataSheet.DataSetName);
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void findDuplicateDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<DataSet> oldSheets = dataSheets;
            DuplicateCells form = new DuplicateCells(dataSheets);
            if (form.windowStatus())
            {
                form.ShowDialog(this);  //locks parent when child is open
            }
            else
            {
                MessageBox.Show("Please go back and select at least one file to search for duplicates in.", "Error");
            }
            dataSheets = oldSheets;
        }

        private void openFileAndResetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openExcelFile(true) == true)
            {
                displaySheet();
                displayTable(comboBox2.SelectedIndex);
                bindData(dataSheets[comboBox2.SelectedIndex], comboBox1.SelectedIndex);
            }
        }

        /// <summary>
        /// Should open up new window for the insert template functionality if exactly one file has been opened in the main window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void insertColumnDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<DataSet> oldSheets = dataSheets;
            List<DataSet> oldSheetsNoFormat = dataSheetsNoFormat;
            InsertTemplate form = new InsertTemplate(dataSheets, dataSheetsNoFormat);
            if (form.windowStatus())
            {
                form.ShowDialog(this);
            }
            dataSheets = oldSheets;
            dataSheetsNoFormat = oldSheetsNoFormat;
        }

        /// <summary>
        /// Should open up window for the map to template functionality if at least one file has been opened in the main window.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mapColumnsIntoTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<DataSet> oldSheets = dataSheets;
            List<DataSet> oldSheetsNoFormat = dataSheetsNoFormat;
            InsertTemplateMap form = new InsertTemplateMap(dataSheets, dataSheetsNoFormat);
            if (form.windowStatus())
            {
                form.ShowDialog(this);
            }
            dataSheets = oldSheets;
            dataSheetsNoFormat = oldSheetsNoFormat;
        }
    }
}
