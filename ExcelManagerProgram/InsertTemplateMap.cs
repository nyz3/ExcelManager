using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using CSOpenXmlExcelToXml;
using ExcelDataReader;
using ExportToExcel;

namespace ExcelManager
{
    public partial class InsertTemplateMap : Form
    {
        private static DataSet chosenFiles; //selected file
        private static List<DataTable> fileSheets;  //list of sheets in source file
        private static DataTable selectedTable; //currently selected table of source
        private static string[] columnKeys; //represents the list of keys of columns names to retrieve column data from
        private static string[] columnValues; //represents the list of values of column names to insert columns into

        private static string templatePath; //string representing the pathname for the template file
        private static List<DataTable> templateCopy; //selected template file
        private static DataTable selectedTemplate;  //selected sheet within template
        private static List<DataColumn> templateColumns; //columns in selected template
        private bool validFiles;    //checks if one and only one file has been selected to allow form to open

        private static DataSet chosenFilesNoFormat; //same as variables above except for the non-formatted file
        private static List<DataTable> fileSheetsNoFormat;
        private static DataTable selectedTableNoFormat;

        private static List<DataTable> templateCopyNoFormat;//same as variables above except for the non-formatted file
        private static DataTable selectedTemplateNoFormat;

        /// <summary>
        /// Ensures that when form is opened, one and only one file is selected for insert. Columns in the given file
        /// </summary>
        /// <param name="data"></param>
        public InsertTemplateMap(List<DataSet> data, List<DataSet> dataNoFormat)
        {
            InitializeComponent();
            if (data.Count < 1)
            {
                validFiles = false;
                var dialog = MessageBox.Show("Please select a file to insert.", "Error");
            }
            else if (data.Count > 1)
            {
                validFiles = false;
                var dialog = MessageBox.Show("Please make sure only one file is selected for inserting at a time. Please use the \"Open File and Reset\" option to try again.", "Error");
            }
            else
            {
                validFiles = true;                                //display opened form because files have been correctly selected
                fileSheets = new List<DataTable>();               //reset all static variables to make sure everything is up to date
                selectedTable = new DataTable();
                columnKeys = null;
                columnValues = null;
                templateCopy = new List<DataTable>();
                selectedTemplate = new DataTable();
                templateColumns = new List<DataColumn>();
                chosenFiles = data[0]; //there should only be one file in data, therefore must be the file we are importing from

                chosenFilesNoFormat = dataNoFormat[0];
                fileSheetsNoFormat = new List<DataTable>();
                selectedTableNoFormat = new DataTable();

                templateCopyNoFormat = new List<DataTable>();
                selectedTemplateNoFormat = new DataTable();

                createSheets();
                displayFileSheets();
            }
        }

        /// <summary>
        /// Returns whether or not form is ready to be displayed.
        /// </summary>
        /// <returns></returns>
        public bool windowStatus()
        {
            return validFiles;
        }

        /// <summary>
        /// Creates a list of sheets from the file to be inserted, stored as a list of datatables.
        /// </summary>
        private void createSheets()
        {
            DataSet selectedFile = chosenFiles;
            foreach (DataTable sheet in selectedFile.Tables)
            {
                fileSheets.Add(sheet);
            }

            DataSet selectedFileNoFormat = chosenFilesNoFormat;
            foreach (DataTable sheet in selectedFileNoFormat.Tables)
            {
                fileSheetsNoFormat.Add(sheet);
            }
        }

        /// <summary>
        /// Displays the various sheets from the inserted file to allow user to select which sheet from file to insert from drop
        /// down menu.
        /// </summary>
        private void displayFileSheets()
        {
            var bindSheets = new BindingSource();
            bindSheets.DataSource = fileSheets;
            comboBox1.DataSource = bindSheets.DataSource;
            comboBox1.DisplayMember = "TableName";
        }

        /// <summary>
        /// Displays the various sheets from the template file to allow user to select which sheet from file to work as template from drop
        /// down menu.
        /// </summary>
        private void displayTemplateSheets()
        {
            var bindTempSheets = new BindingSource();
            bindTempSheets.DataSource = templateCopy;
            comboBox2.DataSource = bindTempSheets.DataSource;
            comboBox2.DisplayMember = "TableName";
        }

        /// <summary>
        /// Allows user to input text file that contains column mappings of keys and values.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            List<string> columnNamesText = new List<string>();  //list of each line of the text file with extra white space removed

            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.Filter = "Text|*.txt;";
                dialog.Multiselect = false; //can only select one text file with list of columns at a time
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    columnNamesText = System.IO.File.ReadAllLines(dialog.FileName).Where(s => s.Trim() != string.Empty).ToList();  //puts all lines from text file (column names) into string array
                    for (int x = 0; x < columnNamesText.Count; x++)
                    {
                        columnNamesText[x] = columnNamesText[x].Trim();
                    }
                    if (columnNamesText.Count != columnNamesText.Distinct().Count())  //checks text file for duplicate key>value columns
                    {
                        MessageBox.Show("Please make sure selected file does not have duplicate (Key, Value) column names. Try again.", "Error");
                        return;
                    }

                    columnKeys = new string[columnNamesText.Count];
                    columnValues = new string[columnNamesText.Count];
                    foreach (string crudeText in columnNamesText)
                    {
                        List<string> split = crudeText.Split('\t').ToList();
                        if (split.Count != 2)
                        {
                            MessageBox.Show("Please make sure each row in text file has a \"Key\" column name and a \"Value\" column name seperated by one and only one tab key.");
                            return;
                        }
                        else if (string.IsNullOrWhiteSpace(split[0]) || string.IsNullOrWhiteSpace(split[1]))
                        {
                            MessageBox.Show("Please make sure there are no \"Key\" and/or \"Value\" column names that are whitespace in the text file.", "Error");
                            return;
                        }
                        else
                        {
                            columnKeys[columnNamesText.IndexOf(crudeText)] = split[0].Trim();
                            if (!columnValues.Contains(split[1].Trim()))
                            {
                                columnValues[columnNamesText.IndexOf(crudeText)] = split[1].Trim();
                            }
                            else
                            {
                                MessageBox.Show("Please make sure there are no duplicate \"Value\" column names (cannot insert two source columns into one destination column.", "Error");
                                return;
                            }
                        }
                    }
                    textBox2.Text = Path.GetFileNameWithoutExtension(dialog.FileName).Replace(Path.GetDirectoryName(dialog.FileName), String.Empty);            //displays file name in read-only box, if name displayed, means keys/value arrays are good to go  
                    displayTextColumns();                       //displays columns from textfile into read-only box          
                }
            }
            catch
            {
                //this should never happen, only scenario would be if text file was somehow corrupted. will come back and check edge cases later.
            }
        }

        /// <summary>
        /// Display column mappings in the format (key ---> value) in the read-only box.
        /// </summary>
        private void displayTextColumns()
        {
            listBox1.Items.Clear();
            for (int x = 0; x < columnKeys.Length; x++)
            {
                string mapping = columnKeys[x] + "   --------->   " + columnValues[x];
                listBox1.Items.Add(mapping);
            }
        }

        /// <summary>
        /// Updates which table/sheet we are working with as the source file.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedTable = fileSheets[comboBox1.SelectedIndex];
        }

        /// <summary>
        /// Open File interface for the template to insert data into.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.Filter = "Excel|*.xlsx; *.xls"; //"Excel Files|(*.xlsx, *.xls)|*.xlsx;*.xls";
                dialog.Multiselect = false;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    templatePath = dialog.FileName;
                    using (var stream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        // Auto-detect format, supports:
                        //  - Binary Excel files (2.0-2003 format; *.xls)
                        //  - OpenXml Excel files (2007 format; *.xlsx)
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
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

                            var resultNoFormat = reader.AsDataSet(); //read dataset without configuration of column names

                            //copy the table to add to the DataSet otherwise only a reference will be copied. 
                            templateCopy = new List<DataTable>(); //refresh templatecopy everytime new template opened
                            templateCopyNoFormat = new List<DataTable>(); //refresh nonformatted templatecopy
                            selectedTemplate = new DataTable(); //refresh selectedtemplate everytime new template opened
                            selectedTemplateNoFormat = new DataTable(); //refresh nonformatted selectedtemplate
                            templateColumns = new List<DataColumn>(); //refresh columns from selected template when new template opened

                            foreach (DataTable table in result.Tables)
                            {
                                DataTable datatable = table.Copy();
                                templateCopy.Add(datatable);
                            }

                            foreach (DataTable table in resultNoFormat.Tables) //nonformatted version of above
                            {
                                DataTable datatable = table.Copy();
                                templateCopyNoFormat.Add(datatable);
                            }
                        }
                    }
                    textBox1.Text = Path.GetFileNameWithoutExtension(dialog.FileName).Replace(Path.GetDirectoryName(dialog.FileName), String.Empty);         //displays name of template file
                    displayTemplateSheets();                //displays list of sheets within template file
                }
            }
            catch (IOException)
            {
                var dialog = MessageBox.Show("The file you are trying to open is being used by another program. Please close the program and try again.", "Error");
            }
        }

        /// <summary>
        /// Displays a read-only list of the unique column names within the template selected.
        /// </summary>
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            {
                templateColumns = new List<DataColumn>(); //refreshes list to make sure old columns are erased when sheet changes
                int index = comboBox2.SelectedIndex;
                selectedTemplate = templateCopy[index];
                selectedTemplateNoFormat = templateCopyNoFormat[index]; //nonformatted template table
                List<string> tempColNamesNoFormat = new List<string>();

                for (int x = 0; x < selectedTemplateNoFormat.Columns.Count; x++)       //read first row (column names) of nonformatted sheet
                {
                    if (!String.IsNullOrWhiteSpace(selectedTemplateNoFormat.Rows[0][x].ToString().Trim()))
                    {
                        tempColNamesNoFormat.Add(selectedTemplateNoFormat.Rows[0][x].ToString().Trim());
                    }
                }

                foreach (DataColumn col in selectedTemplate.Columns)
                {
                    if (tempColNamesNoFormat.Contains(col.ColumnName.Trim()))
                    {
                        templateColumns.Add(col);
                    }
                }

                var bindTemplateTable = new BindingSource();
                bindTemplateTable.DataSource = templateColumns;
                listBox2.DataSource = bindTemplateTable.DataSource;
                listBox2.DisplayMember = "ColumnName";
            }
        }

        /// <summary>
        /// This checks all conditions for insertion and does the actual insertion process. Look at comments for step-by-step information.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0) //if there are no columns to insert (from text file or selected additional columns)
            {
                MessageBox.Show("There are no columns to insert. Please import a text file with column mappings to insert.", "Error");
                return;
            }

            if (String.IsNullOrWhiteSpace(textBox1.Text))           //check if template has been selected
            {
                if (listBox2.Items.Count == 0)                      //check if template sheet has columns to insert into
                {
                    MessageBox.Show("Please select a template sheet that has columns to insert data into.", "Error");
                    return;
                }
                else
                {
                    MessageBox.Show("Please select a valid template file.", "Error");
                    return;
                }
            }

            List<string> uniqueColNames = new List<string>();           //check if source file has duplicate column names
            for (int x = 0; x < selectedTableNoFormat.Columns.Count; x++)
            {
                string temp = selectedTableNoFormat.Rows[0][x].ToString().Trim();
                if (!string.IsNullOrWhiteSpace(temp))
                {
                    if (uniqueColNames.Contains(temp))
                    {
                        MessageBox.Show("Duplicate column names were found in source file sheet. Please make sure there are no duplicates.", "Error");
                        return;
                    }
                    else
                    {
                        uniqueColNames.Add(temp);
                    }
                }
            }

            uniqueColNames = new List<string>();
            for (int x = 0; x < selectedTemplateNoFormat.Columns.Count; x++)
            {
                string temp = selectedTemplateNoFormat.Rows[0][x].ToString().Trim();
                if (!string.IsNullOrWhiteSpace(temp))
                {
                    if (uniqueColNames.Contains(temp))
                    {
                        MessageBox.Show("Duplicate column names were found in template sheet. Please make sure there are no duplicates.", "Error");
                        return;
                    }
                    else
                    {
                        uniqueColNames.Add(temp);
                    }
                }
            }

            DataTable selectedTableCopy = selectedTable.Copy();                            //use copies so original does not get modified when deleting rows
            DataTable selectedTemplateCopy = selectedTemplate.Copy();
            List<string> tempColNamesNoFormat = new List<string>();

            foreach (DataColumn datum in selectedTableCopy.Columns)                         //trim excess off of column names of source
            {
                selectedTableCopy.Columns[datum.ColumnName].ColumnName = datum.ColumnName.Trim();
            }

            foreach (DataColumn datum in selectedTemplateCopy.Columns)                         //trim excess off of column names of template
            {
                selectedTemplateCopy.Columns[datum.ColumnName].ColumnName = datum.ColumnName.Trim();
            }

            for (int x = 0; x < selectedTemplateNoFormat.Columns.Count; x++)                //stores all the column names without formatting (gets rid of default)
            {
                if (!string.IsNullOrWhiteSpace(selectedTemplateNoFormat.Rows[0][x].ToString().Trim()))
                {
                    tempColNamesNoFormat.Add(selectedTemplateNoFormat.Rows[0][x].ToString().Trim());
                }
            }

            for (int x = selectedTemplateCopy.Columns.Count - 1; x >= 0; x--)                  //remove all default template columns
            {
                selectedTemplateCopy.Columns[x].ColumnName = selectedTemplateCopy.Columns[x].ColumnName.Trim();
                if (!tempColNamesNoFormat.Contains(selectedTemplateCopy.Columns[x].ColumnName))
                {
                    selectedTemplateCopy.Columns.Remove(selectedTemplateCopy.Columns[x]);
                }
            }

            bool colExistSource = true;                                             //make sure all inserted columns exist in source file (bc text file)
            foreach (string colName in columnKeys)
            {
                if (!selectedTableCopy.Columns.Contains(colName))
                {
                    colExistSource = false;
                }
            }
            if (!colExistSource)
            {
                MessageBox.Show("Please make sure all key columns from text file to insert exist in the source file.", "Error");
                return;
            }

            List<string> templateColNames = new List<string>();
            foreach (DataColumn tempCol in templateColumns)
            {
                templateColNames.Add(tempCol.ColumnName.Trim());
            }

            bool match = true;                                                      //true if all columns found in template, false otherwise
            foreach (string colName in columnValues)
            {
                if (!templateColNames.Contains(colName))
                {
                    match = false;
                }
            }
            if (!match)
            {
                MessageBox.Show("Please make sure all columns to be inserted are columns that exist in template.", "Error");
                return;
            }

            //now, all conditions have been checked, start process of actually inserting
            //move stuff from "selectedTable" based on column names from  "columnInsert" into template table called "selectedTemplate" with matching column
            //names as in "columnInsert"

            bool foundContent = false;
            for (int i = 0; i < selectedTableCopy.Rows.Count; i++)                     //delete all rows with any cell containing "NotDelete" as value
            {
                DataRow row = selectedTableCopy.Rows[i];
                if (!row.ItemArray.Any(o => o.ToString().Trim().Equals("NotDelete")))           //if row has whitespace or content
                {
                    foundContent = true;
                }
                else                                                                            //if row contains NotDelete
                {
                    if (foundContent)
                    {
                        MessageBox.Show("Please make sure the source file has no extra rows before the last row with any cell designated \"NotDelete\".");
                        return;
                    }
                    else
                    {
                        selectedTableCopy.Rows.Remove(selectedTableCopy.Rows[i]);               //remove all leading rows with "NotDelete" 
                        i--;                                                                    //account for removed row
                    }
                }
            }

            for (int i = selectedTableCopy.Rows.Count - 1; i >= 0; i--)                        //delete all empty rows from end of source file       
            {
                DataRow row = selectedTableCopy.Rows[i];
                if (row.ItemArray.All(o => string.IsNullOrWhiteSpace(o.ToString())))
                {
                    selectedTableCopy.Rows.Remove(selectedTableCopy.Rows[i]);
                }
            }

            for (int i = selectedTemplateCopy.Rows.Count - 1; i >= 0; i--)                     //delete all empty rows from end of template file         
            {
                DataRow row = selectedTemplateCopy.Rows[i];
                if (row.ItemArray.All(o => string.IsNullOrWhiteSpace(o.ToString())))
                {
                    selectedTemplateCopy.Rows.Remove(selectedTemplateCopy.Rows[i]);
                }
                else
                {
                    break;                                                                  //break out of loop when encountering first row with content
                }
            }

            bool emptyTemplate = true;                                          //true if template is "empty" and ready for insert, false otherwise
            for (int i = 0; i < selectedTemplateCopy.Rows.Count; i++)
            {
                DataRow row = selectedTemplateCopy.Rows[i];
                if (!row.ItemArray.Any(o => o.ToString().Trim().Equals("NotDelete")))
                {
                    emptyTemplate = false;
                }
            }
            if (!emptyTemplate)
            {
                MessageBox.Show("Please make sure the template you are inserting into is empty (except for rows containing cells designated \"NotDelete\") and does not have empty rows between designated \"NotDelete\" rows.");
                return;
            }

            //At this point, our sourcefile only contains data under each column name we actually want to add.
            //At this point, our templatefile is guaranteed to be empty and ready for things to be inserted.

            int rowTemplateInsert = selectedTemplateCopy.Rows.Count;                  //first row to insert into from the template

            for (int x = 0; x < selectedTableCopy.Rows.Count; x++)
            {
                selectedTemplateCopy.Rows.Add(selectedTemplateCopy.NewRow()); //add empty rows to allow for data insert
            }

            for (int z = 0; z < columnKeys.Length; z++)            //col is already trimmed   
            {
                string colKey = columnKeys[z];
                string colVal = columnValues[z];
                List<string> dataInsert = new List<string>();

                if (selectedTableCopy.Columns[colKey] != null)
                {
                    int columnIndex = selectedTableCopy.Columns[colKey].Ordinal;
                    for (int x = 0; x < selectedTableCopy.Rows.Count; x++)
                    {
                        dataInsert.Add(selectedTableCopy.Rows[x][columnIndex].ToString().Trim()); //copy all data from a given column into a stringlist
                    }
                }

                if (selectedTemplateCopy.Columns[colVal] != null)
                {
                    int templateColIndex = selectedTemplateCopy.Columns[colVal].Ordinal;
                    for (int y = 0; y < dataInsert.Count; y++)
                    {
                        selectedTemplateCopy.Rows[y + rowTemplateInsert][templateColIndex] = dataInsert[y];   //copy over cell data into template
                    }
                }
            }

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
                catch (IOException)
                {
                    MessageBox.Show("The file to override is currently running. Please close the open file and try again to save.", "Error");
                    return;
                }

                File.Copy(templatePath, dialog.FileName, true);           //creates original template copy with all formatting into FileName

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbooks workBooks = excelApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workBook = workBooks.Open(dialog.FileName);
                Microsoft.Office.Interop.Excel.Worksheet workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets[comboBox2.SelectedItem.ToString()];

                for (int x = 0; x < selectedTemplateCopy.Rows.Count; x++)
                {
                    for (int y = 0; y < selectedTemplateCopy.Columns.Count; y++)
                    {
                        if (!string.IsNullOrWhiteSpace(selectedTemplateCopy.Rows[x][y].ToString()))
                        {
                            workSheet.Cells[x + 2, y + 1] = selectedTemplateCopy.Rows[x][y].ToString();
                        }
                    }
                }

                excelApp.DisplayAlerts = false;
                workBook.SaveAs(dialog.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                    Type.Missing, Type.Missing);

                workBook.Close();
                workBooks.Close();

                var confirmation = new Confirmation();
                if (confirmation.ShowDialog(this) == DialogResult.OK)
                {
                    Close();
                }
            }
        }
    }
}
