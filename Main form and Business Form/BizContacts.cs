using Microsoft.Office.Interop.Excel; //lets us make excel obj
using Microsoft.Office.Interop.Word; //allows us to make word objs
using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics; //allows us to pen up excel from code
using System.IO; //needed for file use
using System.Windows.Forms;

namespace Main_form_and_Business_Form
{
    public partial class BizContacts : Form
    {
        //conntection string. must have @ to make a literal string
        //you can get this from the properties of the db
        string connString = @"Data Source=DESKTOP-3BDGT3N;Initial Catalog=AddressBook;Integrated Security=True";

        //allows us to build the connection betweeen the program and the db
        SqlDataAdapter dataAdapter;

        //table to hold into to fill our datagrid view
        System.Data.DataTable table;


        //var to hold sql connection
        SqlConnection conn;


        string selectionStatement = "Select * from BizContacts";

        public BizContacts()
        {
            InitializeComponent();
        }

        private void BizContacts_Load(object sender, EventArgs e)
        {
            //set default of combo box to 1st item there
            cboSearch.SelectedIndex = 0;

            //sets source of data to the table through our bindingsource
            dataGridView1.DataSource = bindingSource1;

            //sql query to get data
            GetData(selectionStatement);
        }

        private void GetData(string selectCommand)
        {
            // b/c there could be a db connection error. always put a try/catch 
            try
            {
                //put in sql query and connstring to access db table
                dataAdapter = new SqlDataAdapter(selectCommand, connString);

                //create data table obj
                table = new System.Data.DataTable();

                //data and language won't vary depending on location
                table.Locale = System.Globalization.CultureInfo.InvariantCulture;

                //fills the table w/ data
                dataAdapter.Fill(table);

                //set data source to binding source 
                bindingSource1.DataSource = table;

                //prevents first column with IDs from being edited
                dataGridView1.Columns[0].ReadOnly = true;

            }
            catch (SqlException ex)
            {

                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //new sql command obj
            SqlCommand command;

            //field neames in the table followed by paramter names in values
            string insert = @"insert into BizContacts(Date_Added, Company, Website, Title, First_Name, Last_Name, Address,
                                                     City, State, Postal_Code, Email, Mobile, Notes, Image)
                                                                        
                                                     values(@Date_Added, @Company, @Website, @Title, @First_Name, @Last_Name, @Address,
                                                     @City, @State, @Postal_Code, @Email, @Mobile, @Notes, @Image)";

            //opens & closes connection to sql server
            using (conn = new SqlConnection(connString))
            {

                try
                {

                    //start connnection
                    conn.Open();

                    //sql command is the insert statement and the connection string
                    command = new SqlCommand(insert, conn);


                    //set Data_Added var
                    command.Parameters.AddWithValue(@"Date_Added", dateTimePicker1.Value);
                    command.Parameters.AddWithValue(@"Company", txtCompany.Text);
                    command.Parameters.AddWithValue(@"Website", txtWebSite.Text);
                    command.Parameters.AddWithValue(@"Title", txtTitle.Text);
                    command.Parameters.AddWithValue(@"First_Name", txtFirstName.Text);
                    command.Parameters.AddWithValue(@"Last_Name", txtLName.Text);
                    command.Parameters.AddWithValue(@"Address", txtAddress.Text);
                    command.Parameters.AddWithValue(@"City", txtCity.Text);
                    command.Parameters.AddWithValue(@"State", txtState.Text);
                    command.Parameters.AddWithValue(@"Postal_Code", txtZip.Text);
                    command.Parameters.AddWithValue(@"Mobile", txtMobile.Text);
                    command.Parameters.AddWithValue(@"Notes", txtNotes.Text);
                    command.Parameters.AddWithValue(@"Email", txtEmail.Text);

                    if (dlgOpenImage.FileName != "")
                    {
                        //convert image into bytes for saving in sql server
                        command.Parameters.AddWithValue(@"Image", File.ReadAllBytes(dlgOpenImage.FileName));
                    }
                    else
                    {
                        //save null to db
                        command.Parameters.Add("@Image", SqlDbType.VarBinary).Value = DBNull.Value;
                    }



                    //this pushes data into the table
                    command.ExecuteNonQuery();

                    dateTimePicker1.Value = DateTime.Today;
                    txtCompany.Text = "";
                    txtTitle.Text = "";
                    txtWebSite.Text = "";
                    txtTitle.Text = "";
                    txtFirstName.Text = "";
                    txtLName.Text = "";
                    txtAddress.Text = "";
                    txtCity.Text = "";
                    txtState.Text = "";
                    txtZip.Text = "";
                    txtMobile.Text = "";
                    txtNotes.Text = "";
                    txtEmail.Text = "";
                    pictureBox1.Image = null;
                    pictureBox1.Update();


                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }

            //query the table to see the new row
            GetData(selectionStatement);

            //updates datagridview
            dataGridView1.Update();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //creates command builder obj
            ////declare a new sql command            
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter);

            //sql update command
            dataAdapter.UpdateCommand = commandBuilder.GetUpdateCommand();

            try
            {
                //updates table in memory
                bindingSource1.EndEdit();

                //updates the data base
                dataAdapter.Update(table);

                MessageBox.Show("Update Sucessful");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            //get reference from current row
            DataGridViewRow row = dataGridView1.CurrentCell.OwningRow;

            //get id from id field selected record
            string value = row.Cells["ID"].Value.ToString();
            string fname = row.Cells["First_Name"].Value.ToString();
            string lname = row.Cells["Last_Name"].Value.ToString();

            //message box gives you yes/no option and has a ? icon
            DialogResult result = MessageBox.Show($"Do want to delete {fname} {lname} record numer {value}?", "Message", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            string deleteStatement = $"DELETE BizContacts where ID = {value};";

            //if user confirms he want to delete row
            if (result == DialogResult.Yes)
            {
                //connect to db
                using (conn = new SqlConnection(connString))
                {
                    try
                    {
                        //open sql connection
                        conn.Open();

                        SqlCommand comm = new SqlCommand(deleteStatement, conn);

                        //runs the query
                        comm.ExecuteNonQuery();

                        //query the table to see the new row
                        GetData(selectionStatement);

                        //updates datagridview
                        dataGridView1.Update();


                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }

            }

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string columnSearch = string.Empty;

            switch (cboSearch.SelectedIndex.ToString())
            {
                case "0":
                    columnSearch = "First_Name";
                    break;

                case "1":
                    columnSearch = "Last_Name";
                    break;

                default:
                    columnSearch = "Company";
                    break;

            }



            string searchTerm = txtSearch.Text.ToString().Trim();

            if (txtSearch.Text.ToString() == string.Empty)
            {
                MessageBox.Show("Missing Search Term!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string searchQuery = $"SELECT * FROM BizContacts where {columnSearch} like '%{searchTerm}%';";



                using (conn = new SqlConnection(connString))
                {

                    try
                    {

                        conn.Open();

                        GetData(searchQuery);

                        dataGridView1.Update();

                        txtSearch.Text = "";

                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }


            }


        }

        private void btnGetImage_Click(object sender, EventArgs e)
        {
            //open dialog box for finding image and if image is actually choosen
            if (dlgOpenImage.ShowDialog() == DialogResult.OK)
            {
                //load image when choosen from file dialog
                pictureBox1.Load(dlgOpenImage.FileName);

            }



        }

        private void pictureBox1_DoubleClick(object sender, EventArgs e)
        {
            //make a new form
            Form frm = new Form();

            //set background image to image clicked
            frm.BackgroundImage = pictureBox1.Image;

            //set size of form to size of the image
            frm.Size = pictureBox1.Image.Size;

            //show form
            frm.Show();
        }

        private void btnExportOpen_Click(object sender, EventArgs e)
        {
            //creates excel obj
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();

            //creates a work book
            _Workbook workbook = excel.Workbooks.Add(Type.Missing);

            //make a worksheet and sets it to null
            _Worksheet worksheet = null;

            try
            {
                //set active sheet
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Business Contacts";

                // b/c both data grids and excel work sheets are tabular, we must use nexted loops to write from one to another
                //controls the row number
                for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count - 1; rowIndex++)
                {
                    //used to go over columns of each row
                    for (int colIndex = 0; colIndex < dataGridView1.Columns.Count; colIndex++)
                    {
                        if (rowIndex == 0)
                        {
                            //in excel rows row and column start w/ 1,1 not 0,0
                            //write out header text from grid view to excel sheet
                            worksheet.Cells[rowIndex + 1, colIndex + 1] = dataGridView1.Columns[colIndex].HeaderText;
                        }
                        else
                        {
                            //fix row idex at 1 then change column index over possible values from 0 to 5
                            worksheet.Cells[rowIndex + 1, colIndex + 1] = dataGridView1.Rows[rowIndex].Cells[colIndex].Value.ToString();
                        }
                    }

                }

                //user clicks okay to save
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //save file to drive
                    worksheet.SaveAs(saveFileDialog1.FileName);

                    Process.Start("excel.exe", saveFileDialog1.FileName);
                }
            }


            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //close excel
                excel.Quit();

                //empty work book
                workbook = null;

                excel = null;
            }



        }

        private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void btnSaveToText_Click(object sender, EventArgs e)
        {
            //check if someone clicked on ok button
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                {
                    //grap each row in gridview
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        //once you have a row grabbed, go through cells
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            sw.Write(cell.Value); //write value to text file for each row
                        }

                        sw.WriteLine(); //pushes cursor to next line
                    }
                }

                // open file in note pad
                Process.Start("notepad.exe", saveFileDialog1.FileName);
            }
        }

        private void btnOpenWorld_Click(object sender, EventArgs e)
        {
            //make a word obj
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            //make new doc
            Document doc = word.Documents.Add();

            Microsoft.Office.Interop.Word.Range rng = doc.Range(0, 0);

            //makes new table based on datagridview
            Table wdTable = doc.Tables.Add(rng, dataGridView1.Rows.Count, dataGridView1.Columns.Count);

            //make thick outer border
            wdTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleDouble;

            //make cell lines thin
            wdTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

            try
            {
                //make active doc in word
                doc = word.ActiveDocument;

                //i is row index
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    //j is column index
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        //rows several times and writes cell value from grid to word
                        if (i == 0)
                        {
                            wdTable.Cell(i + 1, j + 1).Range.InsertAfter(dataGridView1.Columns[j].HeaderText);
                        }
                        else
                        {
                            wdTable.Cell(i + 1, j + 1).Range.InsertAfter(dataGridView1.Rows[i].Cells[j].Value.ToString());
                        }
                    }
                }

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        //saves file to drive
                        doc.SaveAs(saveFileDialog1.FileName);

                        //open doc in word
                        Process.Start("winword.exe", saveFileDialog1.FileName);
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                //quite world
                word.Quit();
                word = null;

                //clean up world obj & doc obj
                doc = null;
            }
        }
    }
}
