using System;
using System.Data;
using System.Data.SqlClient;
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
        DataTable table;

        SqlCommandBuilder commandBuilder;

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
            GetData("Select * from BizContacts");
        }

        private void GetData(string selectCommand)
        {
            // b/c there could be a db connection error. always put a try/catch 
            try
            {
                //put in sql query and connstring to access db table
                dataAdapter = new SqlDataAdapter(selectCommand, connString);

                //create data table obj
                table = new DataTable();

                //data and language won't vary depending on location
                table.Locale = System.Globalization.CultureInfo.InvariantCulture;

                //fills the table w/ data
                dataAdapter.Fill(table);

                //set data source to binding source 
                bindingSource1.DataSource = table;
            }
            catch (SqlException ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            //new sql command obj
            SqlCommand command;

            //field neames in the table followed by paramter names in values
            string insert = @"insert into BizContacts(Date_Added, Company, Website, Title, First_Name, Last_Name, Address,
                                                     City, State, Postal_Code, Email, Mobile, Notes)
                                                                        
                                                     values(@Date_Added, @Company, @Website, @Title, @First_Name, @Last_Name, @Address,
                                                     @City, @State, @Postal_Code, @Email, @Mobile, @Notes)";

            //opens & closes connection to sql server
            using (SqlConnection conn = new SqlConnection(connString))
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

                    //this pushes data into the table
                    command.ExecuteNonQuery();

                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }


            }

            //query the table to see the new row
            GetData("Select * from BizContacts");

            //updates datagridview
            dataGridView1.Update();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            commandBuilder = new SqlCommandBuilder(dataAdapter);
        }
    }
}
