using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;
using System.Data.OleDb;

namespace WindowsFormsApplication1
{

    public partial class SaveON : Form
    {
       private OleDbConnection connection = new OleDbConnection();
       OleDbDataAdapter da;
       DataTable dt = new DataTable();
        public SaveON()
        {
            InitializeComponent();
            connection.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\UCSC\3rd year\group project ii\saveon\saveon\saveon\update.accdb";
           
            // Set the BoldedDates property to a DateTime array.
            // ... Use array initializer syntax to add tomorrow and two days after.
            monthCalendar1.BoldedDates = new DateTime[]

            
	    {
		DateTime.Today.AddDays(1),
		DateTime.Today.AddDays(2),
		DateTime.Today.AddDays(4)
	    };

            timer1.Start();
        }
        

        //local member variable
        private DateTime dtBoottime = new DateTime();

        //enable the timer in the constructor

        //public SaveONd()
        //{
            //InitializeComponent();
            //timer1.Enabled = true;
        //}
    

        private void SaveON_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'database1DataSet.saveon' table. You can move, or remove it, as needed.
            SelectQuery query = new SelectQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem WHERE Primary = true");
            //create a new management object and pass it 
            //the select query
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);

            foreach (ManagementObject mo in searcher.Get())
            {
                dtBoottime = ManagementDateTimeConverter.ToDateTime(
                    mo.Properties["LastBootUpTime"].Value.ToString());

                //display
                lableDate.Text = dtBoottime.ToShortDateString();
                lableTime.Text = dtBoottime.ToLongTimeString();

            }
            comboBox2.Items.Add("4th floor");
            comboBox2.Items.Add("W002");
            comboBox2.Items.Add("MiniAuditorium");
            comboBox2.SelectedItem = "W002";

            comboBox3.Items.Add("1");
            comboBox3.Items.Add("2");
            comboBox3.Items.Add("3");
            comboBox3.SelectedItem = "4";
        }

        //private void button4_Click(object sender, EventArgs e)
        //{
           // this.Close();

       // }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        //public static DateTime Today { get; set; }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label5_Click_1(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {


            label8.Text = DateTime.Today.ToString(" HH:mm : yyyy");
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
           Application.Exit();
        }

        private void SaveON_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Are you sure you really want to exit ?", "Exit", MessageBoxButtons.YesNo);
            if (dialog == DialogResult.Yes) 
            {
                Application.Exit();
            }
            else if(dialog == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        { 
            DateTime datetime = DateTime.Now;
            this.time_lable.Text = datetime.ToString();


        }

        private void button7_Click(object sender, EventArgs e)
        {

            try 
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection; 
                //string query = "select * from update";
                command.CommandText = "insert into update (Date, Room_No, Floor , Open_hour, Closing_Hour) values ('" + dateTimePicker1 + "', '" + comboBox2 + "', '" + comboBox3 + "', '" + dateTimePicker2 + "', '" + dateTimePicker3 + "')";

              //command.CommandText = query;
              //OleDbDataReader reader = command.ExecuteReader();
                //command.ExecuteNonQuery();
              /*while (reader.Read())
              { 
                  comboBox2.Items.Add(reader["Floor"].ToString());
                  comboBox3.Items.Add(reader["Room_No"].ToString());
                  
              
              }*/
                //command.ExecuteNonQuery();
                MessageBox.Show("data updated");
                connection.Close();

            }
                        
            catch (Exception ex)
            {
            //MessageBox.Show("Error" + ex);
            }

        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string query = "select * from update";
                command.CommandText = query;

                da = new OleDbDataAdapter(command);
                
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                



                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured !" + ex);
            }
        }
    }
}
