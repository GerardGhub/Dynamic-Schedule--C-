using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DSched;

namespace DynamicSched
{
    public partial class Form3 : Form
    {
        String connetionString = @"Server=FM-MMERCADO-L;Initial Catalog=hr_bak;Integrated Security=SSPI;";
        //String connetionString = @"Data Source=192.168.2.9\SQLEXPRESS;Initial Catalog=hr_bak;User ID=sa;Password=Nescafe3in1;MultipleActiveResultSets=true";
        SqlConnection con = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;


        private readonly Form1 frmMain;

        public Form3(Form1 _frm)
        {
            InitializeComponent();
            frmMain = _frm;
        }
        private void con_on()
        {
            con = new SqlConnection();
            con.ConnectionString = connetionString;
            con.Open();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void gbNew_Enter(object sender, EventArgs e)
        {

        }

        private void Form3_Load(object sender, EventArgs e)
        {
            string dateTXX = DateTime.Now.ToString("yyyy-MM-dd");

            dateTimePicker1.Value = Convert.ToDateTime(dateTXX + " 08:00 AM");
            dateTimePicker2.Value = Convert.ToDateTime(dateTXX + " 05:00 PM");




            con_on();
            string q = "SELECT MAX(ss.Shift_Name) FROM dbo.Sched_Shift ss";
            cmd = new SqlCommand(q, con);
            dr = cmd.ExecuteReader();
            if (dr.HasRows)
            {
                dr.Read();
                int i = Convert.ToInt32(dr[0].ToString()) + 1;

                label2.Text = i.ToString();
            }
            con.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con_on();
            string q =
                "INSERT INTO dbo.Sched_Shift (Shift_Name,Sched_TimeIN,Sched_TimeOUT,is_active)VALUES(@a,@b,@c,1)";
            cmd = new SqlCommand(q, con);
            cmd.Parameters.AddWithValue("@a", label2.Text);
            cmd.Parameters.AddWithValue("@b", dateTimePicker1.Value.ToString("hh:mm ttt"));
            cmd.Parameters.AddWithValue("@c", dateTimePicker2.Value.ToString("hh:mm ttt"));

            cmd.ExecuteNonQuery();
            con.Close();

            frmMain.show_shiftsched();
            this.Close();

        }
    }
}
