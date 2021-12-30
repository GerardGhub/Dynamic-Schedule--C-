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
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using DynamicSched;
using Tulpep.NotificationWindow;
using System.Diagnostics;

namespace DSched
{
    public partial class Form1 : Form
    {
        //String connetionString = @"Data Source=192.168.2.9\SQLEXPRESS;Initial Catalog=hr_bak;User ID=sa;Password=Nescafe3in1;MultipleActiveResultSets=true";
        String connetionString = @"Server=FM-MMERCADO-L;Initial Catalog=hr_bak;Integrated Security=SSPI;";
        SqlConnection con = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        SqlDataAdapter da;
        DataView dv;
        DataSet ds;
        DataTable dt;
        TableLogOnInfos crtableLogoninfos = new TableLogOnInfos();
        TableLogOnInfo crtableLogoninfo = new TableLogOnInfo();
        ConnectionInfo crConnectionInfo = new ConnectionInfo();
        Tables CrTables;
        bool pic;
        string mode;
        bool sw;
        int rowindex;
        int globalindex;
        string dtr_type;
        bool edited;

        bool added;
        int lstCount, lstAdded;

        private string selectedShift = "";
        private string timeInSS = "";
        private string timeOutSS = "";


        public Form1()
        {
            InitializeComponent();
        }
        private void con_on()
        {
            con = new SqlConnection();
            con.ConnectionString = connetionString;
            con.Open();
        }
        private void call_me(ReportDocument rd)
        {

            crConnectionInfo.ServerName = @"192.168.2.9\SQLEXPRESS";
            crConnectionInfo.DatabaseName = "hr_bak";
            crConnectionInfo.UserID = "sa";
            crConnectionInfo.Password = "Nescafe3in1";
            CrTables = rd.Database.Tables;

            foreach (CrystalDecisions.CrystalReports.Engine.Table CrTable in CrTables)
            {
                crtableLogoninfo = CrTable.LogOnInfo;
                crtableLogoninfo.ConnectionInfo = crConnectionInfo;
                CrTable.ApplyLogOnInfo(crtableLogoninfo);
            }
        }
        void show_tl()
        {
            ds = new DataSet();
            con_on();
            string q = "select e.employee_id,e.employee_number, (e.lastname + ', ' + e.firstname) as name,s.section_name,d.department_name from employee e " +
             "left join department d on d.department_id = e.department_id left join section s on s.section_id = e.section_id left join overnight_emp_sched oe on oe.employee_id = e.employee_id  where oe.dtr_type = 1 order by name asc";

            da = new SqlDataAdapter(q, con);
            da.Fill(ds);

            dataGridView4.DataSource = ds.Tables[0];
            con.Close();

            //foreach (DataGridViewColumn col in dataGridView4.Columns)
            //{
            //    col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //    col.HeaderCell.Style.Font = new Font("Tahoma", 12F, FontStyle.Regular, GraphicsUnit.Pixel);
            //    col.SortMode = DataGridViewColumnSortMode.NotSortable;
            //    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //    col.DefaultCellStyle.BackColor = Color.FromArgb(227, 242, 253);
            //}
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dataGridView4.Columns[0].HeaderText = "ID";
            dataGridView4.Columns[1].HeaderText = "Number";
            dataGridView4.Columns[2].HeaderText = "Name";
            dataGridView4.Columns[3].HeaderText = "Section";
            dataGridView4.Columns[4].HeaderText = "Department";
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            string ui = File.ReadLines(@"C:\a\userid.txt").First();

            lstCount = 0;
            lstAdded = 0;
            added = false;

            dateTimePicker1.Value = DateTime.Now;
            mode = "New Entry";
            label4.Text = mode;
            show_tl();
            show_emplist();
            label2.Text = dataGridView1.Rows.Count.ToString();
            //check_count();
            dateTimePicker3.Value = DateTime.Now;
            textBox1.Select();


            if (ui == "27" || ui == "28" || ui == "29" || ui == "31") // ui == 17 JAYCE USER ID
            {
                bool sw = false;

                button1.Enabled = sw;
                button4.Enabled = sw;
                button6.Enabled = sw;
                button5.Enabled = sw;

                dataGridView2.ReadOnly = true;
            }

            notif("Dynamic Schedule (Bulk Entry)", "Choose employees on the left side and click to the list view on the right side");
        }

        public void show_shiftsched()
        {
           
                con_on();
                string q = "SELECT ss.Shift_Name, ss.Sched_TimeIN, ss.Sched_TimeOUT FROM dbo.Sched_Shift ss WHERE ss.is_active = 1";
                cmd = new SqlCommand(q, con);
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    dataGridView5.Rows.Clear();
                    listView2.Items.Clear();
                    while (dr.Read())
                    {
                        listView2.Items.Add(dr[0].ToString());
                        listView2.Items[listView2.Items.Count - 1].SubItems.Add(dr[1].ToString());
                        listView2.Items[listView2.Items.Count - 1].SubItems.Add(dr[2].ToString());
                    }
                }
                con.Close();

        }
        void check_count()
        {
            con_on();
            string q = "SELECT DISTINCT COUNT(*) FROM dbo.dynamic_shift_tab dst";
            cmd = new SqlCommand(q, con);
            dr = cmd.ExecuteReader();
            dr.Read(); 
            lstCount = Convert.ToInt32(dr[0].ToString());
            con.Close();
        }
 
        private void generate_grid(int ri) // DYNAMIC SCHED EDIT EDIT EDIT
        {
            try
            {
                globalindex = ri; edited = false;
                ds = new DataSet();
                con_on();
                string q = "SELECT DISTINCT format(CAST(CONVERT(VARCHAR(15), attendance_date) AS DATETIME), 'MMM dd yyyy') AS ATTDATE, DATENAME(DW, attendance_date),in_time as In_Time, out_time as Out_Time, is_overnight, is_rest_day FROM dynamic_shift_tab WHERE emp_no = '" + dataGridView1.Rows[ri].Cells[1].Value + "' and FORMAT(CAST(CONVERT(VARCHAR, attendance_date, 114) AS DATETIME), 'MMMM-yyyy') = '" + comboBox1.Text + "-" + DateTime.Now.ToString("yyyy") + "' ORDER BY ATTDATE ASC"; // 
                da = new SqlDataAdapter(q, con);
                da.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
                con.Close();
                
                dataGridView2.Columns[0].HeaderText = "Date";
                dataGridView2.Columns[1].HeaderText = "Day";
                dataGridView2.Columns[2].HeaderText = "In Time";
                dataGridView2.Columns[3].HeaderText = "Out Time";
                dataGridView2.Columns[4].HeaderText = "Is Overnight";
                dataGridView2.Columns[5].HeaderText = "Is Restday";

                foreach (DataGridViewColumn col in dataGridView2.Columns)
                {
                    col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    col.HeaderCell.Style.Font = new Font("Tahoma", 12F, FontStyle.Regular, GraphicsUnit.Pixel);
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }

                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView2.Columns[4].DefaultCellStyle.BackColor = Color.FromArgb(236, 239, 241);


                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    if (dataGridView2.Rows[i].Cells[1].Value.ToString() == "Sunday")
                    {
                        //dataGridView2.Rows[i].Cells[1].Style.BackColor = Color.Gold;
                        dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Gold;
                    } // 141, 183, 152
                }

                    dataGridView2.ClearSelection();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                dataGridView2.DataSource = null;
            }

        }
        void show_emplist()
        {
            ds = new DataSet();
            con_on();
            string q = "select e.employee_id,e.employee_number, (e.lastname + ', ' + e.firstname) as name,d.department_name from employee e " +
             "left join department d on d.department_id = e.department_id where e.is_resigned = 0 and e.is_active = 1 and  e.employee_number is not null and e.employee_number <> '' order by name asc";

            da = new SqlDataAdapter(q, con);
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dv = new DataView(ds.Tables[0]);

            setup_dg(dataGridView1);
        }

        void setup_dg(DataGridView dgv)
        {
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.HeaderCell.Style.BackColor = Color.FromArgb(46, 46, 46);
                col.HeaderCell.Style.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Pixel);
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //col.DefaultCellStyle.BackColor = Color.MintCream;
                //col.DefaultCellStyle.BackColor = Color.FromArgb(227, 242, 253);
            }
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.Columns[0].HeaderText = "ID";
            dgv.Columns[0].Width = 30;
            dgv.Columns[1].HeaderText = "#";
            dgv.Columns[1].Width = 50;
            dgv.Columns[2].HeaderText = "Name";
            dgv.Columns[2].Width = 140;
            dgv.Columns[3].HeaderText = "Department";
            dgv.Columns[3].Width = 120;
            dgv.RowHeadersVisible = false;
            dgv.ClearSelection();
            label2.Text = dgv.Rows.Count.ToString();

        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            doSearch();
        }
        void doSearch()
        {
            try
            {
                dv.RowFilter = "name like '%" + textBox1.Text + "%'";
                dataGridView1.DataSource = dv;
                label2.Text = dataGridView1.Rows.Count.ToString();
            }
            catch
            {

            }
           
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listView1_Click(object sender, EventArgs e)
        {

        }

        private void listView1_MouseEnter(object sender, EventArgs e)
        {

        }

        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowindex = dataGridView1.CurrentCell.RowIndex;

            if (mode == "New Entry")
            {
                sw = true;
                try
                {
                    pictureBox3.Image = Image.FromFile("c:\\pics " + @"\" + dataGridView1.Rows[rowindex].Cells[0].Value.ToString() + ".jpg");
                }
                catch (Exception Ex)
                {
                    pictureBox3.Image = Image.FromFile("c:\\pics " + @"\Employee.png");
                }
            }
            else if (mode == "Per Month")
            {
                generate_grid(rowindex);
            }
            else if (mode == "Overnight List")
            {
                gen_onight(rowindex);
                button8.Enabled = false;
                pictureBox1.Image = null;
                pictureBox2.Image = null;
                qwe();
            }
        }
        void gen_onight(int ri)
        {
            try
            {
                globalindex = ri;
                ds = new DataSet();
                con_on();
                string q = "SELECT distinct attendance_date,  DATENAME(DW, attendance_date),ISNULL(FORMAT(CAST(CONVERT(nvarchar, in_time, 112) AS datetime), 'hh:mm ttt'), null) as In_Time, ISNULL(FORMAT(CAST(CONVERT(nvarchar, out_time, 112) AS datetime), 'hh:mm ttt'), null) as Out_Time FROM dynamic_shift_tab WHERE emp_no = '" + dataGridView1.Rows[ri].Cells[1].Value + "' and FORMAT(CAST(CONVERT(VARCHAR, attendance_date, 114) AS DATETIME), 'MMMM-yyyy') = '" + comboBox2.Text + "-" + DateTime.Now.ToString("yyyy") + "' and is_overnight = 1 ORDER BY attendance_date ASC";
                da = new SqlDataAdapter(q, con);
                da.Fill(ds);
                dataGridView3.DataSource = ds.Tables[0];
                con.Close();

                foreach (DataGridViewColumn col in dataGridView3.Columns)
                {
                    col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    col.HeaderCell.Style.Font = new Font("Tahoma", 12F, FontStyle.Regular, GraphicsUnit.Pixel);
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                dataGridView3.Columns[0].HeaderText = "Date";
                dataGridView3.Columns[1].HeaderText = "Day";
                dataGridView3.Columns[2].HeaderText = "In Time";
                dataGridView3.Columns[3].HeaderText = "Out Time";
                dataGridView3.ClearSelection();
            }
            catch
            {
                dataGridView3.DataSource = null;
            }

        }
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            pictureBox3.Image = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count != 0)
            {
                DialogResult dialogResult = MessageBox.Show("Save New Entry List for the Month of " + dateTimePicker1.Value.ToString("MMMM") + "?", "Dynamic Sched", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dialogResult == DialogResult.Yes)
                {
                    con_on();
                    for (int i = 0; i <= listView1.Items.Count - 1; i++)
                    {
                        var startDate = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, 1);
                        var cutoffDate = startDate.AddMonths(1).AddDays(-1);

                        while (startDate <= cutoffDate)
                        {
                            string sw = "select * from dynamic_shift_tab where attendance_date = '" + startDate.ToString("yyyy-MM-dd") + "' and emp_no = '" + listView1.Items[i].SubItems[1].Text + "'";
                            cmd = new SqlCommand(sw, con);
                            dr = cmd.ExecuteReader();

                            if (dr.HasRows == false)
                            {
                                string q = "INSERT INTO dynamic_shift_tab (year, emp_no, attendance_date, in_time, out_time, is_active, is_generated) VALUES (@a, @b, @c, @d, @e, 1, 0)";
                                cmd = new SqlCommand(q, con);
                                cmd.Parameters.AddWithValue("@a", startDate.ToString("yyyy"));
                                cmd.Parameters.AddWithValue("@b", listView1.Items[i].SubItems[1].Text);
                                cmd.Parameters.AddWithValue("@c", startDate.ToString("yyyy-MM-dd"));
                                cmd.Parameters.AddWithValue("@d", "08:00 AM");
                                cmd.Parameters.AddWithValue("@e", "05:00 PM");
                                cmd.ExecuteNonQuery();

                                lstAdded++;
                            }
                           
                            
                            startDate = startDate.AddDays(1);
                        }
                        ActivitiesLogs("Added " + listView1.Items[i].SubItems[2].Text + " on Dynamic Sched " + dateTimePicker1.Value.ToString("yyyy-MM") + " ");
                    }

                    added = true;
                    con.Close();
                    MessageBox.Show("New Bulk Entry Saved", "Dynamic Shift", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    button4_Click(sender, e);
                }
            }
            else
            {
                MessageBox.Show("Check your Input", "Dynamic Sched", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }
        void check_newadd()
        {
            con_on();
            string q = "SELECT TOP " + lstAdded + " (e.lastname + ', ' + e.firstname) FROM dbo.dynamic_shift_tab dst LEFT JOIN dbo.employee e ON dst.emp_no = e.employee_number ORDER BY dst.DST_ID DESC";
            cmd = new SqlCommand(q, con);
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[2].Value.ToString() == dr[0].ToString())
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.PaleGreen;
                    }
                }
            }
            con.Close();

        }

        public void notif(string titlex, string txtx)
        {
            PopupNotifier pp = new PopupNotifier();
           // pp.Image = Properties.Resources.iconfinder_130_man_student_2_3099383;
            pp.Image = Image.FromFile(@"notsched.png");
            pp.Size = new Size(new Point(400, 60));
            pp.TitleText = titlex;
            pp.ContentText = txtx;
            pp.TitleColor = Color.Black;
            pp.Popup();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabControl1.TabPages[0])//your specific tabname
            {
                mode = "New Entry";
                show_emplist();
                //notifyIcon1.Text = "Form1 (NotifyIcon example)";
                //notifyIcon1.Visible = true;
                //notifyIcon1.BalloonTipTitle = "Dynamic Schedule (Bulk Entry)";
                //notifyIcon1.BalloonTipText = "Choose employees on the left side and click to the list view on the right side";
                //notifyIcon1.ShowBalloonTip(5000);

                notif("Dynamic Schedule (Bulk Entry)", "Choose employees on the left side and click to the list view on the right side");


                textBox1.Text = ""; textBox1.Select(); pictureBox3.Image = null;
                this.WindowState = FormWindowState.Normal;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[1])
            {
                mode = "Per Month";
                dataGridView2.DataSource = null;
                comboBox1.Text = DateTime.Now.ToString("MMMM");
                show_month_list();
                textBox1.Text = "";
                notifyIcon1.Visible = false; textBox1.Select();
                if (added == true)
                {
                   check_newadd();
                }
                this.WindowState = FormWindowState.Normal;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[2])
            {
                mode = "Overnight List";
                comboBox2.Text = DateTime.Now.ToString("MMMM");
                dataGridView3.DataSource = null;
                show_onight();
                textBox1.Text = "";
                notifyIcon1.Visible = false; textBox1.Select();
                this.WindowState = FormWindowState.Normal;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[3])
            {
                mode = "Two Logs";
                textBox1.Text = "";
                dataGridView1.DataSource = null;
                notifyIcon1.Visible = false;
                this.WindowState = FormWindowState.Normal;
            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[4])
            {

                mode = "Shift Sched";
                textBox1.Text = "";
                dataGridView1.DataSource = null;
                notifyIcon1.Visible = false;
                this.WindowState = FormWindowState.Normal;

            }
            else if (tabControl1.SelectedTab == tabControl1.TabPages[5])
            {
                mode = "Report";
                textBox1.Text = "";
                dataGridView1.DataSource = null;
                notifyIcon1.Visible = false;
                this.WindowState = FormWindowState.Maximized;

            }
            label4.Text = mode;
            label2.Text = dataGridView1.Rows.Count.ToString();
        }
        void show_onight()
        {
            ds = new DataSet();
            con_on();

            string q = "SELECT DISTINCT e.employee_id, dst.emp_no, (e.lastname + ', ' + e.firstname) AS name, d.department_name FROM dynamic_shift_tab dst LEFT JOIN employee e  ON e.employee_number = dst.emp_no " +
  " LEFT JOIN dbo.department d ON d.department_id = e.department_id WHERE is_overnight = 1 and FORMAT(CAST(CONVERT(VARCHAR, dst.attendance_date, 114) AS DATETIME), 'MMMM-yyyy') = '" + comboBox2.Text + "-" + DateTime.Now.ToString("yyyy") + "' ORDER BY name ASC";
            da = new SqlDataAdapter(q, con);
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dv = new DataView(ds.Tables[0]);

            setup_dg(dataGridView1);
        }
        void show_month_list()
        {
            dataGridView2.DataSource = null;
            ds = new DataSet();
            con_on();
            string q = "SELECT DISTINCT e.employee_id, dst.emp_no, (e.lastname + ', ' + e.firstname) AS name, d.department_name FROM dynamic_shift_tab dst LEFT JOIN employee e  ON e.employee_number = dst.emp_no " +
 " LEFT JOIN dbo.department d ON d.department_id = e.department_id WHERE  FORMAT(CAST(CONVERT(VARCHAR, dst.attendance_date, 114) AS DATETIME), 'MMMM-yyyy') = '" + comboBox1.Text + "-" + DateTime.Now.ToString("yyyy") + "' ORDER BY name ASC"; // 
            da = new SqlDataAdapter(q, con);
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dv = new DataView(ds.Tables[0]);

            setup_dg(dataGridView1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            show_month_list();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;
            show_month_list();
            textBox1.Text = "";
            textBox1.Focus();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            notifyIcon1.Visible = false;
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Save Changes For " + dataGridView1.Rows[globalindex].Cells[2].Value + "?", "Dynamic Sched", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dialogResult == DialogResult.Yes)
            { // OT CELL 4

                con_on();
                //string q = "DELETE FROM dynamic_shift_tab where emp_no = @a and FORMAT(CAST(CONVERT(VARCHAR, attendance_date, 114) AS DATETIME), 'MMMM-yyyy') = @b";
                //cmd = new SqlCommand(q, con);
                //cmd.Parameters.AddWithValue("@a", dataGridView1.Rows[globalindex].Cells[1].Value);
                //cmd.Parameters.AddWithValue("@b", comboBox1.Text + "-" + DateTime.Now.ToString("yyyy"));
                //cmd.ExecuteNonQuery();

                int oncount = 0;

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {

                    if (dataGridView2.Rows[i].Cells[2].Value.ToString() == "08:00 AM" && dataGridView2.Rows[i].Cells[3].Value.ToString() == "05:00 PM")
                    {
                        string wew = "DELETE FROM dynamic_shift_tab where emp_no = @a and attendance_date = @b";
                        cmd = new SqlCommand(wew, con);
                        cmd.Parameters.AddWithValue("@a", dataGridView1.Rows[globalindex].Cells[1].Value);
                        cmd.Parameters.AddWithValue("@b", Convert.ToDateTime(dataGridView2.Rows[i].Cells[0].Value).ToString("yyyy-MM-dd"));
                        cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        string w = "UPDATE dynamic_shift_tab SET in_time = @d, out_time = @e, is_overnight = @f, is_rest_day = @g WHERE emp_no = @b and attendance_date = @c";
                        cmd = new SqlCommand(w, con);
                        cmd.Parameters.AddWithValue("@a", DateTime.Now.ToString("yyyy")); // overnight restday
                        cmd.Parameters.AddWithValue("@b", dataGridView1.Rows[globalindex].Cells[1].Value);
                        cmd.Parameters.AddWithValue("@c", Convert.ToDateTime(dataGridView2.Rows[i].Cells[0].Value).ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@d", dataGridView2.Rows[i].Cells[2].Value);
                        cmd.Parameters.AddWithValue("@e", dataGridView2.Rows[i].Cells[3].Value);
                        cmd.Parameters.AddWithValue("@f", dataGridView2.Rows[i].Cells[4].Value);
                        if (dataGridView2.Rows[i].Cells[4].Value.Equals(true))
                        {
                            oncount++;
                        }
                        cmd.Parameters.AddWithValue("@g", dataGridView2.Rows[i].Cells[5].Value);

                        cmd.ExecuteNonQuery();
                    }

                    ActivitiesLogs("Updated " + dataGridView1.Rows[globalindex].Cells[2].Value.ToString() + " Dynamic Sched");
                }

                if (oncount == 0)
                {
                    string jk = "select dtr_type from overnight_emp_sched where employee_id = '" + dataGridView1.Rows[globalindex].Cells[0].Value + "'";
                    cmd = new SqlCommand(jk, con);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == true)
                    {
                        dr.Read();
                        if (dr[0].ToString() != "1")
                        {
                            string ws = "DELETE FROM overnight_emp_sched where employee_id = '" + dataGridView1.Rows[globalindex].Cells[0].Value + "'";
                            cmd = new SqlCommand(ws, con);
                            cmd.ExecuteNonQuery();

                            string ew = "DELETE FROM SchedTrigger where employee_id = '" + dataGridView1.Rows[globalindex].Cells[0].Value + "'";
                            cmd = new SqlCommand(ew, con);
                            cmd.ExecuteNonQuery();

                            ActivitiesLogs("Deleted " + dataGridView1.Rows[globalindex].Cells[2].Value.ToString() + " Overnight Timekeeping Access");
                        }
                    }
                }
                else if (oncount > 0)
                {
                    string jk = "select dtr_type from overnight_emp_sched where employee_id = '" + dataGridView1.Rows[globalindex].Cells[0].Value + "'";
                    cmd = new SqlCommand(jk, con);
                    dr = cmd.ExecuteReader();
                    if (dr.HasRows == false)
                    {
                        DateTime now = new DateTime();
                        now = DateTime.Now;

                        var zxc = new DateTime(now.Year, now.Month, 1);

                        string q = "insert into overnight_emp_sched(employee_id,employee_number,dtr_type) values (@a,@b,2)";

                        cmd = new SqlCommand(q, con);

                        cmd.Parameters.AddWithValue("a", dataGridView1.Rows[globalindex].Cells[0].Value);
                        cmd.Parameters.AddWithValue("b", dataGridView1.Rows[globalindex].Cells[1].Value);
                        cmd.ExecuteNonQuery();

                        string w = "insert into SchedTrigger(employee_id,datetimeStart) values (@a,@b)";

                        cmd = new SqlCommand(w, con);

                        cmd.Parameters.AddWithValue("a", dataGridView1.Rows[globalindex].Cells[0].Value);
                        cmd.Parameters.AddWithValue("b", zxc.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                        cmd.ExecuteNonQuery();
                        ActivitiesLogs("Inserted " + dataGridView1.Rows[globalindex].Cells[2].Value.ToString() + " Overnight Timekeeping Access");

                    }
                }
                ActivitiesLogs("Saved " + dataGridView1.Rows[globalindex].Cells[2].Value + "Dynamic Sched");
                con.Close();
                MessageBox.Show("Saved");
                show_month_list();
                check_newadd();
                textBox1.Clear();
                textBox1.Select();
            }
        }
        public void ActivitiesLogs(string logs)
        {

            try
            {

                //@"c:\a\UserName.txt"
                const string location = @"DSCHEDLogs";

                if (!File.Exists(location))
                {
                    var createText = "New Activities Logs" + Environment.NewLine;
                    File.WriteAllText(location, createText);

                }
                var appendLogs = "Activities Logs: " + logs + " " + DateTime.Now + Environment.NewLine;
                File.AppendAllText(location, appendLogs);
            }
            catch (Exception ex)
            {
                const string location = @"DSCHEDLogs";
                if (!File.Exists(location))
                {
                    TextWriter file = File.CreateText(@"C:\DSCHEDLogs");
                    var createText = "New Activities Logs" + Environment.NewLine;

                    File.WriteAllText(location, createText);

                }

                var appendLogs = ex.Message + logs + DateTime.Now + Environment.NewLine;
                File.AppendAllText(location, appendLogs);


            }

        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Delete Dynamic Schedule Of " + dataGridView1.Rows[globalindex].Cells[2].Value + " For " + comboBox1.Text + "-" + DateTime.Now.ToString("yyyy") + "?", "Dynamic Sched", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dialogResult == DialogResult.Yes)
            {
                con_on();
                string q = "DELETE FROM dynamic_shift_tab where emp_no = @a and FORMAT(CAST(CONVERT(VARCHAR, attendance_date, 114) AS DATETIME), 'MMMM-yyyy') = @b";
                cmd = new SqlCommand(q, con);
                cmd.Parameters.AddWithValue("@a", dataGridView1.Rows[globalindex].Cells[1].Value);
                cmd.Parameters.AddWithValue("@b", comboBox1.Text + "-" + DateTime.Now.ToString("yyyy"));
                cmd.ExecuteNonQuery();

                string w = "DELETE FROM overnight_emp_sched where employee_id = '" + dataGridView1.Rows[globalindex].Cells[0].Value + "'";
                cmd = new SqlCommand(w, con);
                cmd.ExecuteNonQuery();

                string ew = "DELETE FROM SchedTrigger where employee_id = '" + dataGridView1.Rows[globalindex].Cells[0].Value + "'";
                cmd = new SqlCommand(ew, con);
                cmd.ExecuteNonQuery();
                ActivitiesLogs("Deleted " + dataGridView1.Rows[globalindex].Cells[2].Value + " Dynamic Sched");
                con.Close();
                MessageBox.Show("Deleted");
                show_month_list();
                check_newadd();
                textBox1.Select();
               

            }
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            show_onight();
            textBox1.Text = "";
            textBox1.Focus();
            dataGridView3.DataSource = null;
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        void qwe()
        {
            con_on();
            string q = "select dtr_type from overnight_emp_sched where employee_number = @a";
            cmd = new SqlCommand(q, con);
            cmd.Parameters.AddWithValue("@a", dataGridView1.Rows[globalindex].Cells[1].Value);
            dr = cmd.ExecuteReader();

            dr.Read();
            if (dr.HasRows == true)
            {
                if (dr[0].ToString() == "1")
                {
                    pic = true;
                    pictureBox1.Image = imageList1.Images[0];
                    pictureBox2.Image = imageList1.Images[1];
                }
                else if (dr[0].ToString() == "2")
                {
                    pic = true;
                    pictureBox1.Image = imageList1.Images[1];
                    pictureBox2.Image = imageList1.Images[0];
                }
            }
            else
            {
                pictureBox1.Image = imageList1.Images[1];
                pictureBox2.Image = imageList1.Images[1];
                pic = false;
            }
            con.Close();
        }
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            DateTime now = new DateTime();
            now = DateTime.Now;

            var zxc = new DateTime(now.Year, now.Month, 1);

            con_on();
            string q = "insert into overnight_emp_sched(employee_id,employee_number,dtr_type) values (@a,@b,@c)";

            cmd = new SqlCommand(q, con);

            cmd.Parameters.AddWithValue("a", dataGridView1.Rows[globalindex].Cells[0].Value);
            cmd.Parameters.AddWithValue("b", dataGridView1.Rows[globalindex].Cells[1].Value);
            cmd.Parameters.AddWithValue("c", dtr_type);
            cmd.ExecuteNonQuery();

            string w = "insert into SchedTrigger(employee_id,datetimeStart) values (@a,@b)";

            cmd = new SqlCommand(w, con);

            cmd.Parameters.AddWithValue("a", dataGridView1.Rows[globalindex].Cells[0].Value);
            cmd.Parameters.AddWithValue("b", zxc.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            cmd.ExecuteNonQuery();


            con.Close();
            MessageBox.Show("Added");
            button8.Enabled = false; textBox1.Select();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (pic == false)
            {
                pictureBox2.Image = imageList1.Images[0];
                button8.Enabled = true;
                dtr_type = "2";
            }
            else
            {
                MessageBox.Show("Employee already on other option");
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (pic == false)
            {
                pictureBox1.Image = imageList1.Images[0];
                button8.Enabled = true;
                dtr_type = "1";
            }
            else
            {
                MessageBox.Show("Employee already on other option");
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (sw == true && mode == "New Entry")
            {
                listView1.Items.Add(dataGridView1.Rows[rowindex].Cells[0].Value.ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(dataGridView1.Rows[rowindex].Cells[1].Value.ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(dataGridView1.Rows[rowindex].Cells[2].Value.ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(dataGridView1.Rows[rowindex].Cells[3].Value.ToString());
                sw = false;
            }
            else if (mode == "Per Month")
            {

                Image img;

                try
                {
                    img = Image.FromFile("c:\\pics " + @"\" + dataGridView1.Rows[rowindex].Cells[0].Value.ToString() + ".jpg");
                }
                catch (Exception ex)
                {
                    img = Image.FromFile("c:\\pics " + @"\Employee.png");
                }


                Form2 frm2 = new Form2(this, img);
                frm2.ShowDialog();
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = dataGridView2.CurrentCell.RowIndex;

            if (e.ColumnIndex > -1)
            {
                DataGridViewComboBoxCell l_objGridDropbox = new DataGridViewComboBoxCell();
                try
                {


                    if (dataGridView2.Columns[e.ColumnIndex].Name.Contains("In_Time") && edited == false)
                    {
                        DialogResult dialogResult12 = MessageBox.Show("Enable Drop Down List?", "O&G | HR", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                        if (dialogResult12 == DialogResult.Yes)
                        {
                            edited = true;
                            dataGridView2[e.ColumnIndex, e.RowIndex] = l_objGridDropbox;
                            l_objGridDropbox.DataSource = GetDescriptionTable();
                            l_objGridDropbox.ValueMember = "Description";
                            l_objGridDropbox.DisplayMember = "Description";
                            l_objGridDropbox.FlatStyle = FlatStyle.System;
                            dataGridView2.Rows[index].Cells[2].Style.BackColor = Color.IndianRed;
                            dataGridView2.Rows[index].Cells[2].Style.ForeColor = Color.White;
                        }
                        else if (dialogResult12 == DialogResult.No)
                        {
                            edited = true;
                        }
                    }
                    if (dataGridView2.Columns[e.ColumnIndex].Name.Contains("Out_Time") && edited == false)
                    {
                        DialogResult dialogResult12 = MessageBox.Show("Enable Drop Down List?", "O&G | HR", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                        if (dialogResult12 == DialogResult.Yes)
                        {
                            edited = true;
                            dataGridView2[e.ColumnIndex, e.RowIndex] = l_objGridDropbox;
                            l_objGridDropbox.DataSource = GetDescriptionTable();
                            l_objGridDropbox.ValueMember = "Description";
                            l_objGridDropbox.DisplayMember = "Description"; l_objGridDropbox.FlatStyle = FlatStyle.System;
                            dataGridView2.Rows[index].Cells[3].Style.BackColor = Color.IndianRed;
                            dataGridView2.Rows[index].Cells[3].Style.ForeColor = Color.White;
                        }
                        else if (dialogResult12 == DialogResult.No)
                        {
                            edited = true;
                        }
                    }
                }
                catch { }
                // Check the column  cell, in which it click.  
                
            }
        }
        private DataTable GetDescriptionTable()
        {
            DataTable l_dtDescription = new DataTable();
            l_dtDescription.Columns.Add("Description", typeof(string));

            l_dtDescription.Rows.Add("01:00 AM");
            l_dtDescription.Rows.Add("02:00 AM");
            l_dtDescription.Rows.Add("03:00 AM");
            l_dtDescription.Rows.Add("04:00 AM");
            l_dtDescription.Rows.Add("05:00 AM");
            l_dtDescription.Rows.Add("06:00 AM");
            l_dtDescription.Rows.Add("07:00 AM");
            l_dtDescription.Rows.Add("08:00 AM");
            l_dtDescription.Rows.Add("09:00 AM");
            l_dtDescription.Rows.Add("10:00 AM");
            l_dtDescription.Rows.Add("11:00 AM");
            l_dtDescription.Rows.Add("12:00 AM");

            l_dtDescription.Rows.Add("01:00 PM");
            l_dtDescription.Rows.Add("02:00 PM");
            l_dtDescription.Rows.Add("03:00 PM");
            l_dtDescription.Rows.Add("04:00 PM");
            l_dtDescription.Rows.Add("05:00 PM");
            l_dtDescription.Rows.Add("06:00 PM");
            l_dtDescription.Rows.Add("07:00 PM");
            l_dtDescription.Rows.Add("08:00 PM");
            l_dtDescription.Rows.Add("09:00 PM");
            l_dtDescription.Rows.Add("10:00 PM");
            l_dtDescription.Rows.Add("11:00 PM");
            l_dtDescription.Rows.Add("12:00 PM");

            return l_dtDescription;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            
        }

        private void dataGridView2_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            edited = false;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

        
            ReportDocument cryRpt = new ReportDocument();
            cryRpt.Load("C:/Reports/dschedlist.rpt");
            cryRpt.SetParameterValue("@title", "List of Dynamic Schedule for " + dateTimePicker2.Value.ToString("MMM dd, yyyy"));
            cryRpt.SetParameterValue("@date", dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            call_me(cryRpt);
            crystalReportViewer1.ReportSource = cryRpt;
            crystalReportViewer1.Refresh();
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && mode == "Per Month")
            {
                button6_Click(this, e);
            }
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        void show_ss_emp()
        {
            con_on();
            string q = "";
        }
        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView5.Rows.Clear();

            if (listView2.SelectedItems.Count != 0)
            {
                //MessageBox.Show(listView2.SelectedItems[0].Text);
                if (listView2.SelectedItems[0].Text == "1")
                {
                    return;
                }
                else
                {
                    selectedShift = listView2.SelectedItems[0].Text;
                    timeInSS = listView2.SelectedItems[0].SubItems[1].Text; 
                    timeOutSS = listView2.SelectedItems[0].SubItems[2].Text;
                    show_shiftdgv(listView2.SelectedItems[0].Text);
                }

            }
        }

        void show_shiftdgv(string shiftId)
        {
            con_on();
            string q = "SELECT e.employee_number, (e.lastname + ', ' + e.firstname), p.position_name FROM dbo.employee e LEFT JOIN dbo.positions p ON p.position_id = e.position_id WHERE e.ShiftNo = " + shiftId;
            cmd = new SqlCommand(q, con);
            dr = cmd.ExecuteReader();

            if (dr.HasRows)
            {
                dataGridView5.Rows.Clear();
                while (dr.Read())
                {
                    dataGridView5.Rows.Add(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), "ADD");
                }

                dataGridView5.ClearSelection();
            }
            else
            {

            }



            con.Close();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3(this);
            frm3.ShowDialog();
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            show_shiftsched();
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
              //  MessageBox.Show(dataGridView5[3, e.RowIndex].Value.ToString() + " TIME ADD : " + timeInSS + " TIME OUT : " + timeOutSS);

              int dateAdded = 0;
                var startDate = new DateTime(dateTimePicker3.Value.Year, dateTimePicker3.Value.Month, 1);
                var cutoffDate = startDate.AddMonths(1).AddDays(-1);
                string _empNum = dataGridView5[0, e.RowIndex].Value.ToString();
                
                con_on();

                while (startDate <= cutoffDate)
                {
                    string sw = "select DST_ID from dynamic_shift_tab where attendance_date = '" + startDate.ToString("yyyy-MM-dd") + "' and emp_no = '" + _empNum + "'";
                    cmd = new SqlCommand(sw, con);
                    dr = cmd.ExecuteReader();

                    if (dr.HasRows == false)
                    {
                        string q = "INSERT INTO dynamic_shift_tab (year, emp_no, attendance_date, in_time, out_time, is_active, is_generated) VALUES (@a, @b, @c, @d, @e, 1, 0)";
                        cmd = new SqlCommand(q, con);
                        cmd.Parameters.AddWithValue("@a", startDate.ToString("yyyy"));
                        cmd.Parameters.AddWithValue("@b", _empNum);
                        cmd.Parameters.AddWithValue("@c", startDate.ToString("yyyy-MM-dd"));
                        cmd.Parameters.AddWithValue("@d", timeInSS);
                        cmd.Parameters.AddWithValue("@e", timeOutSS);
                        cmd.ExecuteNonQuery();
                        dateAdded++;
                    }


                    startDate = startDate.AddDays(1);
                }

                con.Close();

                MessageBox.Show(
                    "Added " + dateAdded.ToString() + " dates to " + dataGridView5[1, e.RowIndex].Value.ToString() +
                    " schedule from " + timeInSS + " to " + timeOutSS + Environment.NewLine + "For the month of " +
                    dateTimePicker3.Value.ToString("MMMM yyyy"), "Sucess", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(selectedShift) && selectedShift != "1" && listView2.SelectedItems[0].Text != "1")
            {

                selSht.Text = selectedShift;
                var col = new AutoCompleteStringCollection();

                con_on();
                SqlCommand cmd = new SqlCommand("sp_philip", con);
                cmd.Parameters.AddWithValue("@mode", "showemployee");
                cmd.CommandType = CommandType.StoredProcedure;

                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    col.Add(dr[0].ToString());
                }
                con.Close();

                textBox4.AutoCompleteCustomSource = col;


                groupBox13.Visible = true;
            }
            else
            {
                MessageBox.Show("Check your selected shift on the list view", "Scheduling", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            textBox4.Text = "";
            groupBox13.Visible = false;
        }

        private void NotifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox4.Text))
            {
                con_on();

                string q = "UPDATE employee SET employee.ShiftNo = " + selSht.Text +
                           " WHERE (employee.lastname + ', ' + employee.firstname) = '" + textBox4.Text.Trim() + "'";
                cmd = new SqlCommand(q, con);
                if (cmd.ExecuteNonQuery() == 1)
                {
                    MessageBox.Show("Added");
                    textBox4.Text = "";
                    groupBox13.Visible = false;
                }
                con.Close();

            }
        }  
    }
}
