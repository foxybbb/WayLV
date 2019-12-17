using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.Sql;
using MySql.Data.MySqlClient;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            timer1.Interval = 5000;

            timer1.Enabled = false;
            button3.Text = "Enable";
            label2.Text = "Disabled";
        }

       
        public void Auth()
        {
            try{
                //login=demo+&password=demo&Submit=ieiet
                string username = textBox1.Text;
                string password = textBox2.Text;
                string data = "login=" + username + "&password=" + password + "&Submit=ieiet";
                webBrowser1.DocumentCompleted += wb_DocumentCompleted;
                webBrowser1.Navigate("http://www.lvceli.lv/cms/", "_self", System.Text.ASCIIEncoding.ASCII.GetBytes(data), "Content-Type:application/x-www-form-urlencoded");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void Mysql()
        {


            try
            {
                string sqlconnetion = @"Server=" + textBox3.Text + ";Database=" + textBox4.Text + ";Uid=" + textBox5.Text + ";Pwd=" + textBox6.Text + ";";
                MySqlConnection cnn = new MySqlConnection(sqlconnetion);
                cnn.ConnectionString = sqlconnetion;
                cnn.Open();


                MySqlCommand Create_table = new MySqlCommand(@"CREATE TABLE IF NOT EXISTS `" + textBox4.Text + "`.`" + textBox7.Text + "` (" +
                    "`id" + textBox7.Text + "` INT NOT NULL AUTO_INCREMENT," +
                    "`Stacija` VARCHAR(45) NULL," +
                    " `Laiks` VARCHAR(45) NULL," +
                    "`Gaisa temperatura`VARCHAR(45) NULL," +
                    " `Gaisa temperaturas tendence(-1 h)` VARCHAR(45) NULL," +
                    " `Gaisa mitrums` VARCHAR(45) NULL," +
                    "`Rasas punkts` VARCHAR(45) NULL," +
                    "`Nokrisni` VARCHAR(45) NULL," +
                    " `Intensitate mm / h` VARCHAR(45) NULL," +
                    "`Redzamiba`VARCHAR(45) NULL," +
                    " `Cela temperatura1` VARCHAR(45) NULL," +
                    "`Cela temperatura1 tendence(-1h)` VARCHAR(45) NULL," +
                    "`Cela stavoklis1` VARCHAR(45) NULL," +
                    "`Cela bridinajums1` VARCHAR(45) NULL," +
                    "`Sasalsanas punkts1` VARCHAR(45) NULL," +
                    "`Cela temperatura2` VARCHAR(45) NULL," +
                    " `Cela temperatura2 tendence(-1h)` VARCHAR(45) NULL," +
                    "`Cela stavoklis2` VARCHAR(45) NULL," +
                    "`Cela bridinajums2` VARCHAR(45) NULL," +
                    "`Sasalsanas punkts2` VARCHAR(45) NULL," +
                    "PRIMARY KEY(`idwaylv`))ENGINE = InnoDB; ", cnn);
                Create_table.ExecuteNonQuery();


                DataTable dt = null;
                dt = new DataTable();
                dt = scrapHtmlTable();

                MySqlCommand check = new MySqlCommand("SELECT COUNT(*) FROM "+ textBox7.Text + ";", cnn);
                MySqlDataReader reader = check.ExecuteReader();
                reader.Read();
                int count = reader.GetInt32(0);

                if (count == 0)
                {
                    reader.Close();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string data = @"INSERT INTO `"+ textBox4.Text + "`.`"+ textBox7.Text + "` VALUES ('" + (i + 1) + "','" + dt.Rows[i][0] + "', '" + dt.Rows[i][1] + "', '" + dt.Rows[i][2] + "','" + dt.Rows[i][3] + "','" + dt.Rows[i][4] + "','" + dt.Rows[i][5] + "', '" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "','" + dt.Rows[i][8] + "','" + dt.Rows[i][9] + "','" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "','" + dt.Rows[i][13] + "','" + dt.Rows[i][14] + "','" + dt.Rows[i][15] + "', '" + dt.Rows[i][16] + "', '" + dt.Rows[i][17] + "','" + dt.Rows[i][18] + "');";
                        //  string data = @"INSERT INTO mydata.table1 VALUES('123','d','2');";
                        MySqlCommand dat = new MySqlCommand(data, cnn);

                        dat.ExecuteNonQuery();
                    }
                }
                else if (count > 0)
                {

                    reader.Close();
                    for (int i = 1; i <= count; i++)
                    {
                        string delete = @"DELETE FROM `" + textBox7.Text + "` WHERE `id" + textBox7.Text + "` = " + i + ";";
                        MySqlCommand del = new MySqlCommand(delete, cnn);
                        del.ExecuteNonQuery();

                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string data = @"INSERT INTO `"+ textBox4.Text + "`.`"+ textBox7.Text + "` VALUES ('" + (i + 1) + "','" + dt.Rows[i][0] + "', '" + dt.Rows[i][1] + "', '" + dt.Rows[i][2] + "','" + dt.Rows[i][3] + "','" + dt.Rows[i][4] + "','" + dt.Rows[i][5] + "', '" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "','" + dt.Rows[i][8] + "','" + dt.Rows[i][9] + "','" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "','" + dt.Rows[i][13] + "','" + dt.Rows[i][14] + "','" + dt.Rows[i][15] + "', '" + dt.Rows[i][16] + "', '" + dt.Rows[i][17] + "','" + dt.Rows[i][18] + "');";
                        //  string data = @"INSERT INTO mydata.table1 VALUES('123','d','2');";
                        MySqlCommand dat = new MySqlCommand(data, cnn);

                        dat.ExecuteNonQuery();
                    }
                }
                cnn.Close();
            }
            catch (MySql.Data.MySqlClient.MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        void wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser1.Document.Window.ScrollTo(12, 36);//Scroll to the 100 position
        }
        public DataTable scrapHtmlTable()
        {

            HtmlElement mytable = webBrowser1.Document.GetElementById("table-1");

            HtmlDocument doc = new HtmlDocument();
            // doc.LoadHtml(htmlCode);
            //  var headers = doc.DocumentNode.SelectNodes("//tr/th");
            DataTable dt = new DataTable();
            DataRow dr = null;
            dataGridView1.AllowUserToAddRows = false;     //запрешаем пользователю самому добавлять строки

            dt.Columns.Add("Stacija", typeof(string));
            dt.Columns.Add("Laiks", typeof(string));
            dt.Columns.Add("Gaisa temperatūra", typeof(string));
            dt.Columns.Add("Gaisa temperatūras tendence (-1 h)", typeof(string));
            dt.Columns.Add("Gaisa mitrums", typeof(string));
            dt.Columns.Add("Rasas punkts", typeof(string));
            dt.Columns.Add("Nokrišņi", typeof(string));
            dt.Columns.Add("Intensitāte mm/h", typeof(string));
            dt.Columns.Add("Redzamība", typeof(string));
            dt.Columns.Add("Ceļa temperatūra1", typeof(string));
            dt.Columns.Add("Ceļa temperatūra1 tendence (-1h)", typeof(string));
            dt.Columns.Add("Ceļa stāvoklis1", typeof(string));
            dt.Columns.Add("Ceļa brīdinājums1", typeof(string));
            dt.Columns.Add("Sasalšanas punkts1", typeof(string));
            dt.Columns.Add("Ceļa temperatūra2", typeof(string));
            dt.Columns.Add("Ceļa temperatūra2 tendence (-1h)", typeof(string));
            dt.Columns.Add("Ceļa stāvoklis2", typeof(string));
            dt.Columns.Add("Ceļa brīdinājums2", typeof(string));
            dt.Columns.Add("Sasalšanas punkts2", typeof(string));

            dt.AcceptChanges();

            foreach (HtmlElement row in mytable.GetElementsByTagName("tr"))
            {
                dr = dt.NewRow();
                HtmlElementCollection cells = row.GetElementsByTagName("td");
                for (int i = 0; i < cells.Count; i++)
                {
                    dr[i] = cells[i].InnerText;

                }
                dt.Rows.Add(dr);
            }
            dt.AcceptChanges();
            dt.Rows.RemoveAt(0);
            dataGridView1.DataSource = dt;
           return dt;      
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Auth();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            scrapHtmlTable();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            Auth();
            WindowsFormsApp1.LoginHtml
            authPR();
           
            Mysql();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int j = authPR();
            if (j == 0)
            {
                if (timer1.Enabled)
                {
                    timer1.Enabled = false;

                    button3.Text = "Enable";
                    label2.Text = "Disabled";
                }
                else
                {
                    timer1.Enabled = true;
                    button3.Text = "Disable";
                    label2.Text = "Active";

                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            webBrowser1.Document.Window.ScrollTo(12, 36);//Scroll to the 100 position
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                }

                StartRow++;

                //Write datagridview content
                for (i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Mysql();
        }

       
    }
}