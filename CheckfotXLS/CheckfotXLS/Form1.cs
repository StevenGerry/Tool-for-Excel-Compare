using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//Porgrammer: Gerry Huang 2020/05/20

namespace CheckfotXLS
{
    public partial class ExcelMatch : Form
    {
        public ExcelMatch()
        {
            InitializeComponent();
        }

        private void read_xls1()
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Excel Files(*.xls)|*.xls";
            string strPath;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    strPath = fd.FileName;
                    string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";
                    OleDbConnection Con = new OleDbConnection(strCon);
                    string strSql = "select * from [Sheet1$]";
                    OleDbCommand Cmd = new OleDbCommand(strSql, Con);
                    OleDbDataAdapter da = new OleDbDataAdapter(Cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "Sheet1");

                    dataGridView1.DataSource = ds.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void read_xls2()
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Excel Files(*.xls)|*.xls";
            string strPath;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    strPath = fd.FileName;
                    string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//关键是红色区域
                    OleDbConnection Con = new OleDbConnection(strCon);
                    string strSql = "select * from [Sheet1$]";
                    OleDbCommand Cmd = new OleDbCommand(strSql, Con);
                    OleDbDataAdapter da = new OleDbDataAdapter(Cmd);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "Sheet2");

                    dataGridView2.DataSource = ds.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:" + ex.Message);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //read_xls2();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void label3_Click(object sender, EventArgs e)
        {
            read_xls1();
        }

        private void label4_Click(object sender, EventArgs e)
        {
            read_xls2();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int row1 = dataGridView1.Rows.Count;
            int row2 = dataGridView2.Rows.Count;
            int a1 = int.Parse(textBox1.Text.ToString());
            int a2 = int.Parse(textBox2.Text.ToString());
            int b1 = int.Parse(textBox3.Text.ToString());
            int b2 = int.Parse(textBox4.Text.ToString());
            int OK = 0;
            int FAIL = 0;
            try
            {
                for (int i = 0; i < row1-1; i++)
                {
                    String FindTXT = dataGridView1.Rows[i].Cells[a1-1].Value.ToString();
                    String FindData = dataGridView1.Rows[i].Cells[a2-1].Value.ToString();
                    for (int j = 0; j < row2-1; j++)
                    {
                        if (FindTXT == dataGridView2.Rows[j].Cells[b1-1].Value.ToString())
                        {
                            if (FindData == dataGridView2.Rows[j].Cells[b2-1].Value.ToString())
                                OK++;
                            else
                                FAIL++;
                        }
                        label1.Text = "Matching：" + OK.ToString();
                        label2.Text = "Error:" + FAIL.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:"+ex.Message);
            }
        }
    }
}
