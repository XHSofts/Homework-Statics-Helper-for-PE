using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using Docs.Excel;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace zytj
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "请稍后...正在删除旧数据";
            button1.Enabled = false;
            File.Delete(@"./data.txt");
            button1.Text = "请稍后...正在下载新数据";
            WebClient Client = new WebClient();
            Client.DownloadFile("http://47.75.182.230/data.txt", @"./data.txt");
            button1.Text = "数据同步完成，点击再次同步^^";
            button1.Enabled = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = ".";
            openFileDialog1.Filter = "Excel 表格 (*.xlsx, *.xls)|*.xlsx;*.xls";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                textBox1.Text = selectedFileName;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = comboBox1.SelectedItem.ToString();
            if (String.Compare(textBox2.Text, "2019【网二34】") == 0)
            {
                textBox3.Text = "713011431";
            }
            else if (String.Compare(textBox2.Text, "2019【气一34】") == 0) {
                textBox3.Text = "720351151";
            }
            else if (String.Compare(textBox2.Text, "2019【气二56】") == 0)
            {
                textBox3.Text = "764786230";
            }
            else if (String.Compare(textBox2.Text, "2019【气四34】") == 0)
            {
                textBox3.Text = "777830439";
            }
            else if (String.Compare(textBox2.Text, "2019【气三34】") == 0)
            {
                textBox3.Text = "761795783";
            }
            else if (String.Compare(textBox2.Text, "2019【气三56】") == 0)
            {
                textBox3.Text = "601830842";
            }
            else if (String.Compare(textBox2.Text, "2019【气一56】") == 0)
            {
                textBox3.Text = "652910637";
            }
            else if (String.Compare(textBox2.Text, "2019【双健班】") == 0)
            {
                textBox3.Text = "627624513";
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (String.Compare(textBox1.Text, "") == 0)
            {
                MessageBox.Show("请选择记分册");
                return;
            }

            if (String.Compare(textBox2.Text, "") == 0)
            {
                MessageBox.Show("请选择班级");
                return;
            }

            if (String.Compare(button1.Text, "点击同步数据") == 0)
            {
                MessageBox.Show("请同步数据");
                return;
            }

            //Create a new workbook.  
            ExcelWorkbook Wbook = ExcelWorkbook.ReadXLSX(textBox1.Text);

            //锁定学号的行和列
            int start_row = 0;
            int start_col = 0;
            int total_row = Wbook.Worksheets[0].Rows.Count;

            while (true)
            {
                if (Wbook.Worksheets[0].Cells[start_row, start_col].Value != null)
                {
                    if (String.Compare(Wbook.Worksheets[0].Cells[start_row, start_col].Value.ToString(), "学号") == 0)
                    {
                        break;
                    }
                }

                if (start_col < total_row)
                {
                    start_col++;
                }
                else
                {
                    start_col = 0;
                    start_row++;
                }

            }

            richTextBox1.Text = richTextBox1.Text + "[通知] 发现“学号”位于" + (start_row + 1).ToString() + "行" + (start_col + 1).ToString() + "列\n";
            richTextBox1.Text = richTextBox1.Text + "[通知] 正在读取全班同学学号\n";

            string[] student_ids = new string[50];
            int student_amount;
            for (int i = 1; ; i++)
            {
                if (Wbook.Worksheets[0].Cells[start_row + i, start_col].Value == null)
                {
                    student_amount = i;
                    break;
                }
                string cur_student_id = Wbook.Worksheets[0].Cells[start_row + i, start_col].Value.ToString();
                student_ids[i-1] = cur_student_id;
            }
            richTextBox1.Text = richTextBox1.Text + "[通知] 读取完成，您班共" + student_amount.ToString() + "名同学\n";

            using (StreamReader sr = new StreamReader("data.txt"))
            {
                string line;

                // 从文件读取并显示行，直到文件的末尾 
                while ((line = sr.ReadLine()) != null)
                {
                    Console.WriteLine(line);
                }
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
