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
            int row = 0;
            int col = 0;

            while (true)
            {
                if (Wbook.Worksheets[0].Cells[row, col].Value != null)
                {
                    if (String.Compare(Wbook.Worksheets[0].Cells[row, col].Value.ToString(), "学号") == 0)
                    {
                        break;
                    }
                }
                
                if (col < Wbook.Worksheets[0].Rows.Count)
                {
                    col++;
                }
                else
                {
                    col = 0;
                    row++;
                }

            }

            richTextBox1.Text = richTextBox1.Text + "学号在" + (row + 1).ToString() + "行" + (col + 1).ToString() + "列";

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
