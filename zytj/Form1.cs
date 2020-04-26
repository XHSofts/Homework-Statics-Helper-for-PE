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
    }
}
