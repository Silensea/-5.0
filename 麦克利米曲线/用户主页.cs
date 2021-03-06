﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 麦克利米曲线
{
    public partial class 用户主页 : Form
    {
        public 用户主页()
        {
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            InitializeComponent();
        }

        private void ovalShape1_Click(object sender, EventArgs e)
        {
            精准查询 f5 = new 精准查询();
            f5.Owner = this;
            Hide();
            f5.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            推荐 f1 = new 推荐();
            f1.Owner = this;
            Hide();
            f1.ShowDialog();
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.Hide();
                welcome fwel = new welcome();
                fwel.ShowDialog();
                //this.Close();
                //this.Owner.Show();
            }
            catch (Exception) { }
        }

        private void labelSearch_Click(object sender, EventArgs e)
        {
            精准查询 f5 = new 精准查询();
            f5.Owner = this;
            Hide();
            f5.ShowDialog();
        }

        private void labelSelect_Click(object sender, EventArgs e)
        {
            检索目录 f6 = new 检索目录();
            f6.Owner = this;
            Hide();
            f6.ShowDialog();
        }

        private void labelUserC_Click(object sender, EventArgs e)
        {
            用户中心 f_user = new 用户中心();
            f_user.Owner = this;
            Hide();
            f_user.ShowDialog();
        }

        private void 用户主页_Load(object sender, EventArgs e)
        {

        }
    }
}
