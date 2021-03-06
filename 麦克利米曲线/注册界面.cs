﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
//using Microsoft.Office.Interop.Word;

namespace 麦克利米曲线
{
    public partial class 注册界面 : Form
    {
        public 注册界面()
        {
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            InitializeComponent();
        }
        SqlConnection myconn = new SqlConnection("database=InternJi;data source=DESKTOP-E8FREVT\\VACE;integrated security=true");
        public static string user_email = "";
        private void button1_Click(object sender, EventArgs e)
        {
            if (PasswBox1.Text != PasswBox2.Text) 
                MessageBox.Show("两次密码不一致，请重新输入！");
            else if (EmailBox.Text == "" || PasswBox1.Text == "" || PasswBox2.Text == "") 
                MessageBox.Show("用户名、密码不能为空！");
            else
            {
                SqlCommand mycmd = new SqlCommand("register_new", myconn);
                mycmd.CommandType = CommandType.StoredProcedure;
                SqlParameter user_type = new SqlParameter("@user_type ", SqlDbType.SmallInt);
                mycmd.Parameters.Add(user_type);
                SqlParameter email = new SqlParameter("@email ", SqlDbType.VarChar, 40);
                mycmd.Parameters.Add(email);
                SqlParameter password = new SqlParameter("@password ", SqlDbType.VarChar, 20);
                mycmd.Parameters.Add(password);
                email.Value = EmailBox.Text;
                password.Value = PasswBox1.Text;
                if (radioButton1.Checked == true) 
                    user_type.Value = 1;
                else if (radioButton2.Checked == true) 
                    user_type.Value = 2;
                myconn.Open();
                    mycmd.ExecuteNonQuery();
                myconn.Close();
                user_email = EmailBox.Text;
                //Hide();
                if (radioButton1.Checked)
                {
                    注册成功 f1 = new 注册成功();
                    f1.Owner = this.Owner;
                    Close();
                    f1.ShowDialog();
                }
                else
                {
                    Close();
                    Owner.Show();
                }
            }
        }

        private void EmailBox_TextChanged(object sender, EventArgs e)
        {
            if (!(EmailBox.Text.Contains('@') && EmailBox.Text.Contains('.')))
                HintLabel1.Text = "请输入正确的邮箱格式";
            else
                HintLabel1.Text = "";
        }

        private void PasswBox1_TextChanged(object sender, EventArgs e)
        {
            if (PasswBox1.TextLength < 6)
                HintLabel2.Text = "密码过短";
            else if (PasswBox1.TextLength > 20)
                HintLabel2.Text = "密码过长";
            else
                HintLabel2.Text = "";
        }

        private void PasswBox2_TextChanged(object sender, EventArgs e)
        {
            if (PasswBox2.Text != PasswBox1.Text)
                HintLabel3.Text = "请输入相同密码";
            else
                HintLabel3.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Owner.Show();
        }
    }
}
