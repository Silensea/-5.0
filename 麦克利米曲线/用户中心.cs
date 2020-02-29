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
using Microsoft.Office.Interop;
using Word=Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace 麦克利米曲线
{
    public partial class 用户中心 : Form
    {
        public 用户中心()
        {
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            InitializeComponent();
            fill_info(登录界面.ID);
            InfoGroupBox.BringToFront();
        }
        SqlConnection myconn = new SqlConnection("database=InternJi;data source=DESKTOP-E8FREVT\\VACE;integrated security=true");
        string mysql;
        DataSet mydataset = new DataSet();

        private void fill_info(int ID)
        {
            try
            {
                mysql = "select seeker_name,seeker_sex,seeker_tel,seeker_status,seeker_degree,seeker_school,seeker_major,seeker_religion,seeker_describe from seeker where seeker_ID=" + ID;
                SqlDataAdapter myadapter = new SqlDataAdapter(mysql, myconn);
                mydataset.Clear();
                myadapter.Fill(mydataset, "info");
                labelUname.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][0]);
                labelName.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][0]);
                labelSex.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][1]);
                labelTel.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][2]);
                labelStatus.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][3]);
                labelDegree.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][4]);
                labelSchool.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][5]);
                labelMajor.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][6]);
                labelReligion.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][7]);
                labelDesc.Text = Convert.ToString(mydataset.Tables["info"].Rows[0][8]);
            }
            catch (Exception)
            {
                ;//MessageBox.Show("请先完善个人信息！");
                //return;
            }
        }

        private void produce_CV(int ID)
        {
            //新建文档
            Word.Application newapp = new Word.Application();
            Word.Document newdoc;
            object nothing = System.Reflection.Missing.Value;//函数的默认参数
            newdoc = newapp.Documents.Add(ref nothing, ref nothing, ref nothing, ref nothing);//生成一个word文档
            newapp.Visible = true;

            //页面设置
            newdoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            newdoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            newdoc.PageSetup.TopMargin = 57.0f;
            newdoc.PageSetup.BottomMargin = 57.0f;
            newdoc.PageSetup.LeftMargin = 57.0f;
            newdoc.PageSetup.RightMargin = 57.0f;

            //文档填充
            try
            {
                newapp.Selection.Font.Name = "黑体";
                newapp.Selection.Font.Size = 22;
                newapp.Selection.Font.Bold = 1;
                newapp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                newapp.Selection.TypeText("我的简历");
                newapp.Selection.TypeParagraph();
                newapp.Selection.Font.Name = "宋体";
                newapp.Selection.Font.Size = 14;
                newapp.Selection.Font.Bold = 0;
                newapp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("姓名：" + labelName.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("性别：" + labelSex.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("联系电话：" + labelTel.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("政治面貌：" + labelStatus.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("学历：" + labelDegree.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("毕业院校：" + labelSchool.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("专业：" + labelMajor.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("宗教：" + labelReligion.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText("实习经历：");
                newapp.Selection.TypeParagraph();
                newapp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                newapp.Selection.TypeText(dateTimePicker1.Text.ToString() + "——" + dateTimePicker2.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText(textBox4.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText(dateTimePicker3.Text.ToString() + "——" + dateTimePicker4.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.TypeText(textBox5.Text.ToString());
                newapp.Selection.TypeParagraph();
                newapp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                newapp.Selection.TypeText("自我能力陈述和评价：" + textBoxDesc.Text.ToString());
            }
            catch (Exception)
            {
                MessageBox.Show("请先完善个人信息！");
                return;
            }
            //文档保存
            try
            {
                object name = "D:\\" + ID + ".doc";
                newdoc.SaveAs(ref name, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing);
                MessageBox.Show("已保存至" + name.ToString(), "生成简历");

            }
            catch (Exception)
            {
                MessageBox.Show("文件导出异常,请重试!");
                return;
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            mysql = "update seeker set seeker_name='" + textBoxName.Text + "' , seeker_sex='" + comboBoxSex.SelectedItem + "' , seeker_tel='" + textBoxTel.Text + "' , seeker_status='" + comboBoxStatus.SelectedItem + "' , seeker_degree='" + comboBoxDegree.SelectedItem + "' , seeker_school='" + comboBoxSchool.SelectedItem + "' , seeker_major='" + comboBoxMajor.SelectedItem + "', seeker_religion='" + comboBoxReligion.SelectedItem + "', seeker_describe='" + textBoxDesc.Text + "' where seeker_ID=" + 登录界面.ID;
            SqlCommand mycmd = new SqlCommand(mysql, myconn);
            myconn.Open();
            try
            {
                mycmd.ExecuteNonQuery();
                MessageBox.Show("修改成功！", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            myconn.Close();
            fill_info(登录界面.ID);
           
            /*EditGroupBox.Hide();
            InfoGroupBox.Show();*/
            InfoGroupBox.BringToFront();
        }

        private void InfoLabel_Click(object sender, EventArgs e)
        {
            fill_info(登录界面.ID);
            InfoGroupBox.BringToFront();
            /*InfoGroupBox.Show();
            RecordGroupBox.Hide();
            CVGroupBox.Hide();
            EditGroupBox.Hide();*/
        }

        private void CVLabel_Click(object sender, EventArgs e)
        {
            CVGroupBox.BringToFront();
            /*CVGroupBox.Show();
            RecordGroupBox.Hide();
            InfoGroupBox.Hide();
            EditGroupBox.Hide();*/
        }

        private void RecordLabel_Click(object sender, EventArgs e)
        {
            RecordGroupBox.BringToFront();
            mysql = "select apply_time as 申请时间, CV_path as 投递路径, rec_post as 申请职位, emp_status as 简历状态 from employ where seeker_ID=" + 登录界面.ID;
            SqlDataAdapter myadapter1 = new SqlDataAdapter(mysql, myconn);
            myadapter1.Fill(mydataset, "简历信息");
            dataGridView1.DataSource = mydataset.Tables["简历信息"];

            SqlCommand mycmd = new SqlCommand("seeker_CV", myconn);
            mycmd.CommandType = CommandType.StoredProcedure;
            SqlParameter ID = new SqlParameter("@ID", SqlDbType.Int);
            mycmd.Parameters.Add(ID);
            ID.Value = 登录界面.ID;
            SqlParameter count = new SqlParameter("@count", SqlDbType.Int);
            mycmd.Parameters.Add(count);
            count.Direction = ParameterDirection.Output;
            SqlParameter accept = new SqlParameter("@accept", SqlDbType.Int);
            mycmd.Parameters.Add(accept);
            accept.Direction = ParameterDirection.Output;
            SqlParameter unaccept = new SqlParameter("@unaccept", SqlDbType.Int);
            mycmd.Parameters.Add(unaccept);
            unaccept.Direction = ParameterDirection.Output;
            myconn.Open();
            try
            {
                mycmd.ExecuteScalar();
                label10.Text = "共有" + count.Value.ToString() + "条投递记录，" + accept.Value.ToString() + "条已通过，" + unaccept.Value.ToString() + "条未通过";
            }
            catch (Exception)
            {
                MessageBox.Show("没有相关数据信息！");
                myconn.Close();
                return;
            }
            myconn.Close();
        }

        private void HomeLabel_Click(object sender, EventArgs e)
        {
            Close();
            Owner.Show();
            //求职者主页 f_see = new 求职者主页();
            //f_see.ShowDialog();
        }

        private void CVProButton_Click(object sender, EventArgs e)
        {
            produce_CV(登录界面.ID);
        }

        private void LogoutButton_Click(object sender, EventArgs e)
        {
            Close();
            welcome f_wel = new welcome();
            f_wel.ShowDialog();
            
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            EditGroupBox.BringToFront();
            /*InfoGroupBox.Hide();
            CVGroupBox.Hide();
            RecordGroupBox.Hide();
            EditGroupBox.Show();*/
        }
    }
}
