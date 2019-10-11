using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;

namespace EXCEL操作
{
    public partial class Form1 : Form
    {
        public bool inse_Switch = false;
        public Form1()
        {
          
            InitializeComponent();




        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {
             
        }
        //创建表格
        private void button5_Click(object sender, EventArgs e)
        {
            var File_Path = "学生信息.xls";
            string sql = "CREATE TABLE 学生信息([学号] INT,[姓名] VarChar,[班级] VarChar,[电话号码] VarChar,[状态] VarChar)";
            //执行SQL语句
            Excel2.Update(sql, File_Path);
        }
        //查询
        private void button1_Click(object sender, EventArgs e)
        {
            //判断输入框是否为空
            if(this.textBox1.Text=="")
            {
                //查询全部
                //Excel路径
                var File_path = "学生信息.xls";
                //构建SQL语句
                string sql = "select 学号,姓名,班级,电话号码 from [学生信息$]  where 状态 = '正常'";
                //执行SQL语句
                this.dataGridView1.DataSource = Excel.Excel2.GetDataTable(sql, File_path);
            }
            else
            {
                //按照条件查询
                //Excel路径
                var File_path = "学生信息.xls";
                //构建SQL语句
                string sql = "select 学号,姓名,班级,电话号码 from [学生信息$]  where 状态 ="+this.textBox1.Text;
                //执行SQL语句
                this.dataGridView1.DataSource = Excel.Excel2.GetDataTable(sql, File_path);
            }
        }
        //插入数据
        private void button2_Click(object sender, EventArgs e)
        {
            inse_Switch = true;
            this.groupBox1.Enabled = true;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        //提交数据
        private void button6_Click(object sender, EventArgs e)
        {
            //Excel路径
            var File_Path = "学生信息.xls";
            //SQL语句
            string sql = "";
            //
            if (this.textBox2.Text.Trim() == "")

	{
		 MessageBox.Show("请输入姓名");
                return;
	}
            if (inse_Switch==true)
            {
                //SQL语句//插入数据
                sql = "insert into [学生信息$](学号,姓名,班级,电话号码,状态)  values ({0},'{1}','{2}','{3}','{4}')";
                sql = string.Format(sql, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text, this.textBox5.Text, "正常");
                inse_Switch = false;
            }
            else
            {
                //修改数据
                sql = "update [学生信息$] set  姓名 ='{0}' ,班级 = '{1}',电话号码= '{2}',状态 = '正常'   where   学号 = {3}";
                sql = string.Format(sql, this.textBox3.Text, this.textBox4.Text, this.textBox5.Text, this.textBox1.Text);
            }

           
            //提交
            Excel2.Update(sql, File_Path);
            //查询全部信息
            this.textBox1.Text = "";//赋值为空，查询全部信息
            button1_Click(null, null);//触发按钮，查询信息

        }
        //删除按钮//其实在文件中还是存在的
        private void button3_Click(object sender, EventArgs e)
        {
            //指定文件位置
            var filepath = "学生信息.xls";
            //判断学号是否为空
            if (this.textBox1.Text.Trim() == "")
            {
                MessageBox.Show("请输入学号");
                return;
            }
            else
            {
                //构建查询语句
                string sql = "update [学生信息$] set 状态 = '删除' where 学号 = {0}";
                sql = string.Format(sql, this.textBox1.Text);
                //更新数据
                Excel2.Update(sql, filepath);
                //清空查询编号
                this.textBox1.Text = "";
                //执行查询操作
                button1_Click(null, null);
            }
        }
        //修改
        private void button4_Click(object sender, EventArgs e)
        {
            
            this.groupBox1.Enabled = true;
        }
    }
}
