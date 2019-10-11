using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//引用数据库操作类
using System.Data;
using System.Data.OleDb;

namespace Excel
{
    public class Excel2
    {
        //在根目录下的bin文件下创建一个Excel表格
        
        public static int Update (string sql,string path)
        {
            //构建连接语句
            string sConnectionString = 
                "Provider = Microsoft.ACE.OLEDB.12.0;" + 
                "Data Source = " + path + ";" + 
                "Extended Properties = 'Excel 8.0;HDR = Yes;IMEX = 0'";
            //IMEX=0 为汇出模式，这个模式Excle只能用作"写入"用途
            //IMEX=1 为汇入模式，这个模式Excle只能用作"读取"用途
            //IMEX=2 为链接模式, 这个模式Excle同时支持"读写"用途
            //HDR=Yes 创建表头
            //使用连接语句连接Excel-连接上数据库Connection
            using (OleDbConnection ole_cnn = new OleDbConnection(sConnectionString))
            {
                //打开链接
                ole_cnn.Open();
                //创建编辑器开始编辑SQL命令
                using (OleDbCommand ole_cmd = ole_cnn.CreateCommand())
                {
                    //执行sql语句
                    ole_cmd.CommandText = sql;
                    //执行SQL命令- 返回受影响的函数
                    return ole_cmd.ExecuteNonQuery();
                }
            }
            
        }

        //输入在数据库里面查询到的内容，再通过Excel表格输出
        public static DataTable GetDataTable(string sql,string path)
        {
            //构建连接语句
            string sConnectionString =
                "Provider = Microsoft.ACE.OLEDB.12.0;" +
                "Data Source = " + path + ";" +
                "Extended Properties = 'Excel 8.0;HDR = Yes;IMEX = 0'";
            //使用连接语句连接Excel-连接上数据库Connection
            using (OleDbConnection ole_cnn = new OleDbConnection(sConnectionString))
            {
                //打开链接
                ole_cnn.Open();
                //创建编辑器开始编辑SQL命令
                using (OleDbCommand ole_cmd = ole_cnn.CreateCommand())
                {
                    //执行sql语句
                    ole_cmd.CommandText = sql;
                  //执行语句--接受数据
                    using (OleDbDataAdapter dapter = new OleDbDataAdapter(ole_cmd))
                    {
                        //创建对应内存表格结构数据-为了接受返回表格
                        DataSet dr = new DataSet();//创建空白表格
                        //填充表格数据
                        dapter.Fill(dr);
                        //返回查询到的内容
                        return dr.Tables[0];
                    }
                }
            }
        }
    }
}
