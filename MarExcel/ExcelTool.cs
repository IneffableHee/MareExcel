using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;

namespace MarExcel
{
    class ExcelTool
    {
        public string GetExcelUrl(Button btn, TextBox tb)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = @"excel文件|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            string path = openFileDialog.FileName;
            if (path != "")
            {
                tb.Text = path;
                tb.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#66A791"));
                string again;
                again = "重新导入";
                btn.Content = again;
            }
            return path;
        }

        public DataTable GetExcelTableByOleDB(string strExcelPath, string tablename)
        {
            try
            {
                System.Console.WriteLine("开始--");
                DataTable dtExcel = new DataTable(tablename);
                System.Console.WriteLine("DataTable dtExcel = new DataTable()");
                //获取文件扩展名
                System.Console.WriteLine("获取文件扩展名");
                string strExtension = System.IO.Path.GetExtension(strExcelPath);
                string strFileName = System.IO.Path.GetFileName(strExcelPath);
                System.Console.WriteLine("Excel的连接");

                //数据表
                dtExcel = ReadExcelByOLEDB(strExcelPath, strExtension, strFileName, dtExcel);
                return dtExcel;
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                System.Console.WriteLine(ex.ToString());
                return null;
            }
        }

        public DataTable ReadExcelByOLEDB(string strExcelPath, string strExtension, string strFileName, DataTable dt)
        {
            DataSet ds = new DataSet();
            //Excel的连接
            OleDbConnection objConn = null;
            switch (strExtension)
            {
                case ".xls":
                    //PrintMessage("Excel 2003");
                    objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"");
                    break;
                case ".xlsx":
                    //PrintMessage("Excel 2007");
                    objConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;\"");
                    break;
                default:
                    //PrintMessage("Excel null");
                    objConn = null;
                    break;
            }
            if (objConn == null)
            {
                //PrintMessage("objConn == null");
                return null;
            }
            //PrintMessage("Excel Open");
            try
            {
                objConn.Open();
            }
            catch (System.Net.Sockets.SocketException ex)
            {
                //PrintMessage(ex.ToString());
            }

            //获取Excel中所有Sheet表的信息
            DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            //获取Excel的第一个Sheet表名
            string tableName = schemaTable.Rows[0][2].ToString().Trim();
            //PrintMessage(tableName);
            //MessageBox.Show(tableName);
            string strSql = "select * from [" + tableName + "]";
            //PrintMessage(strSql);
            //获取Excel指定Sheet表中的信息
            OleDbDataAdapter myData = new OleDbDataAdapter(strSql, objConn);
            myData.Fill(ds);//填充数据
            objConn.Close();
            //PrintMessage(" ReadExcelByOLEDB---   行数：" + ds.Tables[0].Rows.Count + ";" + "列数" + ds.Tables[0].Columns.Count);
            //PrintMessage("\nReadExcelByOLEDB End--");
            dt = ds.Tables[0];
            return dt;
        }

        //检测导入的表是否正确
        public bool CheckExcel(DataTable khxxData, DataTable dkmxData)
        {
            if (khxxData.Rows.Count <= 0)
            {
                if (dkmxData.Rows.Count <= 0)
                {
                    System.Console.WriteLine("客户信息表、贷款明细表为空表，请重新导入！");
                    return false;
                }
                System.Console.WriteLine("客户信息表为空表，请重新导入！");
                return false;
            }

            ExcelTool et = new ExcelTool();
            string[] strKhxx = new string[] { "分区一", "分区二", "分区三", "证件号码" , "客户名称",
                                           "客户类型", "客户编码" , "客户号", "评定审批日期", "信用等级",
                                           "授信额度", "年审审批日期", "上次年审信用等级", "评定批准金额" };
            string[] strDkmm = new string[] { "分区一", "分区二", "分区三", "证件号码" , "客户名称",
                                           "信用等级", "授信额度" , "五级分类", "贷款日期", "到期日期",
                                           "结欠金额", "贷款金额" };

            if ((CheckTableStr(khxxData, strKhxx, "客户信息表")) == 0)
            {
                System.Console.WriteLine("缺少字段！");
                return false;
            }
            else if ((CheckTableStr(dkmxData, strDkmm, "贷款明细表")) == 0)
            {
                System.Console.WriteLine("缺少字段！");
                return false;
            }
            else
            {
                System.Console.WriteLine("检测成功！");
                return true;
            }

            //et.CreateTableColumnsByRowName(dt,);

        }

        //检测是否缺少关键字段
        private int CheckTableStr(DataTable dt, string[] checkStr, string tbName)
        {
            int flag = 1;
            string outStr = tbName;
            outStr += "缺少关键字段：";
            foreach (string s in checkStr)
            {
                int sig = 1;
                if (s != null)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {

                        if (s == dt.Columns[i].ColumnName.ToString())
                        {
                            sig = 0;
                            break;
                        }
                    }
                }
                if (sig == 1)
                {
                    flag = 0;
                    outStr += s;
                    outStr += "、";
                }
            }
            if (flag == 0)
            {
                outStr = outStr.TrimEnd('、');
                outStr += "，请检查后重新导入数据表！";
                System.Console.WriteLine(outStr);
            }
            return flag;
        }

        public DataTable CreateTableColumnsByRowName(DataTable dt, string[] str, int len)
        {

            DataView dv = dt.DefaultView;
            DataTable dt2 = dv.ToTable(true, str);
            foreach (DataRow dr in dt2.Rows)
            {
                for (int i = 0; i < len; i++)
                {
                    System.Console.Write(dr[i].ToString() + "\t");
                }
                System.Console.Write("\n");
            }

            return dt2;
        }

        public ArrayList GetTableColumnByIndex(DataTable dt, int arg)
        {
            ArrayList ls = new ArrayList();
            foreach (DataRow dr in dt.Rows)
            {
                if (!ls.Contains(dr[arg]))
                {
                    ls.Add(dr[arg]);
                }
            }
            return ls;
        }
    }
}
