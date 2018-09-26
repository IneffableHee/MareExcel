using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.Forms.MessageBox;

namespace MarExcel
{
    class XYGC
    {
        private DataTable khxxData;
        private DataTable dkmxData;
        private DataTable areaData;
        private string tempUrl;
        private string saveUrl;
        private string saveName;
        private Excel.Application excelApp;
        private Excel.Sheets sheets;
        private Excel.Workbook workbook;

        public XYGC()
        {
            System.Console.WriteLine("XYGC");
        }

        public XYGC(DataTable khxxDt, DataTable dkmxDt, DataTable areaDt, string saveName)
        {
            System.Console.WriteLine("XYGC");
            khxxData = khxxDt;
            dkmxData = dkmxDt;
            areaData = areaDt;
            tempUrl = @"G:\\Desktop\\源数据文件\\template.xlsx";
            saveUrl = "G:\\Desktop\\" + saveName + ".xlsx";
            Console.WriteLine(saveUrl);
            ExcelAppInit();
        }

        public void ExcelAppInit()
        {
            try
            {
                excelApp = new Excel.Application();//实例化Excel对象
                if (excelApp == null)
                {
                    MessageBox.Show("Excel无法启动");
                    return;
                }
                object missing = System.Reflection.Missing.Value;//获取缺少的object类型值
                excelApp.Application.DisplayAlerts = false;//不显示提示对话框  
                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                workbook = excelApp.Workbooks.Open(tempUrl,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);//打开Excel 
                sheets = workbook.Worksheets;//实例表格

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }
        }

        public void WriteData(DataTable data, string sheetname, int row, int column)
        {
            Excel.Worksheet worksheet = (Excel.Worksheet)sheets[sheetname];
            int t = row;
            foreach (DataRow dr in data.Rows)
            {
                for (int i = 0; i < data.Columns.Count; i++)
                {
                    int col = column + i;
                    if (dr[i] != null)
                    {
                        Console.WriteLine("set data row:" + t + ",col:" + col);
                        worksheet.Cells[t, column + i].value = dr[i];
                    }
                }
                t++;
            }
            //worksheet.Cells[row, column].value = data;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

        }

        public void SaveExcel()
        {
            try
            {
                Console.WriteLine("saveUrl111" + saveUrl);
                workbook.SaveAs(saveUrl);//保存工作表
                Console.WriteLine("saveUrl222");
                workbook.Close(false);//关闭工作表
                Console.WriteLine("saveUrl333");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
                Console.WriteLine("saveUrl444");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                Console.WriteLine("saveUrl555");
                excelApp.Quit();
                GC.Collect();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message.ToString());
            }

        }

        public DataTable SelectArea(DataTable xzdt)
        {

            return xzdt;
        }

        //生成信用工程表
        //khxxData：客户信息表
        //dkmxData：贷款明细表
        //town：乡镇名称
        public void XXGCDataStatistics(DataTable khxxData, DataTable dkmxData, DataTable areaData)
        {
            ExcelTool et = new ExcelTool();
            ArrayList xiang = et.GetTableColumnByIndex(areaData, 0);
            if (xiang.Count > 1)
            {
                //多个乡镇
                foreach (object o in xiang)
                {
                    string s = "分区一 = '" + o.ToString() + "'";
                    Console.WriteLine(s);
                    DataView dv = areaData.DefaultView;
                    dv.RowFilter = s;
                    DataTable newTable = dv.ToTable();
                    foreach (DataRow dr in newTable.Rows)
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            System.Console.Write(dr[i].ToString() + "\t");
                        }
                        System.Console.Write("\n");
                    }

                    //XYGC xygc = new XYGC(khxxData, dkmxData, newTable, o.ToString());
                    //xygc.setExcelData();
                }

            }
            else
            {
                //XYGC xygc = new XYGC(khxxData, dkmxData, areaData, xiang[0].ToString());
                //xygc.setExcelData();
            }
        }

        //public void WriteData(string data, string sheetname, int row, int column)
        //{
        //    ExcelTool et = new ExcelTool();
        //    et.WriteData("test", "G:\\Desktop\\源数据文件\\贵州省农村信用社信用乡（镇、街道）、信用村（社区）、信用组评定验收、年审套表.xlsx", 10, 10);
        //    try
        //    {

        //        Excel.Application excelApp = new Excel.Application();//实例化Excel对象
        //        if (excelApp == null)
        //        {
        //            MessageBox.Show("Excel无法启动");
        //            return;
        //        }
        //        object missing = System.Reflection.Missing.Value;//获取缺少的object类型值
        //        excelApp.Application.DisplayAlerts = false;//不显示提示对话框  
        //        excelApp.Visible = false;
        //        excelApp.ScreenUpdating = false;

        //        workbook = excelApp.Workbooks.Open(sheetname,
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //        Type.Missing, Type.Missing);//打开Excel 
        //        Excel.Sheets sheets = workbook.Worksheets;//实例表格

        //        Excel.Worksheet worksheet = (Excel.Worksheet)sheets["1-2"];//第一个表格

        //        worksheet.Cells[row, column].value = data;
        //        workbook.SaveAs("G:\\Desktop\\test.xlsx");//保存工作表
        //        workbook.Close(false, missing, missing);//关闭工作表
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
        //        excelApp.Quit();
        //        GC.Collect();
        //    }
        //    catch (Exception e)
        //    {
        //        System.Windows.Forms.MessageBox.Show(e.Message.ToString());
        //    }
        //}
    }
}
