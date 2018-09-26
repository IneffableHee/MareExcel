using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MarExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private string dkmxUrl;
        private string khxxUrl;
        private string savePath;
        private XYGC XYGC;
        public MainWindow()
        {
            InitializeComponent();
            XYGC = new XYGC();
        }

        public void DragWindow(object sender, MouseButtonEventArgs args)
        {
            this.DragMove();
        }

        private void Btn_Run_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Hello~~~~");
            string khxxUrl = @"G:\Desktop\源数据文件\客户信息.xls";
            string dkxxUrl = @"G:\Desktop\源数据文件\贷款信息.xls";

            if (khxxUrl == "" || dkxxUrl == "")
            {
                MessageBox.Show("请导入源数据表");
                return;
            }

            System.Console.WriteLine("开始生成--打开文件");
            System.Console.WriteLine("打开客户信息表 ");
            ExcelTool et = new ExcelTool();
            DataTable khxxData = et.GetExcelTableByOleDB(khxxUrl, "客户信息");
            for (int i = 4; i < 21; i++)
            {
                System.Threading.Thread.Sleep(10);
                System.Console.WriteLine(i + "\r\n");
            }
            System.Console.WriteLine("客户信息表行数：" + khxxData.Rows.Count + ";" + "列数" + khxxData.Columns.Count);
            System.Console.WriteLine("打开贷款明细表 ");
            DataTable dkmxData = et.GetExcelTableByOleDB(dkxxUrl, "贷款明细");
            for (int i = 20; i < 41; i++)
            {
                System.Threading.Thread.Sleep(10);
                System.Console.WriteLine(i + "\r\n");
            }
            System.Console.WriteLine("贷款明细表行数：" + dkmxData.Rows.Count + ";" + "列数" + dkmxData.Columns.Count);
            if (et.CheckExcel(khxxData, dkmxData) == false)
            {
                //数据导入失败，提示后重新导入。
                return;
            }

            //获取行政区划表
            string[] str = new string[] { "分区一", "分区二", "分区三" };
            DataTable xzdt = et.CreateTableColumnsByRowName(khxxData, str, 3);
            DataTable areaData = XYGC.SelectArea(xzdt);
            for (int i = 40; i < 51; i++)
            {
                System.Threading.Thread.Sleep(10);
                System.Console.WriteLine(i + "------\r\n");
            }
            XYGC.XXGCDataStatistics(khxxData, dkmxData, areaData);
            for (int i = 50; i < 75; i++)
            {
                System.Threading.Thread.Sleep(10);
                System.Console.WriteLine(i + "\r\n");
            }
        }

        private void BtnSetPath_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn)) return;
            ExcelTool et = new ExcelTool();
            savePath =  et.GetExcelUrl(btn, TBoxPath);
        }

        private void BtnDkmxImport_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn)) return;
            ExcelTool et = new ExcelTool();
            dkmxUrl = et.GetExcelUrl(btn, TBoxDkmxUrl);
        }

        private void BtnKhxxImport_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn)) return;
            ExcelTool et = new ExcelTool();
            khxxUrl = et.GetExcelUrl(btn, TBoxKhxxUrl);
        }

        private void BtnMini_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("您确定要退出吗？", "关闭程序", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                Application.Current.Shutdown();
            }
        }
    }
}
