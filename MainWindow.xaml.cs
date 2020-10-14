using System;
using System.Collections.Generic;
using System.IO;
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
using System.Xml;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.Diagnostics;
using System.Threading;

namespace BuildReportApp
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadXMLConfig();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (wb != null)
                    wb.Close(true, Missing.Value, Missing.Value);
                if (excel != null)
                {
                    excel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                }
            }
            catch(Exception ex)
            {

            }
            finally
            {
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                }
                GC.Collect();
                this.Close();
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }
        string sourceFilePath = AppDomain.CurrentDomain.BaseDirectory + "Template.xls";
        Microsoft.Office.Interop.Excel.Application excel;
        Workbook wb;
        private void btnSaveReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //先关掉Excel
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                    GC.Collect();
                }

                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.Filter = "xls files(*.xls)|*.xls";
                dlg.FileName = "ND1633IQC测试报告" +  System.DateTime.Now.ToString("yyyy-MM-dd") + ".xls";
                //dlg.InitialDirectory = "D:\\";
                dlg.AddExtension = false;
                dlg.RestoreDirectory = true;
                System.Nullable<bool> result = dlg.ShowDialog();
                if (result == true)
                {
                    string filePath = dlg.FileName.ToString();
                    if(File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    File.Copy(sourceFilePath, filePath);
                    tbFilePath.Text = filePath;
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.Visible = false;
                    excel.UserControl = true;
                    object missing = System.Reflection.Missing.Value;
                    wb = excel.Application.Workbooks.Open(tbFilePath.Text.Trim(), missing, missing, missing, missing, missing, missing, missing, missing,
                                  missing, missing, missing, missing, missing, missing);
                    Worksheet worksheet = (Worksheet)wb.Worksheets.get_Item(1);//取得第一个工作簿
                    worksheet.Cells[1, 2] = DateTime.Now.ToString("yyyy/MM/dd");
                    worksheet.Cells[2, 2] = "1633-PH";
                    excel.DisplayAlerts = false;
                    wb.Save();

                    AutoClosedMsgBox.Show("报告生成成功！", "提示", 1000, 64);
                    SnList.Clear();
                    index = 8;
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message,"提示", MessageBoxButton.OK, MessageBoxImage.Information);
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                }
            }
        }
        int index = 8;
        List<string> SnList = new List<string>();
        private void btnSaveData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(tbFilePath.Text.Trim()))
                {
                    MessageBox.Show("报告保存路径不能为空！请先生成测试报告！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (string.IsNullOrEmpty(tbSn.Text.Trim()))
                {
                    MessageBox.Show("PCB条码不能为空！请先扫条码！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (!File.Exists(tbFilePath.Text.Trim()))
                {
                    MessageBox.Show(string.Format("报告 {0} 不存在，请确认！", tbFilePath.Text.Trim()), "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (SnList.Contains(tbSn.Text.Trim()))
                {
                    MessageBox.Show(string.Format("条码 {0} 已存在！", tbSn.Text.Trim()), "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                btnSaveData.IsEnabled = false;
                btnBuildBatch.IsEnabled = false;
                //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                //excel.Visible = false;
                //excel.UserControl = true;
                //object missing = System.Reflection.Missing.Value;
                //Workbook wb = excel.Application.Workbooks.Open(tbFilePath.Text.Trim(), missing, missing, missing, missing, missing, missing, missing, missing,
                //              missing, missing, missing, missing, missing, missing);
                if (excel != null && wb != null)
                {
                    Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);//取得第一个工作簿
                                                                        //ws.Cells[1, 2] = DateTime.Now.ToString("yyyy/MM/dd");
                                                                        //ws.Cells[2, 2] = "1633-PH";
                    WriteTestData(ws, tbSn.Text.Trim());
                    //int offset = 1;
                    //ws.Cells[index, offset] = tbSn.Text.Trim(); offset++;
                    //Random ovrandom = new Random(); ws.Cells[index, offset] = GetRandomNumber(ovrandom, OVItem.minValue, OVItem.maxValue, OVItem.decimalLen); offset++;
                    //Random ocdsgrandom = new Random(); ws.Cells[index, offset] = GetRandomNumber(ocdsgrandom, OCDSGItem.minValue, OCDSGItem.maxValue, OCDSGItem.decimalLen); offset++;
                    //Random cuvrandom = new Random();
                    //foreach (var it in CUVItemList)
                    //{
                    //    Thread.Sleep(1);
                    //    ws.Cells[index, offset] = GetRandomNumber(cuvrandom, it.minValue, it.maxValue, it.decimalLen);
                    //    offset++;
                    //}
                    //Random covrandom = new Random();
                    //foreach (var it in COVItemList)
                    //{
                    //    Thread.Sleep(1);
                    //    ws.Cells[index, offset] = GetRandomNumber(covrandom, it.minValue, it.maxValue, it.decimalLen);
                    //    offset++;
                    //}
                    //Random cclrandom = new Random();
                    //Thread.Sleep(1);
                    //ws.Cells[index, offset] = GetRandomNumber(cclrandom, CCLItem.minValue, CCLItem.maxValue, CCLItem.decimalLen); offset++;
                    //Random balrandom = new Random();
                    //foreach (var it in BALItemList)
                    //{
                    //    Thread.Sleep(1);
                    //    ws.Cells[index, offset] = GetRandomNumber(balrandom, it.minValue, it.maxValue, it.decimalLen);
                    //    offset++;
                    //}
                    //Random nscrandom = new Random();
                    //double totalNSC = 0;
                    //foreach (var it in NSCItemList)
                    //{
                    //    Thread.Sleep(1);
                    //    double val = GetRandomNumber(nscrandom, it.minValue, it.maxValue, it.decimalLen);
                    //    ws.Cells[index, offset] = val;
                    //    totalNSC += val;
                    //    offset++;
                    //}
                    //ws.Cells[index, offset] = totalNSC;
                    //offset++;
                    //double totalSSC = 0;
                    //Random sscrandom = new Random();
                    //foreach (var it in SSCItemList)
                    //{
                    //    Thread.Sleep(1);
                    //    double val = GetRandomNumber(sscrandom, it.minValue, it.maxValue, it.decimalLen);
                    //    ws.Cells[index, offset] = val;
                    //    totalSSC += val;
                    //    offset++;
                    //}
                    //ws.Cells[index, offset] = totalSSC;
                    //offset++;
                    //ws.Cells[index, offset] = "pass";
                    index++;
                    excel.DisplayAlerts = false;
                    wb.Save();
                    AutoClosedMsgBox.Show("数据保存成功！", "提示", 1000, 64);
                    SnList.Add(tbSn.Text.Trim());
                    btnBuildBatch.IsEnabled = true;
                    btnSaveData.IsEnabled = true;
                    tbSn.Text = string.Empty;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                btnBuildBatch.IsEnabled = true;
                btnSaveData.IsEnabled = true;
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                }
                GC.Collect();
            }
        }

        private bool WriteTestData(Worksheet ws,string barcode)
        {
            try
            {
                int offset = 1;
                ws.Cells[index, offset] = barcode; offset++;
                Random ovrandom = new Random(); ws.Cells[index, offset] = GetRandomNumber(ovrandom, OVItem.minValue, OVItem.maxValue, OVItem.decimalLen); offset++;
                Random ocdsgrandom = new Random(); ws.Cells[index, offset] = GetRandomNumber(ocdsgrandom, OCDSGItem.minValue, OCDSGItem.maxValue, OCDSGItem.decimalLen); offset++;
                Random cuvrandom = new Random();
                foreach (var it in CUVItemList)
                {
                    Thread.Sleep(1);
                    ws.Cells[index, offset] = GetRandomNumber(cuvrandom, it.minValue, it.maxValue, it.decimalLen);
                    offset++;
                }
                Random covrandom = new Random();
                foreach (var it in COVItemList)
                {
                    Thread.Sleep(1);
                    ws.Cells[index, offset] = GetRandomNumber(covrandom, it.minValue, it.maxValue, it.decimalLen);
                    offset++;
                }
                Random cclrandom = new Random();
                Thread.Sleep(1);
                ws.Cells[index, offset] = GetRandomNumber(cclrandom, CCLItem.minValue, CCLItem.maxValue, CCLItem.decimalLen); offset++;
                Random balrandom = new Random();
                foreach (var it in BALItemList)
                {
                    Thread.Sleep(1);
                    ws.Cells[index, offset] = GetRandomNumber(balrandom, it.minValue, it.maxValue, it.decimalLen);
                    offset++;
                }
                Random nscrandom = new Random();
                double totalNSC = 0;
                foreach (var it in NSCItemList)
                {
                    Thread.Sleep(1);
                    double val = GetRandomNumber(nscrandom, it.minValue, it.maxValue, it.decimalLen);
                    ws.Cells[index, offset] = val;
                    totalNSC += val;
                    offset++;
                }
                ws.Cells[index, offset] = totalNSC;
                offset++;
                double totalSSC = 0;
                Random sscrandom = new Random();
                foreach (var it in SSCItemList)
                {
                    Thread.Sleep(1);
                    double val = GetRandomNumber(sscrandom, it.minValue, it.maxValue, it.decimalLen);
                    ws.Cells[index, offset] = val;
                    totalSSC += val;
                    offset++;
                }
                ws.Cells[index, offset] = totalSSC;
                offset++;
                ws.Cells[index, offset] = "pass";

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        private void btnOpenReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //先关掉Excel
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                    GC.Collect();
                }
                wb = null;excel = null; 
                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.Filter = "xls files(*.xls)|*.xls";
                dlg.AddExtension = false;
                dlg.RestoreDirectory = true;
                System.Nullable<bool> result = dlg.ShowDialog();
                if (result == true)
                {
                    string filePath = dlg.FileName.ToString();
                    if (!filePath.Contains("ND1633IQC测试报告"))
                    {
                        if (MessageBoxResult.No == MessageBox.Show("选择的文件的文件名不包含特定字符，该文件可能与模板不一致，是否继续打开？", "提示", MessageBoxButton.YesNo, MessageBoxImage.Information))
                        {
                            return;
                        }
                    }
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.Visible = false;
                    excel.UserControl = true;
                    object missing = System.Reflection.Missing.Value;
                    wb = excel.Application.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing,
                                  missing, missing, missing, missing, missing, missing);
                    Worksheet worksheet = (Worksheet)wb.Worksheets.get_Item(1);//取得第一个工作簿
                    bool isfitTemplate = false;
                    if (((Range)worksheet.Cells[1, 1]).Text == "time" && ((Range)worksheet.Cells[2, 1]).Text == "model" && ((Range)worksheet.Cells[2, 2]).Text == "1633-PH")
                    {
                        tbFilePath.Text = filePath;
                        isfitTemplate = true;
                        int rowsNum = worksheet.UsedRange.Cells.Rows.Count;
                        index = 8;
                        SnList.Clear();
                        for (int i = index; i <= rowsNum; i++)
                        {
                            string sn = ((Range)worksheet.Cells[i, 1]).Text;
                            if (!string.IsNullOrEmpty(sn))
                            {
                                SnList.Add(sn);
                            }
                            index++;
                        }
                    }
                    else
                    {
                        isfitTemplate = false;
                    }

                    excel.DisplayAlerts = false;
                    wb.Save();
                    if (isfitTemplate)
                        AutoClosedMsgBox.Show("报告打开成功！", "提示", 1000, 64);
                    else
                        MessageBox.Show("打开的Excel文件格式和模板格式不一致，请检查！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                    GC.Collect();
                }
            }
        }
        public double GetRandomNumber(Random random,double minimum, double maximum, int Len)   //Len小数点保留位数
        {
            return Math.Round(random.NextDouble() * (maximum - minimum) + minimum, Len);
        }

        TestItem OVItem = new TestItem();
        TestItem OCDSGItem = new TestItem();
        List<TestItem> CUVItemList = new List<TestItem>();
        List<TestItem> COVItemList = new List<TestItem>();
        TestItem CCLItem = new TestItem();
        List<TestItem> BALItemList = new List<TestItem>();
        List<TestItem> NSCItemList = new List<TestItem>();
        List<TestItem> SSCItemList = new List<TestItem>();

        public void LoadXMLConfig()
        {
            try
            {
                string strFileName = AppDomain.CurrentDomain.BaseDirectory + "config_info.xml";
                if (!File.Exists(strFileName))
                {
                    return;
                }

                XmlDocument xmlDoc = new XmlDocument();

                xmlDoc.Load(strFileName);

                XmlNode xn = xmlDoc.SelectSingleNode("root");
                XmlNodeList nodelist = xn.ChildNodes;
                foreach (XmlNode item in nodelist)
                {
                    if (item.LocalName == "OV")
                    {
                        XmlNodeList OVList = item.ChildNodes;
                        foreach(XmlNode ov in OVList)
                        {
                            if(ov.LocalName == "item")
                            {
                                string cellIndexStr = ov.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = ov.SelectSingleNode("min").InnerText;
                                string maxValueStr = ov.SelectSingleNode("max").InnerText;
                                string decimalLenStr = ov.SelectSingleNode("decimal").InnerText;
                                OVItem.cellIndex = int.Parse(cellIndexStr);
                                OVItem.minValue = double.Parse(minValueStr);
                                OVItem.maxValue = double.Parse(maxValueStr);
                                OVItem.decimalLen = int.Parse(decimalLenStr);
                            }
                        }
                    }
                    else if (item.LocalName == "OCDSG")
                    {
                        XmlNodeList OCDSGList = item.ChildNodes;
                        foreach (XmlNode ocdsg in OCDSGList)
                        {
                            if (ocdsg.LocalName == "item")
                            {
                                string cellIndexStr = ocdsg.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = ocdsg.SelectSingleNode("min").InnerText;
                                string maxValueStr = ocdsg.SelectSingleNode("max").InnerText;
                                string decimalLenStr = ocdsg.SelectSingleNode("decimal").InnerText;
                                OCDSGItem.cellIndex = int.Parse(cellIndexStr);
                                OCDSGItem.minValue = double.Parse(minValueStr);
                                OCDSGItem.maxValue = double.Parse(maxValueStr);
                                OCDSGItem.decimalLen = int.Parse(decimalLenStr);
                            }
                        }
                    }
                    else if (item.LocalName == "CUV")
                    {
                        XmlNodeList CUVList = item.ChildNodes;
                        foreach (XmlNode cuv in CUVList)
                        {
                            if (cuv.LocalName == "item")
                            {
                                string cellIndexStr = cuv.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = cuv.SelectSingleNode("min").InnerText;
                                string maxValueStr = cuv.SelectSingleNode("max").InnerText;
                                string decimalLenStr = cuv.SelectSingleNode("decimal").InnerText;
                                TestItem cuvItem = new TestItem();
                                cuvItem.cellIndex = int.Parse(cellIndexStr);
                                cuvItem.minValue = double.Parse(minValueStr);
                                cuvItem.maxValue = double.Parse(maxValueStr);
                                cuvItem.decimalLen = int.Parse(decimalLenStr);
                                CUVItemList.Add(cuvItem);
                            }
                        }
                    }
                    else if (item.LocalName == "COV")
                    {
                        XmlNodeList COVList = item.ChildNodes;
                        foreach (XmlNode cov in COVList)
                        {
                            if (cov.LocalName == "item")
                            {
                                string cellIndexStr = cov.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = cov.SelectSingleNode("min").InnerText;
                                string maxValueStr = cov.SelectSingleNode("max").InnerText;
                                string decimalLenStr = cov.SelectSingleNode("decimal").InnerText;
                                TestItem covItem = new TestItem();
                                covItem.cellIndex = int.Parse(cellIndexStr);
                                covItem.minValue = double.Parse(minValueStr);
                                covItem.maxValue = double.Parse(maxValueStr);
                                covItem.decimalLen = int.Parse(decimalLenStr);
                                COVItemList.Add(covItem);
                            }
                        }
                    }
                    else if (item.LocalName == "CCL")
                    {
                        XmlNodeList CCLList = item.ChildNodes;
                        foreach (XmlNode ccl in CCLList)
                        {
                            if (ccl.LocalName == "item")
                            {
                                string cellIndexStr = ccl.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = ccl.SelectSingleNode("min").InnerText;
                                string maxValueStr = ccl.SelectSingleNode("max").InnerText;
                                string decimalLenStr = ccl.SelectSingleNode("decimal").InnerText;
                                CCLItem.cellIndex = int.Parse(cellIndexStr);
                                CCLItem.minValue = double.Parse(minValueStr);
                                CCLItem.maxValue = double.Parse(maxValueStr);
                                CCLItem.decimalLen = int.Parse(decimalLenStr);
                            }
                        }
                    }
                    else if (item.LocalName == "BAL")
                    {
                        XmlNodeList BALList = item.ChildNodes;
                        foreach (XmlNode bal in BALList)
                        {
                            if (bal.LocalName == "item")
                            {
                                string cellIndexStr = bal.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = bal.SelectSingleNode("min").InnerText;
                                string maxValueStr = bal.SelectSingleNode("max").InnerText;
                                string decimalLenStr = bal.SelectSingleNode("decimal").InnerText;
                                TestItem balItem = new TestItem();
                                balItem.cellIndex = int.Parse(cellIndexStr);
                                balItem.minValue = double.Parse(minValueStr);
                                balItem.maxValue = double.Parse(maxValueStr);
                                balItem.decimalLen = int.Parse(decimalLenStr);
                                BALItemList.Add(balItem);
                            }
                        }
                    }
                    else if (item.LocalName == "NSC")
                    {
                        XmlNodeList NSCList = item.ChildNodes;
                        foreach (XmlNode nsc in NSCList)
                        {
                            if (nsc.LocalName == "item")
                            {
                                string cellIndexStr = nsc.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = nsc.SelectSingleNode("min").InnerText;
                                string maxValueStr = nsc.SelectSingleNode("max").InnerText;
                                string decimalLenStr = nsc.SelectSingleNode("decimal").InnerText;
                                TestItem nscItem = new TestItem();
                                nscItem.cellIndex = int.Parse(cellIndexStr);
                                nscItem.minValue = double.Parse(minValueStr);
                                nscItem.maxValue = double.Parse(maxValueStr);
                                nscItem.decimalLen = int.Parse(decimalLenStr);
                                NSCItemList.Add(nscItem);
                            }
                        }
                    }
                    else if (item.LocalName == "SSC")
                    {
                        XmlNodeList SSCList = item.ChildNodes;
                        foreach (XmlNode ssc in SSCList)
                        {
                            if (ssc.LocalName == "item")
                            {
                                string cellIndexStr = ssc.SelectSingleNode("cellIndex").InnerText;
                                string minValueStr = ssc.SelectSingleNode("min").InnerText;
                                string maxValueStr = ssc.SelectSingleNode("max").InnerText;
                                string decimalLenStr = ssc.SelectSingleNode("decimal").InnerText;
                                TestItem sscItem = new TestItem();
                                sscItem.cellIndex = int.Parse(cellIndexStr);
                                sscItem.minValue = double.Parse(minValueStr);
                                sscItem.maxValue = double.Parse(maxValueStr);
                                sscItem.decimalLen = int.Parse(decimalLenStr);
                                SSCItemList.Add(sscItem);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void tbSn_KeyUp(object sender, KeyEventArgs e)
        {
            if (tbSn.Text.Trim().Length == 18)
            {
                btnSaveData_Click(null, null);
            }
        }


        private void btnBuildBatch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(tbFilePath.Text.Trim()))
                {
                    MessageBox.Show("报告保存路径不能为空！请先生成测试报告！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (string.IsNullOrEmpty(tbSnBase.Text.Trim()))
                {
                    MessageBox.Show("PCB条码格式不能为空！请先输入条码格式！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                int minVal, maxVal;
                if (!Int32.TryParse(tbMinSerialNum.Text.Trim(), out minVal))
                {
                    MessageBox.Show("输入的流水号最小值格式不正确！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (!Int32.TryParse(tbMaxSerialNum.Text.Trim(), out maxVal))
                {
                    MessageBox.Show("输入的流水号最大值格式不正确！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (maxVal < minVal)
                {
                    MessageBox.Show("输入的流水号最小值大于流水号最大值！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                ////先关闭主线程打开的Excel
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                    GC.Collect();
                }
                wb = null;excel = null;
                SnList.Clear();
                btnBuildBatch.IsEnabled = false;
                btnSaveData.IsEnabled = false;
                labMsg.Content = "批量数据生成中，请等待......";
                pbBuilding.Value = 0;
                double val = maxVal - minVal + 1;
                double step = 100.00 / val;
                int len = tbMinSerialNum.Text.Trim().Length;
                //Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);
                gridBuilding.Visibility = Visibility.Visible;
                string SnBase = tbSnBase.Text.Trim();
                string filePath = tbFilePath.Text.Trim();

                //新创建线程，处理数据的生成
                Task.Factory.StartNew(new System.Action(() =>
                {
                    Thread.Sleep(1000);
                    excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.Visible = false;
                    excel.UserControl = true;
                    object missing = System.Reflection.Missing.Value;
                    wb = excel.Application.Workbooks.Open(filePath, missing, missing, missing, missing, missing, missing, missing, missing,
                                  missing, missing, missing, missing, missing, missing);
                    Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);//取得第一个工作簿

                    int rowsNum = ws.UsedRange.Cells.Rows.Count;
                    index = 8;
                    SnList.Clear();
                    for (int i = index; i <= rowsNum; i++)
                    {
                        string sn = ((Range)ws.Cells[i, 1]).Text;
                        if (!string.IsNullOrEmpty(sn))
                        {
                            SnList.Add(sn);
                        }
                        index++;
                    }

                    while (minVal <= maxVal)
                    {
                        string minValStr = minVal.ToString().PadLeft(len, '0');
                        string Sn = SnBase + minValStr;
                        if (SnList.Contains(Sn))
                        {
                            Dispatcher.Invoke(new System.Action(() =>
                            {
                                MessageBox.Show(string.Format("条码 {0} 已存在！", tbSn.Text.Trim()), "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                            }));
                            continue;
                        }
                        WriteTestData(ws, Sn);
                        excel.DisplayAlerts = false;
                        wb.Save();
                        Dispatcher.Invoke(new System.Action(() => { pbBuilding.Value += step; }));

                        index++;
                        SnList.Add(Sn);
                        minVal++;
                    }

                    Dispatcher.Invoke(new System.Action(() =>
                    {
                        //MessageBox.Show("批量数据生成成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        labMsg.Content = "批量数据生成成功！";
                        Thread.Sleep(2000);
                        gridBuilding.Visibility = Visibility.Collapsed;
                        btnBuildBatch.IsEnabled = true;
                        btnSaveData.IsEnabled = true;
                        tbMaxSerialNum.Text = string.Empty;
                        tbMinSerialNum.Text = string.Empty;
                        tbSnBase.Text = string.Empty;

                    }));
                }));
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                gridBuilding.Visibility = Visibility.Collapsed;
                btnBuildBatch.IsEnabled = true;
                btnSaveData.IsEnabled = true;
                Process[] procs = Process.GetProcessesByName("excel");
                foreach (Process pro in procs)
                {
                    pro.Kill();//没有更好的方法,只有杀掉进程
                    GC.Collect();
                }
            }
        }
    }
}
