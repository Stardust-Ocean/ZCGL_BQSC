using DSkin.Controls;
using DSkin.Forms;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;
using static 打印文件.ZebraGesigner;

namespace 打印文件
{
    public partial class Form1 : DSkinForm
    {

        #region
        public static List<string> ProgramParametersNum = new List<string>();
        public static string[] PrintMessageArray = new string[30];


        public static string my1HKMC { get; private set; }
        public static string my296390GI100 { get; private set; }
        public static string my3PartNumber { get; private set; }
        public static string my4ECUManufacturingDate { get; private set; }
        public static string my5SoftwareID { get; private set; }
        public static string my6HardwareID { get; private set; }
        public static string my7CANDBID { get; private set; }
        public static string my8Supplierinformation { get; private set; }
        public static string my9SYSTEMManufacturingDate { get; private set; }
        public static string my10VehicleName { get; private set; }
        public static string my11Manufacturinglocation { get; private set; }
        public static string my12ManufacturinglineNo { get; private set; }
        public static string my13SerialNumber { get; private set; }
        public static string my14PrintName { get; private set; }
        public static string my15Printpath { get; private set; }
        string Pdfpath = "";
        PrintDocument pd = new PrintDocument();

        #endregion
        public Form1()
        {
            InitializeComponent();
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            PrintDocument print = new PrintDocument();
            string sDefault = print.PrinterSettings.PrinterName;
            //默认打印机名
            // DSkinMessageBox.Show(sDefault);
            //foreach (string sPrint in PrinterSettings.InstalledPrinters)//获取所有打印机名称
            //{
            //    comboBox1.Items.Add(sPrint);
            //    if (sPrint == sDefault)
            //        comboBox1.SelectedIndex = comboBox1.Items.IndexOf(sPrint);
            //}
            //#region
        }


        public static void PrintMessageArrayRead()
        {

            string[] temp = { };

            string FilePathBaseMessage = System.Windows.Forms.Application.StartupPath + @"打印参数.txt";
            if (File.Exists(FilePathBaseMessage) == true)
            {
                PrintMessageArray = File.ReadAllLines(FilePathBaseMessage, Encoding.UTF8);

                my1HKMC = PrintMessageArray[0];// msg3
                my296390GI100 = PrintMessageArray[1];//msg4
                my3PartNumber = PrintMessageArray[2];// msg3
                my4ECUManufacturingDate = PrintMessageArray[3];//msg4
                my5SoftwareID = PrintMessageArray[4];//msg5
                my6HardwareID = PrintMessageArray[5];//msg6
                my7CANDBID = PrintMessageArray[6];//msg7
                my8Supplierinformation = PrintMessageArray[7];//msg8
                my9SYSTEMManufacturingDate = PrintMessageArray[8];//msg9
                my10VehicleName = PrintMessageArray[9];//msg10
                my11Manufacturinglocation = PrintMessageArray[10];//msg11
                my12ManufacturinglineNo = PrintMessageArray[11];//msg12
                my13SerialNumber = PrintMessageArray[12];// msg13
                my14PrintName = PrintMessageArray[13];// 打印机名称
                my15Printpath = PrintMessageArray[14];// 木板路径
            }
            else
            {

                my1HKMC = "HKMC";// msg1
                my296390GI100 = "96390-GI100";//msg2
                my3PartNumber = "96390-GI100";// msg3
                my4ECUManufacturingDate = "20191116";//msg4
                my5SoftwareID = "0.01";//msg5
                my6HardwareID = "0.01";//msg6
                my7CANDBID = "0.02";//msg7
                my8Supplierinformation = "SZD3";//msg8
                my9SYSTEMManufacturingDate = "20191116";//msg9
                my10VehicleName = "NE";//msg10
                my11Manufacturinglocation = "KOC";//msg11
                my12ManufacturinglineNo = "1 LINE";//msg12
                my13SerialNumber = "0001";// msg13
                my14PrintName = "ZDesigner ZT210-200dpi ZPL";
                my15Printpath = System.Windows.Forms.Application.StartupPath + @"mytest3.prn";
                PrintMessageArraySave();
            }

        }
        public static void PrintMessageArraySave()
        {
            string[] mStrs1 = ProgramParametersNum.ToArray();//strArray=[str0,str1,str2]

            string FilePathBaseMessage = System.Windows.Forms.Application.StartupPath + @"print.prn";
            StreamWriter mStreamWriterBaseMessage = new StreamWriter(FilePathBaseMessage, false, Encoding.UTF8);
            PrintMessageArray[0] = my1HKMC;
            PrintMessageArray[1] = my296390GI100;
            PrintMessageArray[2] = my3PartNumber;
            PrintMessageArray[3] = my4ECUManufacturingDate;
            PrintMessageArray[4] = my5SoftwareID;
            PrintMessageArray[5] = my6HardwareID;
            PrintMessageArray[6] = my7CANDBID;
            PrintMessageArray[7] = my8Supplierinformation;
            PrintMessageArray[8] = my9SYSTEMManufacturingDate;
            PrintMessageArray[9] = my10VehicleName;
            PrintMessageArray[11] = my12ManufacturinglineNo;
            PrintMessageArray[12] = my13SerialNumber;
            PrintMessageArray[13] = my14PrintName;
            PrintMessageArray[14] = my15Printpath;

            for (int i = 0; i < PrintMessageArray.Length; i++)
            {
                mStreamWriterBaseMessage.WriteLine(PrintMessageArray[i]);
            }
            mStreamWriterBaseMessage.Close();
            mStreamWriterBaseMessage = null;
        }

        /// <summary>
        /// 解析PRN
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dSkinButton3_Click(object sender, EventArgs e)
        {
            string str = "";
            string FilePathBaseMessage = @Pdfpath;
            if (File.Exists(FilePathBaseMessage) == true)
            {
                PrintMessageArray = File.ReadAllLines(FilePathBaseMessage, Encoding.ASCII);

                for (int i = 0; i < PrintMessageArray.Length; i++)
                {
                    str += PrintMessageArray[i] + "\r\n";
                }


                MessageBox.Show(str);

                for (int i = 0; i < PrintMessageArray.Length; i++)
                {
                    if (!(PrintMessageArray[i].IndexOf("msg10") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg10", "NE");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg11") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg11", "KOC");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg12") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg12", "1 LINE");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg13") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg13", "0001");
                    }



                    if (!(PrintMessageArray[i].IndexOf("msg1") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg1", "HKMC");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg2") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg2", "96390-GI100");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg3") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg3", "96390-GI100");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg4") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg4", "20191116");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg5") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg5", "0.01");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg6") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg6", "0.01");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg7") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg7", "0.02");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg8") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg8", "SZD3");
                    }
                    if (!(PrintMessageArray[i].IndexOf("msg9") < 0))
                    {
                        PrintMessageArray[i] = PrintMessageArray[i].Replace("msg9", "20191116");
                    }


                }
                for (int i = 0; i < PrintMessageArray.Length; i++)
                {
                    str += PrintMessageArray[i] + "\r\n";
                }


                MessageBox.Show(str);

            }
        }
        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CommAddBtn_Click(object sender, EventArgs e)
        {

            //导入excel_Click();
           // getColumnDB(@"D:\打印\面包装车分货.xls");
            ImportExcel(@"D:\打印\面包装车分货.xls");
           
        }
        public string DirName = ConfigurationManager.AppSettings["DirName"];
        /// <summary>
        /// 打印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dSkinButton2_Click(object sender, EventArgs e)
        {
            try
            {

                //文件中包含名
                string FileName = dSkinGridList1.SelectedItem.Cells[1].Value.ToString();
                string f1name = Regex.Replace(FileName.Replace("（", "(").Replace("）", ")"), @"\([^\(]*\)", "") + ".prn";
                string smallDir = "";
                GetFile(DirName, f1name, ref smallDir);
                //  DSkinMessageBox.Show("文件在" + smallDir);
                var FileSum = dSkinGridList1.SelectedItem.Cells[3].Value.ToString();
                int asum = int.Parse(FileSum) / 2;
                for (int i = 0; i < asum; i++)
                {
                    @Pdfpath = smallDir;
                    string ZeberNmae = ConfigurationManager.AppSettings["ZeberNmae"];
                    if (PrintCode.SendFileToPrinter(ZeberNmae, smallDir))
                    {
                        Console.WriteLine("文件已成功发送至打印队列！", "提示信息");
                    }
                }
                // DSkinMessageBox.Show(FileSum);


            }
            catch (Exception ex)
            {
                DSkinMessageBox.Show("请选择要打印的数字");
            }
        }

        public List<string> ColumnDB = new List<string>();
        public void getColumnDB(string ExcelName)
        {
            //创建 Excel对象
            Microsoft.Office.Interop.Excel.Application App = new Microsoft.Office.Interop.Excel.Application();
            //获取缺少的object类型值
            object missing = Missing.Value;
            //打开指定的Excel文件
            Workbook openwb = App.Workbooks.Open(ExcelName, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //获取选选择的工作表
            Worksheet ws = ((Worksheet)openwb.Worksheets["Sheet1"]);//方法一：指定工作表名称读取
            //Worksheet ws = (Worksheet)openwb.Worksheets.get_Item(1);//方法二：通过工作表下标读取
            //获取工作表中的行数
            int rows = ws.UsedRange.Rows.Count;
            //获取工作表中的列数
            int columns = ws.UsedRange.Columns.Count;
            int column = Convert.ToInt16(2);
            //提取对应行列的数据并将其存入数组中
            for (int i = 2; i < rows; i++)
            {
                string a = ((Range)ws.Cells[i, column]).Text.ToString();
                // DSkinMessageBox.Show("读取的数据:" + a);//测试是否获得数据
                string f1name = Regex.Replace(a.Replace("（", "(").Replace("）", ")"), @"\([^\(]*\)", "") + ".prn";

                string smallDir = "";
                GetFile(DirName, f1name, ref smallDir);
                if (string.IsNullOrEmpty(smallDir))
                {
                    //DSkinMessageBox.Show(f1name+"文件不存在:" + smallDir);
                }
                //else
                //{
                //    DSkinMessageBox.Show("读取的数据:" + smallDir);
                //}


                // dSkinGridList1.Rows.AddRow(a);
            }
        }


        /// <summary>
        /// 读取excel
        /// </summary>
        private void 导入excel_Click()
        {
            string path = "";
            System.Windows.Forms.OpenFileDialog fd = new System.Windows.Forms.OpenFileDialog();
            fd.Title = "选择文件";//选择框名称
            fd.Filter = "xls files (*.xls)|*.xls";//选择文件的类型为Xls表格
            if (fd.ShowDialog() == System.Windows.Forms.DialogResult.OK)//当点击确定
            {
                path = fd.FileName.Trim();//文件路径
                path = path.Replace("\\", "/");  
            }

    
                dSkinGridList1.Rows.Clear();
           
                string mystring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='";
                mystring += path.ToString();
                mystring += "';User ID=admin;Password=;Extended properties='Excel 8.0;IMEX=1;HDR=NO;'";
                OleDbConnection cnnxls = new OleDbConnection(mystring);
                OleDbDataAdapter myDa = new OleDbDataAdapter("Select * from [Sheet1$]", cnnxls);
                DataSet myDs = new DataSet();
                myDa.Fill(myDs);//数据存放在myDs中了
                System.Data.DataTable dataTable = myDs.Tables[0];
                for (int i = 10; i < dataTable.Columns.Count; i++)
                {
                    DSkinGridListColumn dSkinGridListColumn = new DSkinGridListColumn();
                    dSkinGridListColumn.Name = dataTable.Columns[i].ColumnName;
                    dSkinGridListColumn.Width = 100;
                    dSkinGridListColumn.DataPropertyName = dataTable.Columns[i].ColumnName;
                    dSkinGridList1.Columns.Add(dSkinGridListColumn);
                }
                for (int j = 0; j < dataTable.Rows.Count; j++)
                {
                    if (dataTable.Rows[j].ItemArray[3].ToString() != "")
                    {
                        dSkinGridList1.Rows.AddRow(dataTable.Rows[j].ItemArray);
                    }
                }
        }
        static void Director(string dir)
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    Director(fsinfo.FullName);//递归调用
                }
                else
                {
                    //DSkinMessageBox.Show(fsinfo.FullName);//输出文件的全部路径
                }
            }
        }


        /// <summary>
        /// 获取路径下所有文件以及子文件夹中文件
        /// </summary>
        /// <param name="path">全路径根目录</param>
        /// <param name="FileList">存放所有文件的全路径</param>
        /// <param name="RelativePath"></param>
        /// <returns></returns>
        public static void GetFile(string path, string fileName, ref string smallDir)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] fil = dir.GetFiles();
            DirectoryInfo[] dii = dir.GetDirectories();
            foreach (FileInfo f in fil)
            {
                string name = Path.GetFileName(f.FullName.ToString());
                //Console.WriteLine(name);
                if (name.Contains(fileName))
                {
                    Console.WriteLine(Path.GetDirectoryName(f.FullName));
                    string temp = Path.GetFullPath(f.FullName);
                    smallDir = temp;
                    return;
                }
            }
            //获取子文件夹内的文件列表，递归遍历
            foreach (DirectoryInfo d in dii)
            {
                if (smallDir == "")
                    GetFile(d.FullName, fileName, ref smallDir);
            }
        }


        static void GetFileName(string DirName, string FileName)
        {
            //文件夹信息
            DirectoryInfo dir = new DirectoryInfo(DirName);
            //如果非根路径且是系统文件夹则跳过
            if (null != dir.Parent && dir.Attributes.ToString().IndexOf("System") > -1)
            {
                return;
            }
            //取得所有文件
            FileInfo[] finfo = dir.GetFiles();
            string fname = string.Empty;
            for (int i = 0; i < finfo.Length; i++)
            {
                fname = finfo[i].Name;
                //判断文件是否包含查询名
                if (fname.IndexOf(FileName) > -1)
                {
                    Console.WriteLine(finfo[i].FullName);
                }
            }
            //取得所有子文件夹
            DirectoryInfo[] dinfo = dir.GetDirectories();
            for (int i = 0; i < dinfo.Length; i++)
            {
                //查找子文件夹中是否有符合要求的文件
                GetFileName(dinfo[i].FullName, FileName);
            }
        }





        public DataSet ImportExcel(string filePath)
        {
            DataSet ds = null;
            try
            {
                FileStream fileStream = new FileStream(filePath, FileMode.Open);
                XSSFWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = null;
                IRow row = null;
                ds = new DataSet();
                System.Data.DataTable dt = null;
                for (int i = 0; i < workbook.Count; i++)
                {
                    dt = new System.Data.DataTable();
                    dt.TableName = "table" + i.ToString();
                    //获取sheet表
                    sheet = workbook.GetSheetAt(i);
                    //起始行索引
                    int rowIndex = sheet.FirstRowNum;
                    //获取行数
                    int rowCount = sheet.LastRowNum;
                    //获取第一行
                    IRow firstRow = sheet.GetRow(rowIndex);
                    //起始列索引
                    int colIndex = firstRow.FirstCellNum;
                    //获取列数
                    int colCount = firstRow.LastCellNum;
                    DataColumn dc = null;
                    //获取列数
                    for (int j = colIndex; j < colCount; j++ )
                    {
                        dc = new DataColumn(firstRow.GetCell(j).StringCellValue);
                        dt.Columns.Add(dc);
                    }
                    //跳过第一行列名
                    rowIndex++;
                    for (int k = rowIndex; k <= rowCount; k++)
                    {
                        DataRow dr = dt.NewRow();
                        row = sheet.GetRow(k);
                        for (int l = colIndex; l < colCount; l++)
                        {
                            if (row.GetCell(l) == null)
                            {
                                continue;
                            }
                            dr[l] = row.GetCell(l).StringCellValue;

                        }
                        dt.Rows.Add(dr);
                    }
                    ds.Tables.Add(dt);
                }
                sheet = null;
                workbook = null;
                fileStream.Close();
                fileStream.Dispose();
            }
            catch (Exception ex)
            {
                throw;
            }
            return ds;
        }

    }
}
