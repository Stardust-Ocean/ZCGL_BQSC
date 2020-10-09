using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.IO;

namespace 打印文件
{
    class ZebraGesigner
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)]
            public string pDataType;
        }

        public class PrintCode
        {
            [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);


            [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool ClosePrinter(IntPtr hPrinter);


            [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);


            [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndDocPrinter(IntPtr hPrinter);


            [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool StartPagePrinter(IntPtr hPrinter);


            [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool EndPagePrinter(IntPtr hPrinter);


            [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
            public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);


            /// <summary>
            /// 该方法把非托管内存中的字节数组发送到打印机的打印队列
            /// </summary>
            /// <param name="szPrinterName">打印机名称</param>
            /// <param name="pBytes">非托管内存指针</param>
            /// <param name="dwCount">字节数</param>
            /// <returns>成功返回true，失败时为false</returns>
            public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
            {
                Int32 dwError = 0, dwWritten = 0;
                IntPtr hPrinter = new IntPtr(0);
                DOCINFOA di = new DOCINFOA();
                bool bSuccess = false;

                di.pDocName = "My C#.NET RAW Document";
                di.pDataType = "RAW";

                try
                {
                    // 打开打印机
                    if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
                    {
                        // 启动文档打印
                        if (StartDocPrinter(hPrinter, 1, di))
                        {
                            // 开始打印
                            if (StartPagePrinter(hPrinter))
                            {
                                // 向打印机输出字节  
                                bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                                EndPagePrinter(hPrinter);
                            }
                            EndDocPrinter(hPrinter);
                        }
                        ClosePrinter(hPrinter);
                    }
                    if (bSuccess == false)
                    {
                        dwError = Marshal.GetLastWin32Error();
                    }
                }
                catch (Win32Exception ex)
                {
                    WriteLog(ex.Message);
                    bSuccess = false;
                }
                return bSuccess;
            }


            /// <summary>
            /// 发送文件到打印机方法
            /// </summary>
            /// <param name="szPrinterName">打印机名称</param>
            /// <param name="szFileName">打印文件的路径</param>
            /// <returns></returns>
            public static bool SendFileToPrinter(string szPrinterName, string szFileName)
            {
                bool bSuccess = false;
                try
                {
                    // 打开文件  
                    FileStream fs = new FileStream(szFileName, FileMode.Open);

                    // 将文件内容读作二进制
                    BinaryReader br = new BinaryReader(fs);

                    // 定义字节数组
                    Byte[] bytes = new Byte[fs.Length];

                    // 非托管指针  
                    IntPtr pUnmanagedBytes = new IntPtr(0);

                    int nLength;

                    nLength = Convert.ToInt32(fs.Length);

                    // 读取文件内容到字节数组
                    bytes = br.ReadBytes(nLength);

                    // 为这些字节分配一些非托管内存
                    pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);

                    // 将托管字节数组复制到非托管内存指针
                    Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);

                    // 将非托管字节发送到打印机
                    bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);

                    // 释放先前分配的非托管内存
                    Marshal.FreeCoTaskMem(pUnmanagedBytes);

                    fs.Close();
                    fs.Dispose();
                }
                catch (Win32Exception ex)
                {
                    WriteLog(ex.Message);
                    bSuccess = false;
                }
                return bSuccess;
            }

            /// <summary>
            /// 将字符串发送到打印机方法
            /// </summary>
            /// <param name="szPrinterName">打印机名称</param>
            /// <param name="szString">打印的字符串</param>
            /// <returns></returns>
            public static bool SendStringToPrinter(string szPrinterName, string szString)
            {
                bool flag = false;
                try
                {
                    IntPtr pBytes;
                    Int32 dwCount;

                    // 获取字符串长度  
                    dwCount = szString.Length;

                    // 将字符串复制到非托管 COM 任务分配的内存非托管内存块，并转换为 ANSI 文本
                    pBytes = Marshal.StringToCoTaskMemAnsi(szString);

                    // 将已转换的 ANSI 字符串发送到打印机
                    flag = SendBytesToPrinter(szPrinterName, pBytes, dwCount);

                    // 释放先前分配的非托管内存
                    Marshal.FreeCoTaskMem(pBytes);
                }
                catch (Win32Exception ex)
                {
                    WriteLog(ex.Message);
                    flag = false;
                }
                return flag;
            }

            /// <summary>
            /// 写入日志方法
            /// </summary>
            /// <param name="msg">记录信息</param>
            public static void WriteLog(string msg)
            {
                string str = string.Empty;
                string path = AppDomain.CurrentDomain.BaseDirectory + "log\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt";

                FileStream filestream = new FileStream(path, FileMode.OpenOrCreate);

                str += "************" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "************\r\n";
                str += msg;
                str += "************" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "************\r\n";

                FileStream fs = new FileStream(path, FileMode.Append);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
                sw.WriteLine(str);

                sw.Flush();

                sw.Close();
                sw.Dispose();

                fs.Close();
                fs.Dispose();
            }
        }
    }
}


