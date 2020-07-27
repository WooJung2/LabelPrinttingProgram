using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Common
{
    public class CommonConnect
    {
        #region 변수 및 프로퍼티

        // Structure and API declarions:
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public class DOCINFOA
        {
            [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
            [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
            [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
        }

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

        private static string _PrintType;
        private static string _PrintIP;
        private static string _SerialName;
        private static SerialPort _SerialPort;

        public static string PrintType
        {
            get
            {
                return _PrintType;
            }
            set
            {
                _PrintType = value;
            }
        }

        public static string PrintIP
        {
            get
            {
                return _PrintIP;
            }
            set
            {
                _PrintIP = value;
            }
        }

        public static string SerialName
        {
            get
            {
                return _SerialName;
            }
            set
            {
                _SerialName = value;
            }
        }

        public static SerialPort SerialPort
        {
            get
            {
                return _SerialPort;
            }
            set
            {
                _SerialPort = value;
            }
        }

        #endregion 변수 및 프로퍼티

        // SendBytesToPrinter()
        // When the function is given a printer name and an unmanaged array
        // of bytes, the function sends those bytes to the print queue.
        // Returns true on success, false on failure.
        public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            Int32 dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            DOCINFOA di = new DOCINFOA();
            bool bSuccess = false; // Assume failure unless you specifically succeed.

            di.pDocName = "My C#.NET RAW Document";
            di.pDataType = "RAW";

            // Open the printer.
            if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (StartPagePrinter(hPrinter))
                    {
                        // Write your bytes.
                        bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        EndPagePrinter(hPrinter);
                    }
                    EndDocPrinter(hPrinter);
                }
                ClosePrinter(hPrinter);
            }
            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        public static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            // Create a BinaryReader on the file.
            BinaryReader br = new BinaryReader(fs);
            // Dim an array of bytes big enough to hold the file's contents.
            Byte[] bytes = new Byte[fs.Length];
            bool bSuccess = false;
            // Your unmanaged pointer.
            IntPtr pUnmanagedBytes = new IntPtr(0);
            int nLength;

            nLength = Convert.ToInt32(fs.Length);
            // Read the contents of the file into the array.
            bytes = br.ReadBytes(nLength);
            // Allocate some unmanaged memory for those bytes.
            pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);
            // Copy the managed byte array into the unmanaged array.
            Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);
            // Send the unmanaged bytes to the printer.
            bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);
            // Free the unmanaged memory that you allocated earlier.
            Marshal.FreeCoTaskMem(pUnmanagedBytes);
            return bSuccess;
        }

        public static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            Int32 dwCount;
            // How many characters are in the string?
            dwCount = szString.Length;
            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }

        /// <summary>
        /// ZEBRA 프린트기를 통한 인쇄
        /// </summary>
        /// <param name="parameter"></param>
        public static string zebraPrintting(Dictionary<int, string> parameter)
        {
            string strLabel = string.Empty;

            if (parameter == null || parameter.Count == 0) return "210";

            if (parameter[6].Length > 6)
            {
                strLabel = "CT~~CD,~CC^~CT~" + "\t" +
                              "^XA ^PQ1" + "\t" +
                              "^BY2,2.0^FS" + "\t" +
                              "^SEE:UHANGUL.DAT^FS" + "\t" +
                              "^CWQ,E:KFONT3.TTF^CI26^FS" + "\t" +
                              "^LH40,15 ^BQN,3,7 ^FD!!{0}^FS          ^CF0,40" + "\t" +
                              "^LH245,15 ^AQN,30,30^FD {1}^FS             ^CF0,28" + "\t" +
                              "^LH245,60 ^FD Part: ^FS ^LH330,60 ^FD {2}^FS" + "\t" +
                              "^LH245,95 ^FD Date: ^FS ^LH330,95 ^FD {3}^FS" + "\t" +
                              "^LH245,130 ^FD S / N: ^FS ^LH330,130 ^FD {4}^FS" + "\t" +
                              "^LH245,165 ^FD Vendor: ^FS ^LH330,165 ^AQN,30,30 ^FD {5}^FS" + "\t" +
                              "^LH50,240 ^FD {0}^FS" + "\t" +
                              "^XZ";
            }
            else
            {
                strLabel = "CT~~CD,~CC^~CT~" + "\t" +
                              "^XA ^PQ1" + "\t" +
                              "^BY2,2.0^FS" + "\t" +
                              "^SEE:UHANGUL.DAT^FS" + "\t" +
                              "^CWQ,E:KFONT3.TTF^CI26^FS" + "\t" +
                              "^LH40,15 ^BQN,3,7 ^FD!!{0}^FS          ^CF0,40" + "\t" +
                              "^LH240,15 ^AQN,30,30^FD {1}^FS             ^CF0,28" + "\t" +
                              "^LH220,60 ^FD Part: ^FS ^LH330,60 ^FD {2}^FS" + "\t" +
                              "^LH220,95 ^FD Date: ^FS ^LH330,95 ^FD {3}^FS" + "\t" +
                              "^LH220,130 ^FD S / N: ^FS ^LH330,130 ^FD {4}^FS" + "\t" +
                              "^LH220,165 ^FD Vendor: ^FS ^LH330,165 ^AQN,30,30 ^FD {5}^FS" + "\t" +
                              "^LH50,200 ^FD {0}^FS" + "\t" +
                              "^XZ";
            }

            if (_PrintType.Equals("TCP/IP"))
            {
                if (string.IsNullOrEmpty(PrintIP))
                    return "200";   //IP를 찾을 수 없습니다.

                if (parameter.Count != 7)
                    return "210";   //인쇄 항목이 맞지 않습니다.

                strLabel = string.Format(strLabel, parameter[0], parameter[1], parameter[2], parameter[3], parameter[4], parameter[5]);

                System.Net.Sockets.TcpClient client = new System.Net.Sockets.TcpClient();
                client.Connect(PrintIP, 9100);
                System.IO.StreamWriter writer = new System.IO.StreamWriter(client.GetStream(), Encoding.Default);
                //writer.Write(strLabel);
                writer.Flush();
                writer.Close();
                client.Close();

                return "100";   //인쇄를 완료하였습니다.
            }
            else
            {
                if (parameter.Count != 6)
                    return "210";   //인쇄 항목이 맞지 않습니다.

                strLabel = string.Format(strLabel, parameter[0], parameter[1], parameter[2], parameter[3], parameter[4], parameter[5]);
                //SerialPort = new SerialPort(_SerialName);
                //SerialPort.PortName = _SerialName;
                //SerialPort.BaudRate = 9600;
                //SerialPort.DataBits = 8;
                //SerialPort.Parity = Parity.None;
                //SerialPort.StopBits = StopBits.One;
                //SerialPort.Encoding = Encoding.Default;

                //SerialPort.Open();
                //bool Chk = OpenSerialPort();
                //if (Chk)
                //{
                //    SerialPort.WriteLine(strLabel);
                //    // port.Write("^XA^FO50,50^A0,40,40^FDSolbar Tech^FS^XZ");
                //    SerialPort.Close();
                //    return "100";   //인쇄를 완료하였습니다.
                //}
                //else
                //{
                //    return "220";
                //}

                //SerialPort.Open();

                
                if (!SerialPort.IsOpen)
                {
                    SerialPort.Open();
                }
                SerialPort.Write(strLabel);
                SerialPort.Close();
                return "100";   //인쇄를 완료하였습니다.
            }
        }

        //private static bool OpenSerialPort()
        //{
        //    try
        //    {
        //        SerialPort = new SerialPort(_SerialName);

        //        if (SerialPort != null && SerialPort.IsOpen)
        //        {
        //            SerialPort.Close();
        //        }

        //        if (!SerialPort.IsOpen)
        //        {
        //            SerialPort.PortName = _SerialName;
        //            SerialPort.BaudRate = 9600;
        //            SerialPort.DataBits = 8;
        //            SerialPort.Parity = Parity.None;
        //            SerialPort.StopBits = StopBits.One;
        //            SerialPort.Encoding = Encoding.Default;

        //            SerialPort.Open();
        //            return true;
        //        }
        //        else
        //        {
        //            CommonHelper.OpenMessageBox("포트가 이미 열려있습니다.", MessageButtonType.OK, MessageIconType.Error);
        //            return false;
        //        }
        //    }
        //    catch(IOException ex)
        //    {
        //        CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
        //        return false;
        //    }
        //    catch (ArgumentException ex)
        //    {
        //        CommonHelper.OpenMessageBox(ex.Message, MessageButtonType.OK, MessageIconType.Error);
        //        return false;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        /// <summary>
        /// 복합기, 프린트기를 통한 인쇄
        /// </summary>
        /// <param name="parameter"></param>
        public static string laserPrintting(Dictionary<string, string> parameter)
        {
            string strLabel = "" +
                              "";

            PrintDialog pd = new PrintDialog();
            bool? resultbool = pd.ShowDialog();

            if (resultbool == true)
            {
                CommonConnect.SendStringToPrinter(pd.PrintQueue.Name, strLabel);
            }

            return "100";   //인쇄를 완료하였습니다.
        }
    }
}
