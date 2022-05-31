namespace PunchIn_System.Activity_Manage_System
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;

    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using LibUsbDotNet;
    using LibUsbDotNet.Main;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.Drawing;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System.Runtime.InteropServices;

    using System.Diagnostics;
    using System.IO;
    using System.Data.OleDb;
    using Microsoft.Office.Interop;
    using Microsoft.VisualBasic;

    public partial class ActivityPageSignIn : Form
    {
        public ActivityPageSignIn()
        {
            InitializeComponent();
            x = this.Width;
            y = this.Height;
            setTag(this);
        }

        //顯示現在時間
        private void timer1_Tick(object sender, EventArgs e)
        {
            label4.Text = DateTime.Now.ToString();
        }

        //紀錄每個Controls元件的Size
        private float x;
        private float y;
        private void setTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ";" + con.Height + ";" + con.Left + ";" + con.Top + ";" + con.Font.Size;
                if (con.Controls.Count > 0)
                {
                    setTag(con);
                }
            }
        }
        //視窗最大化時內容可以等比例放大
        private void setControls(float newx, float newy, Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                if (con.Tag != null)
                {
                    string[] mytag = con.Tag.ToString().Split(new char[] { ';' });
                    //根據窗體縮放的比例確定控制元件的值
                    con.Width = Convert.ToInt32(System.Convert.ToSingle(mytag[0]) * newx);//寬度
                    con.Height = Convert.ToInt32(System.Convert.ToSingle(mytag[1]) * newy);//高度
                    con.Left = Convert.ToInt32(System.Convert.ToSingle(mytag[2]) * newx);//左邊距
                    con.Top = Convert.ToInt32(System.Convert.ToSingle(mytag[3]) * newy);//頂邊距
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;//字型大小
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                    if (con.Controls.Count > 0)
                    {
                        setControls(newx, newy, con);
                    }
                }
            }
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            float newx = (this.Width) / x;
            float newy = (this.Height) / y;
            setControls(newx, newy, this);
        }

        private void send_notification_Click(object sender, EventArgs e)
        {
            Notification_Setting notification_Setting = new Notification_Setting();
            notification_Setting.ShowDialog();
        }

        private void back_to_ActivityList(object sender, EventArgs e)
        {
            this.Hide();
            ActivityList activityList = new ActivityList();
            activityList.ShowDialog();
        }

        private void Start_to_SignIn_Click(object sender, EventArgs e)
        {

            try
            {
                LibUSB usb = new LibUSB();
                string id = usb.returnid();
                id = '"' + id + '"';
                MessageBox.Show(id);
                XSSFWorkbook wookbook;
                //string filepath = @"..\..\..\..\..\excel\"+ActivityList.filename;
                var path = Directory.GetCurrentDirectory() + @"\excel\" + ActivityList.filename + ".xlsx";
                FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
                byte[] bytes = new byte[file.Length];
                file.Read(bytes, 0, (int)file.Length);
                MemoryStream ms = new MemoryStream(bytes);
                wookbook = new XSSFWorkbook(ms);
                int row = LibUSB.Getrow(wookbook, "工作表1", id, "簽到時間") + 1;
                int col = LibUSB.Getcol(wookbook, "工作表1", id, "簽到時間") + 1;
                file.Close();
                //MessageBox.Show("here");
                //string pathfile = @"..\..\..\..\..\excel\"+ActivityList.filename;
                Excel.Application excelApp;
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                //Excel.Range wRange;
                // 開啟一個新的應用程式
                excelApp = new Excel.Application();
                // 讓Excel文件可見
                excelApp.Visible = false;
                // 停用警告訊息
                excelApp.DisplayAlerts = false;
                // 加入新的活頁簿
                wBook = excelApp.Workbooks.Open(path);
                // 設定活頁簿焦點
                //wBook.Activate();
                // 引用第一個工作表
                wSheet = wBook.Worksheets["工作表1"];
                // 命名工作表的名稱
                //wSheet.Name = "工作表1";
                // 設定工作表焦點
                //wSheet.Activate();

                if (excelApp.Cells[row, col].Value == null)
                {
                    DateTime myDate = DateTime.Now;
                    string myDateString = myDate.ToString("yyyy-MM-dd HH:mm:ss");
                    excelApp.Cells[row, col] = myDateString;
                    MessageBox.Show(id + "簽到成功");
                    
                }
                else
                {
                    DateTime myDate = DateTime.Now;
                    string myDateString = myDate.ToString("yyyy-MM-dd HH:mm:ss");
                    excelApp.Cells[row, col + 1] = myDateString;
                    MessageBox.Show(id + "簽退成功");
                }
                wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //excelApp.Save(pathFile);
                wBook.Close();
                wBook = null;
                wSheet = null;
                excelApp.Quit();
                excelApp = null;
                label3.Text = "實到人數:" + countfun();
                label3.Refresh();
                //MessageBox.Show("here");
                //killexcel();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error");
            }
        }
        public static void killexcel() {
            string command = "/c taskkill /F /IM excel.exe";
            Process proc = new Process();
            proc.StartInfo.FileName = "CMD.exe";
            proc.StartInfo.Arguments = command;
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo.CreateNoWindow = true;
            proc.Start();
            proc.WaitForExit();
        }
        public void ActivityPageSignIn_Load(object sender, EventArgs e)
        {
            /*
            OleDbConnection myconnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='C:\Users\rex\Downloads\PunchIn_System\excel\活動七.xlsx';Extended Properties='Excel 12.0;HDR=YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("Select * from [工作表1$]", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            myconnection.Close();
            */
            label2.Text = "應到人數:" + counting_participants();
            label3.Text = "實到人數:" + countfun();
            label2.Refresh();
            label3.Refresh();
            killexcel();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + @"\excel\"+ ActivityList.filename+".xlsx";
            MessageBox.Show(path);
            OleDbConnection myconnection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+path+";Extended Properties='Excel 12.0;HDR=YES';");
            //;Extended Properties='Excel 12.0;HDR=YES';"
            OleDbDataAdapter oda = new OleDbDataAdapter("Select * from [工作表1$]", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            myconnection.Close();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //var path = Directory.GetCurrentDirectory() + @"\excel\";
            //string[] lines = System.IO.File.ReadAllLines(path+"活動列表.txt");
            try
            {
                Excel.Application excelApp;
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                //Excel.Range wRange;
                // 開啟一個新的應用程式
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                // 停用警告訊息
                excelApp.DisplayAlerts = false;
                wBook = excelApp.Workbooks.Open(Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx");
                wSheet = wBook.Worksheets["工作表1"];
                string messagestring = "";
                for (int i = 1; ; i++)
                {
                    //MessageBox.Show("do");


                    if (excelApp.Cells[i,1].Value.Equals(ActivityList.filename))
                    {
                        if (excelApp.Cells[i,5].Value.Equals("班級"))
                        {
                            messagestring = "課堂名稱: " + excelApp.Cells[i, 1].Value + "\n活動性質: " + excelApp.Cells[i, 2].Value;
                        }
                        else if (excelApp.Cells[i, 5].Value.Equals("上班"))
                        {
                            messagestring = "公司名稱: " + excelApp.Cells[i, 1].Value + "\n活動性質: " + excelApp.Cells[i, 2].Value;
                        }
                        else
                        {
                            messagestring = "活動名稱: " + excelApp.Cells[i, 1].Value + "\n活動性質: " + excelApp.Cells[i, 2].Value;
                        }
                        
                        if (excelApp.Cells[i, 2].Value.Equals("集會"))
                        {
                            messagestring += "\n活動時間: " + excelApp.Cells[i, 3].Value + "\n活動主辦人: " + excelApp.Cells[i, 4].Value;
                        }
                        if (excelApp.Cells[i, 2].Value.Equals("考試"))
                        {
                            messagestring += "\n活動時間: " + excelApp.Cells[i, 3].Value + "\n考試地點: " + excelApp.Cells[i, 4].Value;
                        }
                        MessageBox.Show(messagestring, "活動資料", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                    

                }

                // save the application  
                wBook.SaveAs(Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                wBook.Close();
                excelApp.Quit();
                
                wBook = null;
                wSheet = null;
                excelApp = null;
                ActivityPageSignIn.killexcel();
                
                
            }
            catch (Exception d)
            {
                d.ToString();
            }
            /*
            foreach (string line in System.IO.File.ReadLines(path + "活動列表.txt"))
            {
                string[] subs = line.Trim().Split(" ");
                //MessageBox.Show(subs[0]+"**");
                // Use a tab to indent each line of the file.
                string messagestring = "";
                if (subs[0].Equals(ActivityList.filename)) {
                    messagestring = "活動名稱: " + subs[0] + "\n活動性質: " + subs[1];
                    if (subs[1].Equals("集會")) {
                        messagestring += "\n活動時間: " + subs[2] + "\n活動主辦人: " + subs[3];
                    }
                    if (subs[1].Equals("考試"))
                    {
                        messagestring += "\n活動時間: " + subs[2] + "\n考試地點: " + subs[3];
                    }
                }
                MessageBox.Show(messagestring,"活動資料", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;
            }
            */
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string spath = "";
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            spath = path.SelectedPath;
            //MessageBox.Show(spath);
            string myValue = Interaction.InputBox("匯出檔案名", "檔案匯出", ActivityList.filename);
            string targetfilepath = spath + @"\" + myValue+".xlsx";
            //MessageBox.Show(targetfilepath);
            string sourcepath = Directory.GetCurrentDirectory() + @"\excel\" + ActivityList.filename+".xlsx";
            if (spath != "" && myValue != "")
            {
                System.IO.File.Copy(sourcepath, targetfilepath, true);
                MessageBox.Show("檔案匯出至 " + targetfilepath);
            }
            else
            {
                MessageBox.Show("檔名或所選資料夾路徑值為空");
            }
            

        }
    }

    class LibUSB
    {
        private UsbDevice usbDevice;
        private UsbEndpointReader epReader;
        private UsbEndpointWriter epWriter;

        public void Open(int vid, int pid)
        {
            UsbDeviceFinder usbFinder = new UsbDeviceFinder(vid, pid);
            usbDevice = UsbDevice.OpenUsbDevice(usbFinder);


            if (usbDevice == null)
            {
                // 多次尝试连接USB设备
                int count = 0;
                while (count < 10 && usbDevice == null)
                {
                    Thread.Sleep(100);
                    usbDevice = UsbDevice.OpenUsbDevice(usbFinder);
                    count++;
                }
            }
            if (usbDevice != null)
            {
                //Console.WriteLine("打開USB裝置成功");
            }
            if (usbDevice == null)
            {
                //Console.WriteLine("打開USB裝置失敗");
            }
            else
            {
                // If this is a "whole" usb device (libusb-win32, linux libusb)
                // it will have an IUsbDevice interface. If not (WinUSB) the 
                // variable will be null indicating this is an interface of a 
                // device.
                IUsbDevice wholeUSBDevice = usbDevice as IUsbDevice;
                if (wholeUSBDevice != null)
                {
                    // This is a "whole" USB device. Before it can be used, 
                    // the desired configuration and interface must be selected.

                    // Select config #1
                    wholeUSBDevice.SetConfiguration(1);

                    // Claim interface #0.
                    wholeUSBDevice.ClaimInterface(0);
                }

                // open read endpoint 1.
                epReader = usbDevice.OpenEndpointReader(ReadEndpointID.Ep01);

                // open write endpoint 1.
                epWriter = usbDevice.OpenEndpointWriter(WriteEndpointID.Ep01);
            }
        }


        public void Close()
        {
            if (IsOpen())
            {
                // If this is a "whole" usb device(libusb-win32, linux libusb-1.0)
                // it exposes an IUsbDevice interface. If not (WinUSB) the 
                // 'wholeUsbDevice' variable will be null indicating this is 
                // an interface of a device; it does not require or support 
                // configuration and interface selection.
                IUsbDevice wholeUsbDevice = usbDevice as IUsbDevice;
                if (!ReferenceEquals(wholeUsbDevice, null))
                {
                    // Release interface #0.
                    wholeUsbDevice.ReleaseInterface(0);
                }

                usbDevice.Close();
                usbDevice = null;
            }

            // Free usb resource
            UsbDevice.Exit();
        }

        public bool IsOpen()
        {
            if (usbDevice == null)
            {
                return false;
            }

            return usbDevice.IsOpen;
        }

        //public List<byte> SendCommand(string cmd)
        public List<byte> SendCommand()
        {
            List<byte> result = new List<byte>();
            /*int byteWrite;
            ErrorCode ec = epWriter.Write(Encoding.Default.GetBytes(cmd), 3000, out byteWrite);
            if (ec != ErrorCode.None)
            {
                Console.WriteLine($"Write Error, {UsbDevice.LastErrorString}");
                return null;
            }*/

            byte[] readBuffer = new byte[1024];
            int byteRead;
            ErrorCode ec = ErrorCode.None;
            while (ec == ErrorCode.None)
            {
                // If the device hasn't sent data in the last 1000 milliseconds,
                // a timeout error (ec = IoTimedOut) will occur. 
                ec = epReader.Read(readBuffer, 3000, out byteRead);
                // 只有当超时的时候才会有byteRead为0，也就是结束
                if (byteRead != 0)
                {
                    byte[] buffer = new byte[byteRead];
                    Array.Copy(readBuffer, buffer, byteRead);
                    result.AddRange(buffer);
                }
                else
                {
                    Console.WriteLine("結束讀取");
                    return result;
                }
            }
            return null;
        }
        public String returnid()
        {
            LibUSB usb = new LibUSB();
            usb.Open(0x16C0, 0x27DB);
            List<byte> result = new List<byte>();
            result = usb.SendCommand();
            //Console.WriteLine("讀入的數字");
            result.ForEach(i => Console.Write("{0} ", i));
            //Console.WriteLine();
            //int[] numarr = new int[12];
            //int j = 0;
            String id = "";
            for (int i = 0; i < result.Count; i++)
            {
                if ((i + 1) % 3 == 0)
                {
                    if (result[i] > 30)
                    {
                        //Console.Write(result[i] % 10);
                        //numarr[j++]=result[i]%10;
                        id += (result[i] % 10).ToString();
                    }
                    else
                    {
                       // Console.Write(result[i]);
                        //numarr[j++] = result[i];
                        id += result[i].ToString();
                    }

                }
            }
            usb.Close();
            return id;
        }
        
        public static string GetToolConfigCell(XSSFWorkbook wookbook, string sheetName, string rowName, string columnName)
        {

            ISheet sheet = wookbook.GetSheet(sheetName);

            if (sheet == null)
            {
                Console.WriteLine("The sheet is null");
                return null;

            }

            int rowIndex = -1;
            int colIndex = -1;

            for (int r = 0; r < sheet.PhysicalNumberOfRows; r++)
            {
                IRow row = sheet.GetRow(r);
                Console.WriteLine(row.Cells[0].CellType == CellType.String);
                if (row == null || row.Cells.Count == 0)
                    continue;

                if (row.Cells[0].CellType == CellType.String && string.Compare(row.Cells[0].StringCellValue, rowName, true) == 0)
                {
                    rowIndex = r;
                    Console.WriteLine(r);
                    break;
                }
            }
            IRow firstRow = sheet.GetRow(0);
            for (int c = 0; c < firstRow.Cells.Count; c++)
            {
                if (firstRow.Cells[c].CellType == CellType.String && string.Compare(firstRow.Cells[c].StringCellValue, columnName, true) == 0)
                {
                    colIndex = c;
                    Console.WriteLine(c);
                    break;
                }
            }

            if (rowIndex != -1 && colIndex != -1)
            {
                string cellData = "";
                ICell cell = sheet.GetRow(rowIndex).Cells[colIndex];

                switch (cell.CellType)
                {
                    case CellType.Blank:
                        cellData = "";
                        break;
                    case CellType.Boolean:
                        cellData = Convert.ToString(cell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        cellData = Convert.ToString(cell.ErrorCellValue);
                        break;
                    case CellType.Formula:
                        cellData = Convert.ToString(cell.StringCellValue);
                        break;
                    case CellType.Numeric:
                        cellData = Convert.ToString(cell.NumericCellValue);
                        Console.WriteLine("GG");
                        break;
                    case CellType.String:
                        cellData = Convert.ToString(cell.StringCellValue);
                        break;
                    case CellType.Unknown:
                        cellData = "";
                        break;
                }


                return cellData;
            }
            else
            {

                return null;
            }
        }
        public static int Getrow(XSSFWorkbook wookbook, string sheetName, string rowName, string columnName)
        {
            ISheet sheet = wookbook.GetSheet(sheetName);

            if (sheet == null)
            {
                Console.WriteLine("The sheet is null");

            }

            int rowIndex = -1;
            int colIndex = -1;

            for (int r = 0; r < sheet.PhysicalNumberOfRows; r++)
            {
                IRow row = sheet.GetRow(r);
                Console.WriteLine(row.Cells[0].CellType == CellType.String);
                if (row == null || row.Cells.Count == 0)
                    continue;

                if (row.Cells[0].CellType == CellType.String && string.Compare(row.Cells[0].StringCellValue, rowName, true) == 0)
                {
                    rowIndex = r;
                    Console.WriteLine(r);
                    break;
                }
            }
            IRow firstRow = sheet.GetRow(0);
            for (int c = 0; c < firstRow.Cells.Count; c++)
            {
                if (firstRow.Cells[c].CellType == CellType.String && string.Compare(firstRow.Cells[c].StringCellValue, columnName, true) == 0)
                {
                    colIndex = c;
                    Console.WriteLine(c);
                    break;
                }
            }
            return rowIndex;
        }
        public static int Getcol(XSSFWorkbook wookbook, string sheetName, string rowName, string columnName)
        {
            ISheet sheet = wookbook.GetSheet(sheetName);

            if (sheet == null)
            {
                Console.WriteLine("The sheet is null");

            }

            int rowIndex = -1;
            int colIndex = -1;

            for (int r = 0; r < sheet.PhysicalNumberOfRows; r++)
            {
                IRow row = sheet.GetRow(r);
                Console.WriteLine(row.Cells[0].CellType == CellType.String);
                if (row == null || row.Cells.Count == 0)
                    continue;

                if (row.Cells[0].CellType == CellType.String && string.Compare(row.Cells[0].StringCellValue, rowName, true) == 0)
                {
                    rowIndex = r;
                    Console.WriteLine(r);
                    break;
                }
            }
            IRow firstRow = sheet.GetRow(0);
            for (int c = 0; c < firstRow.Cells.Count; c++)
            {
                if (firstRow.Cells[c].CellType == CellType.String && string.Compare(firstRow.Cells[c].StringCellValue, columnName, true) == 0)
                {
                    colIndex = c;
                    Console.WriteLine(c);
                    break;
                }
            }
            return colIndex;
        }

    }
    
}
