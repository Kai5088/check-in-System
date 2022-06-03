

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
    using Excel = Microsoft.Office.Interop.Excel;
    using System.Drawing;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    public partial class Activity_background_system : Form
    {
        public Activity_background_system()
        {
            InitializeComponent();
        }

        private void Back_to_ActivityList(object sender, EventArgs e)
        {
            this.Hide();
            ActivityList activityList = new ActivityList();
            activityList.ShowDialog();
        }

        private void build_assembly(object sender, EventArgs e)
        {
            if(Login.Login_Form.identity.Equals("老師")|| Login.Login_Form.identity.Equals("老闆") || Login.Login_Form.identity.Equals("系統管理員"))
            {
                //this.Hide();
                //panel1.Hide();
                panel2.Show();
            }
            else
            {
                MessageBox.Show("很抱歉你沒有權限執行此操作", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

        }

        private void build_Assembly_btn(object sender, EventArgs e)
        {
            this.Hide();
            string event_name = textBox1.Text;
            string event_hoster = textBox2.Text;
            string event_time = dateTimePicker1.Text+comboBox1.Text+":"+comboBox2.Text;
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
                wBook = excelApp.Workbooks.Add();
                wSheet = wBook.Worksheets["工作表1"];
                excelApp.Cells[1, 1] = "ID";
                excelApp.Cells[1, 2] = "姓名";
                excelApp.Cells[1, 3] = "簽到時間";
                excelApp.Cells[1, 4] = "簽退時間";

                /*int i = 2;
                for (int x = 0; x < listBox2.Items.Count; x++)
                {
                    if (listBox2.GetSelected(x) == true)
                    {
                        string str = listBox2.Items[x].ToString();
                        string[] sub = str.Split(" ");
                        excelApp.Cells[i, 1] = sub[0];
                        excelApp.Cells[i, 2] = sub[1];
                        i++;
                        //MessageBox.Show(listBox1.Items[x].ToString());
                    }
                }*/

                string path = Directory.GetCurrentDirectory() + @"\excel\";
                // save the application  
                wBook.SaveAs(path + event_name, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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

            //MessageBox.Show(event_name+" "+event_hoster+" "+event_time+" "+member_num.ToString()); 
            MessageBoxButtons msgButton = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show("是否馬上建立簽到表單", "提示", msgButton);
            if (result == DialogResult.Yes)
            {
                MessageBox.Show("馬上開始簽到");
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
                    var path = Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx";
                    wBook = excelApp.Workbooks.Open(path);
                    wSheet = wBook.Worksheets["工作表1"];
                    for (int i = 1; ; i++)
                    {
                        if (excelApp.Cells[i, 1].Value == null)
                        {
                            excelApp.Cells[i, 1] = event_name;
                            excelApp.Cells[i, 2] = "集會";
                            excelApp.Cells[i, 3] = event_time;
                            excelApp.Cells[i, 4] = event_hoster;
                            excelApp.Cells[i, 5] = "馬上簽到";
                            break;
                        }
                        else
                        {
                            continue;
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
                ActivityList.filename = event_name;
                ActivityPageSignIn activityPageSignIn = new ActivityPageSignIn();
                activityPageSignIn.timer1.Interval = 1000;
                activityPageSignIn.timer1.Start();
                activityPageSignIn.ShowDialog();
            }
            else
            {
                MessageBox.Show("此集會放入活動列表");
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
                    var path = Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx";
                    wBook = excelApp.Workbooks.Open(path);
                    wSheet = wBook.Worksheets["工作表1"];
                    for (int i = 1; ; i++) {
                        if (excelApp.Cells[i, 1].Value == null) {
                            excelApp.Cells[i, 1] = event_name;
                            excelApp.Cells[i, 2] = "集會";
                            excelApp.Cells[i, 3] = event_time;
                            excelApp.Cells[i, 4] = event_hoster;
                            excelApp.Cells[i, 5] = "一般";
                            break;
                        }
                        else
                        {
                            continue;
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
                string path = Directory.GetCurrentDirectory() + @"\excel\";
                using StreamWriter file = new(path+"活動列表.txt", append: true);
                file.WriteLine(event_name+" "+"集會"+" "+ event_time+" "+ event_hoster);
                file.Close();
                */
                ActivityList activityList = new ActivityList();
                activityList.ShowDialog();

            }
        }

        private void build_Exam_btn(object sender, EventArgs e)
        {
            this.Hide();
            string test_subject = textBox4.Text;
            string test_place=textBox3.Text;
            string test_time = dateTimePicker2.Text+ comboBox6.Text+":"+ comboBox5.Text;
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
                wBook = excelApp.Workbooks.Add();
                wSheet = wBook.Worksheets["工作表1"];
                excelApp.Cells[1, 1] = "ID";
                excelApp.Cells[1, 2] = "姓名";
                excelApp.Cells[1, 3] = "簽到時間";
                excelApp.Cells[1, 4] = "簽退時間";
                int i = 2;
                for (int x = 0; x < listBox1.Items.Count; x++)
                {
                    if (listBox1.GetSelected(x) == true)
                    {
                        string str = listBox1.Items[x].ToString();
                        string[] sub=str.Split(" ");
                        excelApp.Cells[i,1]=sub[0];
                        excelApp.Cells[i,2]=sub[1];
                        i++;
                        //MessageBox.Show(listBox1.Items[x].ToString());
                    }
                }

                string path = Directory.GetCurrentDirectory() + @"\excel\";
                // save the application  
                wBook.SaveAs(path + test_subject, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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
            
            //MessageBox.Show(listBox1.SelectedValue.ToString());
            
            MessageBoxButtons msgButton = MessageBoxButtons.YesNo;
            DialogResult result = MessageBox.Show("是否馬上建立簽到表單", "提示", msgButton);
            if (result == DialogResult.Yes)
            {
                MessageBox.Show("馬上開始簽到");
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
                    var path = Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx";
                    wBook = excelApp.Workbooks.Open(path);
                    wSheet = wBook.Worksheets["工作表1"];
                    for (int i = 1; ; i++)
                    {
                        if (excelApp.Cells[i, 1].Value == null)
                        {
                            excelApp.Cells[i, 1] = test_subject;
                            excelApp.Cells[i, 2] = "考試";
                            excelApp.Cells[i, 3] = test_time;
                            excelApp.Cells[i, 4] = test_place;
                            excelApp.Cells[i, 5] = "馬上簽到";
                            break;
                        }
                        else
                        {
                            continue;
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
                ActivityList.filename = test_subject;
                ActivityPageSignIn activityPageSignIn = new ActivityPageSignIn();
                activityPageSignIn.timer1.Interval = 1000;
                activityPageSignIn.timer1.Start();
                activityPageSignIn.ShowDialog();
            }
            else
            {
                MessageBox.Show("此考試放入活動列表");
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
                    var path = Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx";
                    wBook = excelApp.Workbooks.Open(path);
                    wSheet = wBook.Worksheets["工作表1"];
                    for (int i = 1; ; i++)
                    {
                        if (excelApp.Cells[i, 1].Value == null)
                        {
                            excelApp.Cells[i, 1] = test_subject;
                            excelApp.Cells[i, 2] = "考試";
                            excelApp.Cells[i, 3] = test_time;
                            excelApp.Cells[i, 4] = test_place;
                            excelApp.Cells[i, 5] = "一般";
                            break;
                        }
                        else
                        {
                            continue;
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
                string path = Directory.GetCurrentDirectory() + @"\excel\";
                using StreamWriter file = new(path + "活動列表.txt", append: true);
                file.WriteLine(test_subject + " " + "考試"+" "+test_time+" "+ test_place);
                file.Close();
                */
                ActivityList activityList = new ActivityList();
                activityList.ShowDialog();

            }
        }

        private void build_Exam(object sender, EventArgs e)
        {
            if (Login.Login_Form.identity.Equals("老師") || Login.Login_Form.identity.Equals("系統管理員")) {
                panel1.Hide();
                panel3.Show();
            }
            else
            {
                MessageBox.Show("很抱歉你沒有權限執行此操作", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }


        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
        //建立上課打卡
        private void build_class(object sender, EventArgs e)
        {
            if (Login.Login_Form.identity.Equals("老師") || Login.Login_Form.identity.Equals("系統管理員"))
            {
                this.Hide();
                Search_Class_Before_buildclass search_Class_Before_Buildclass = new Search_Class_Before_buildclass();
                search_Class_Before_Buildclass.ShowDialog();
            }
            else
            {
                MessageBox.Show("很抱歉你沒有權限執行此操作", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        //建立上班打卡
        private void build_work(object sender, EventArgs e)
        {
            if (Login.Login_Form.identity.Equals("老闆") || Login.Login_Form.identity.Equals("系統管理員"))
            {
                this.Hide();
                Search_Work search_work = new Search_Work();
                search_work.ShowDialog();
            }
            else
            {
                MessageBox.Show("很抱歉你沒有權限執行此操作", "提示", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            /*
            ActivityPageSignIn activityPageSignIn = new ActivityPageSignIn();
            activityPageSignIn.timer1.Interval = 1000;
            activityPageSignIn.timer1.Start();
            activityPageSignIn.ShowDialog();
            */
        }

        private void Back_to_ActivityChoose(object sender, EventArgs e)
        {
            panel2.Hide();
            panel3.Hide();
            panel1.Show();
        }

        private void Activity_background_system_Load(object sender, EventArgs e)
        {
            var path = Directory.GetCurrentDirectory() + @"\excel\";
            string[] lines = System.IO.File.ReadAllLines(path + "名單檔案.txt");
            foreach (string line in lines)
            {
                try
                {
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
                    var path1 = Directory.GetCurrentDirectory() + @"\excel\一般\";
                    //MessageBox.Show(line);
                    wBook = excelApp.Workbooks.Open(path1 + line);
                    
                    // 設定活頁簿焦點
                    //wBook.Activate();
                    // 引用第一個工作表
                    wSheet = wBook.Worksheets["工作表1"];
                    // 命名工作表的名稱
                    //wSheet.Name = "工作表1";
                    // 設定工作表焦點
                    //wSheet.Activate();
                    //excelApp.Save(pathFile);
                    //MessageBox.Show(excelApp.Cells[1, 1].Value);
                    for (int j = 1; ; j++)
                    {
                        if (excelApp.Cells[j,1].Value == null)
                        {
                            
                            break;
                        }
                        else
                        {
                            //MessageBox.Show("do");
                            
                            listBox1.Items.Add(excelApp.Cells[j, 1].Value + " " + excelApp.Cells[j, 2].Value);
                            listBox2.Items.Add(excelApp.Cells[j, 1].Value + " " + excelApp.Cells[j, 2].Value);
                        }

                    }
                    wBook.SaveAs(path1 + line, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wBook.Close();
                    excelApp.Quit();
                    
                    wBook = null;
                    wSheet = null;

                    excelApp = null;
                }
                catch (Exception ex)
                {

                }
                // Use a tab to indent each line of the file.

                
                //Console.WriteLine("\t" + line);
            }
        }

        //搜尋參與人
        private void label14_Click(object sender, EventArgs e)
        {
            Search_participants search_Participants = new Search_participants();
            search_Participants.ShowDialog();
        }
    }
}
