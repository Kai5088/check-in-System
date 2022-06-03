using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PunchIn_System.Activity_Manage_System;
using PunchIn_System.Login;
using Excel = Microsoft.Office.Interop.Excel;

namespace PunchIn_System.Activity_Manage_System
{
    public partial class ActivityList : Form
    {
        public ActivityList()
        {
            InitializeComponent();
            InitialListView();
        }

        private void InitialListView()
        {
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.LabelEdit = false;
            listView1.FullRowSelect = true;
            listView1.Columns.Add("活動編號", 100);
            listView1.Columns.Add("活動名稱", 200);
            listView1.Columns.Add("活動性質", 100);
            listView1.Columns.Add("活動時間", 200);
            listView1.Columns.Add("活動狀態", 100);
            var path = Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx";
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
                wBook = excelApp.Workbooks.Open(path);
                wSheet = wBook.Worksheets["工作表1"];
                int j = 1;
                for (int i = 1; ;i++)
                {
                    //MessageBox.Show(i.ToString());
                    if (excelApp.Cells[i, 1].Value != null)
                    {
                        if (!(excelApp.Cells[i,5].Value.Equals("班級"))&&!(excelApp.Cells[i, 5].Value.Equals("上班")) && !(excelApp.Cells[i, 5].Value.Equals("馬上簽到")))
                        {
                            var item = new ListViewItem($"No.{j++}");
                            item.SubItems.Add(excelApp.Cells[i, 1].Value);
                            item.SubItems.Add(excelApp.Cells[i, 2].Value);
                            item.SubItems.Add(excelApp.Cells[i, 3].Value);
                            listView1.Items.Add(item);
                        }
                        
                    }
                    else {
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
            //string[] lines = System.IO.File.ReadAllLines(path+"活動列表.txt");
            //int i = 0;
            /*foreach (string line in System.IO.File.ReadLines(path + "活動列表.txt"))
            {
                string[] subs = line.Trim().Split(" ");
                //MessageBox.Show(subs[0]+"**");
                // Use a tab to indent each line of the file.
                var item = new ListViewItem($"No.{i + 1}");
                item.SubItems.Add(subs[0]);
                item.SubItems.Add(subs[1]);
                item.SubItems.Add(subs[2]);
                listView1.Items.Add(item);
                i++;

            }*/
            /*
            foreach (string line in lines)
            {
                string[] subs = line.Split(' ');
                // Use a tab to indent each line of the file.
                var item = new ListViewItem($"No.{i + 1}");
                item.SubItems.Add(subs[0]);
                //item.SubItems.Add(subs[1]);
                listView1.Items.Add(item);
                i++;
                //Console.WriteLine("\t" + line);
            }
            */
        }

        //按下選擇後跳轉至打卡介面並開始顯示時間
        ActivityPageSignIn activityPageExam = new ActivityPageSignIn();
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            activityPageExam.timer1.Interval = 1000;
            activityPageExam.timer1.Start();
            activityPageExam.ShowDialog();
        }

        //返回
        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            //System_Choose system_Choose = new System_Choose();
            //system_Choose.ShowDialog();
            Login_Form login = new Login_Form();
            login.ShowDialog();
            
        }

        //建立活動
        private void Hold_Event(object sender, EventArgs e)
        {
            string path = Directory.GetCurrentDirectory();
            //MessageBox.Show(path);
            this.Hide();
            Activity_background_system activity_Background_System = new Activity_background_system();
            activity_Background_System.ShowDialog();
        }

        //搜尋活動
        private void Search_Event(object sender, EventArgs e)
        {
            this.Hide();
            Search_Event search_Event = new Search_Event();
            search_Event.ShowDialog();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string colunmName = listView1.Columns[0].Text;//获取第一列的标题名称

                //获取选择行第一列的值
                string colunmVal1 = listView1.SelectedItems[0].SubItems[0].Text;

                //获取选择行第二列的值
                filename = listView1.SelectedItems[0].SubItems[1].Text;
                
                //MessageBox.Show(returnfilename());
                this.Hide();
                activityPageExam.timer1.Interval = 1000;
                activityPageExam.timer1.Start();
                activityPageExam.ShowDialog();
            }
        }
        public static string filename="";
        
    }
}
