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
    public partial class Search_Class_Before_buildclass : Form
    {
        public Search_Class_Before_buildclass()
        {
            InitializeComponent();
        }

        private void Subject_Grade_Class_Search(object sender, EventArgs e)
        {
            panel1.Hide();
            try
            {
                var path = Directory.GetCurrentDirectory() + @"\excel\班級\";
                string[] lines = System.IO.File.ReadAllLines(path + "班級.txt");
                foreach (string line in lines)
                {
                    comboBox5.Items.Add(line.Replace(".xlsx",""));
                }
            }
            catch (Exception ex)
            {

            }
            panel3.Show();
        }

        private void Search_teacher_subject(object sender, EventArgs e)
        {
            panel1.Hide();
            panel2.Show();
        }

        private void back_to_Activity_background_system(object sender, EventArgs e)
        {
            this.Hide();
            Activity_background_system activity_Background_System = new Activity_background_system();
            activity_Background_System.ShowDialog();
        }

        private void build_class_btn(object sender, EventArgs e)
        {
            string cla = comboBox5.GetItemText(this.comboBox5.SelectedItem);
            ActivityList.filename = cla;
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
                Excel.Application excelApp1;
                Excel.Workbook wBook1;
                Excel.Worksheet wSheet1;
                //Excel.Range wRange;
                // 開啟一個新的應用程式
                excelApp1 = new Excel.Application();
                excelApp1.Visible = false;
                // 停用警告訊息
                excelApp1.DisplayAlerts = false;
                //var path = Directory.GetCurrentDirectory() + @"\excel\活動列表.xlsx";
                wBook = excelApp.Workbooks.Add();
                wBook1 = excelApp1.Workbooks.Open(Directory.GetCurrentDirectory() + @"\excel\班級\" + cla + ".xlsx");
                wSheet = wBook.Worksheets["工作表1"];
                wSheet1 = wBook1.Worksheets["工作表1"];
                excelApp.Cells[1, 1] = "ID";
                excelApp.Cells[1, 2] = "姓名";
                excelApp.Cells[1, 3] = "簽到時間";
                excelApp.Cells[1, 4] = "簽退時間";
                int j = 2;
                
                for (int i = 1; ; i++)
                {
                    if (excelApp1.Cells[i, 1].Value != null)
                    {
                        excelApp.Cells[j, 1] = excelApp1.Cells[i, 1];
                        excelApp.Cells[j++, 2] = excelApp1.Cells[i, 2];
                    }
                    else {
                        break;
                    }
                }
                
                // save the application  
                wBook.SaveAs(Directory.GetCurrentDirectory() + @"\excel\"+cla + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook1.SaveAs(Directory.GetCurrentDirectory() + @"\excel\班級\" + cla + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                wBook.Close();
                wBook1.Close();
                excelApp.Quit();
                excelApp1.Quit();
                wBook = null;
                wSheet = null;
                wBook1 = null;
                wSheet1 = null;
                excelApp = null;
                excelApp1 = null;
                ActivityPageSignIn.killexcel();
            }
            catch (Exception d)
            {
                d.ToString();
            }
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
                        excelApp.Cells[i, 1] = cla;
                        excelApp.Cells[i, 2] = "上課";
                        excelApp.Cells[i, 3] = "";
                        excelApp.Cells[i, 4] = "";
                        excelApp.Cells[i, 5] = "班級";
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
            this.Hide();
            ActivityPageSignIn activityPageSignIn = new ActivityPageSignIn();
            activityPageSignIn.ShowDialog();
        }
    }
}
