using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PunchIn_System.Login
{
    using Excel = Microsoft.Office.Interop.Excel;
    public partial class Login_Form : Form
    {
        public static string identity = "";
        public Login_Form()
        {
            InitializeComponent();
        }

        //登入
        private void button1_Click(object sender, EventArgs e)
        {
            string useracc = textBox1.Text;
            string userpw =textBox2.Text;
            
            try {
                Boolean enter = false;
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
                var path1 = Directory.GetCurrentDirectory() + @"\excel\系統管理員帳號.xlsx";
                //MessageBox.Show(line);
                wBook = excelApp.Workbooks.Open(path1);

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
                    if (excelApp.Cells[j, 1].Value == null)
                    {
                        MessageBox.Show("帳號或密碼輸入錯誤","ERROR",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox1.Refresh();
                        textBox2.Refresh();
                        break;
                    }
                    else
                    {
                        //MessageBox.Show("do");
                        if (excelApp.Cells[j, 1].Value == useracc)
                        {
                            if (excelApp.Cells[j, 2].Value == userpw)
                            {
                                identity = excelApp.Cells[j, 4].Value;
                                enter = true;
                                break;
                            }
                        }


                    }

                }
                wBook.SaveAs(path1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close();
                excelApp.Quit();

                wBook = null;
                wSheet = null;

                excelApp = null;
                Activity_Manage_System.ActivityPageSignIn.killexcel();
                if (enter == true)
                {
                    this.Hide();
                    System_Choose system_Choose = new System_Choose();
                    system_Choose.ShowDialog();
                }
                
            }
            catch(Exception ex) { 
            }
            
            
        }

        //離開系統
        private void label3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //clear field
        private void label2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
        }
    }
}
