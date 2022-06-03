
using PunchIn_System;
using PunchIn_System.Activity_Manage_System;
using PunchIn_System.Login;

namespace Management
{
    using Excel = Microsoft.Office.Interop.Excel;
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            x = this.Width;
            y = this.Height;
            setTag(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            panel1.Show();
            panel2.Hide();
            panel3.Hide();
            panel4.Hide();
            panel5.Hide();
            panel6.Hide();
        }
        private void button8_Click(object sender, EventArgs e)
        {
            string message = "";
            try
            {
                var path = Directory.GetCurrentDirectory() + @"\excel\系統管理員帳號.xlsx";
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

                for (int x = 0; x < listBox1.Items.Count; x++)
                {
                    if (listBox1.GetSelected(x) == true)
                    {
                        string str = listBox1.Items[x].ToString();
                        for (int i = 1; ; i++)
                        {
                            if (excelApp.Cells[i, 1].Value.Equals(str))
                            {
                                excelApp.Cells[i, 3] = "使用中";
                                message += str + " ";
                                //MessageBox.Show("復用帳號" + message);
                                break;
                            }

                        }
                    }
                    
                }

                // save the application  
                wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                wBook.Close();
                excelApp.Quit();


                MessageBox.Show("復用帳號" + message);
                wBook = null;
                wSheet = null;
                excelApp = null;

                ActivityPageSignIn.killexcel();

            }
            catch (Exception d)
            {
                d.ToString();
            }
            this.listBox1.Items.Clear();
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
                        break;
                    }
                    else
                    {
                        //MessageBox.Show("do");
                        if (excelApp.Cells[j, 3].Value == "停用")
                        {
                            listBox1.Items.Add(excelApp.Cells[j, 1].Value);
                        }


                    }

                }
                wBook.SaveAs(path1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close();
                excelApp.Quit();

                wBook = null;
                wSheet = null;

                excelApp = null;
                ActivityPageSignIn.killexcel();
            }
            catch (Exception ex)
            {

            }
            
            this.listBox1.Refresh();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            panel1.Hide();
            panel2.Show();
            panel3.Hide();
            panel4.Hide();
            panel5.Hide();
            panel6.Hide();
            this.listBox1.Items.Clear();
            this.listBox1.Refresh();
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
                        break;
                    }
                    else
                    {
                        //MessageBox.Show("do");
                        if (excelApp.Cells[j, 3].Value == "停用")
                        {
                            listBox1.Items.Add(excelApp.Cells[j, 1].Value);
                        }


                    }

                }
                wBook.SaveAs(path1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close();
                excelApp.Quit();

                wBook = null;
                wSheet = null;

                excelApp = null;
                ActivityPageSignIn.killexcel();
            }
            catch (Exception ex)
            {

            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            string message = "";
            try
            {
                var path = Directory.GetCurrentDirectory() + @"\excel\系統管理員帳號.xlsx";
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
                
                for (int x = 0; x < listBox2.Items.Count; x++)
                {
                    if (listBox2.GetSelected(x) == true)
                    {
                        string str = listBox2.Items[x].ToString();
                        for (int i=1; ; i++)
                        {
                            if(excelApp.Cells[i,1].Value.Equals(str))
                            {
                                excelApp.Cells[i,3] = "停用";
                                message += str+" ";
                                //MessageBox.Show("停用帳號" + message);
                                break;
                            }
                            
                        }                                                                                              
                    }
                    //MessageBox.Show(x.ToString());
                }
                
                // save the application  
                wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                wBook.Close();
                excelApp.Quit();
                
                
                MessageBox.Show("停用帳號" + message);
                wBook = null;
                wSheet = null;
                excelApp = null;
                
                ActivityPageSignIn.killexcel();
                
            }
            catch (Exception d)
            {
                d.ToString();
            }
            this.listBox2.Items.Clear();
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
                        break;
                    }
                    else
                    {
                        //MessageBox.Show("do");
                        if (excelApp.Cells[j, 3].Value == "使用中")
                        {
                            listBox2.Items.Add(excelApp.Cells[j, 1].Value);
                        }


                    }

                }
                wBook.SaveAs(path1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close();
                excelApp.Quit();

                wBook = null;
                wSheet = null;

                excelApp = null;
                ActivityPageSignIn.killexcel();
            }
            catch (Exception ex)
            {

            }
            this.listBox2.Refresh();
        }
        private void button3_Click(object sender, EventArgs e)
        {

            panel1.Hide();
            panel2.Hide();
            panel3.Show();
            panel4.Hide();
            panel5.Hide();
            panel6.Hide();
            this.listBox2.Items.Clear();
            this.listBox2.Refresh();
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
                        break;
                    }
                    else
                    {
                        //MessageBox.Show("do");
                        if (excelApp.Cells[j, 3].Value == "使用中")
                        {
                            listBox2.Items.Add(excelApp.Cells[j, 1].Value);
                        }
                        

                    }

                }
                wBook.SaveAs(path1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close();
                excelApp.Quit();
                
                wBook = null;
                wSheet = null;

                excelApp = null;
                ActivityPageSignIn.killexcel();
            }
            catch (Exception ex)
            {

            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel1.Hide();
            panel2.Hide();
            panel3.Hide();
            panel4.Show();
            panel5.Hide();
            panel6.Hide();
            this.listBox3.Items.Clear();

            //this.listBox4.Items.Clear();

            this.listBox3.Refresh();

            //this.listBox4.Refresh();

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
                        break;
                    }
                    else
                    {
                        //MessageBox.Show("do");
                        if(excelApp.Cells[j, 3].Value == "使用中")
                        {
                            listBox3.Items.Add(excelApp.Cells[j, 1].Value);
                        }
                        if (excelApp.Cells[j, 3].Value == "停用")
                        {
                            //listBox4.Items.Add(excelApp.Cells[j, 1].Value);
                        }
                        
                    }

                }
                wBook.SaveAs(path1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close();
                excelApp.Quit();
               
                wBook = null;
                wSheet = null;

                excelApp = null;
                ActivityPageSignIn.killexcel();
            }
            catch (Exception ex)
            {

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel1.Hide();
            panel2.Hide();
            panel3.Hide();
            panel4.Hide();
            panel5.Show();
            panel6.Hide();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.Title = "請選擇Excel檔";
            file.Filter = "所有檔案(*.xlsx)|*.xlsx*";
            if (file.ShowDialog() == DialogResult.OK)
            {
                this.textBox3.Text = file.SafeFileName;
                string selectedFileName = file.FileName;
                source_path= selectedFileName;
                fname = this.textBox3.Text;
                //MessageBox.Show(fname);
                textBox3.Refresh();
                //this.textBox3.Text=selectedFileName;
                //MessageBox.Show(selectedFileName);
            }
        }
        public static string source_path = "";
        public static string fname="";
        
         
        private void button9_Click(object sender, EventArgs e)
        {
            string category = "";
            string path = Directory.GetCurrentDirectory() + @"\excel\";
            category = comboBox4.GetItemText(this.comboBox4.SelectedItem);
            path += category + @"\";
            string target_path = path + fname;
            //MessageBox.Show(fname);
            //MessageBox.Show("source_path:" + source_path);
            //MessageBox.Show("target_path:" + target_path);
            System.IO.File.Copy(source_path, target_path, true);
            MessageBox.Show("名單匯入成功");
            textBox3.Text = "";
            textBox3.Refresh();
            if (category.Equals("一般"))
            {
                string path1 = Directory.GetCurrentDirectory() + @"\excel\";
                using StreamWriter file = new(path1 + "名單檔案.txt", append: true);
                file.WriteLine(fname);
                file.Close();
            }
            if (category.Equals("班級"))
            {
                string path1 = Directory.GetCurrentDirectory() + @"\excel\班級\";
                using StreamWriter file = new(path1 + "班級.txt", append: true);
                file.WriteLine(fname);
                file.Close();
            }
            if (category.Equals("上班"))
            {
                string path1 = Directory.GetCurrentDirectory() + @"\excel\上班\";
                using StreamWriter file = new(path1 + "上班.txt", append: true);
                file.WriteLine(fname);
                file.Close();
            }


        }
        private void button10_Click(object sender, EventArgs e)
        {
            var path= Directory.GetCurrentDirectory() + @"\excel\系統管理員帳號.xlsx";
            string acc = textBox1.Text;
            string pw = textBox2.Text;
            string acc_identity = comboBox3.GetItemText(this.comboBox3.SelectedItem);
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
                for (int i = 1; ; i++)
                {
                    if (excelApp.Cells[i, 1].Value != null)
                    {
                        if (acc.Equals(excelApp.Cells[i, 1].Value))
                        {
                            MessageBox.Show("此帳號名稱已經被使用");
                            textBox1.Text = "";
                            textBox2.Text = "";
                            textBox1.Refresh();
                            textBox2.Refresh();
                            break;
                        }
                    }
                    else
                    {
                        excelApp.Cells[i, 1] = acc;
                        excelApp.Cells[i,2] = pw;
                        excelApp.Cells[i, 3] = "使用中";
                        excelApp.Cells[i, 4] = acc_identity;
                        MessageBox.Show("帳號" + acc + "成功建立");
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox1.Refresh();
                        textBox2.Refresh();
                        break;
                    }

                }
                //MessageBox.Show("帳號" + acc + "成功建立");
                // save the application  
                wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                //MessageBox.Show("帳號" + acc + "成功建立");
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
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label12.Text = DateTime.Now.ToString();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private float x;//定義當前窗體的寬度
        private float y;//定義當前窗體的高度
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
        private void setControls(float newx, float newy, Control cons)
        {
            //遍歷窗體中的控制元件，重新設定控制元件的值
            foreach (Control con in cons.Controls)
            {
                //獲取控制元件的Tag屬性值，並分割後儲存字串陣列
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

        private void Form1_Resize_1(object sender, EventArgs e)
        {
            float newx = (this.Width) / x;
            float newy = (this.Height) / y;
            setControls(newx, newy, this);
        }

        private void back_to_ActivityList(object sender, EventArgs e)
        {
            this.Hide();
            //System_Choose system_Choose = new System_Choose();
            //system_Choose.ShowDialog();
            Login_Form login = new Login_Form();
            login.ShowDialog();
        }
    }
}