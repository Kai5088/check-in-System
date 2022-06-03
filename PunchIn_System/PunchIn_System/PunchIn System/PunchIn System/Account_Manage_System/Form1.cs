
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
                var path = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                Excel.Application excelApp;
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                //Excel.Range wRange;
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                wBook = excelApp.Workbooks.Open(path);
                wSheet = wBook.Worksheets["�u�@��1"];

                for (int x = 0; x < listBox1.Items.Count; x++)
                {
                    if (listBox1.GetSelected(x) == true)
                    {
                        string str = listBox1.Items[x].ToString();
                        for (int i = 1; ; i++)
                        {
                            if (excelApp.Cells[i, 1].Value.Equals(str))
                            {
                                excelApp.Cells[i, 3] = "�ϥΤ�";
                                message += str + " ";
                                //MessageBox.Show("�_�αb��" + message);
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


                MessageBox.Show("�_�αb��" + message);
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
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                // ��Excel���i��
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                // �[�J�s������ï
                var path1 = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                //MessageBox.Show(line);
                wBook = excelApp.Workbooks.Open(path1);

                // �]�w����ï�J�I
                //wBook.Activate();
                // �ޥβĤ@�Ӥu�@��
                wSheet = wBook.Worksheets["�u�@��1"];
                // �R�W�u�@���W��
                //wSheet.Name = "�u�@��1";
                // �]�w�u�@��J�I
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
                        if (excelApp.Cells[j, 3].Value == "����")
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
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                // ��Excel���i��
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                // �[�J�s������ï
                var path1 = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                //MessageBox.Show(line);
                wBook = excelApp.Workbooks.Open(path1);

                // �]�w����ï�J�I
                //wBook.Activate();
                // �ޥβĤ@�Ӥu�@��
                wSheet = wBook.Worksheets["�u�@��1"];
                // �R�W�u�@���W��
                //wSheet.Name = "�u�@��1";
                // �]�w�u�@��J�I
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
                        if (excelApp.Cells[j, 3].Value == "����")
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
                var path = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                Excel.Application excelApp;
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                //Excel.Range wRange;
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                wBook = excelApp.Workbooks.Open(path);
                wSheet = wBook.Worksheets["�u�@��1"];
                
                for (int x = 0; x < listBox2.Items.Count; x++)
                {
                    if (listBox2.GetSelected(x) == true)
                    {
                        string str = listBox2.Items[x].ToString();
                        for (int i=1; ; i++)
                        {
                            if(excelApp.Cells[i,1].Value.Equals(str))
                            {
                                excelApp.Cells[i,3] = "����";
                                message += str+" ";
                                //MessageBox.Show("���αb��" + message);
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
                
                
                MessageBox.Show("���αb��" + message);
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
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                // ��Excel���i��
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                // �[�J�s������ï
                var path1 = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                //MessageBox.Show(line);
                wBook = excelApp.Workbooks.Open(path1);

                // �]�w����ï�J�I
                //wBook.Activate();
                // �ޥβĤ@�Ӥu�@��
                wSheet = wBook.Worksheets["�u�@��1"];
                // �R�W�u�@���W��
                //wSheet.Name = "�u�@��1";
                // �]�w�u�@��J�I
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
                        if (excelApp.Cells[j, 3].Value == "�ϥΤ�")
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
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                // ��Excel���i��
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                // �[�J�s������ï
                var path1 = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                //MessageBox.Show(line);
                wBook = excelApp.Workbooks.Open(path1);

                // �]�w����ï�J�I
                //wBook.Activate();
                // �ޥβĤ@�Ӥu�@��
                wSheet = wBook.Worksheets["�u�@��1"];
                // �R�W�u�@���W��
                //wSheet.Name = "�u�@��1";
                // �]�w�u�@��J�I
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
                        if (excelApp.Cells[j, 3].Value == "�ϥΤ�")
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
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                // ��Excel���i��
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                // �[�J�s������ï
                var path1 = Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
                //MessageBox.Show(line);
                wBook = excelApp.Workbooks.Open(path1);

                // �]�w����ï�J�I
                //wBook.Activate();
                // �ޥβĤ@�Ӥu�@��
                wSheet = wBook.Worksheets["�u�@��1"];
                // �R�W�u�@���W��
                //wSheet.Name = "�u�@��1";
                // �]�w�u�@��J�I
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
                        if(excelApp.Cells[j, 3].Value == "�ϥΤ�")
                        {
                            listBox3.Items.Add(excelApp.Cells[j, 1].Value);
                        }
                        if (excelApp.Cells[j, 3].Value == "����")
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
            file.Title = "�п��Excel��";
            file.Filter = "�Ҧ��ɮ�(*.xlsx)|*.xlsx*";
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
            MessageBox.Show("�W��פJ���\");
            textBox3.Text = "";
            textBox3.Refresh();
            if (category.Equals("�@��"))
            {
                string path1 = Directory.GetCurrentDirectory() + @"\excel\";
                using StreamWriter file = new(path1 + "�W���ɮ�.txt", append: true);
                file.WriteLine(fname);
                file.Close();
            }
            if (category.Equals("�Z��"))
            {
                string path1 = Directory.GetCurrentDirectory() + @"\excel\�Z��\";
                using StreamWriter file = new(path1 + "�Z��.txt", append: true);
                file.WriteLine(fname);
                file.Close();
            }
            if (category.Equals("�W�Z"))
            {
                string path1 = Directory.GetCurrentDirectory() + @"\excel\�W�Z\";
                using StreamWriter file = new(path1 + "�W�Z.txt", append: true);
                file.WriteLine(fname);
                file.Close();
            }


        }
        private void button10_Click(object sender, EventArgs e)
        {
            var path= Directory.GetCurrentDirectory() + @"\excel\�t�κ޲z���b��.xlsx";
            string acc = textBox1.Text;
            string pw = textBox2.Text;
            string acc_identity = comboBox3.GetItemText(this.comboBox3.SelectedItem);
            try
            {
                Excel.Application excelApp;
                Excel.Workbook wBook;
                Excel.Worksheet wSheet;
                //Excel.Range wRange;
                // �}�Ҥ@�ӷs�����ε{��
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                // ����ĵ�i�T��
                excelApp.DisplayAlerts = false;
                wBook = excelApp.Workbooks.Open(path);
                wSheet = wBook.Worksheets["�u�@��1"];
                for (int i = 1; ; i++)
                {
                    if (excelApp.Cells[i, 1].Value != null)
                    {
                        if (acc.Equals(excelApp.Cells[i, 1].Value))
                        {
                            MessageBox.Show("���b���W�٤w�g�Q�ϥ�");
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
                        excelApp.Cells[i, 3] = "�ϥΤ�";
                        excelApp.Cells[i, 4] = acc_identity;
                        MessageBox.Show("�b��" + acc + "���\�إ�");
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox1.Refresh();
                        textBox2.Refresh();
                        break;
                    }

                }
                //MessageBox.Show("�b��" + acc + "���\�إ�");
                // save the application  
                wBook.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                //MessageBox.Show("�b��" + acc + "���\�إ�");
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

        private float x;//�w�q��e���骺�e��
        private float y;//�w�q��e���骺����
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
            //�M�����餤�������A���s�]�w����󪺭�
            foreach (Control con in cons.Controls)
            {
                //��������Tag�ݩʭȡA�ä��Ϋ��x�s�r��}�C
                if (con.Tag != null)
                {
                    string[] mytag = con.Tag.ToString().Split(new char[] { ';' });
                    //�ھڵ����Y�񪺤�ҽT�w����󪺭�
                    con.Width = Convert.ToInt32(System.Convert.ToSingle(mytag[0]) * newx);//�e��
                    con.Height = Convert.ToInt32(System.Convert.ToSingle(mytag[1]) * newy);//����
                    con.Left = Convert.ToInt32(System.Convert.ToSingle(mytag[2]) * newx);//����Z
                    con.Top = Convert.ToInt32(System.Convert.ToSingle(mytag[3]) * newy);//����Z
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;//�r���j�p
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