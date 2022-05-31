using Management;
using PunchIn_System.Activity_Manage_System;

namespace PunchIn_System
{
    public partial class System_Choose : Form
    {
        public System_Choose()
        {
            InitializeComponent();
        }

        //離開系統
        private void label3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        //活動發起人
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            ActivityList activityList = new ActivityList();
            activityList.Show();
        }
        //系統管理員
        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(Login.Login_Form.identity);
            if (Login.Login_Form.identity.Equals("系統管理員"))
            {
                this.Hide();
                Form1 form1 = new Form1();
                form1.ShowDialog();
            }
            else
            {
                MessageBox.Show("很抱歉你沒有權限執行此操作","提示",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }
        }
    }
}