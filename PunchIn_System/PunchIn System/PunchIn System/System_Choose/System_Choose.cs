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

        //���}�t��
        private void label3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        //���ʵo�_�H
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            ActivityList activityList = new ActivityList();
            activityList.Show();
        }
        //�t�κ޲z��
        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(Login.Login_Form.identity);
            if (Login.Login_Form.identity.Equals("�t�κ޲z��"))
            {
                this.Hide();
                Form1 form1 = new Form1();
                form1.ShowDialog();
            }
            else
            {
                MessageBox.Show("�ܩ�p�A�S���v�����榹�ާ@","����",MessageBoxButtons.OK,MessageBoxIcon.Stop);
            }
        }
    }
}