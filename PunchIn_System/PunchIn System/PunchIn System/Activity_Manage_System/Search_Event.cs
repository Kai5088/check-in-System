using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PunchIn_System.Activity_Manage_System
{
    public partial class Search_Event : Form
    {
        public Search_Event()
        {
            InitializeComponent();
        }

        //返回活動清單
        private void Back_to_ActivityList(object sender, EventArgs e)
        {
            this.Hide();
            ActivityList activityList = new ActivityList();
            activityList.ShowDialog();
        }

        private void Search_Work(object sender, EventArgs e)
        {
            this.Hide();
            Search_Work search_Work = new Search_Work();
            search_Work.ShowDialog();
        }

        private void Search_Class(object sender, EventArgs e)
        {
            this.Hide();
            Search_Class search_Class = new Search_Class();
            search_Class.ShowDialog();
        }
    }
}
