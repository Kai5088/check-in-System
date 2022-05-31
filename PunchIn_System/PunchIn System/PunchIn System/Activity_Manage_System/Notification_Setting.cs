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
    public partial class Notification_Setting : Form
    {
        public Notification_Setting()
        {
            InitializeComponent();
        }

        private void Cancel_to_SendEmail(object sender, EventArgs e)
        {
            this.Hide();
           
        }

        private void Check_to_SendEmail(object sender, EventArgs e)
        {

        }
    }
}
