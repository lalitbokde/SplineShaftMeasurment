using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CCD_Measurment_of_SSSA
{
    public partial class Frm_MenuList : Form
    {
        public Frm_MenuList()
        {
            InitializeComponent();
        }

        private void bttn_CreateUser_Click(object sender, EventArgs e)
        {
            UserRegistration _objUser = new UserRegistration();
            _objUser.ShowDialog();
        }

        private void bttn_Report_Click(object sender, EventArgs e)
        {

        }

        private void bttn_report_Click_1(object sender, EventArgs e)
        {
            Frm_DatewiseReport _objreport = new Frm_DatewiseReport();
            _objreport.ShowDialog();
        }
    }
}
