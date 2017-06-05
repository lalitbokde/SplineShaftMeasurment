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

    public partial class LoginForm : Form
    {

        Cls_UserManagement _objUsermanagement = new Cls_UserManagement();
        public string adminrole;
        public string result { get; set; }

        public LoginForm()
        {
            InitializeComponent();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            txt_username.Clear();
            txt_password.Clear();
            result = "cancel";
            this.Close();
           
        }

        private void btn_login_Click(object sender, EventArgs e)
        {
            txt_password.Focus();
            if (txt_username.Text!="" && txt_password.Text!="")
            {
                if (_objUsermanagement.login(txt_username.Text, txt_password.Text))
                {

                    result = "loginSuccess";

                    this.Close();

                }

                else
                {
                    MessageBox.Show("Incorrect ID or Password", "JMS Says ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txt_username.Text = "";
                    txt_password.Text = "";
                    txt_username.Focus();
                    result = "loginFail";

                }
            }
            else
            {
                result = "loginFail";
                MessageBox.Show("Please Enter Correct ID or Password", "JMS Says ", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        private void lnk_lbl_NewUser_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void txt_username_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_password.Focus();
            }
        }

        private void txt_password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_login.Focus();
            }
        }

        private void LoginForm_Load(object sender, EventArgs e)
        {
            txt_username.Select();

            if (DateTime.Now.Hour < 12)
            {
                txt_shift.Text = "Morning";
            }
            else if (DateTime.Now.Hour < 17)
            {
                txt_shift.Text = "Afternoon";
            }
            else
            {
                txt_shift.Text = "Evening";
            }
        }
    }
}
