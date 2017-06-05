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
    public partial class UserRegistration : Form
    {
        Cls_UserManagement _objUsermanagement = new Cls_UserManagement();

        public UserRegistration()
        {
            InitializeComponent();
        }


        String SRNO = "";

        private void bttn_User_Click(object sender, EventArgs e)
        {
            try
            {
                if (txt_password.Text != "" && txt_username.Text != "" && cmb_role.Text != "")
                {
                    _objUsermanagement.AddUser(txt_username.Text, txt_password.Text, cmb_role.Text);
                    MessageBox.Show("User Created");
                    FillGrid();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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
                cmb_role.Focus();
            }
        }

        private void cmb_role_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                bttn_User.Focus();
            }
        }

        private void UserRegistration_Load(object sender, EventArgs e)
        {
            Reset();
        }

        public void Reset()
        {
            SRNO ="";
            txt_username.ResetText();
            txt_password.ResetText();
            cmb_role.ResetText();
            FillGrid();
        }

        private void FillGrid()
        {
            dgv_users.Rows.Clear();
            int j = 0;
            foreach (DataRow dr in _objUsermanagement.getAllUsers().Rows)
            {
                dgv_users.Rows.Add();
                dgv_users.Rows[j].Cells["ID"].Value = dr["ID"];
                dgv_users.Rows[j].Cells["user_name"].Value = dr["user_name"];
                dgv_users.Rows[j].Cells["user_password"].Value = dr["user_password"];
                dgv_users.Rows[j].Cells["Role_Name"].Value = dr["Role_Name"];

                j++;
            }
        }

        private void bttn_cancel_Click(object sender, EventArgs e)
        {

        }

        private void btn_Update_Click(object sender, EventArgs e)
        {
            try
            {
                _objUsermanagement.UpdateModel(SRNO, txt_username.Text, txt_password.Text, cmb_role.Text);
                MessageBox.Show("Record Updated");
                Reset();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            _objUsermanagement.deleteUser(SRNO);
            MessageBox.Show("Record Deleted");
            Reset();
        }

        private void dgv_users_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgv_users.Rows[e.RowIndex];
                SRNO = row.Cells["ID"].Value.ToString();
                txt_username.Text = row.Cells["user_name"].Value.ToString();
                txt_password.Text = row.Cells["user_password"].Value.ToString();
                cmb_role.Text = row.Cells["Role_Name"].Value.ToString();
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Reset();
        }
    }
}
