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
    public partial class ModelMaster : Form
    {

        SQLHelper _helper = new SQLHelper();
        BAL _bal = new BAL();
        Validation _objValidation = new Validation();

        public ModelMaster()
        {
            InitializeComponent();
        }

        private void ModelMaster_Load(object sender, EventArgs e)
        {
            FillGrid();
        }

        public void FillGrid()
        {
            dgv_ModelMaster.Rows.Clear();
            int j = 0;
            foreach (DataRow dr in _bal.SelectModel().Rows)
            {
                dgv_ModelMaster.Rows.Add();
                dgv_ModelMaster.Rows[j].Cells["SrNo"].Value = dgv_ModelMaster.Rows.Count;
                dgv_ModelMaster.Rows[j].Cells["ID"].Value = dr["ID"];
                dgv_ModelMaster.Rows[j].Cells["Model"].Value = dr["model_name"];
                dgv_ModelMaster.Rows[j].Cells["master_value"].Value = dr["master_value"];
                dgv_ModelMaster.Rows[j].Cells["shim_verification_tolerence"].Value = dr["shim_verification_tolerence"];

                j++;
            }
        }



        private void bttn_save_Click(object sender, EventArgs e)
        {
            _bal.AddModel(txt_shim.Text, dtp_date.Text, txt_masterValue.Text, txt_tolerence_value.Text);
            Reset();
        }

        private void bttn_delete_Click(object sender, EventArgs e)
        {
            _bal.deleteModel(txt_id.Text);
            Reset();
        }

        private void bttn_update_Click(object sender, EventArgs e)
        {
            try
            {
                _bal.UpdateModel(txt_id.Text, txt_shim.Text, dtp_date.Text, txt_masterValue.Text, txt_tolerence_value.Text);
                Reset();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dgv_ModelMaster_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgv_ModelMaster.Rows[e.RowIndex];
                txt_shim.Text = row.Cells["Model"].Value.ToString();
                txt_id.Text = row.Cells["Id"].Value.ToString();
                txt_masterValue.Text = row.Cells["master_value"].Value.ToString();
                txt_tolerence_value.Text = row.Cells["shim_verification_tolerence"].Value.ToString();
            }
        }

        public void Reset()
        {
            txt_id.ResetText();
            txt_shim.ResetText();
            dtp_date.ResetText();
            txt_masterValue.ResetText();
            txt_tolerence_value.ResetText();
            FillGrid();
        }

        private void bttn_reset_Click(object sender, EventArgs e)
        {
            Reset();
        }

        private void bttn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txt_masterValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && (!char.IsDigit(e.KeyChar))
       && (e.KeyChar != '.') && (e.KeyChar != '-'))
                e.Handled = true;

            // only allow one decimal point
            if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
                e.Handled = true;

            // only allow minus sign at the beginning
            if (e.KeyChar == '-' && (sender as TextBox).Text.Length > 0)
                e.Handled = true;
        }

        private void txt_tolerence_value_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && (!char.IsDigit(e.KeyChar))
       && (e.KeyChar != '.') && (e.KeyChar != '-'))
                e.Handled = true;

            // only allow one decimal point
            if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
                e.Handled = true;

            // only allow minus sign at the beginning
            if (e.KeyChar == '-' && (sender as TextBox).Text.Length > 0)
                e.Handled = true;
        }

        private void txt_shim_verification_tolerence_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                bttn_save.Focus();
            }
        }

        private void txt_shim_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_masterValue.Focus();
            }
        }

        private void txt_masterValue_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_tolerence_value.Focus();
            }
        }
        
    }
}
