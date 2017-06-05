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
    public partial class ShimMaster : Form
    {
        SQLHelper _helper = new SQLHelper();
        BAL _bal = new BAL();
        Validation _objValidation = new Validation();

        public ShimMaster()
        {
            InitializeComponent();
        }

        private void bttn_reset_Click(object sender, EventArgs e)
        {
            txt_shim.ResetText();
            dtp_date.ResetText();
        }

        private void bttn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bttn_delete_Click(object sender, EventArgs e)
        {
            _bal.deleteShim(txt_id.Text);

            try
            {
                dgv_ShimMaster.Rows.RemoveAt(dgv_ShimMaster.CurrentCell.RowIndex);
                //clearall();
                //gettotalvalues();
            }
            catch (Exception)
            {
                MessageBox.Show("Unable To Delete", "Jms says", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void bttn_save_Click(object sender, EventArgs e)
        {
            _bal.AddShim(txt_shim.Text, dtp_date.Text);
            txt_shim.ResetText();
            dtp_date.ResetText();

            DataGridViewRow row = this.dgv_ShimMaster.RowTemplate;
            dgv_ShimMaster.Rows.Add();

            int gcount = Convert.ToInt32(this.dgv_ShimMaster.Rows.Count);
            gcount = gcount - 1;
            // dgv_ModelMaster.Rows[gcount].Cells["Id"].Value = txt_id.Text;
            dgv_ShimMaster.Rows[gcount].Cells["Shim"].Value = txt_shim.Text;
            dgv_ShimMaster.Rows[gcount].Cells["Date"].Value = dtp_date.Text;
        }

        private void bttn_update_Click(object sender, EventArgs e)
        {
            try
            {
                _bal.UpdateShim(txt_id.Text, txt_shim.Text, dtp_date.Text);

                dgv_ShimMaster.Rows[dgv_ShimMaster.CurrentRow.Index].Cells["ID"].Value = txt_id.Text;
                dgv_ShimMaster.Rows[dgv_ShimMaster.CurrentRow.Index].Cells["Shim"].Value = txt_shim.Text;
                // dgv_ModelMaster.Rows[dgv_ModelMaster.CurrentRow.Index].Cells["MachineID"].Value = dtp_date.Text;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dgv_ShimMaster_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dgv_ShimMaster.Rows[e.RowIndex];
                txt_shim.Text = row.Cells["Shim"].Value.ToString();
               //  dtp_date.Text = row.Cells["Date"].Value.ToString();
                txt_id.Text = row.Cells["ID"].Value.ToString();
            }
        }

        private void ShimMaster_Load(object sender, EventArgs e)
        {
            DataTable dt = _bal.SelectShim();
            int j = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dgv_ShimMaster.Rows.Add();
                dgv_ShimMaster.Rows[j].Cells["SrNo"].Value = dgv_ShimMaster.Rows.Count;
                dgv_ShimMaster.Rows[j].Cells["ID"].Value = dt.Rows[i]["ID"].ToString();
                dgv_ShimMaster.Rows[j].Cells["Shim"].Value = dt.Rows[i]["shim"].ToString();
                j++;
            }
        }
    }
}
