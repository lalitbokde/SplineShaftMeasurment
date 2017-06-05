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
    public partial class Spline_Shaft_MeasurmentTabDesing : Form
    {
        SQLHelper _helper = new SQLHelper();
        BAL _bal = new BAL();
        Validation _objValidation = new Validation();
        LoginForm _login = new LoginForm();

        double shim;
        string master_value, CCD_Value, shim_verification_tolerence;
        bool verifyAll = false;
        public Spline_Shaft_MeasurmentTabDesing()
        {
            InitializeComponent();
        }

        private void PTOMesurementTabDesing_Load(object sender, EventArgs e)
        {
            dgv_PTOMesurement.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 12F, FontStyle.Bold);
            this.Hide();

        login:
            _login.ShowDialog();
            if (_login.result == "loginSuccess")
            {
                this.Show();
                if (_bal.checkAdmin(_login.txt_username.Text))
                {
                    bttn_Menu.Visible = true;
                }
                txt_username.Text = _login.txt_username.Text;
                txt_shift.Text = _login.txt_shift.Text;

                _login.txt_username.Text = "";
                _login.txt_shift.Text = "";
                _login.txt_password.Text = "";
            }
            else if (_login.result == "loginFail")
            {
                goto login;
            }
            else
            {
                Application.Exit();
            }
            //lbl_Time.Text = DateTime.Now.ToString("dd/MM/yyyy ");
            txt_time.Text = DateTime.Now.ToString("hh:mm tt");

            FillGrid();
            cmb_selectmodel.DataSource = _bal.getModel(cmb_selectmodel.Text);
            cmb_selectmodel.Focus();
        }

        private void bttn_checkvalue_Click(object sender, EventArgs e)
        {

        }

        private void bttn_checkshims_Click(object sender, EventArgs e)
        {

            int count0point1microm = 0;
            int count0point2microm = 0;
            int count0point5microm = 0;

            shim = Math.Abs(Convert.ToDouble(txt_masterValue_page2.Text) * 1000);

            while (shim >= 50)
            {
                if (shim >= 500)
                {
                    shim = shim - 500;
                    count0point5microm++;
                    continue;
                }

                if (shim >= 200)
                {
                    shim = shim - 200;
                    count0point2microm++;
                    continue;
                }


                if (shim >= 100)
                {
                    shim = shim - 100;
                    count0point1microm++;
                    continue;
                }


            }

            string MicroM1 = "0.1 mm :";
            string MicroM2 = "0.2 mm :";
            string MicroM3 = "0.5 mm :";

            LB_Total_shims.Items.Clear();

            if (count0point1microm != 0)
                LB_Total_shims.Items.Add(MicroM1 + count0point1microm.ToString());
            if (count0point2microm != 0)
                LB_Total_shims.Items.Add(MicroM2 + count0point2microm.ToString());
            if (count0point5microm != 0)
                LB_Total_shims.Items.Add(MicroM3 + count0point5microm.ToString());
        }

        private void bttn_checkall_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(cmb_selectmodel.Text))
            {
                MessageBox.Show("Please enter a value");
                cmb_selectmodel.Focus();
            }
            if (String.IsNullOrEmpty(txt_masterValue_page1.Text))
            {
                MessageBox.Show("Please enter a value");
                txt_masterValue_page1.Focus();
            }


            if (String.IsNullOrEmpty(txt_dailgaugereading_page1.Text))
            {
                MessageBox.Show("Please enter a value");
                txt_dailgaugereading_page1.Focus();
            }
            if (String.IsNullOrEmpty(txt_masterValue_page1.Text))
            {
                MessageBox.Show("Please enter a value");
                txt_masterValue_page1.Focus();
            }
        }

        private void bttn_checkreset_Click(object sender, EventArgs e)
        {
            txt_dailgaugereading_page1.Text = _objValidation.Gettingdecimals(txt_dailgaugereading_page1.Text = "0");
            txt_masterValue_page1.Text = _objValidation.Gettingdecimals(txt_masterValue_page1.Text = master_value.ToString());
        }

        private void MasterTab_Click(object sender, EventArgs e)
        {

        }

        private void bttn_checkreset_driving_Click(object sender, EventArgs e)
        {
            //txt_dialgaugereading_driving_page1.Text = _objValidation.Gettingdecimals(txt_dialgaugereading_driving_page1.Text = "0");
            //txt_masterValueofY.Text = _objValidation.Gettingdecimals(txt_masterValueofY.Text = master_value_of_X.ToString());
        }



        private void bttn_checkvalue_Click_1(object sender, EventArgs e)
        {
        }

        private void bttn_checkshims_Click_1(object sender, EventArgs e)
        {


        }

        private void bttn_checkvalue_driving_Click(object sender, EventArgs e)
        {
        }

        private void bttn_checkshims_driving_Click(object sender, EventArgs e)
        {


        }

        private void bttn_DialGaugeReading_Click(object sender, EventArgs e)
        {

            shim = Math.Abs(Convert.ToDouble(txt_masterValue_page2.Text));
            if (((Convert.ToDouble(txt_masterValue_page2.Text) * 1000) - (Convert.ToDouble(txt_verifieddailgaugereading_verification_page3.Text) * 1000)) < Convert.ToDouble(shim_verification_tolerence) * 1000)
            {
                Verified.Visible = true;
                NotVerified.Visible = false;

            }
            else
            {
                Verified.Visible = false;
                NotVerified.Visible = true;
            }
            // txt_shimsRequired.Text = txt_dialgaugereading.Text;
        }
        public void DrivenGaugereadeingCheck()
        {
            try
            {
                if (txt_masterValue_page2.Text != "" && txt_verifieddailgaugereading_verification_page3.Text != "")
                {
                    shim = Math.Abs(Convert.ToDouble(txt_masterValue_page2.Text));
                    if (Math.Abs(((Convert.ToDouble(txt_shimsRequiredDrivenVerify_page3.Text) * 1000) - (Convert.ToDouble(txt_verifieddailgaugereading_verification_page3.Text) * 1000))) < Convert.ToDouble(shim_verification_tolerence) * 1000)
                    {
                        Verified.Visible = true;
                        NotVerified.Visible = false;
                        
                    }
                    else
                    {
                        Verified.Visible = false;
                        NotVerified.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Verified.Visible = false;
                NotVerified.Visible = false;
            }
        }

        private bool VerifyAll()
        {
            verifyAll = true;

            if (String.IsNullOrEmpty(cmb_selectmodel.Text))
            {
                lbl_selectmodel.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_selectmodel.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_dailgaugereading_page1.Text) || Convert.ToDouble(txt_dailgaugereading_page1.Text) != 0)
            {
                lbl_gauge_reading.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_gauge_reading.ForeColor = Color.Black;
            }


            if (String.IsNullOrEmpty(txt_masterValue_page1.Text))
            {
                lbl_mastervalueofx.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_mastervalueofx.ForeColor = Color.Black;
            }




            //third txt_dailgaugereading_driven_page3

            if (String.IsNullOrEmpty(txt_verifieddailgaugereading_verification_page3.Text))
            {
                lbl_verifieddialgaugereading.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_verifieddialgaugereading.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_dailgaugereading_page3.Text))
            {
                lbl_gauge_reading_page3.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_gauge_reading_page3.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_manualReading_page3.Text))
            {
                lbl_manualReading_page3.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_manualReading_page3.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_shimsRequiredDrivenVerify_page3.Text))
            {
                lbl_shimRequiredDriven_verify.ForeColor = Color.Red;

                verifyAll = false;
            }
            else
            {
                lbl_shimRequiredDriven_verify.ForeColor = Color.Black;
            }

            return verifyAll;

        }

        private void bttn_next_Click(object sender, EventArgs e)
        {
            if (ValidFirst())
            {
                this.PtoMeasurement.SelectedTab = tab_shimmeasurement;

                txt_masterValue_page2.Text = txt_masterValue_page1.Text;
                txt_dailgaugereading_page2.Select();
            }
        }
        private bool ValidFirst()
        {
            bool checkvalue = true;
            if (String.IsNullOrEmpty(cmb_selectmodel.Text))
            {
                lbl_selectmodel.ForeColor = Color.Red;

                checkvalue = false;
            }
            else
            {
                lbl_selectmodel.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_dailgaugereading_page1.Text))
            {
                lbl_gauge_reading.ForeColor = Color.Red;

                checkvalue = false;
            }
            else
            {
                lbl_gauge_reading.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_masterValue_page1.Text))
            {
                lbl_mastervalueofx.ForeColor = Color.Red;

                checkvalue = false;
            }
            else
            {
                lbl_mastervalueofx.ForeColor = Color.Black;
            }



            return checkvalue;
        }

        private void txt_dailgaugereading_deriven_Validated(object sender, EventArgs e)
        {
            //txt_shimsRequired_page2.Text = _objValidation.Gettingdecimals(txt_shimsRequired_page2.Text);
        }

        private void txt_dialgaugereading_driving_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_mastervalue_KeyPress(object sender, KeyPressEventArgs e)
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



        private void txt_mastervalue_driving_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_shimsRequired_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_shimsreq_driving_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_dialgaugereading_KeyPress(object sender, KeyPressEventArgs e)
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

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
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

        private Dictionary<TabPage, Color> TabColors = new Dictionary<TabPage, Color>();
        private void SetTabHeader(TabPage page, Color color)
        {
            TabColors[page] = color;
            PtoMeasurement.Invalidate();
        }
        private void PtoMeasurement_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabPage page = PtoMeasurement.TabPages[e.Index];
            e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(239, 127, 26)), e.Bounds);

            Rectangle paddedBounds = e.Bounds;
            int yOffset = (e.State == DrawItemState.Selected) ? -2 : 1;
            paddedBounds.Offset(1, yOffset);
            Font f = new Font("Arial", 12, FontStyle.Bold);
            TextRenderer.DrawText(e.Graphics, page.Text, f, paddedBounds, Color.White);

          
        }

        private void bttn_saveAll_Click(object sender, EventArgs e)
        {
            if (VerifyAll())
            {
                string total_shim_required = "", shim_required_driving = "";
                foreach (var item in LB_Total_shims.Items)
                {
                    total_shim_required += item + "  ";
                }

                _bal.AddPTOMeasurementData(cmb_selectmodel.SelectedValue.ToString(), txt_masterValue_page2.Text, txt_dailgaugereading_page2.Text, txt_dailgaugereading_page2.Text, txt_manualReading_page3.Text, txt_shimsRequiredDrivenVerify_page3.Text, total_shim_required, txt_verifieddailgaugereading_verification_page3.Text, txt_username.Text, DateTime.Now);
                FillGrid();
                Masterresetall();
                MessageBox.Show("Data Saved");
                this.PtoMeasurement.SelectedTab = MasterTab;
            }
            else
            {
                MessageBox.Show("Verify All values");
            }
        }

        private void Masterresetall()
        {
            txt_masterValue_page1.ResetText();
            txt_dailgaugereading_page3.ResetText();

            txt_manualReading_page3.ResetText();
            txt_dailgaugereading_page2.Text = "0";
            txt_manualReading_page2.Text = "0";
            txt_shimsRequiredDrivenVerify_page2.ResetText();
            txt_dailgaugereading_page1.ResetText();
            txt_verifieddailgaugereading_verification_page3.ResetText();
            txt_shimsRequiredDrivenVerify_page3.ResetText();
            LB_Total_shims.Items.Clear();
            Verified.Visible = false;
            NotVerified.Visible = false;
        }

        private void FillGrid()
        {
            dgv_PTOMesurement.Rows.Clear();
            int i = 0;
            foreach (DataRow dt in _bal.getPtoMeasurementData().Rows)
            {
                if (Convert.ToDateTime(dt["CreateDate"].ToString()).ToString("dd/MM/yyyy") == DateTime.Now.ToString("dd/MM/yyyy"))
                {
                    dgv_PTOMesurement.Rows.Add();
                    dgv_PTOMesurement.Rows[i].Cells["modelid"].Value = _bal.getModel(Convert.ToInt32(dt["modelid"].ToString())).Rows[0]["model_name"].ToString();
                    dgv_PTOMesurement.Rows[i].Cells["master_value_"].Value = dt["master_value"];
                    dgv_PTOMesurement.Rows[i].Cells["dialguage_reading"].Value = dt["dialguage_reading"];
                    dgv_PTOMesurement.Rows[i].Cells["actual_reading"].Value = dt["actual_reading"];
                    dgv_PTOMesurement.Rows[i].Cells["manual_reading"].Value = dt["manual_reading"];
                    dgv_PTOMesurement.Rows[i].Cells["shim_required"].Value = dt["shim_required"];
                    dgv_PTOMesurement.Rows[i].Cells["total_shims"].Value = dt["total_shims"];
                    dgv_PTOMesurement.Rows[i].Cells["shim_required_verification"].Value = dt["shim_required_verification"];
                    dgv_PTOMesurement.Rows[i].Cells["operator_name"].Value = dt["operator_name"];
                    dgv_PTOMesurement.Rows[i].Cells["createdate"].Value = dt["createdate"];
                    i++;
                }
            }
        }
        
        private void tab_shimmeasurement_Enter(object sender, EventArgs e)
        {
            if (!ValidFirst())
            {
                bttn_exportexcel.Visible = false;

                this.PtoMeasurement.SelectedTab = MasterTab;
            }
            else
            {
                bttn_exportexcel.Visible = false;
            }
        }

        private void tab_shimVerification_Enter(object sender, EventArgs e)
        {
            if (!ValidFirst())
            {
                bttn_exportexcel.Visible = false;

                this.PtoMeasurement.SelectedTab = MasterTab;
            }
            else if (!ValidSecond())
            {
                bttn_exportexcel.Visible = false;

                this.PtoMeasurement.SelectedTab = tab_shimmeasurement;
            }
            else
            {
                bttn_exportexcel.Visible = false;
                txt_verifieddailgaugereading_verification_page3.Focus();
            }
        }

        private void lbl_addModel_Click(object sender, EventArgs e)
        {
            ModelMaster _objShow = new ModelMaster();
            _objShow.ShowDialog();
            cmb_selectmodel.DataSource = _bal.getModel(cmb_selectmodel.Text);
        }

        private void txt_dailgaugereading_driven_page1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToDouble(txt_dailgaugereading_page1.Text) == 0)
                {
                    txt_masterValue_page1.Text = master_value; // "4.512"
                    txt_dailgaugereading_page1.Text = "0.000";
                }
            }
            catch
            {
                if (txt_dailgaugereading_page1.Text != "")
                {
                    MessageBox.Show("Please Enter Zero Value from Dial Gauage");
                }
            }
        }

        private void txt_dialgaugereading_driving_page1_TextChanged(object sender, EventArgs e)
        {
            //try
            //{
            //    if (Convert.ToDouble(txt_dialgaugereading_driving_page1.Text) == 0)
            //    {
            //        txt_dialgaugereading_driving_page1.Text = "0.000";
            //        txt_masterValueofY.Text = master_value_of_Y.ToString();
            //    }
            //}
            //catch
            //{
            //    if (txt_dialgaugereading_driving_page1.Text != "")
            //    {
            //        MessageBox.Show("Please Enter Zero Value from Dial Gauage");
            //    }
            //}
        }

        private void groupBox_drivingshaft_Enter(object sender, EventArgs e)
        {

        }

        private void txt_dailgaugereading_driven_page3_TextChanged(object sender, EventArgs e)
        {
            DrivenGaugereadeingCheck();
        }

        private void txt_dialgaugereading_driving_page3_TextChanged(object sender, EventArgs e)
        {

        }

        private void PTOMesurementTabDesing_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyData == Keys.F2)
            {
                if (PtoMeasurement.SelectedTab == MasterTab)
                {
                    if (ValidFirst())
                    {
                        this.PtoMeasurement.SelectedTab = tab_shimmeasurement;
                    }
                }
                else if (PtoMeasurement.SelectedTab == tab_shimmeasurement)
                {
                    if (ValidSecond())
                    {
                        this.PtoMeasurement.SelectedTab = tab_shimVerification;

                        txt_verifieddailgaugereading_verification_page3.Focus();
                    }
                }
                else if (PtoMeasurement.SelectedTab == tab_shimVerification)
                {
                    this.PtoMeasurement.SelectedTab = tab_report;
                    bttn_exportexcel.Visible = true;

                }
            }

            if (e.KeyData == Keys.F1)
            {
                if (PtoMeasurement.SelectedTab == MasterTab)
                {

                }
                else if (PtoMeasurement.SelectedTab == tab_shimmeasurement)
                {

                    this.PtoMeasurement.SelectedTab = MasterTab;

                }
                else if (PtoMeasurement.SelectedTab == tab_shimVerification)
                {
                    this.PtoMeasurement.SelectedTab = tab_shimmeasurement;
                }
                else if (PtoMeasurement.SelectedTab == tab_report)
                {
                    this.PtoMeasurement.SelectedTab = tab_shimVerification;

                }

            }

            if (e.KeyCode == Keys.S && e.Modifiers == Keys.Control)
            {
                bttn_saveAll_Click(this, new EventArgs());
            }

            if (e.KeyCode == Keys.M && e.Modifiers == Keys.Control)
            {
                lbl_addModel_Click(this, new EventArgs());
            }
        }

        private void txt_mastervalue_driving_page3_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_dialgaugereading_driving_page1_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_mastervalue_driven_page2_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_mastervalue_driven_page2_KeyPress_1(object sender, KeyPressEventArgs e)
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

        private void lbl_F1_Click(object sender, EventArgs e)
        {

        }

        private void txt_dailgaugereading_driven_page2_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_dialgaugereading_driving_page2_KeyPress(object sender, KeyPressEventArgs e)
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

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void bttn_Menu_Click(object sender, EventArgs e)
        {
            Frm_MenuList _objMenu = new Frm_MenuList();
            _objMenu.ShowDialog();
        }

        private void txt_splineshaftmodelno_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_dailgaugereading_page1.Focus();
            }
        }

        private void MasterTab_Enter(object sender, EventArgs e)
        {
            bttn_exportexcel.Visible = false;

            txt_dailgaugereading_page1.Focus();
        }

        private void txt_dailgaugereading_driven_page3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                System.Threading.Thread.Sleep(5000);
                if (Verified.Visible == true)
                {
                    if (VerifyAll())
                    {
                        string total_shim_required = "", shim_required_driving = "";
                        foreach (var item in LB_Total_shims.Items)
                        {
                            total_shim_required += item + "  ";
                        }

                        _bal.AddPTOMeasurementData(cmb_selectmodel.SelectedValue.ToString(), txt_masterValue_page2.Text, txt_dailgaugereading_page2.Text, txt_dailgaugereading_page2.Text, txt_manualReading_page3.Text, txt_shimsRequiredDrivenVerify_page3.Text, total_shim_required, txt_verifieddailgaugereading_verification_page3.Text, txt_username.Text, DateTime.Now);
                        FillGrid();
                        Masterresetall();
                        MessageBox.Show("Data Saved");
                        this.PtoMeasurement.SelectedTab = MasterTab;
                    }
                    else
                    {
                        MessageBox.Show("Verify All values");
                    }
                }
            }
        }

        private void txt_dailgaugereading_driven_page1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txt_dailgaugereading_page1.Text == "")
                {
                    return;
                }
                else
                {
                    bttn_next.Focus();
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            PTOMesurementTabDesing_Load(this, new EventArgs());
        }

        private void cmb_selectmodel_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmb_selectmodel_SelectedValueChanged(object sender, EventArgs e)
        {
            DataTable dtmodel = _bal.getModelDetailsById(cmb_selectmodel.SelectedValue.ToString());
            master_value = _objValidation.Gettingdecimals(dtmodel.Rows[0]["master_value"].ToString());
            shim_verification_tolerence = _objValidation.Gettingdecimals(dtmodel.Rows[0]["shim_verification_tolerence"].ToString());

        }

        private void bttn_exportexcel_Click(object sender, EventArgs e)
        {
            if (dgv_PTOMesurement.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet1;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //xlWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                int i = 0;
                int j = 0;
                for (i = 0; i < 1; i++)
                {
                    for (j = 0; j <= dgv_PTOMesurement.ColumnCount - 1; j++)
                    {
                        string header = dgv_PTOMesurement.Columns[j].HeaderText;
                        xlWorkSheet.Cells[i + 1, j + 1] = header;
                    }
                }

                for (i = 0; i <= dgv_PTOMesurement.RowCount - 1; i++)
                {
                    for (j = 0; j <= dgv_PTOMesurement.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dgv_PTOMesurement[j, i];
                        xlWorkSheet.Cells[i + 2, j + 1] = cell.Value;
                    }
                }
                // xlWorkSheet1.Cells[dgv_problem_master.RowCount, dgv_problem_master.ColumnCount] = "http://csharp.net-informations.com";
                //xlWorkSheet1.Cells[dgv_problem_master.RowCount, dgv_problem_master.ColumnCount] = "Chart Image";

                //xlWorkSheet1.Shapes.AddPicture(@"F:\vaishali\LearningMOBI\WebProject\img1.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 50, 50 + dgv_problem_master.RowCount, 300, 45); 
                string time = DateTime.Now.ToString("HH:mm").Replace(':', '_');
                string excelsheet_name = Environment.CurrentDirectory + "\\ExcelFiles\\" + "Spline_Shaft_Measurment" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + "_" + time + ".xls";
                xlWorkBook.SaveAs(excelsheet_name, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                // releaseObject(xlWorkSheet1);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Excel file created , you can find the file " + excelsheet_name);
            }
            else
            {
                MessageBox.Show("There is no data to export");

            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void tab_report_Enter(object sender, EventArgs e)
        {
            bttn_exportexcel.Visible = true;

        }


        private void txt_dialgaugereading_driving_page1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                bttn_next.Focus();
            }
        }

        private void txt_Dial_Gauge_Reading_of_Y_page2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txt_masterValue_page2.Text != "" && txt_masterValue_page2.Text != "")
                {
                    txt_masterValue_page2.Text = _objValidation.Gettingdecimals(((Convert.ToDouble(txt_masterValue_page2.Text) - Convert.ToDouble(txt_masterValue_page2.Text)) - (0.2)).ToString());


                    int count0point1microm = 0;
                    int count0point2microm = 0;
                    int count0point5microm = 0;

                    shim = Math.Abs(Convert.ToDouble(txt_masterValue_page2.Text) * 1000);

                    while (shim >= 50)
                    {

                        if (shim >= 500)
                        {
                            shim = shim - 500;
                            count0point5microm++;
                            continue;
                        }

                        if (shim >= 200)
                        {
                            shim = shim - 200;
                            count0point2microm++;
                            continue;
                        }


                        if (shim >= 100)
                        {
                            shim = shim - 100;
                            count0point1microm++;
                            continue;
                        }


                    }

                    string MicroM1 = "0.1 mm :";
                    string MicroM2 = "0.2 mm :";
                    string MicroM3 = "0.5 mm :";

                    LB_Total_shims.Items.Clear();

                    if (count0point1microm != 0)
                        LB_Total_shims.Items.Add(MicroM1 + count0point1microm.ToString());
                    if (count0point2microm != 0)
                        LB_Total_shims.Items.Add(MicroM2 + count0point2microm.ToString());
                    if (count0point5microm != 0)
                        LB_Total_shims.Items.Add(MicroM3 + count0point5microm.ToString());

                }
            }
            catch { }
        }


        private bool ValidSecond()
        {
            bool checkvalue = true;
            if (String.IsNullOrEmpty(txt_dailgaugereading_page2.Text))
            {
                lbl_gauge_reading_page2.ForeColor = Color.Red;

                checkvalue = false;
            }
            else
            {
                lbl_gauge_reading_page2.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_manualReading_page2.Text))
            {
                lbl_manualReading_page2.ForeColor = Color.Red;

                checkvalue = false;
            }
            else
            {
                lbl_manualReading_page2.ForeColor = Color.Black;
            }

            if (String.IsNullOrEmpty(txt_shimsRequiredDrivenVerify_page2.Text))
            {
                lbl_shimRequiredDriven_verify_page2.ForeColor = Color.Red;

                checkvalue = false;
            }
            else
            {
                lbl_shimRequiredDriven_verify_page2.ForeColor = Color.Black;
            }

            return checkvalue;
        }



        private void bttn_next_page2_Click(object sender, EventArgs e)
        {
            if (ValidSecond())
            {
                this.PtoMeasurement.SelectedTab = tab_shimVerification;

                txt_verifieddailgaugereading_verification_page3.Select();
                txt_dailgaugereading_page3.Text = txt_dailgaugereading_page2.Text;
                txt_manualReading_page3.Text = txt_manualReading_page2.Text;
                txt_shimsRequiredDrivenVerify_page3.Text = txt_shimsRequiredDrivenVerify_page2.Text;
                
            }
        }


        public void TotalShimRequiredCalculations()
        {
            try
            {
                txt_masterValue_page2.Text = "79";

                if (txt_masterValue_page2.Text != "" && txt_dailgaugereading_page2.Text != "" && txt_manualReading_page2.Text != "")
                {
                    txt_shimsRequiredDrivenVerify_page2.Text = _objValidation.Gettingdecimals((Convert.ToString(Convert.ToDouble(txt_manualReading_page2.Text) + ((Convert.ToDouble(txt_masterValue_page2.Text) - (Convert.ToDouble(txt_dailgaugereading_page2.Text)))) - 148.000)).ToString());

                    int count0point1microm = 0;
                    int count0point2microm = 0;
                    int count0point4microm = 0;

                    shim = Math.Abs(Convert.ToDouble(txt_shimsRequiredDrivenVerify_page2.Text) * 1000);

                    while (shim >= 50)
                    {
                        if (shim >= 400)
                        {
                            shim = shim - 400;
                            count0point4microm++;
                            continue;
                        }
                        if (shim >= 200)
                        {
                            shim = shim - 200;
                            count0point2microm++;
                            continue;
                        }
                        if (shim >= 100 || shim >= 50)
                        {
                            shim = shim - 100;
                            count0point1microm++;
                            continue;
                        }
                    }

                    string MicroM1 = "0.1 mm :";
                    string MicroM2 = "0.2 mm :";
                    string MicroM3 = "0.4 mm :";

                    LB_Total_shims.Items.Clear();

                    if (count0point1microm != 0)
                        LB_Total_shims.Items.Add(MicroM1 + count0point1microm.ToString());
                    if (count0point2microm != 0)
                        LB_Total_shims.Items.Add(MicroM2 + count0point2microm.ToString());
                    if (count0point4microm != 0)
                        LB_Total_shims.Items.Add(MicroM3 + count0point4microm.ToString());

                }
            }
            catch
            {

            }
        }


        private void txt_dailgaugereading_page2_TextChanged(object sender, EventArgs e)
        {
            TotalShimRequiredCalculations();
        }

        private void txt_dailgaugereading_page2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txt_dailgaugereading_page2.Text == "")
                {
                    return;
                }
                else
                {
                    txt_manualReading_page2.Focus();
                }
            }
        }

        private void txt_dailgaugereading_page2_Validated(object sender, EventArgs e)
        {

        }

        private void txt_dailgaugereading_page2_KeyPress(object sender, KeyPressEventArgs e)
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

        private void txt_manualReading_page2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txt_manualReading_page2.Text == "")
                {
                    return;
                }
                else
                {
                    bttn_next_page2.Focus();
                }
            }
        }

        private void txt_manualReading_page2_TextChanged(object sender, EventArgs e)
        {
            TotalShimRequiredCalculations();
        }

        private void txt_manualReading_page2_KeyPress(object sender, KeyPressEventArgs e)
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

        private void bttn_save_Click(object sender, EventArgs e)
        {

        }


        private void cmb_selectmodel_KeyDown(object sender, KeyEventArgs e)

        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_dailgaugereading_page1.Focus();
            }
        }

    }
}
