namespace CCD_Measurment_of_SSSA
{
    partial class Frm_DatewiseReport
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgv_PTOMesurement = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.bttn_exportexcel = new System.Windows.Forms.Button();
            this.bttn_View = new System.Windows.Forms.Button();
            this.dtp_endDate = new System.Windows.Forms.DateTimePicker();
            this.dtp_StartDate = new System.Windows.Forms.DateTimePicker();
            this.modelid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.master_value = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dialguage_reading = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.actual_reading = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.manual_reading = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.shim_required = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.total_shims = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.shim_required_verification = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.operator_name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.createdate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_PTOMesurement)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgv_PTOMesurement
            // 
            this.dgv_PTOMesurement.AllowUserToAddRows = false;
            this.dgv_PTOMesurement.AllowUserToDeleteRows = false;
            this.dgv_PTOMesurement.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv_PTOMesurement.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_PTOMesurement.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.modelid,
            this.master_value,
            this.dialguage_reading,
            this.actual_reading,
            this.manual_reading,
            this.shim_required,
            this.total_shims,
            this.shim_required_verification,
            this.operator_name,
            this.createdate});
            this.dgv_PTOMesurement.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_PTOMesurement.Location = new System.Drawing.Point(0, 65);
            this.dgv_PTOMesurement.Name = "dgv_PTOMesurement";
            this.dgv_PTOMesurement.ReadOnly = true;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgv_PTOMesurement.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_PTOMesurement.Size = new System.Drawing.Size(1003, 513);
            this.dgv_PTOMesurement.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::CCD_Measurment_of_SSSA.Properties.Resources.mahindra_1;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.Controls.Add(this.bttn_exportexcel);
            this.panel1.Controls.Add(this.bttn_View);
            this.panel1.Controls.Add(this.dtp_endDate);
            this.panel1.Controls.Add(this.dtp_StartDate);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1003, 65);
            this.panel1.TabIndex = 2;
            // 
            // bttn_exportexcel
            // 
            this.bttn_exportexcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bttn_exportexcel.BackgroundImage = global::CCD_Measurment_of_SSSA.Properties.Resources.excel_icon;
            this.bttn_exportexcel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bttn_exportexcel.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.bttn_exportexcel.FlatAppearance.BorderSize = 0;
            this.bttn_exportexcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttn_exportexcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_exportexcel.ForeColor = System.Drawing.Color.White;
            this.bttn_exportexcel.Location = new System.Drawing.Point(511, 14);
            this.bttn_exportexcel.Name = "bttn_exportexcel";
            this.bttn_exportexcel.Size = new System.Drawing.Size(72, 46);
            this.bttn_exportexcel.TabIndex = 24;
            this.bttn_exportexcel.UseVisualStyleBackColor = true;
            this.bttn_exportexcel.Click += new System.EventHandler(this.bttn_exportexcel_Click);
            // 
            // bttn_View
            // 
            this.bttn_View.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.bttn_View.BackColor = System.Drawing.Color.Transparent;
            this.bttn_View.BackgroundImage = global::CCD_Measurment_of_SSSA.Properties.Resources._298px_HILLBLU_button_background1;
            this.bttn_View.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bttn_View.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.bttn_View.FlatAppearance.BorderSize = 0;
            this.bttn_View.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.bttn_View.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.bttn_View.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttn_View.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_View.ForeColor = System.Drawing.Color.White;
            this.bttn_View.Location = new System.Drawing.Point(424, 16);
            this.bttn_View.Name = "bttn_View";
            this.bttn_View.Size = new System.Drawing.Size(81, 39);
            this.bttn_View.TabIndex = 10;
            this.bttn_View.Text = "View";
            this.bttn_View.UseVisualStyleBackColor = false;
            this.bttn_View.Click += new System.EventHandler(this.bttn_View_Click);
            // 
            // dtp_endDate
            // 
            this.dtp_endDate.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtp_endDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtp_endDate.Location = new System.Drawing.Point(218, 20);
            this.dtp_endDate.Name = "dtp_endDate";
            this.dtp_endDate.Size = new System.Drawing.Size(200, 26);
            this.dtp_endDate.TabIndex = 1;
            // 
            // dtp_StartDate
            // 
            this.dtp_StartDate.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtp_StartDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtp_StartDate.Location = new System.Drawing.Point(12, 20);
            this.dtp_StartDate.Name = "dtp_StartDate";
            this.dtp_StartDate.Size = new System.Drawing.Size(200, 26);
            this.dtp_StartDate.TabIndex = 0;
            // 
            // modelid
            // 
            this.modelid.HeaderText = "Model Name";
            this.modelid.Name = "modelid";
            this.modelid.ReadOnly = true;
            // 
            // master_value
            // 
            this.master_value.HeaderText = "Master Value";
            this.master_value.Name = "master_value";
            this.master_value.ReadOnly = true;
            // 
            // dialguage_reading
            // 
            this.dialguage_reading.HeaderText = "Dial Duage Reading";
            this.dialguage_reading.Name = "dialguage_reading";
            this.dialguage_reading.ReadOnly = true;
            // 
            // actual_reading
            // 
            this.actual_reading.HeaderText = "Actual Reading";
            this.actual_reading.Name = "actual_reading";
            this.actual_reading.ReadOnly = true;
            // 
            // manual_reading
            // 
            this.manual_reading.HeaderText = "Manual Reading";
            this.manual_reading.Name = "manual_reading";
            this.manual_reading.ReadOnly = true;
            // 
            // shim_required
            // 
            this.shim_required.HeaderText = "Shim Required";
            this.shim_required.Name = "shim_required";
            this.shim_required.ReadOnly = true;
            // 
            // total_shims
            // 
            this.total_shims.HeaderText = "Total Shims";
            this.total_shims.Name = "total_shims";
            this.total_shims.ReadOnly = true;
            // 
            // shim_required_verification
            // 
            this.shim_required_verification.HeaderText = "Shim Required Verification";
            this.shim_required_verification.Name = "shim_required_verification";
            this.shim_required_verification.ReadOnly = true;
            // 
            // operator_name
            // 
            this.operator_name.HeaderText = "Operator Name";
            this.operator_name.Name = "operator_name";
            this.operator_name.ReadOnly = true;
            // 
            // createdate
            // 
            this.createdate.HeaderText = "Date";
            this.createdate.Name = "createdate";
            this.createdate.ReadOnly = true;
            // 
            // Frm_DatewiseReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1003, 578);
            this.Controls.Add(this.dgv_PTOMesurement);
            this.Controls.Add(this.panel1);
            this.Name = "Frm_DatewiseReport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Frm_DatewiseReport";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Frm_DatewiseReport_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_PTOMesurement)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv_PTOMesurement;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DateTimePicker dtp_endDate;
        private System.Windows.Forms.DateTimePicker dtp_StartDate;
        private System.Windows.Forms.Button bttn_View;
        private System.Windows.Forms.Button bttn_exportexcel;
        private System.Windows.Forms.DataGridViewTextBoxColumn modelid;
        private System.Windows.Forms.DataGridViewTextBoxColumn master_value;
        private System.Windows.Forms.DataGridViewTextBoxColumn dialguage_reading;
        private System.Windows.Forms.DataGridViewTextBoxColumn actual_reading;
        private System.Windows.Forms.DataGridViewTextBoxColumn manual_reading;
        private System.Windows.Forms.DataGridViewTextBoxColumn shim_required;
        private System.Windows.Forms.DataGridViewTextBoxColumn total_shims;
        private System.Windows.Forms.DataGridViewTextBoxColumn shim_required_verification;
        private System.Windows.Forms.DataGridViewTextBoxColumn operator_name;
        private System.Windows.Forms.DataGridViewTextBoxColumn createdate;
    }
}