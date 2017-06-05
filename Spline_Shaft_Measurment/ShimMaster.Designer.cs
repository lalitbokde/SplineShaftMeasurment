namespace CCD_Measurment_of_SSSA
{
    partial class ShimMaster
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
            this.lbl_shimMaster = new System.Windows.Forms.Label();
            this.lbl_date = new System.Windows.Forms.Label();
            this.txt_shim = new System.Windows.Forms.TextBox();
            this.dtp_date = new System.Windows.Forms.DateTimePicker();
            this.bttn_save = new System.Windows.Forms.Button();
            this.bttn_update = new System.Windows.Forms.Button();
            this.bttn_delete = new System.Windows.Forms.Button();
            this.dgv_ShimMaster = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SrNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Shim = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bttn_close = new System.Windows.Forms.Button();
            this.bttn_reset = new System.Windows.Forms.Button();
            this.lbl_id = new System.Windows.Forms.Label();
            this.txt_id = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_ShimMaster)).BeginInit();
            this.SuspendLayout();
            // 
            // lbl_shimMaster
            // 
            this.lbl_shimMaster.AutoSize = true;
            this.lbl_shimMaster.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_shimMaster.Location = new System.Drawing.Point(100, 51);
            this.lbl_shimMaster.Name = "lbl_shimMaster";
            this.lbl_shimMaster.Size = new System.Drawing.Size(40, 15);
            this.lbl_shimMaster.TabIndex = 0;
            this.lbl_shimMaster.Text = "Shim";
            // 
            // lbl_date
            // 
            this.lbl_date.AutoSize = true;
            this.lbl_date.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_date.Location = new System.Drawing.Point(103, 81);
            this.lbl_date.Name = "lbl_date";
            this.lbl_date.Size = new System.Drawing.Size(37, 15);
            this.lbl_date.TabIndex = 1;
            this.lbl_date.Text = "Date";
            // 
            // txt_shim
            // 
            this.txt_shim.Location = new System.Drawing.Point(153, 48);
            this.txt_shim.Name = "txt_shim";
            this.txt_shim.Size = new System.Drawing.Size(193, 20);
            this.txt_shim.TabIndex = 2;
            // 
            // dtp_date
            // 
            this.dtp_date.Location = new System.Drawing.Point(153, 81);
            this.dtp_date.Name = "dtp_date";
            this.dtp_date.Size = new System.Drawing.Size(193, 20);
            this.dtp_date.TabIndex = 3;
            // 
            // bttn_save
            // 
            this.bttn_save.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_save.Location = new System.Drawing.Point(21, 130);
            this.bttn_save.Name = "bttn_save";
            this.bttn_save.Size = new System.Drawing.Size(75, 23);
            this.bttn_save.TabIndex = 4;
            this.bttn_save.Text = "Save";
            this.bttn_save.UseVisualStyleBackColor = true;
            this.bttn_save.Click += new System.EventHandler(this.bttn_save_Click);
            // 
            // bttn_update
            // 
            this.bttn_update.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_update.Location = new System.Drawing.Point(107, 130);
            this.bttn_update.Name = "bttn_update";
            this.bttn_update.Size = new System.Drawing.Size(75, 23);
            this.bttn_update.TabIndex = 5;
            this.bttn_update.Text = "Update";
            this.bttn_update.UseVisualStyleBackColor = true;
            this.bttn_update.Click += new System.EventHandler(this.bttn_update_Click);
            // 
            // bttn_delete
            // 
            this.bttn_delete.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_delete.Location = new System.Drawing.Point(197, 130);
            this.bttn_delete.Name = "bttn_delete";
            this.bttn_delete.Size = new System.Drawing.Size(75, 23);
            this.bttn_delete.TabIndex = 6;
            this.bttn_delete.Text = "Detete";
            this.bttn_delete.UseVisualStyleBackColor = true;
            this.bttn_delete.Click += new System.EventHandler(this.bttn_delete_Click);
            // 
            // dgv_ShimMaster
            // 
            this.dgv_ShimMaster.AllowUserToAddRows = false;
            this.dgv_ShimMaster.AllowUserToDeleteRows = false;
            this.dgv_ShimMaster.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv_ShimMaster.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_ShimMaster.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.SrNo,
            this.Shim});
            this.dgv_ShimMaster.Location = new System.Drawing.Point(12, 172);
            this.dgv_ShimMaster.Name = "dgv_ShimMaster";
            this.dgv_ShimMaster.ReadOnly = true;
            this.dgv_ShimMaster.Size = new System.Drawing.Size(453, 170);
            this.dgv_ShimMaster.TabIndex = 7;
            this.dgv_ShimMaster.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_ShimMaster_CellClick);
            // 
            // ID
            // 
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Visible = false;
            // 
            // SrNo
            // 
            this.SrNo.HeaderText = "SrNo";
            this.SrNo.Name = "SrNo";
            this.SrNo.ReadOnly = true;
            // 
            // Shim
            // 
            this.Shim.HeaderText = "Shim";
            this.Shim.Name = "Shim";
            this.Shim.ReadOnly = true;
            // 
            // bttn_close
            // 
            this.bttn_close.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_close.Location = new System.Drawing.Point(283, 130);
            this.bttn_close.Name = "bttn_close";
            this.bttn_close.Size = new System.Drawing.Size(75, 23);
            this.bttn_close.TabIndex = 8;
            this.bttn_close.Text = "Close";
            this.bttn_close.UseVisualStyleBackColor = true;
            this.bttn_close.Click += new System.EventHandler(this.bttn_close_Click);
            // 
            // bttn_reset
            // 
            this.bttn_reset.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_reset.Location = new System.Drawing.Point(373, 130);
            this.bttn_reset.Name = "bttn_reset";
            this.bttn_reset.Size = new System.Drawing.Size(75, 23);
            this.bttn_reset.TabIndex = 9;
            this.bttn_reset.Text = "Reset";
            this.bttn_reset.UseVisualStyleBackColor = true;
            this.bttn_reset.Click += new System.EventHandler(this.bttn_reset_Click);
            // 
            // lbl_id
            // 
            this.lbl_id.AutoSize = true;
            this.lbl_id.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_id.Location = new System.Drawing.Point(119, 19);
            this.lbl_id.Name = "lbl_id";
            this.lbl_id.Size = new System.Drawing.Size(21, 15);
            this.lbl_id.TabIndex = 21;
            this.lbl_id.Text = "ID";
            this.lbl_id.Visible = false;
            // 
            // txt_id
            // 
            this.txt_id.Location = new System.Drawing.Point(153, 18);
            this.txt_id.Name = "txt_id";
            this.txt_id.Size = new System.Drawing.Size(193, 20);
            this.txt_id.TabIndex = 20;
            this.txt_id.Visible = false;
            // 
            // ShimMaster
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(477, 354);
            this.Controls.Add(this.lbl_id);
            this.Controls.Add(this.txt_id);
            this.Controls.Add(this.bttn_reset);
            this.Controls.Add(this.bttn_close);
            this.Controls.Add(this.dgv_ShimMaster);
            this.Controls.Add(this.bttn_delete);
            this.Controls.Add(this.bttn_update);
            this.Controls.Add(this.bttn_save);
            this.Controls.Add(this.dtp_date);
            this.Controls.Add(this.txt_shim);
            this.Controls.Add(this.lbl_date);
            this.Controls.Add(this.lbl_shimMaster);
            this.Name = "ShimMaster";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ShimMaster";
            this.Load += new System.EventHandler(this.ShimMaster_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_ShimMaster)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_shimMaster;
        private System.Windows.Forms.Label lbl_date;
        private System.Windows.Forms.TextBox txt_shim;
        private System.Windows.Forms.DateTimePicker dtp_date;
        private System.Windows.Forms.Button bttn_save;
        private System.Windows.Forms.Button bttn_update;
        private System.Windows.Forms.Button bttn_delete;
        private System.Windows.Forms.DataGridView dgv_ShimMaster;
        private System.Windows.Forms.Button bttn_close;
        private System.Windows.Forms.Button bttn_reset;
        private System.Windows.Forms.Label lbl_id;
        private System.Windows.Forms.TextBox txt_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn SrNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn Shim;
    }
}