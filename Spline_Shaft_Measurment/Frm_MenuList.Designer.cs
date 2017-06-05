namespace CCD_Measurment_of_SSSA
{
    partial class Frm_MenuList
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Frm_MenuList));
            this.panel1 = new System.Windows.Forms.Panel();
            this.bttn_report = new System.Windows.Forms.Button();
            this.bttn_CreateUser = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::CCD_Measurment_of_SSSA.Properties.Resources.mahindra_1;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(407, 65);
            this.panel1.TabIndex = 10;
            // 
            // bttn_report
            // 
            this.bttn_report.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.bttn_report.BackColor = System.Drawing.Color.Transparent;
            this.bttn_report.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("bttn_report.BackgroundImage")));
            this.bttn_report.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bttn_report.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.bttn_report.FlatAppearance.BorderSize = 0;
            this.bttn_report.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.bttn_report.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.bttn_report.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttn_report.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_report.ForeColor = System.Drawing.Color.Black;
            this.bttn_report.Location = new System.Drawing.Point(12, 161);
            this.bttn_report.Name = "bttn_report";
            this.bttn_report.Size = new System.Drawing.Size(384, 62);
            this.bttn_report.TabIndex = 9;
            this.bttn_report.Text = "Report";
            this.bttn_report.UseVisualStyleBackColor = false;
            this.bttn_report.Click += new System.EventHandler(this.bttn_report_Click_1);
            // 
            // bttn_CreateUser
            // 
            this.bttn_CreateUser.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.bttn_CreateUser.BackColor = System.Drawing.Color.Transparent;
            this.bttn_CreateUser.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("bttn_CreateUser.BackgroundImage")));
            this.bttn_CreateUser.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bttn_CreateUser.FlatAppearance.BorderColor = System.Drawing.Color.Gray;
            this.bttn_CreateUser.FlatAppearance.BorderSize = 0;
            this.bttn_CreateUser.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.bttn_CreateUser.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            this.bttn_CreateUser.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bttn_CreateUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bttn_CreateUser.ForeColor = System.Drawing.Color.Black;
            this.bttn_CreateUser.Location = new System.Drawing.Point(12, 97);
            this.bttn_CreateUser.Name = "bttn_CreateUser";
            this.bttn_CreateUser.Size = new System.Drawing.Size(384, 58);
            this.bttn_CreateUser.TabIndex = 8;
            this.bttn_CreateUser.Text = "Create User";
            this.bttn_CreateUser.UseVisualStyleBackColor = false;
            this.bttn_CreateUser.Click += new System.EventHandler(this.bttn_CreateUser_Click);
            // 
            // Frm_MenuList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(407, 292);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.bttn_report);
            this.Controls.Add(this.bttn_CreateUser);
            this.Name = "Frm_MenuList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Menu";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button bttn_CreateUser;
        private System.Windows.Forms.Button bttn_report;
        private System.Windows.Forms.Panel panel1;
    }
}