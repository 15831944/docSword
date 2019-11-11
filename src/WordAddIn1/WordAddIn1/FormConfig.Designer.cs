namespace OfficeAssist
{
    partial class FormConfig
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblDbUrl = new System.Windows.Forms.Label();
            this.txtConfigDbUrl = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtConfigTempLoc = new System.Windows.Forms.TextBox();
            this.btnConfigSave = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnConfigSave);
            this.groupBox1.Controls.Add(this.txtConfigTempLoc);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtConfigDbUrl);
            this.groupBox1.Controls.Add(this.lblDbUrl);
            this.groupBox1.Location = new System.Drawing.Point(1, 1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(553, 366);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // lblDbUrl
            // 
            this.lblDbUrl.AutoSize = true;
            this.lblDbUrl.Location = new System.Drawing.Point(7, 25);
            this.lblDbUrl.Name = "lblDbUrl";
            this.lblDbUrl.Size = new System.Drawing.Size(127, 15);
            this.lblDbUrl.TabIndex = 0;
            this.lblDbUrl.Text = "数据库连接字符串";
            // 
            // txtConfigDbUrl
            // 
            this.txtConfigDbUrl.Location = new System.Drawing.Point(140, 22);
            this.txtConfigDbUrl.Name = "txtConfigDbUrl";
            this.txtConfigDbUrl.Size = new System.Drawing.Size(407, 25);
            this.txtConfigDbUrl.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 63);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "临时文件存放目录";
            // 
            // txtConfigTempLoc
            // 
            this.txtConfigTempLoc.Location = new System.Drawing.Point(140, 59);
            this.txtConfigTempLoc.Name = "txtConfigTempLoc";
            this.txtConfigTempLoc.Size = new System.Drawing.Size(407, 25);
            this.txtConfigTempLoc.TabIndex = 3;
            // 
            // btnConfigSave
            // 
            this.btnConfigSave.Location = new System.Drawing.Point(465, 337);
            this.btnConfigSave.Name = "btnConfigSave";
            this.btnConfigSave.Size = new System.Drawing.Size(75, 23);
            this.btnConfigSave.TabIndex = 4;
            this.btnConfigSave.Text = "保存";
            this.btnConfigSave.UseVisualStyleBackColor = true;
            this.btnConfigSave.Click += new System.EventHandler(this.btnConfigSave_Click);
            // 
            // FormConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(553, 365);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormConfig";
            this.Text = "配置";
            this.Load += new System.EventHandler(this.FormConfig_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtConfigTempLoc;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtConfigDbUrl;
        private System.Windows.Forms.Label lblDbUrl;
        private System.Windows.Forms.Button btnConfigSave;
    }
}