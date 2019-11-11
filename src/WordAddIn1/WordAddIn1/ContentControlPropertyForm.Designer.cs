namespace OfficeAssist
{
    partial class ContentControlPropertyForm
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
            this.btnCntCtrlOK = new System.Windows.Forms.Button();
            this.btnCntCtrlCancel = new System.Windows.Forms.Button();
            this.txtBoxCntCtrlTag = new System.Windows.Forms.TextBox();
            this.txtBoxCntCtrlName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnCntCtrlOK
            // 
            this.btnCntCtrlOK.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.btnCntCtrlOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnCntCtrlOK.Location = new System.Drawing.Point(55, 94);
            this.btnCntCtrlOK.Name = "btnCntCtrlOK";
            this.btnCntCtrlOK.Size = new System.Drawing.Size(75, 23);
            this.btnCntCtrlOK.TabIndex = 0;
            this.btnCntCtrlOK.Text = "确定";
            this.btnCntCtrlOK.UseVisualStyleBackColor = true;
            // 
            // btnCntCtrlCancel
            // 
            this.btnCntCtrlCancel.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.btnCntCtrlCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCntCtrlCancel.Location = new System.Drawing.Point(187, 94);
            this.btnCntCtrlCancel.Name = "btnCntCtrlCancel";
            this.btnCntCtrlCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCntCtrlCancel.TabIndex = 1;
            this.btnCntCtrlCancel.Text = "取消";
            this.btnCntCtrlCancel.UseVisualStyleBackColor = true;
            // 
            // txtBoxCntCtrlTag
            // 
            this.txtBoxCntCtrlTag.Location = new System.Drawing.Point(55, 47);
            this.txtBoxCntCtrlTag.Name = "txtBoxCntCtrlTag";
            this.txtBoxCntCtrlTag.Size = new System.Drawing.Size(235, 25);
            this.txtBoxCntCtrlTag.TabIndex = 3;
            // 
            // txtBoxCntCtrlName
            // 
            this.txtBoxCntCtrlName.Location = new System.Drawing.Point(55, 12);
            this.txtBoxCntCtrlName.Name = "txtBoxCntCtrlName";
            this.txtBoxCntCtrlName.Size = new System.Drawing.Size(235, 25);
            this.txtBoxCntCtrlName.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "标记";
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "名称";
            // 
            // ContentControlPropertyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(302, 129);
            this.Controls.Add(this.btnCntCtrlCancel);
            this.Controls.Add(this.btnCntCtrlOK);
            this.Controls.Add(this.txtBoxCntCtrlTag);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtBoxCntCtrlName);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ContentControlPropertyForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "属性";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCntCtrlOK;
        private System.Windows.Forms.Button btnCntCtrlCancel;
        public System.Windows.Forms.TextBox txtBoxCntCtrlTag;
        public System.Windows.Forms.TextBox txtBoxCntCtrlName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}