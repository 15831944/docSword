namespace WordAddIn1
{
    partial class relManageForm
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
            System.Windows.Forms.TreeNode treeNode27 = new System.Windows.Forms.TreeNode("引用1");
            System.Windows.Forms.TreeNode treeNode28 = new System.Windows.Forms.TreeNode("引用2");
            System.Windows.Forms.TreeNode treeNode29 = new System.Windows.Forms.TreeNode("定义1", new System.Windows.Forms.TreeNode[] {
            treeNode27,
            treeNode28});
            System.Windows.Forms.TreeNode treeNode30 = new System.Windows.Forms.TreeNode("引用1");
            System.Windows.Forms.TreeNode treeNode31 = new System.Windows.Forms.TreeNode("定义2", new System.Windows.Forms.TreeNode[] {
            treeNode30});
            System.Windows.Forms.TreeNode treeNode32 = new System.Windows.Forms.TreeNode("引用1");
            System.Windows.Forms.TreeNode treeNode33 = new System.Windows.Forms.TreeNode("引用2");
            System.Windows.Forms.TreeNode treeNode34 = new System.Windows.Forms.TreeNode("定义3", new System.Windows.Forms.TreeNode[] {
            treeNode32,
            treeNode33});
            System.Windows.Forms.TreeNode treeNode35 = new System.Windows.Forms.TreeNode("关联引用", new System.Windows.Forms.TreeNode[] {
            treeNode29,
            treeNode31,
            treeNode34});
            System.Windows.Forms.TreeNode treeNode36 = new System.Windows.Forms.TreeNode("运算1");
            System.Windows.Forms.TreeNode treeNode37 = new System.Windows.Forms.TreeNode("运算2");
            System.Windows.Forms.TreeNode treeNode38 = new System.Windows.Forms.TreeNode("运算3");
            System.Windows.Forms.TreeNode treeNode39 = new System.Windows.Forms.TreeNode("运算", new System.Windows.Forms.TreeNode[] {
            treeNode36,
            treeNode37,
            treeNode38});
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Location = new System.Drawing.Point(28, 75);
            this.treeView1.Name = "treeView1";
            treeNode27.Name = "节点14";
            treeNode27.Text = "引用1";
            treeNode28.Name = "节点15";
            treeNode28.Text = "引用2";
            treeNode29.Name = "节点8";
            treeNode29.Text = "定义1";
            treeNode30.Name = "节点16";
            treeNode30.Text = "引用1";
            treeNode31.Name = "节点9";
            treeNode31.Text = "定义2";
            treeNode32.Name = "节点17";
            treeNode32.Text = "引用1";
            treeNode33.Name = "节点18";
            treeNode33.Text = "引用2";
            treeNode34.Name = "节点10";
            treeNode34.Text = "定义3";
            treeNode35.Name = "节点6";
            treeNode35.Text = "关联引用";
            treeNode36.Name = "节点11";
            treeNode36.Text = "运算1";
            treeNode37.Name = "节点12";
            treeNode37.Text = "运算2";
            treeNode38.Name = "节点13";
            treeNode38.Text = "运算3";
            treeNode39.Name = "节点7";
            treeNode39.Text = "运算";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode35,
            treeNode39});
            this.treeView1.Size = new System.Drawing.Size(372, 286);
            this.treeView1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(211, 46);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(82, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "查找";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.treeView1);
            this.groupBox1.Location = new System.Drawing.Point(0, 1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(446, 402);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(325, 379);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(75, 23);
            this.button4.TabIndex = 6;
            this.button4.Text = "更新";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(173, 379);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 5;
            this.button3.Text = "删除";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(28, 379);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 4;
            this.button2.Text = "新建";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(28, 44);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(177, 25);
            this.textBox1.TabIndex = 2;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(316, 46);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(84, 23);
            this.button5.TabIndex = 7;
            this.button5.Text = "重置";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // relManageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(454, 457);
            this.Controls.Add(this.groupBox1);
            this.Name = "relManageForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button5;
    }
}