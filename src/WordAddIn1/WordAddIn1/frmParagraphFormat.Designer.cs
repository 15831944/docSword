namespace OfficeAssist
{
    partial class frmParagraphFormat
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
            this.label1 = new System.Windows.Forms.Label();
            this.cmbAlignStyle = new System.Windows.Forms.ComboBox();
            this.numIndentSpecial = new System.Windows.Forms.NumericUpDown();
            this.cmbIndentSpecial = new System.Windows.Forms.ComboBox();
            this.cmbIndentSpecialUnit = new System.Windows.Forms.ComboBox();
            this.cmbIndentRightUnit = new System.Windows.Forms.ComboBox();
            this.numIndentRight = new System.Windows.Forms.NumericUpDown();
            this.cmbIndentLeftUnit = new System.Windows.Forms.ComboBox();
            this.numIndentLeft = new System.Windows.Forms.NumericUpDown();
            this.chkSymIndent = new System.Windows.Forms.CheckBox();
            this.chkAutoAlignRight = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageIndentLineSpace = new System.Windows.Forms.TabPage();
            this.chkParaLineSpace = new System.Windows.Forms.CheckBox();
            this.chkParaLineSpaceAfter = new System.Windows.Forms.CheckBox();
            this.chkSpaceAfterAuto = new System.Windows.Forms.CheckBox();
            this.chkSpaceBeforeAuto = new System.Windows.Forms.CheckBox();
            this.chkParaLineSpaceBefore = new System.Windows.Forms.CheckBox();
            this.chkIndentSpecial = new System.Windows.Forms.CheckBox();
            this.chkIndentRight = new System.Windows.Forms.CheckBox();
            this.chkIndentLeft = new System.Windows.Forms.CheckBox();
            this.chkAlignStyle = new System.Windows.Forms.CheckBox();
            this.chkAlignMesh = new System.Windows.Forms.CheckBox();
            this.chkNoBlank = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.numBeforeParaSpacing = new System.Windows.Forms.NumericUpDown();
            this.numAfterParaSpacing = new System.Windows.Forms.NumericUpDown();
            this.numLineSpacing = new System.Windows.Forms.NumericUpDown();
            this.cmbBeforeParaSpacingUnit = new System.Windows.Forms.ComboBox();
            this.cmbLineSpacingRule = new System.Windows.Forms.ComboBox();
            this.cmbAfterParaSpacingUnit = new System.Windows.Forms.ComboBox();
            this.cmbLineSpacingUnit = new System.Windows.Forms.ComboBox();
            this.tabPageLineBreakPageBreak = new System.Windows.Forms.TabPage();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.chkCancelBreakWords = new System.Windows.Forms.CheckBox();
            this.chkCancelLineNum = new System.Windows.Forms.CheckBox();
            this.chkBreakPageBeforePara = new System.Windows.Forms.CheckBox();
            this.chkParaNoBreakPage = new System.Windows.Forms.CheckBox();
            this.chkKeepNext = new System.Windows.Forms.CheckBox();
            this.chkAloneParaCtrl = new System.Windows.Forms.CheckBox();
            this.tabPageChineseStyle = new System.Windows.Forms.TabPage();
            this.chkTextAlignStyle = new System.Windows.Forms.CheckBox();
            this.cmbTextAlignStyle = new System.Windows.Forms.ComboBox();
            this.label22 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.chkAutoAdjustNumLineSpacing = new System.Windows.Forms.CheckBox();
            this.chkAutoAdjustLineSpacing = new System.Windows.Forms.CheckBox();
            this.chkAllowCompressCma = new System.Windows.Forms.CheckBox();
            this.chkAllowCmaOverLimit = new System.Windows.Forms.CheckBox();
            this.chkAllowAsciiBreakLineInPara = new System.Windows.Forms.CheckBox();
            this.chkCtrlFromChinese = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numIndentSpecial)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numIndentRight)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numIndentLeft)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPageIndentLineSpace.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numBeforeParaSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numAfterParaSpacing)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLineSpacing)).BeginInit();
            this.tabPageLineBreakPageBreak.SuspendLayout();
            this.tabPageChineseStyle.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "缩进";
            // 
            // cmbAlignStyle
            // 
            this.cmbAlignStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAlignStyle.Enabled = false;
            this.cmbAlignStyle.FormattingEnabled = true;
            this.cmbAlignStyle.Location = new System.Drawing.Point(122, 37);
            this.cmbAlignStyle.Name = "cmbAlignStyle";
            this.cmbAlignStyle.Size = new System.Drawing.Size(64, 20);
            this.cmbAlignStyle.TabIndex = 1;
            this.cmbAlignStyle.SelectedIndexChanged += new System.EventHandler(this.cmbAlignStyle_SelectedIndexChanged);
            // 
            // numIndentSpecial
            // 
            this.numIndentSpecial.Enabled = false;
            this.numIndentSpecial.Location = new System.Drawing.Point(262, 150);
            this.numIndentSpecial.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numIndentSpecial.Name = "numIndentSpecial";
            this.numIndentSpecial.Size = new System.Drawing.Size(64, 21);
            this.numIndentSpecial.TabIndex = 2;
            this.numIndentSpecial.ValueChanged += new System.EventHandler(this.numIndentSpecial_ValueChanged);
            // 
            // cmbIndentSpecial
            // 
            this.cmbIndentSpecial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbIndentSpecial.Enabled = false;
            this.cmbIndentSpecial.FormattingEnabled = true;
            this.cmbIndentSpecial.Location = new System.Drawing.Point(122, 151);
            this.cmbIndentSpecial.Name = "cmbIndentSpecial";
            this.cmbIndentSpecial.Size = new System.Drawing.Size(134, 20);
            this.cmbIndentSpecial.TabIndex = 1;
            this.cmbIndentSpecial.SelectedIndexChanged += new System.EventHandler(this.cmbIndentSpecial_SelectedIndexChanged);
            // 
            // cmbIndentSpecialUnit
            // 
            this.cmbIndentSpecialUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbIndentSpecialUnit.Enabled = false;
            this.cmbIndentSpecialUnit.FormattingEnabled = true;
            this.cmbIndentSpecialUnit.Location = new System.Drawing.Point(332, 151);
            this.cmbIndentSpecialUnit.Name = "cmbIndentSpecialUnit";
            this.cmbIndentSpecialUnit.Size = new System.Drawing.Size(57, 20);
            this.cmbIndentSpecialUnit.TabIndex = 1;
            this.cmbIndentSpecialUnit.SelectedIndexChanged += new System.EventHandler(this.cmbIndentSpecialUnit_SelectedIndexChanged);
            // 
            // cmbIndentRightUnit
            // 
            this.cmbIndentRightUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbIndentRightUnit.Enabled = false;
            this.cmbIndentRightUnit.FormattingEnabled = true;
            this.cmbIndentRightUnit.Location = new System.Drawing.Point(192, 120);
            this.cmbIndentRightUnit.Name = "cmbIndentRightUnit";
            this.cmbIndentRightUnit.Size = new System.Drawing.Size(64, 20);
            this.cmbIndentRightUnit.TabIndex = 1;
            this.cmbIndentRightUnit.SelectedIndexChanged += new System.EventHandler(this.cmbIndentRightUnit_SelectedIndexChanged);
            // 
            // numIndentRight
            // 
            this.numIndentRight.Enabled = false;
            this.numIndentRight.Location = new System.Drawing.Point(122, 119);
            this.numIndentRight.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numIndentRight.Name = "numIndentRight";
            this.numIndentRight.Size = new System.Drawing.Size(64, 21);
            this.numIndentRight.TabIndex = 2;
            this.numIndentRight.ValueChanged += new System.EventHandler(this.numIndentRight_ValueChanged);
            // 
            // cmbIndentLeftUnit
            // 
            this.cmbIndentLeftUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbIndentLeftUnit.Enabled = false;
            this.cmbIndentLeftUnit.FormattingEnabled = true;
            this.cmbIndentLeftUnit.Location = new System.Drawing.Point(192, 91);
            this.cmbIndentLeftUnit.Name = "cmbIndentLeftUnit";
            this.cmbIndentLeftUnit.Size = new System.Drawing.Size(64, 20);
            this.cmbIndentLeftUnit.TabIndex = 1;
            this.cmbIndentLeftUnit.SelectedIndexChanged += new System.EventHandler(this.cmbIndentLeftUnit_SelectedIndexChanged);
            // 
            // numIndentLeft
            // 
            this.numIndentLeft.Enabled = false;
            this.numIndentLeft.Location = new System.Drawing.Point(122, 90);
            this.numIndentLeft.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numIndentLeft.Name = "numIndentLeft";
            this.numIndentLeft.Size = new System.Drawing.Size(64, 21);
            this.numIndentLeft.TabIndex = 2;
            this.numIndentLeft.ValueChanged += new System.EventHandler(this.numIndentLeft_ValueChanged);
            // 
            // chkSymIndent
            // 
            this.chkSymIndent.AutoSize = true;
            this.chkSymIndent.Location = new System.Drawing.Point(44, 181);
            this.chkSymIndent.Name = "chkSymIndent";
            this.chkSymIndent.Size = new System.Drawing.Size(72, 16);
            this.chkSymIndent.TabIndex = 3;
            this.chkSymIndent.Text = "对称缩进";
            this.chkSymIndent.ThreeState = true;
            this.chkSymIndent.UseVisualStyleBackColor = true;
            this.chkSymIndent.CheckedChanged += new System.EventHandler(this.chkSymIndent_CheckedChanged);
            // 
            // chkAutoAlignRight
            // 
            this.chkAutoAlignRight.AutoSize = true;
            this.chkAutoAlignRight.Location = new System.Drawing.Point(44, 205);
            this.chkAutoAlignRight.Name = "chkAutoAlignRight";
            this.chkAutoAlignRight.Size = new System.Drawing.Size(240, 16);
            this.chkAutoAlignRight.TabIndex = 3;
            this.chkAutoAlignRight.Text = "如果定义了文档网络，则自动调整右缩进";
            this.chkAutoAlignRight.ThreeState = true;
            this.chkAutoAlignRight.UseVisualStyleBackColor = true;
            this.chkAutoAlignRight.CheckedChanged += new System.EventHandler(this.chkAutoAlignRight_CheckedChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageIndentLineSpace);
            this.tabControl1.Controls.Add(this.tabPageLineBreakPageBreak);
            this.tabControl1.Controls.Add(this.tabPageChineseStyle);
            this.tabControl1.Location = new System.Drawing.Point(0, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(434, 441);
            this.tabControl1.TabIndex = 4;
            // 
            // tabPageIndentLineSpace
            // 
            this.tabPageIndentLineSpace.Controls.Add(this.chkParaLineSpace);
            this.tabPageIndentLineSpace.Controls.Add(this.chkParaLineSpaceAfter);
            this.tabPageIndentLineSpace.Controls.Add(this.chkSpaceAfterAuto);
            this.tabPageIndentLineSpace.Controls.Add(this.chkSpaceBeforeAuto);
            this.tabPageIndentLineSpace.Controls.Add(this.chkParaLineSpaceBefore);
            this.tabPageIndentLineSpace.Controls.Add(this.chkIndentSpecial);
            this.tabPageIndentLineSpace.Controls.Add(this.chkIndentRight);
            this.tabPageIndentLineSpace.Controls.Add(this.chkIndentLeft);
            this.tabPageIndentLineSpace.Controls.Add(this.chkAlignStyle);
            this.tabPageIndentLineSpace.Controls.Add(this.chkAlignMesh);
            this.tabPageIndentLineSpace.Controls.Add(this.chkNoBlank);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbIndentSpecial);
            this.tabPageIndentLineSpace.Controls.Add(this.chkAutoAlignRight);
            this.tabPageIndentLineSpace.Controls.Add(this.label12);
            this.tabPageIndentLineSpace.Controls.Add(this.label5);
            this.tabPageIndentLineSpace.Controls.Add(this.label10);
            this.tabPageIndentLineSpace.Controls.Add(this.label11);
            this.tabPageIndentLineSpace.Controls.Add(this.label9);
            this.tabPageIndentLineSpace.Controls.Add(this.label1);
            this.tabPageIndentLineSpace.Controls.Add(this.chkSymIndent);
            this.tabPageIndentLineSpace.Controls.Add(this.numBeforeParaSpacing);
            this.tabPageIndentLineSpace.Controls.Add(this.numIndentLeft);
            this.tabPageIndentLineSpace.Controls.Add(this.numAfterParaSpacing);
            this.tabPageIndentLineSpace.Controls.Add(this.numIndentRight);
            this.tabPageIndentLineSpace.Controls.Add(this.numLineSpacing);
            this.tabPageIndentLineSpace.Controls.Add(this.numIndentSpecial);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbBeforeParaSpacingUnit);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbIndentLeftUnit);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbLineSpacingRule);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbAfterParaSpacingUnit);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbIndentRightUnit);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbAlignStyle);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbLineSpacingUnit);
            this.tabPageIndentLineSpace.Controls.Add(this.cmbIndentSpecialUnit);
            this.tabPageIndentLineSpace.Location = new System.Drawing.Point(4, 22);
            this.tabPageIndentLineSpace.Name = "tabPageIndentLineSpace";
            this.tabPageIndentLineSpace.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageIndentLineSpace.Size = new System.Drawing.Size(426, 415);
            this.tabPageIndentLineSpace.TabIndex = 0;
            this.tabPageIndentLineSpace.Text = "缩进与间距";
            this.tabPageIndentLineSpace.UseVisualStyleBackColor = true;
            // 
            // chkParaLineSpace
            // 
            this.chkParaLineSpace.AutoSize = true;
            this.chkParaLineSpace.Location = new System.Drawing.Point(44, 313);
            this.chkParaLineSpace.Name = "chkParaLineSpace";
            this.chkParaLineSpace.Size = new System.Drawing.Size(48, 16);
            this.chkParaLineSpace.TabIndex = 9;
            this.chkParaLineSpace.Text = "行距";
            this.chkParaLineSpace.UseVisualStyleBackColor = true;
            this.chkParaLineSpace.CheckedChanged += new System.EventHandler(this.chkParaLineSpace_CheckedChanged);
            // 
            // chkParaLineSpaceAfter
            // 
            this.chkParaLineSpaceAfter.AutoSize = true;
            this.chkParaLineSpaceAfter.Location = new System.Drawing.Point(44, 285);
            this.chkParaLineSpaceAfter.Name = "chkParaLineSpaceAfter";
            this.chkParaLineSpaceAfter.Size = new System.Drawing.Size(48, 16);
            this.chkParaLineSpaceAfter.TabIndex = 9;
            this.chkParaLineSpaceAfter.Text = "段后";
            this.chkParaLineSpaceAfter.UseVisualStyleBackColor = true;
            this.chkParaLineSpaceAfter.CheckedChanged += new System.EventHandler(this.chkParaLineSpaceAfter_CheckedChanged);
            // 
            // chkSpaceAfterAuto
            // 
            this.chkSpaceAfterAuto.AutoSize = true;
            this.chkSpaceAfterAuto.Enabled = false;
            this.chkSpaceAfterAuto.Location = new System.Drawing.Point(262, 287);
            this.chkSpaceAfterAuto.Name = "chkSpaceAfterAuto";
            this.chkSpaceAfterAuto.Size = new System.Drawing.Size(48, 16);
            this.chkSpaceAfterAuto.TabIndex = 9;
            this.chkSpaceAfterAuto.Text = "自动";
            this.chkSpaceAfterAuto.UseVisualStyleBackColor = true;
            this.chkSpaceAfterAuto.CheckedChanged += new System.EventHandler(this.chkSpaceAfterAuto_CheckedChanged);
            // 
            // chkSpaceBeforeAuto
            // 
            this.chkSpaceBeforeAuto.AutoSize = true;
            this.chkSpaceBeforeAuto.Enabled = false;
            this.chkSpaceBeforeAuto.Location = new System.Drawing.Point(262, 257);
            this.chkSpaceBeforeAuto.Name = "chkSpaceBeforeAuto";
            this.chkSpaceBeforeAuto.Size = new System.Drawing.Size(48, 16);
            this.chkSpaceBeforeAuto.TabIndex = 9;
            this.chkSpaceBeforeAuto.Text = "自动";
            this.chkSpaceBeforeAuto.UseVisualStyleBackColor = true;
            this.chkSpaceBeforeAuto.CheckedChanged += new System.EventHandler(this.chkSpaceBeforeAuto_CheckedChanged);
            // 
            // chkParaLineSpaceBefore
            // 
            this.chkParaLineSpaceBefore.AutoSize = true;
            this.chkParaLineSpaceBefore.Location = new System.Drawing.Point(44, 257);
            this.chkParaLineSpaceBefore.Name = "chkParaLineSpaceBefore";
            this.chkParaLineSpaceBefore.Size = new System.Drawing.Size(48, 16);
            this.chkParaLineSpaceBefore.TabIndex = 9;
            this.chkParaLineSpaceBefore.Text = "段前";
            this.chkParaLineSpaceBefore.UseVisualStyleBackColor = true;
            this.chkParaLineSpaceBefore.CheckedChanged += new System.EventHandler(this.chkParaLineSpaceBefore_CheckedChanged);
            // 
            // chkIndentSpecial
            // 
            this.chkIndentSpecial.AutoSize = true;
            this.chkIndentSpecial.Location = new System.Drawing.Point(44, 153);
            this.chkIndentSpecial.Name = "chkIndentSpecial";
            this.chkIndentSpecial.Size = new System.Drawing.Size(72, 16);
            this.chkIndentSpecial.TabIndex = 8;
            this.chkIndentSpecial.Text = "特殊格式";
            this.chkIndentSpecial.UseVisualStyleBackColor = true;
            this.chkIndentSpecial.CheckedChanged += new System.EventHandler(this.chkIndentSpecial_CheckedChanged);
            // 
            // chkIndentRight
            // 
            this.chkIndentRight.AutoSize = true;
            this.chkIndentRight.Location = new System.Drawing.Point(44, 122);
            this.chkIndentRight.Name = "chkIndentRight";
            this.chkIndentRight.Size = new System.Drawing.Size(48, 16);
            this.chkIndentRight.TabIndex = 7;
            this.chkIndentRight.Text = "右侧";
            this.chkIndentRight.UseVisualStyleBackColor = true;
            this.chkIndentRight.CheckedChanged += new System.EventHandler(this.chkIndentRight_CheckedChanged);
            // 
            // chkIndentLeft
            // 
            this.chkIndentLeft.AutoSize = true;
            this.chkIndentLeft.Location = new System.Drawing.Point(44, 93);
            this.chkIndentLeft.Name = "chkIndentLeft";
            this.chkIndentLeft.Size = new System.Drawing.Size(48, 16);
            this.chkIndentLeft.TabIndex = 6;
            this.chkIndentLeft.Text = "左侧";
            this.chkIndentLeft.UseVisualStyleBackColor = true;
            this.chkIndentLeft.CheckedChanged += new System.EventHandler(this.chkIndentLeft_CheckedChanged);
            // 
            // chkAlignStyle
            // 
            this.chkAlignStyle.AutoSize = true;
            this.chkAlignStyle.Location = new System.Drawing.Point(44, 39);
            this.chkAlignStyle.Name = "chkAlignStyle";
            this.chkAlignStyle.Size = new System.Drawing.Size(72, 16);
            this.chkAlignStyle.TabIndex = 5;
            this.chkAlignStyle.Text = "对齐方式";
            this.chkAlignStyle.UseVisualStyleBackColor = true;
            this.chkAlignStyle.CheckedChanged += new System.EventHandler(this.chkAlignStyle_CheckedChanged);
            // 
            // chkAlignMesh
            // 
            this.chkAlignMesh.AutoSize = true;
            this.chkAlignMesh.Location = new System.Drawing.Point(44, 373);
            this.chkAlignMesh.Name = "chkAlignMesh";
            this.chkAlignMesh.Size = new System.Drawing.Size(216, 16);
            this.chkAlignMesh.TabIndex = 4;
            this.chkAlignMesh.Text = "如果定义了文档网格，则对齐到网格";
            this.chkAlignMesh.ThreeState = true;
            this.chkAlignMesh.UseVisualStyleBackColor = true;
            this.chkAlignMesh.CheckedChanged += new System.EventHandler(this.chkAlignMesh_CheckedChanged);
            // 
            // chkNoBlank
            // 
            this.chkNoBlank.AutoSize = true;
            this.chkNoBlank.Location = new System.Drawing.Point(44, 345);
            this.chkNoBlank.Name = "chkNoBlank";
            this.chkNoBlank.Size = new System.Drawing.Size(192, 16);
            this.chkNoBlank.TabIndex = 4;
            this.chkNoBlank.Text = "在相同样式的段落间不添加空格";
            this.chkNoBlank.ThreeState = true;
            this.chkNoBlank.UseVisualStyleBackColor = true;
            this.chkNoBlank.CheckedChanged += new System.EventHandler(this.chkNoBlank_CheckedChanged);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label12.Location = new System.Drawing.Point(40, 226);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(371, 12);
            this.label12.TabIndex = 0;
            this.label12.Text = "_____________________________________________________________";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label5.Location = new System.Drawing.Point(42, 66);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(371, 12);
            this.label5.TabIndex = 0;
            this.label5.Text = "_____________________________________________________________";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label10.Location = new System.Drawing.Point(42, 12);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(371, 12);
            this.label10.TabIndex = 0;
            this.label10.Text = "_____________________________________________________________";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(10, 226);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(29, 12);
            this.label11.TabIndex = 0;
            this.label11.Text = "间距";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(13, 12);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(29, 12);
            this.label9.TabIndex = 0;
            this.label9.Text = "常规";
            // 
            // numBeforeParaSpacing
            // 
            this.numBeforeParaSpacing.Enabled = false;
            this.numBeforeParaSpacing.Location = new System.Drawing.Point(122, 256);
            this.numBeforeParaSpacing.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numBeforeParaSpacing.Name = "numBeforeParaSpacing";
            this.numBeforeParaSpacing.Size = new System.Drawing.Size(64, 21);
            this.numBeforeParaSpacing.TabIndex = 2;
            this.numBeforeParaSpacing.ValueChanged += new System.EventHandler(this.numBeforeParaSpacing_ValueChanged);
            // 
            // numAfterParaSpacing
            // 
            this.numAfterParaSpacing.Enabled = false;
            this.numAfterParaSpacing.Location = new System.Drawing.Point(122, 285);
            this.numAfterParaSpacing.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numAfterParaSpacing.Name = "numAfterParaSpacing";
            this.numAfterParaSpacing.Size = new System.Drawing.Size(64, 21);
            this.numAfterParaSpacing.TabIndex = 2;
            this.numAfterParaSpacing.ValueChanged += new System.EventHandler(this.numAfterParaSpacing_ValueChanged);
            // 
            // numLineSpacing
            // 
            this.numLineSpacing.Enabled = false;
            this.numLineSpacing.Location = new System.Drawing.Point(262, 312);
            this.numLineSpacing.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numLineSpacing.Name = "numLineSpacing";
            this.numLineSpacing.Size = new System.Drawing.Size(64, 21);
            this.numLineSpacing.TabIndex = 2;
            this.numLineSpacing.ValueChanged += new System.EventHandler(this.numLineSpacing_ValueChanged);
            // 
            // cmbBeforeParaSpacingUnit
            // 
            this.cmbBeforeParaSpacingUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBeforeParaSpacingUnit.Enabled = false;
            this.cmbBeforeParaSpacingUnit.FormattingEnabled = true;
            this.cmbBeforeParaSpacingUnit.Location = new System.Drawing.Point(192, 256);
            this.cmbBeforeParaSpacingUnit.Name = "cmbBeforeParaSpacingUnit";
            this.cmbBeforeParaSpacingUnit.Size = new System.Drawing.Size(64, 20);
            this.cmbBeforeParaSpacingUnit.TabIndex = 1;
            this.cmbBeforeParaSpacingUnit.SelectedIndexChanged += new System.EventHandler(this.cmbBeforeParaSpacingUnit_SelectedIndexChanged);
            // 
            // cmbLineSpacingRule
            // 
            this.cmbLineSpacingRule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLineSpacingRule.Enabled = false;
            this.cmbLineSpacingRule.FormattingEnabled = true;
            this.cmbLineSpacingRule.Location = new System.Drawing.Point(122, 312);
            this.cmbLineSpacingRule.Name = "cmbLineSpacingRule";
            this.cmbLineSpacingRule.Size = new System.Drawing.Size(134, 20);
            this.cmbLineSpacingRule.TabIndex = 1;
            this.cmbLineSpacingRule.SelectedIndexChanged += new System.EventHandler(this.cmbLineSpacingRule_SelectedIndexChanged);
            // 
            // cmbAfterParaSpacingUnit
            // 
            this.cmbAfterParaSpacingUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAfterParaSpacingUnit.Enabled = false;
            this.cmbAfterParaSpacingUnit.FormattingEnabled = true;
            this.cmbAfterParaSpacingUnit.Location = new System.Drawing.Point(192, 285);
            this.cmbAfterParaSpacingUnit.Name = "cmbAfterParaSpacingUnit";
            this.cmbAfterParaSpacingUnit.Size = new System.Drawing.Size(64, 20);
            this.cmbAfterParaSpacingUnit.TabIndex = 1;
            this.cmbAfterParaSpacingUnit.SelectedIndexChanged += new System.EventHandler(this.cmbAfterParaSpacingUnit_SelectedIndexChanged);
            // 
            // cmbLineSpacingUnit
            // 
            this.cmbLineSpacingUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLineSpacingUnit.Enabled = false;
            this.cmbLineSpacingUnit.FormattingEnabled = true;
            this.cmbLineSpacingUnit.Location = new System.Drawing.Point(332, 313);
            this.cmbLineSpacingUnit.Name = "cmbLineSpacingUnit";
            this.cmbLineSpacingUnit.Size = new System.Drawing.Size(57, 20);
            this.cmbLineSpacingUnit.TabIndex = 1;
            this.cmbLineSpacingUnit.SelectedIndexChanged += new System.EventHandler(this.cmbLineSpacingUnit_SelectedIndexChanged);
            // 
            // tabPageLineBreakPageBreak
            // 
            this.tabPageLineBreakPageBreak.Controls.Add(this.label18);
            this.tabPageLineBreakPageBreak.Controls.Add(this.label17);
            this.tabPageLineBreakPageBreak.Controls.Add(this.label2);
            this.tabPageLineBreakPageBreak.Controls.Add(this.label3);
            this.tabPageLineBreakPageBreak.Controls.Add(this.chkCancelBreakWords);
            this.tabPageLineBreakPageBreak.Controls.Add(this.chkCancelLineNum);
            this.tabPageLineBreakPageBreak.Controls.Add(this.chkBreakPageBeforePara);
            this.tabPageLineBreakPageBreak.Controls.Add(this.chkParaNoBreakPage);
            this.tabPageLineBreakPageBreak.Controls.Add(this.chkKeepNext);
            this.tabPageLineBreakPageBreak.Controls.Add(this.chkAloneParaCtrl);
            this.tabPageLineBreakPageBreak.Location = new System.Drawing.Point(4, 22);
            this.tabPageLineBreakPageBreak.Name = "tabPageLineBreakPageBreak";
            this.tabPageLineBreakPageBreak.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageLineBreakPageBreak.Size = new System.Drawing.Size(426, 415);
            this.tabPageLineBreakPageBreak.TabIndex = 1;
            this.tabPageLineBreakPageBreak.Text = "换行与分页";
            this.tabPageLineBreakPageBreak.UseVisualStyleBackColor = true;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label18.Location = new System.Drawing.Point(95, 136);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(317, 12);
            this.label18.TabIndex = 5;
            this.label18.Text = "____________________________________________________";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(8, 136);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(89, 12);
            this.label17.TabIndex = 4;
            this.label17.Text = "格式设置例外项";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label2.Location = new System.Drawing.Point(37, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(371, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "_____________________________________________________________";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "分页";
            // 
            // chkCancelBreakWords
            // 
            this.chkCancelBreakWords.AutoSize = true;
            this.chkCancelBreakWords.Location = new System.Drawing.Point(39, 183);
            this.chkCancelBreakWords.Name = "chkCancelBreakWords";
            this.chkCancelBreakWords.Size = new System.Drawing.Size(72, 16);
            this.chkCancelBreakWords.TabIndex = 6;
            this.chkCancelBreakWords.Text = "取消断字";
            this.chkCancelBreakWords.ThreeState = true;
            this.chkCancelBreakWords.UseVisualStyleBackColor = true;
            this.chkCancelBreakWords.CheckedChanged += new System.EventHandler(this.chkCancelBreakWords_CheckedChanged);
            // 
            // chkCancelLineNum
            // 
            this.chkCancelLineNum.AutoSize = true;
            this.chkCancelLineNum.Location = new System.Drawing.Point(39, 161);
            this.chkCancelLineNum.Name = "chkCancelLineNum";
            this.chkCancelLineNum.Size = new System.Drawing.Size(72, 16);
            this.chkCancelLineNum.TabIndex = 6;
            this.chkCancelLineNum.Text = "取消行号";
            this.chkCancelLineNum.ThreeState = true;
            this.chkCancelLineNum.UseVisualStyleBackColor = true;
            this.chkCancelLineNum.CheckedChanged += new System.EventHandler(this.chkCancelLineNum_CheckedChanged);
            // 
            // chkBreakPageBeforePara
            // 
            this.chkBreakPageBeforePara.AutoSize = true;
            this.chkBreakPageBeforePara.Location = new System.Drawing.Point(39, 102);
            this.chkBreakPageBeforePara.Name = "chkBreakPageBeforePara";
            this.chkBreakPageBeforePara.Size = new System.Drawing.Size(72, 16);
            this.chkBreakPageBeforePara.TabIndex = 6;
            this.chkBreakPageBeforePara.Text = "段前分页";
            this.chkBreakPageBeforePara.ThreeState = true;
            this.chkBreakPageBeforePara.UseVisualStyleBackColor = true;
            this.chkBreakPageBeforePara.CheckedChanged += new System.EventHandler(this.chkBreakPageBeforePara_CheckedChanged);
            // 
            // chkParaNoBreakPage
            // 
            this.chkParaNoBreakPage.AutoSize = true;
            this.chkParaNoBreakPage.Location = new System.Drawing.Point(39, 80);
            this.chkParaNoBreakPage.Name = "chkParaNoBreakPage";
            this.chkParaNoBreakPage.Size = new System.Drawing.Size(84, 16);
            this.chkParaNoBreakPage.TabIndex = 6;
            this.chkParaNoBreakPage.Text = "段中不分页";
            this.chkParaNoBreakPage.ThreeState = true;
            this.chkParaNoBreakPage.UseVisualStyleBackColor = true;
            this.chkParaNoBreakPage.CheckedChanged += new System.EventHandler(this.chkParaNoBreakPage_CheckedChanged);
            // 
            // chkKeepNext
            // 
            this.chkKeepNext.AutoSize = true;
            this.chkKeepNext.Location = new System.Drawing.Point(39, 58);
            this.chkKeepNext.Name = "chkKeepNext";
            this.chkKeepNext.Size = new System.Drawing.Size(84, 16);
            this.chkKeepNext.TabIndex = 6;
            this.chkKeepNext.Text = "与下段同页";
            this.chkKeepNext.ThreeState = true;
            this.chkKeepNext.UseVisualStyleBackColor = true;
            this.chkKeepNext.CheckedChanged += new System.EventHandler(this.chkKeepNext_CheckedChanged);
            // 
            // chkAloneParaCtrl
            // 
            this.chkAloneParaCtrl.AutoSize = true;
            this.chkAloneParaCtrl.Location = new System.Drawing.Point(39, 36);
            this.chkAloneParaCtrl.Name = "chkAloneParaCtrl";
            this.chkAloneParaCtrl.Size = new System.Drawing.Size(72, 16);
            this.chkAloneParaCtrl.TabIndex = 6;
            this.chkAloneParaCtrl.Text = "孤行控制";
            this.chkAloneParaCtrl.ThreeState = true;
            this.chkAloneParaCtrl.UseVisualStyleBackColor = true;
            this.chkAloneParaCtrl.CheckedChanged += new System.EventHandler(this.chkAlongParaCtrl_CheckedChanged);
            // 
            // tabPageChineseStyle
            // 
            this.tabPageChineseStyle.Controls.Add(this.chkTextAlignStyle);
            this.tabPageChineseStyle.Controls.Add(this.cmbTextAlignStyle);
            this.tabPageChineseStyle.Controls.Add(this.label22);
            this.tabPageChineseStyle.Controls.Add(this.label19);
            this.tabPageChineseStyle.Controls.Add(this.label21);
            this.tabPageChineseStyle.Controls.Add(this.label20);
            this.tabPageChineseStyle.Controls.Add(this.chkAutoAdjustNumLineSpacing);
            this.tabPageChineseStyle.Controls.Add(this.chkAutoAdjustLineSpacing);
            this.tabPageChineseStyle.Controls.Add(this.chkAllowCompressCma);
            this.tabPageChineseStyle.Controls.Add(this.chkAllowCmaOverLimit);
            this.tabPageChineseStyle.Controls.Add(this.chkAllowAsciiBreakLineInPara);
            this.tabPageChineseStyle.Controls.Add(this.chkCtrlFromChinese);
            this.tabPageChineseStyle.Location = new System.Drawing.Point(4, 22);
            this.tabPageChineseStyle.Name = "tabPageChineseStyle";
            this.tabPageChineseStyle.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageChineseStyle.Size = new System.Drawing.Size(426, 415);
            this.tabPageChineseStyle.TabIndex = 2;
            this.tabPageChineseStyle.Text = "中文版式";
            this.tabPageChineseStyle.UseVisualStyleBackColor = true;
            // 
            // chkTextAlignStyle
            // 
            this.chkTextAlignStyle.AutoSize = true;
            this.chkTextAlignStyle.Location = new System.Drawing.Point(39, 218);
            this.chkTextAlignStyle.Name = "chkTextAlignStyle";
            this.chkTextAlignStyle.Size = new System.Drawing.Size(96, 16);
            this.chkTextAlignStyle.TabIndex = 11;
            this.chkTextAlignStyle.Text = "文本对齐方式";
            this.chkTextAlignStyle.UseVisualStyleBackColor = true;
            this.chkTextAlignStyle.CheckedChanged += new System.EventHandler(this.chkTextAlignStyle_CheckedChanged);
            // 
            // cmbTextAlignStyle
            // 
            this.cmbTextAlignStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTextAlignStyle.Enabled = false;
            this.cmbTextAlignStyle.FormattingEnabled = true;
            this.cmbTextAlignStyle.Location = new System.Drawing.Point(141, 216);
            this.cmbTextAlignStyle.Name = "cmbTextAlignStyle";
            this.cmbTextAlignStyle.Size = new System.Drawing.Size(87, 20);
            this.cmbTextAlignStyle.TabIndex = 10;
            this.cmbTextAlignStyle.SelectedIndexChanged += new System.EventHandler(this.cmbTextAlignStyle_SelectedIndexChanged);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label22.Location = new System.Drawing.Point(58, 117);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(353, 12);
            this.label22.TabIndex = 8;
            this.label22.Text = "__________________________________________________________";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.ForeColor = System.Drawing.SystemColors.ControlDark;
            this.label19.Location = new System.Drawing.Point(37, 3);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(371, 12);
            this.label19.TabIndex = 8;
            this.label19.Text = "_____________________________________________________________";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(8, 117);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(53, 12);
            this.label21.TabIndex = 7;
            this.label21.Text = "字符间距";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(8, 3);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(29, 12);
            this.label20.TabIndex = 7;
            this.label20.Text = "换行";
            // 
            // chkAutoAdjustNumLineSpacing
            // 
            this.chkAutoAdjustNumLineSpacing.AutoSize = true;
            this.chkAutoAdjustNumLineSpacing.Location = new System.Drawing.Point(39, 185);
            this.chkAutoAdjustNumLineSpacing.Name = "chkAutoAdjustNumLineSpacing";
            this.chkAutoAdjustNumLineSpacing.Size = new System.Drawing.Size(168, 16);
            this.chkAutoAdjustNumLineSpacing.TabIndex = 9;
            this.chkAutoAdjustNumLineSpacing.Text = "自动调整中文与数字的间距";
            this.chkAutoAdjustNumLineSpacing.ThreeState = true;
            this.chkAutoAdjustNumLineSpacing.UseVisualStyleBackColor = true;
            this.chkAutoAdjustNumLineSpacing.CheckedChanged += new System.EventHandler(this.chkAutoAdjustNumLineSpacing_CheckedChanged);
            // 
            // chkAutoAdjustLineSpacing
            // 
            this.chkAutoAdjustLineSpacing.AutoSize = true;
            this.chkAutoAdjustLineSpacing.Location = new System.Drawing.Point(39, 163);
            this.chkAutoAdjustLineSpacing.Name = "chkAutoAdjustLineSpacing";
            this.chkAutoAdjustLineSpacing.Size = new System.Drawing.Size(168, 16);
            this.chkAutoAdjustLineSpacing.TabIndex = 9;
            this.chkAutoAdjustLineSpacing.Text = "自动调整中文与西文的间距";
            this.chkAutoAdjustLineSpacing.ThreeState = true;
            this.chkAutoAdjustLineSpacing.UseVisualStyleBackColor = true;
            this.chkAutoAdjustLineSpacing.CheckedChanged += new System.EventHandler(this.chkAutoAdjustLineSpacing_CheckedChanged);
            // 
            // chkAllowCompressCma
            // 
            this.chkAllowCompressCma.AutoSize = true;
            this.chkAllowCompressCma.Location = new System.Drawing.Point(39, 141);
            this.chkAllowCompressCma.Name = "chkAllowCompressCma";
            this.chkAllowCompressCma.Size = new System.Drawing.Size(120, 16);
            this.chkAllowCompressCma.TabIndex = 9;
            this.chkAllowCompressCma.Text = "允许行首标点压缩";
            this.chkAllowCompressCma.ThreeState = true;
            this.chkAllowCompressCma.UseVisualStyleBackColor = true;
            this.chkAllowCompressCma.CheckedChanged += new System.EventHandler(this.chkAllowCompressCma_CheckedChanged);
            // 
            // chkAllowCmaOverLimit
            // 
            this.chkAllowCmaOverLimit.AutoSize = true;
            this.chkAllowCmaOverLimit.Location = new System.Drawing.Point(39, 71);
            this.chkAllowCmaOverLimit.Name = "chkAllowCmaOverLimit";
            this.chkAllowCmaOverLimit.Size = new System.Drawing.Size(120, 16);
            this.chkAllowCmaOverLimit.TabIndex = 9;
            this.chkAllowCmaOverLimit.Text = "允许标点溢出边界";
            this.chkAllowCmaOverLimit.ThreeState = true;
            this.chkAllowCmaOverLimit.UseVisualStyleBackColor = true;
            this.chkAllowCmaOverLimit.CheckedChanged += new System.EventHandler(this.chkAllowCmaOverLimit_CheckedChanged);
            // 
            // chkAllowAsciiBreakLineInPara
            // 
            this.chkAllowAsciiBreakLineInPara.AutoSize = true;
            this.chkAllowAsciiBreakLineInPara.Location = new System.Drawing.Point(39, 49);
            this.chkAllowAsciiBreakLineInPara.Name = "chkAllowAsciiBreakLineInPara";
            this.chkAllowAsciiBreakLineInPara.Size = new System.Drawing.Size(156, 16);
            this.chkAllowAsciiBreakLineInPara.TabIndex = 9;
            this.chkAllowAsciiBreakLineInPara.Text = "允许西文在单词中间换行";
            this.chkAllowAsciiBreakLineInPara.ThreeState = true;
            this.chkAllowAsciiBreakLineInPara.UseVisualStyleBackColor = true;
            this.chkAllowAsciiBreakLineInPara.CheckedChanged += new System.EventHandler(this.chkAllowAsciiBreakLineInPara_CheckedChanged);
            // 
            // chkCtrlFromChinese
            // 
            this.chkCtrlFromChinese.AutoSize = true;
            this.chkCtrlFromChinese.Location = new System.Drawing.Point(39, 27);
            this.chkCtrlFromChinese.Name = "chkCtrlFromChinese";
            this.chkCtrlFromChinese.Size = new System.Drawing.Size(156, 16);
            this.chkCtrlFromChinese.TabIndex = 9;
            this.chkCtrlFromChinese.Text = "按中文习惯控制首尾字符";
            this.chkCtrlFromChinese.ThreeState = true;
            this.chkCtrlFromChinese.UseVisualStyleBackColor = true;
            this.chkCtrlFromChinese.CheckedChanged += new System.EventHandler(this.chkCtrlFromChinese_CheckedChanged);
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(258, 443);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "确定";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(355, 443);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // frmParagraphFormat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(435, 474);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.tabControl1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmParagraphFormat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "段落";
            this.Load += new System.EventHandler(this.frmParagraphFormat_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numIndentSpecial)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numIndentRight)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numIndentLeft)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPageIndentLineSpace.ResumeLayout(false);
            this.tabPageIndentLineSpace.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numBeforeParaSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numAfterParaSpacing)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLineSpacing)).EndInit();
            this.tabPageLineBreakPageBreak.ResumeLayout(false);
            this.tabPageLineBreakPageBreak.PerformLayout();
            this.tabPageChineseStyle.ResumeLayout(false);
            this.tabPageChineseStyle.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbAlignStyle;
        private System.Windows.Forms.NumericUpDown numIndentSpecial;
        private System.Windows.Forms.ComboBox cmbIndentSpecial;
        private System.Windows.Forms.ComboBox cmbIndentSpecialUnit;
        private System.Windows.Forms.ComboBox cmbIndentRightUnit;
        private System.Windows.Forms.NumericUpDown numIndentRight;
        private System.Windows.Forms.ComboBox cmbIndentLeftUnit;
        private System.Windows.Forms.NumericUpDown numIndentLeft;
        private System.Windows.Forms.CheckBox chkSymIndent;
        private System.Windows.Forms.CheckBox chkAutoAlignRight;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageIndentLineSpace;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.NumericUpDown numBeforeParaSpacing;
        private System.Windows.Forms.NumericUpDown numAfterParaSpacing;
        private System.Windows.Forms.NumericUpDown numLineSpacing;
        private System.Windows.Forms.ComboBox cmbBeforeParaSpacingUnit;
        private System.Windows.Forms.ComboBox cmbLineSpacingRule;
        private System.Windows.Forms.ComboBox cmbAfterParaSpacingUnit;
        private System.Windows.Forms.ComboBox cmbLineSpacingUnit;
        private System.Windows.Forms.TabPage tabPageLineBreakPageBreak;
        private System.Windows.Forms.TabPage tabPageChineseStyle;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox chkAlignMesh;
        private System.Windows.Forms.CheckBox chkNoBlank;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkCancelBreakWords;
        private System.Windows.Forms.CheckBox chkCancelLineNum;
        private System.Windows.Forms.CheckBox chkBreakPageBeforePara;
        private System.Windows.Forms.CheckBox chkParaNoBreakPage;
        private System.Windows.Forms.CheckBox chkKeepNext;
        private System.Windows.Forms.CheckBox chkAloneParaCtrl;
        private System.Windows.Forms.ComboBox cmbTextAlignStyle;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.CheckBox chkAutoAdjustNumLineSpacing;
        private System.Windows.Forms.CheckBox chkAutoAdjustLineSpacing;
        private System.Windows.Forms.CheckBox chkAllowCompressCma;
        private System.Windows.Forms.CheckBox chkAllowCmaOverLimit;
        private System.Windows.Forms.CheckBox chkAllowAsciiBreakLineInPara;
        private System.Windows.Forms.CheckBox chkCtrlFromChinese;
        private System.Windows.Forms.CheckBox chkIndentSpecial;
        private System.Windows.Forms.CheckBox chkIndentRight;
        private System.Windows.Forms.CheckBox chkIndentLeft;
        private System.Windows.Forms.CheckBox chkAlignStyle;
        private System.Windows.Forms.CheckBox chkParaLineSpace;
        private System.Windows.Forms.CheckBox chkParaLineSpaceAfter;
        private System.Windows.Forms.CheckBox chkParaLineSpaceBefore;
        private System.Windows.Forms.CheckBox chkTextAlignStyle;
        private System.Windows.Forms.CheckBox chkSpaceAfterAuto;
        private System.Windows.Forms.CheckBox chkSpaceBeforeAuto;
    }
}