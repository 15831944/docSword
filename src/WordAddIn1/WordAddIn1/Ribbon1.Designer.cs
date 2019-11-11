namespace OfficeAssist
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpConfig = this.Factory.CreateRibbonGroup();
            this.btnLogin = this.Factory.CreateRibbonButton();
            this.chkAutoLogin = this.Factory.CreateRibbonCheckBox();
            this.grpComOp = this.Factory.CreateRibbonGroup();
            this.box8 = this.Factory.CreateRibbonBox();
            this.btnAddHeaderLine = this.Factory.CreateRibbonButton();
            this.btnRemoveHeaderLine = this.Factory.CreateRibbonButton();
            this.box9 = this.Factory.CreateRibbonBox();
            this.btnAddFooterLine = this.Factory.CreateRibbonButton();
            this.btnClearFooterLine = this.Factory.CreateRibbonButton();
            this.btnRibInsertSeparateTblContent = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.btnStrictCenter = this.Factory.CreateRibbonButton();
            this.btnCenterAllPics = this.Factory.CreateRibbonButton();
            this.btnRibCenterTables = this.Factory.CreateRibbonButton();
            this.separator7 = this.Factory.CreateRibbonSeparator();
            this.RibbtnOpenCurDocDir = this.Factory.CreateRibbonButton();
            this.chkBoxUpdTblCntOnSaving = this.Factory.CreateRibbonCheckBox();
            this.chkBoxUpdTblCntOnClose = this.Factory.CreateRibbonCheckBox();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.ribbtnUnitedHeaders = this.Factory.CreateRibbonButton();
            this.ribbtnUnitedFooters = this.Factory.CreateRibbonButton();
            this.lblCurParaOutLine = this.Factory.CreateRibbonLabel();
            this.rbBtnCalculate = this.Factory.CreateRibbonButton();
            this.grpAutoNumbering = this.Factory.CreateRibbonGroup();
            this.ribBtnFillSn = this.Factory.CreateRibbonButton();
            this.ribBtnFillSn2EndRow = this.Factory.CreateRibbonButton();
            this.ribBtnFillSelection = this.Factory.CreateRibbonButton();
            this.grpOutline = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.box4 = this.Factory.CreateRibbonBox();
            this.ribBtnOutLevel1 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevel2 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevel3 = this.Factory.CreateRibbonButton();
            this.box3 = this.Factory.CreateRibbonBox();
            this.ribBtnOutLevel4 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevel5 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevel6 = this.Factory.CreateRibbonButton();
            this.box5 = this.Factory.CreateRibbonBox();
            this.ribBtnOutLevel7 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevel8 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevel9 = this.Factory.CreateRibbonButton();
            this.ribBtnOutLevelTextBody = this.Factory.CreateRibbonButton();
            this.ribBtnViewOutlineLevel = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnOutlineSamePrev = this.Factory.CreateRibbonButton();
            this.btnOutlineLow1Prev = this.Factory.CreateRibbonButton();
            this.btnOutlineHigh1Prev = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnOutlinePromote = this.Factory.CreateRibbonButton();
            this.btnOutlineDemote = this.Factory.CreateRibbonButton();
            this.chkOnlyNonTextBodyPara = this.Factory.CreateRibbonCheckBox();
            this.grpPane = this.Factory.CreateRibbonGroup();
            this.btnCopyHeadingStyles = this.Factory.CreateRibbonButton();
            this.btnPasteHeadingStyles = this.Factory.CreateRibbonButton();
            this.ribbtnCopyHeadingsStructure = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.ribbtnSaveCurHeadingStyle2Style = this.Factory.CreateRibbonButton();
            this.chkHeadingsStylesPersist = this.Factory.CreateRibbonCheckBox();
            this.grpQuickBookmark = this.Factory.CreateRibbonGroup();
            this.btnNavAddBkmk = this.Factory.CreateRibbonButton();
            this.ribBtnRemoveJetNav = this.Factory.CreateRibbonButton();
            this.btnClearBkmk = this.Factory.CreateRibbonButton();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.box6 = this.Factory.CreateRibbonBox();
            this.btnNavFirst = this.Factory.CreateRibbonButton();
            this.btnNavLast = this.Factory.CreateRibbonButton();
            this.box7 = this.Factory.CreateRibbonBox();
            this.btnNavPrev = this.Factory.CreateRibbonButton();
            this.btnNavNext = this.Factory.CreateRibbonButton();
            this.ribBtnJump2Toc = this.Factory.CreateRibbonButton();
            this.ribBtnPrevEditPos = this.Factory.CreateRibbonButton();
            this.ribBtnNextEditPos = this.Factory.CreateRibbonButton();
            this.groupLocalVer = this.Factory.CreateRibbonGroup();
            this.btnLocalVerMileStone = this.Factory.CreateRibbonButton();
            this.ribbtnOpenVerDir = this.Factory.CreateRibbonButton();
            this.chkGenLocalVer = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnTogglePanePos = this.Factory.CreateRibbonButton();
            this.toggleTaskWin = this.Factory.CreateRibbonButton();
            this.grpFuncPages = this.Factory.CreateRibbonGroup();
            this.ribBtnCheckUpdate = this.Factory.CreateRibbonButton();
            this.chkAutoCheckUpdate = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ribBtnTutorial = this.Factory.CreateRibbonButton();
            this.ribBtnHelp = this.Factory.CreateRibbonButton();
            this.ribbtnAbout = this.Factory.CreateRibbonButton();
            this.ribLoadSoloLic = this.Factory.CreateRibbonButton();
            this.RibbtnRegister = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpConfig.SuspendLayout();
            this.grpComOp.SuspendLayout();
            this.box8.SuspendLayout();
            this.box9.SuspendLayout();
            this.grpAutoNumbering.SuspendLayout();
            this.grpOutline.SuspendLayout();
            this.box2.SuspendLayout();
            this.box4.SuspendLayout();
            this.box3.SuspendLayout();
            this.box5.SuspendLayout();
            this.box1.SuspendLayout();
            this.grpPane.SuspendLayout();
            this.grpQuickBookmark.SuspendLayout();
            this.box6.SuspendLayout();
            this.box7.SuspendLayout();
            this.groupLocalVer.SuspendLayout();
            this.group2.SuspendLayout();
            this.grpFuncPages.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpConfig);
            this.tab1.Groups.Add(this.grpComOp);
            this.tab1.Groups.Add(this.grpAutoNumbering);
            this.tab1.Groups.Add(this.grpOutline);
            this.tab1.Groups.Add(this.grpPane);
            this.tab1.Groups.Add(this.grpQuickBookmark);
            this.tab1.Groups.Add(this.groupLocalVer);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.grpFuncPages);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "doc利器";
            this.tab1.Name = "tab1";
            // 
            // grpConfig
            // 
            this.grpConfig.Items.Add(this.btnLogin);
            this.grpConfig.Items.Add(this.chkAutoLogin);
            this.grpConfig.Label = "文库";
            this.grpConfig.Name = "grpConfig";
            // 
            // btnLogin
            // 
            this.btnLogin.Image = ((System.Drawing.Image)(resources.GetObject("btnLogin.Image")));
            this.btnLogin.Label = "登录";
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.ScreenTip = "登录/注销文库系统";
            this.btnLogin.ShowImage = true;
            this.btnLogin.SuperTip = "若存在文库系统，此功能实现登录/注销文库系统功能";
            this.btnLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLogin_Click);
            // 
            // chkAutoLogin
            // 
            this.chkAutoLogin.Label = "自动";
            this.chkAutoLogin.Name = "chkAutoLogin";
            this.chkAutoLogin.ScreenTip = "自动登录文库系统";
            this.chkAutoLogin.SuperTip = "若前次成功登录文库系统，勾选后，则下次即自动登录文库系统";
            this.chkAutoLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkAutoLogin_Click);
            // 
            // grpComOp
            // 
            this.grpComOp.Items.Add(this.box8);
            this.grpComOp.Items.Add(this.box9);
            this.grpComOp.Items.Add(this.btnRibInsertSeparateTblContent);
            this.grpComOp.Items.Add(this.separator3);
            this.grpComOp.Items.Add(this.btnStrictCenter);
            this.grpComOp.Items.Add(this.btnCenterAllPics);
            this.grpComOp.Items.Add(this.btnRibCenterTables);
            this.grpComOp.Items.Add(this.separator7);
            this.grpComOp.Items.Add(this.RibbtnOpenCurDocDir);
            this.grpComOp.Items.Add(this.chkBoxUpdTblCntOnSaving);
            this.grpComOp.Items.Add(this.chkBoxUpdTblCntOnClose);
            this.grpComOp.Items.Add(this.separator5);
            this.grpComOp.Items.Add(this.ribbtnUnitedHeaders);
            this.grpComOp.Items.Add(this.ribbtnUnitedFooters);
            this.grpComOp.Items.Add(this.lblCurParaOutLine);
            this.grpComOp.Items.Add(this.rbBtnCalculate);
            this.grpComOp.Label = "常用";
            this.grpComOp.Name = "grpComOp";
            // 
            // box8
            // 
            this.box8.Items.Add(this.btnAddHeaderLine);
            this.box8.Items.Add(this.btnRemoveHeaderLine);
            this.box8.Name = "box8";
            // 
            // btnAddHeaderLine
            // 
            this.btnAddHeaderLine.Image = ((System.Drawing.Image)(resources.GetObject("btnAddHeaderLine.Image")));
            this.btnAddHeaderLine.Label = "设置页眉线";
            this.btnAddHeaderLine.Name = "btnAddHeaderLine";
            this.btnAddHeaderLine.ScreenTip = "设置当前节页眉线";
            this.btnAddHeaderLine.ShowImage = true;
            this.btnAddHeaderLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddHeaderLine_Click);
            // 
            // btnRemoveHeaderLine
            // 
            this.btnRemoveHeaderLine.Image = ((System.Drawing.Image)(resources.GetObject("btnRemoveHeaderLine.Image")));
            this.btnRemoveHeaderLine.Label = "清除页眉线";
            this.btnRemoveHeaderLine.Name = "btnRemoveHeaderLine";
            this.btnRemoveHeaderLine.ScreenTip = "清除当前节页眉线";
            this.btnRemoveHeaderLine.ShowImage = true;
            this.btnRemoveHeaderLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemoveHeaderLine_Click);
            // 
            // box9
            // 
            this.box9.Items.Add(this.btnAddFooterLine);
            this.box9.Items.Add(this.btnClearFooterLine);
            this.box9.Name = "box9";
            // 
            // btnAddFooterLine
            // 
            this.btnAddFooterLine.Image = ((System.Drawing.Image)(resources.GetObject("btnAddFooterLine.Image")));
            this.btnAddFooterLine.Label = "设置页脚线";
            this.btnAddFooterLine.Name = "btnAddFooterLine";
            this.btnAddFooterLine.ScreenTip = "设置当前节页脚线";
            this.btnAddFooterLine.ShowImage = true;
            this.btnAddFooterLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddFooterLine_Click);
            // 
            // btnClearFooterLine
            // 
            this.btnClearFooterLine.Image = ((System.Drawing.Image)(resources.GetObject("btnClearFooterLine.Image")));
            this.btnClearFooterLine.Label = "清除页脚线";
            this.btnClearFooterLine.Name = "btnClearFooterLine";
            this.btnClearFooterLine.ScreenTip = "清除当前节页脚线";
            this.btnClearFooterLine.ShowImage = true;
            this.btnClearFooterLine.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearFooterLine_Click);
            // 
            // btnRibInsertSeparateTblContent
            // 
            this.btnRibInsertSeparateTblContent.Image = ((System.Drawing.Image)(resources.GetObject("btnRibInsertSeparateTblContent.Image")));
            this.btnRibInsertSeparateTblContent.Label = "插入独立目录节";
            this.btnRibInsertSeparateTblContent.Name = "btnRibInsertSeparateTblContent";
            this.btnRibInsertSeparateTblContent.ScreenTip = "在当前位置插入独立目录节";
            this.btnRibInsertSeparateTblContent.ShowImage = true;
            this.btnRibInsertSeparateTblContent.SuperTip = "在当前位置插入独立目录节，插入的目录节将原文分为3个独立节，之前部分即第1节为封面，插入部分即第2节为目录节，插入位置之后即第3节为正文节";
            this.btnRibInsertSeparateTblContent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRibInsertSeparateTblContent_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // btnStrictCenter
            // 
            this.btnStrictCenter.Description = "整页标准居中";
            this.btnStrictCenter.Image = ((System.Drawing.Image)(resources.GetObject("btnStrictCenter.Image")));
            this.btnStrictCenter.Label = "添加页居中";
            this.btnStrictCenter.Name = "btnStrictCenter";
            this.btnStrictCenter.ScreenTip = "整页标准居中，占整页且居于页正中";
            this.btnStrictCenter.ShowImage = true;
            this.btnStrictCenter.SuperTip = "可应用于保持整页（不受上下文段落影响）标准居中的页面如封面等";
            this.btnStrictCenter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStrictCenter_Click);
            // 
            // btnCenterAllPics
            // 
            this.btnCenterAllPics.Image = ((System.Drawing.Image)(resources.GetObject("btnCenterAllPics.Image")));
            this.btnCenterAllPics.Label = "居中图片";
            this.btnCenterAllPics.Name = "btnCenterAllPics";
            this.btnCenterAllPics.ScreenTip = "将选择范围或全文内所有的独立成行的图片居中（内嵌在文字中的图片不作处理）";
            this.btnCenterAllPics.ShowImage = true;
            this.btnCenterAllPics.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCenterAllPics_Click);
            // 
            // btnRibCenterTables
            // 
            this.btnRibCenterTables.Image = ((System.Drawing.Image)(resources.GetObject("btnRibCenterTables.Image")));
            this.btnRibCenterTables.Label = "居中表格";
            this.btnRibCenterTables.Name = "btnRibCenterTables";
            this.btnRibCenterTables.ScreenTip = "将选择范围或全文内所有的表格进行居中（不影响表格内部的对齐设置）";
            this.btnRibCenterTables.ShowImage = true;
            this.btnRibCenterTables.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRibCenterTables_Click);
            // 
            // separator7
            // 
            this.separator7.Name = "separator7";
            // 
            // RibbtnOpenCurDocDir
            // 
            this.RibbtnOpenCurDocDir.Image = ((System.Drawing.Image)(resources.GetObject("RibbtnOpenCurDocDir.Image")));
            this.RibbtnOpenCurDocDir.Label = "打开文档目录";
            this.RibbtnOpenCurDocDir.Name = "RibbtnOpenCurDocDir";
            this.RibbtnOpenCurDocDir.ScreenTip = "打开当前已保存文档所在目录";
            this.RibbtnOpenCurDocDir.ShowImage = true;
            this.RibbtnOpenCurDocDir.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbtnOpenCurDocDir_Click);
            // 
            // chkBoxUpdTblCntOnSaving
            // 
            this.chkBoxUpdTblCntOnSaving.Label = "保存时更新目录";
            this.chkBoxUpdTblCntOnSaving.Name = "chkBoxUpdTblCntOnSaving";
            this.chkBoxUpdTblCntOnSaving.ScreenTip = "勾选时则每次保存前更新目录，以保证目录最新";
            this.chkBoxUpdTblCntOnSaving.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkBoxUpdTblCntOnSaving_Click);
            // 
            // chkBoxUpdTblCntOnClose
            // 
            this.chkBoxUpdTblCntOnClose.Label = "关闭时更新目录";
            this.chkBoxUpdTblCntOnClose.Name = "chkBoxUpdTblCntOnClose";
            this.chkBoxUpdTblCntOnClose.ScreenTip = "勾选时则关闭前更新目录，以保证目录最新";
            this.chkBoxUpdTblCntOnClose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkBoxUpdTblCntOnClose_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // ribbtnUnitedHeaders
            // 
            this.ribbtnUnitedHeaders.Image = ((System.Drawing.Image)(resources.GetObject("ribbtnUnitedHeaders.Image")));
            this.ribbtnUnitedHeaders.Label = "统一页眉";
            this.ribbtnUnitedHeaders.Name = "ribbtnUnitedHeaders";
            this.ribbtnUnitedHeaders.ScreenTip = "将当前节的页眉统一到其它选中的目标节";
            this.ribbtnUnitedHeaders.ShowImage = true;
            this.ribbtnUnitedHeaders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbtnUnitedHeaders_Click);
            // 
            // ribbtnUnitedFooters
            // 
            this.ribbtnUnitedFooters.Image = ((System.Drawing.Image)(resources.GetObject("ribbtnUnitedFooters.Image")));
            this.ribbtnUnitedFooters.Label = "统一页脚";
            this.ribbtnUnitedFooters.Name = "ribbtnUnitedFooters";
            this.ribbtnUnitedFooters.ScreenTip = "将当前节的页脚统一到其它选中的目标节";
            this.ribbtnUnitedFooters.ShowImage = true;
            this.ribbtnUnitedFooters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbtnUnitedFooters_Click);
            // 
            // lblCurParaOutLine
            // 
            this.lblCurParaOutLine.Label = "当前段落级：";
            this.lblCurParaOutLine.Name = "lblCurParaOutLine";
            this.lblCurParaOutLine.ScreenTip = "显示当前段落（若有多选则指第一个段落）的大纲级别（正表示正文，1-9表示相应大纲级别）";
            this.lblCurParaOutLine.Visible = false;
            // 
            // rbBtnCalculate
            // 
            this.rbBtnCalculate.Image = ((System.Drawing.Image)(resources.GetObject("rbBtnCalculate.Image")));
            this.rbBtnCalculate.Label = "计算";
            this.rbBtnCalculate.Name = "rbBtnCalculate";
            this.rbBtnCalculate.ShowImage = true;
            this.rbBtnCalculate.SuperTip = "计算选中的段落或表格的数值(合计、均值等基本统计）";
            this.rbBtnCalculate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rbBtnCalculate_Click);
            // 
            // grpAutoNumbering
            // 
            this.grpAutoNumbering.Items.Add(this.ribBtnFillSn);
            this.grpAutoNumbering.Items.Add(this.ribBtnFillSn2EndRow);
            this.grpAutoNumbering.Items.Add(this.ribBtnFillSelection);
            this.grpAutoNumbering.Label = "填充";
            this.grpAutoNumbering.Name = "grpAutoNumbering";
            // 
            // ribBtnFillSn
            // 
            this.ribBtnFillSn.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnFillSn.Image")));
            this.ribBtnFillSn.Label = "智能填充";
            this.ribBtnFillSn.Name = "ribBtnFillSn";
            this.ribBtnFillSn.ScreenTip = "智能填充";
            this.ribBtnFillSn.ShowImage = true;
            this.ribBtnFillSn.SuperTip = "若在表格中，没有选择则将当前单元格的内容按顺序填充到表末行；若选择了范围则将选择区第一段落累加填充至选择范围内的最后段落；支持日期和任意编号，序号增加以最右边的数" +
    "字为基数开始";
            this.ribBtnFillSn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnFillSn_Click);
            // 
            // ribBtnFillSn2EndRow
            // 
            this.ribBtnFillSn2EndRow.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnFillSn2EndRow.Image")));
            this.ribBtnFillSn2EndRow.Label = "填充至表末行";
            this.ribBtnFillSn2EndRow.Name = "ribBtnFillSn2EndRow";
            this.ribBtnFillSn2EndRow.ScreenTip = "若在表格中，则将当前单元格的内容累加填充到表末行";
            this.ribBtnFillSn2EndRow.ShowImage = true;
            this.ribBtnFillSn2EndRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnFillSn2EndRow_Click);
            // 
            // ribBtnFillSelection
            // 
            this.ribBtnFillSelection.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnFillSelection.Image")));
            this.ribBtnFillSelection.Label = "填充选择区";
            this.ribBtnFillSelection.Name = "ribBtnFillSelection";
            this.ribBtnFillSelection.ScreenTip = "将当前选择区第一段落内容累加填充到选择区最后一个段落";
            this.ribBtnFillSelection.ShowImage = true;
            this.ribBtnFillSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnFillSelection_Click);
            // 
            // grpOutline
            // 
            this.grpOutline.Items.Add(this.box2);
            this.grpOutline.Items.Add(this.separator2);
            this.grpOutline.Items.Add(this.btnOutlineSamePrev);
            this.grpOutline.Items.Add(this.btnOutlineLow1Prev);
            this.grpOutline.Items.Add(this.btnOutlineHigh1Prev);
            this.grpOutline.Items.Add(this.separator1);
            this.grpOutline.Items.Add(this.box1);
            this.grpOutline.Label = "大纲级别";
            this.grpOutline.Name = "grpOutline";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.box4);
            this.box2.Items.Add(this.box3);
            this.box2.Items.Add(this.box5);
            this.box2.Items.Add(this.ribBtnOutLevelTextBody);
            this.box2.Items.Add(this.ribBtnViewOutlineLevel);
            this.box2.Name = "box2";
            // 
            // box4
            // 
            this.box4.Items.Add(this.ribBtnOutLevel1);
            this.box4.Items.Add(this.ribBtnOutLevel2);
            this.box4.Items.Add(this.ribBtnOutLevel3);
            this.box4.Name = "box4";
            // 
            // ribBtnOutLevel1
            // 
            this.ribBtnOutLevel1.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel1.Image")));
            this.ribBtnOutLevel1.Label = "1";
            this.ribBtnOutLevel1.Name = "ribBtnOutLevel1";
            this.ribBtnOutLevel1.ShowImage = true;
            this.ribBtnOutLevel1.ShowLabel = false;
            this.ribBtnOutLevel1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel1_Click);
            // 
            // ribBtnOutLevel2
            // 
            this.ribBtnOutLevel2.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel2.Image")));
            this.ribBtnOutLevel2.Label = "2";
            this.ribBtnOutLevel2.Name = "ribBtnOutLevel2";
            this.ribBtnOutLevel2.ShowImage = true;
            this.ribBtnOutLevel2.ShowLabel = false;
            this.ribBtnOutLevel2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel2_Click);
            // 
            // ribBtnOutLevel3
            // 
            this.ribBtnOutLevel3.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel3.Image")));
            this.ribBtnOutLevel3.Label = "3";
            this.ribBtnOutLevel3.Name = "ribBtnOutLevel3";
            this.ribBtnOutLevel3.ShowImage = true;
            this.ribBtnOutLevel3.ShowLabel = false;
            this.ribBtnOutLevel3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel3_Click);
            // 
            // box3
            // 
            this.box3.Items.Add(this.ribBtnOutLevel4);
            this.box3.Items.Add(this.ribBtnOutLevel5);
            this.box3.Items.Add(this.ribBtnOutLevel6);
            this.box3.Name = "box3";
            // 
            // ribBtnOutLevel4
            // 
            this.ribBtnOutLevel4.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel4.Image")));
            this.ribBtnOutLevel4.Label = "4";
            this.ribBtnOutLevel4.Name = "ribBtnOutLevel4";
            this.ribBtnOutLevel4.ShowImage = true;
            this.ribBtnOutLevel4.ShowLabel = false;
            this.ribBtnOutLevel4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel4_Click);
            // 
            // ribBtnOutLevel5
            // 
            this.ribBtnOutLevel5.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel5.Image")));
            this.ribBtnOutLevel5.Label = "5";
            this.ribBtnOutLevel5.Name = "ribBtnOutLevel5";
            this.ribBtnOutLevel5.ShowImage = true;
            this.ribBtnOutLevel5.ShowLabel = false;
            this.ribBtnOutLevel5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel5_Click);
            // 
            // ribBtnOutLevel6
            // 
            this.ribBtnOutLevel6.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel6.Image")));
            this.ribBtnOutLevel6.Label = "6";
            this.ribBtnOutLevel6.Name = "ribBtnOutLevel6";
            this.ribBtnOutLevel6.ShowImage = true;
            this.ribBtnOutLevel6.ShowLabel = false;
            this.ribBtnOutLevel6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel6_Click);
            // 
            // box5
            // 
            this.box5.Items.Add(this.ribBtnOutLevel7);
            this.box5.Items.Add(this.ribBtnOutLevel8);
            this.box5.Items.Add(this.ribBtnOutLevel9);
            this.box5.Name = "box5";
            // 
            // ribBtnOutLevel7
            // 
            this.ribBtnOutLevel7.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel7.Image")));
            this.ribBtnOutLevel7.Label = "7";
            this.ribBtnOutLevel7.Name = "ribBtnOutLevel7";
            this.ribBtnOutLevel7.ShowImage = true;
            this.ribBtnOutLevel7.ShowLabel = false;
            this.ribBtnOutLevel7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel7_Click);
            // 
            // ribBtnOutLevel8
            // 
            this.ribBtnOutLevel8.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel8.Image")));
            this.ribBtnOutLevel8.Label = "8";
            this.ribBtnOutLevel8.Name = "ribBtnOutLevel8";
            this.ribBtnOutLevel8.ShowImage = true;
            this.ribBtnOutLevel8.ShowLabel = false;
            this.ribBtnOutLevel8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel8_Click);
            // 
            // ribBtnOutLevel9
            // 
            this.ribBtnOutLevel9.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevel9.Image")));
            this.ribBtnOutLevel9.Label = "9";
            this.ribBtnOutLevel9.Name = "ribBtnOutLevel9";
            this.ribBtnOutLevel9.ShowImage = true;
            this.ribBtnOutLevel9.ShowLabel = false;
            this.ribBtnOutLevel9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevel9_Click);
            // 
            // ribBtnOutLevelTextBody
            // 
            this.ribBtnOutLevelTextBody.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnOutLevelTextBody.Image")));
            this.ribBtnOutLevelTextBody.Label = "正";
            this.ribBtnOutLevelTextBody.Name = "ribBtnOutLevelTextBody";
            this.ribBtnOutLevelTextBody.ShowImage = true;
            this.ribBtnOutLevelTextBody.ShowLabel = false;
            this.ribBtnOutLevelTextBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnOutLevelTextBody_Click);
            // 
            // ribBtnViewOutlineLevel
            // 
            this.ribBtnViewOutlineLevel.Label = "查";
            this.ribBtnViewOutlineLevel.Name = "ribBtnViewOutlineLevel";
            this.ribBtnViewOutlineLevel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnViewOutlineLevel_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnOutlineSamePrev
            // 
            this.btnOutlineSamePrev.Image = ((System.Drawing.Image)(resources.GetObject("btnOutlineSamePrev.Image")));
            this.btnOutlineSamePrev.Label = "同前级";
            this.btnOutlineSamePrev.Name = "btnOutlineSamePrev";
            this.btnOutlineSamePrev.ScreenTip = "设置当前选择段落大纲级别与前面最近章节同级";
            this.btnOutlineSamePrev.ShowImage = true;
            this.btnOutlineSamePrev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOutlineSamePrev_Click);
            // 
            // btnOutlineLow1Prev
            // 
            this.btnOutlineLow1Prev.Image = ((System.Drawing.Image)(resources.GetObject("btnOutlineLow1Prev.Image")));
            this.btnOutlineLow1Prev.Label = "低前一级";
            this.btnOutlineLow1Prev.Name = "btnOutlineLow1Prev";
            this.btnOutlineLow1Prev.ScreenTip = "设置当前选择段落大纲级别为前面最近章节低一级";
            this.btnOutlineLow1Prev.ShowImage = true;
            this.btnOutlineLow1Prev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOutlineLow1Prev_Click);
            // 
            // btnOutlineHigh1Prev
            // 
            this.btnOutlineHigh1Prev.Image = ((System.Drawing.Image)(resources.GetObject("btnOutlineHigh1Prev.Image")));
            this.btnOutlineHigh1Prev.Label = "高前一级";
            this.btnOutlineHigh1Prev.Name = "btnOutlineHigh1Prev";
            this.btnOutlineHigh1Prev.ScreenTip = "设置当前选择段落大纲级别为前面最近章节高一级";
            this.btnOutlineHigh1Prev.ShowImage = true;
            this.btnOutlineHigh1Prev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOutlineHigh1Prev_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.btnOutlinePromote);
            this.box1.Items.Add(this.btnOutlineDemote);
            this.box1.Items.Add(this.chkOnlyNonTextBodyPara);
            this.box1.Name = "box1";
            // 
            // btnOutlinePromote
            // 
            this.btnOutlinePromote.Image = ((System.Drawing.Image)(resources.GetObject("btnOutlinePromote.Image")));
            this.btnOutlinePromote.Label = "批量升级";
            this.btnOutlinePromote.Name = "btnOutlinePromote";
            this.btnOutlinePromote.ScreenTip = "设置当前选择段落大纲级别增加一级（升级）";
            this.btnOutlinePromote.ShowImage = true;
            this.btnOutlinePromote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOutlinePromote_Click);
            // 
            // btnOutlineDemote
            // 
            this.btnOutlineDemote.Image = ((System.Drawing.Image)(resources.GetObject("btnOutlineDemote.Image")));
            this.btnOutlineDemote.Label = "批量降级";
            this.btnOutlineDemote.Name = "btnOutlineDemote";
            this.btnOutlineDemote.ScreenTip = "设置当前选择段落大纲级别降低一级（降级）";
            this.btnOutlineDemote.ShowImage = true;
            this.btnOutlineDemote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOutlineDemote_Click);
            // 
            // chkOnlyNonTextBodyPara
            // 
            this.chkOnlyNonTextBodyPara.Checked = true;
            this.chkOnlyNonTextBodyPara.Label = "排除正文";
            this.chkOnlyNonTextBodyPara.Name = "chkOnlyNonTextBodyPara";
            this.chkOnlyNonTextBodyPara.ScreenTip = "对批量升级/降级是否排除正文（只针对章节）或包括正文";
            // 
            // grpPane
            // 
            this.grpPane.Items.Add(this.btnCopyHeadingStyles);
            this.grpPane.Items.Add(this.btnPasteHeadingStyles);
            this.grpPane.Items.Add(this.ribbtnCopyHeadingsStructure);
            this.grpPane.Items.Add(this.separator4);
            this.grpPane.Items.Add(this.ribbtnSaveCurHeadingStyle2Style);
            this.grpPane.Items.Add(this.chkHeadingsStylesPersist);
            this.grpPane.Label = "章节样式";
            this.grpPane.Name = "grpPane";
            // 
            // btnCopyHeadingStyles
            // 
            this.btnCopyHeadingStyles.Image = ((System.Drawing.Image)(resources.GetObject("btnCopyHeadingStyles.Image")));
            this.btnCopyHeadingStyles.Label = "复制章节样式";
            this.btnCopyHeadingStyles.Name = "btnCopyHeadingStyles";
            this.btnCopyHeadingStyles.ScreenTip = "复制选择范围内或文档章节的样式";
            this.btnCopyHeadingStyles.ShowImage = true;
            this.btnCopyHeadingStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopyHeadingStyles_Click);
            // 
            // btnPasteHeadingStyles
            // 
            this.btnPasteHeadingStyles.Image = ((System.Drawing.Image)(resources.GetObject("btnPasteHeadingStyles.Image")));
            this.btnPasteHeadingStyles.Label = "粘贴章节样式";
            this.btnPasteHeadingStyles.Name = "btnPasteHeadingStyles";
            this.btnPasteHeadingStyles.ScreenTip = "将复制的章节样式粘贴应用到选择区或文档的章节";
            this.btnPasteHeadingStyles.ShowImage = true;
            this.btnPasteHeadingStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPasteHeadingStyles_Click);
            // 
            // ribbtnCopyHeadingsStructure
            // 
            this.ribbtnCopyHeadingsStructure.Image = ((System.Drawing.Image)(resources.GetObject("ribbtnCopyHeadingsStructure.Image")));
            this.ribbtnCopyHeadingsStructure.Label = "复制章节结构";
            this.ribbtnCopyHeadingsStructure.Name = "ribbtnCopyHeadingsStructure";
            this.ribbtnCopyHeadingsStructure.ScreenTip = "复制章节结构";
            this.ribbtnCopyHeadingsStructure.ShowImage = true;
            this.ribbtnCopyHeadingsStructure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbtnCopyHeadingsStructure_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // ribbtnSaveCurHeadingStyle2Style
            // 
            this.ribbtnSaveCurHeadingStyle2Style.Image = ((System.Drawing.Image)(resources.GetObject("ribbtnSaveCurHeadingStyle2Style.Image")));
            this.ribbtnSaveCurHeadingStyle2Style.Label = "保存章节样式";
            this.ribbtnSaveCurHeadingStyle2Style.Name = "ribbtnSaveCurHeadingStyle2Style";
            this.ribbtnSaveCurHeadingStyle2Style.ScreenTip = "将选择区或当前文档的章节样式保存到当前文档的样式表中";
            this.ribbtnSaveCurHeadingStyle2Style.ShowImage = true;
            this.ribbtnSaveCurHeadingStyle2Style.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbtnSaveCurHeadingStyle2Style_Click);
            // 
            // chkHeadingsStylesPersist
            // 
            this.chkHeadingsStylesPersist.Label = "保存到样式库";
            this.chkHeadingsStylesPersist.Name = "chkHeadingsStylesPersist";
            this.chkHeadingsStylesPersist.ScreenTip = "将选择区或当前文档的章节样式保存到Normal模板的样式表中";
            // 
            // grpQuickBookmark
            // 
            this.grpQuickBookmark.Items.Add(this.btnNavAddBkmk);
            this.grpQuickBookmark.Items.Add(this.ribBtnRemoveJetNav);
            this.grpQuickBookmark.Items.Add(this.btnClearBkmk);
            this.grpQuickBookmark.Items.Add(this.separator6);
            this.grpQuickBookmark.Items.Add(this.box6);
            this.grpQuickBookmark.Items.Add(this.box7);
            this.grpQuickBookmark.Items.Add(this.ribBtnJump2Toc);
            this.grpQuickBookmark.Items.Add(this.ribBtnPrevEditPos);
            this.grpQuickBookmark.Items.Add(this.ribBtnNextEditPos);
            this.grpQuickBookmark.Label = "快捷导航";
            this.grpQuickBookmark.Name = "grpQuickBookmark";
            // 
            // btnNavAddBkmk
            // 
            this.btnNavAddBkmk.Image = ((System.Drawing.Image)(resources.GetObject("btnNavAddBkmk.Image")));
            this.btnNavAddBkmk.Label = "添加";
            this.btnNavAddBkmk.Name = "btnNavAddBkmk";
            this.btnNavAddBkmk.ScreenTip = "在当前位置设置快捷书签";
            this.btnNavAddBkmk.ShowImage = true;
            this.btnNavAddBkmk.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavAddBkmk_Click);
            // 
            // ribBtnRemoveJetNav
            // 
            this.ribBtnRemoveJetNav.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnRemoveJetNav.Image")));
            this.ribBtnRemoveJetNav.Label = "删除";
            this.ribBtnRemoveJetNav.Name = "ribBtnRemoveJetNav";
            this.ribBtnRemoveJetNav.ScreenTip = "删除当前位置的快捷书签";
            this.ribBtnRemoveJetNav.ShowImage = true;
            this.ribBtnRemoveJetNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnRemoveJetNav_Click);
            // 
            // btnClearBkmk
            // 
            this.btnClearBkmk.Image = ((System.Drawing.Image)(resources.GetObject("btnClearBkmk.Image")));
            this.btnClearBkmk.Label = "清除";
            this.btnClearBkmk.Name = "btnClearBkmk";
            this.btnClearBkmk.ScreenTip = "清除所有快捷书签";
            this.btnClearBkmk.ShowImage = true;
            this.btnClearBkmk.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearBkmk_Click);
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // box6
            // 
            this.box6.Items.Add(this.btnNavFirst);
            this.box6.Items.Add(this.btnNavLast);
            this.box6.Name = "box6";
            // 
            // btnNavFirst
            // 
            this.btnNavFirst.Image = ((System.Drawing.Image)(resources.GetObject("btnNavFirst.Image")));
            this.btnNavFirst.Label = "|<";
            this.btnNavFirst.Name = "btnNavFirst";
            this.btnNavFirst.ScreenTip = "跳转到第一个快捷书签";
            this.btnNavFirst.ShowImage = true;
            this.btnNavFirst.ShowLabel = false;
            this.btnNavFirst.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavFirst_Click);
            // 
            // btnNavLast
            // 
            this.btnNavLast.Image = ((System.Drawing.Image)(resources.GetObject("btnNavLast.Image")));
            this.btnNavLast.Label = ">|";
            this.btnNavLast.Name = "btnNavLast";
            this.btnNavLast.ScreenTip = "跳转到最后一个快捷书签";
            this.btnNavLast.ShowImage = true;
            this.btnNavLast.ShowLabel = false;
            this.btnNavLast.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavLast_Click);
            // 
            // box7
            // 
            this.box7.Items.Add(this.btnNavPrev);
            this.box7.Items.Add(this.btnNavNext);
            this.box7.Name = "box7";
            // 
            // btnNavPrev
            // 
            this.btnNavPrev.Image = ((System.Drawing.Image)(resources.GetObject("btnNavPrev.Image")));
            this.btnNavPrev.Label = "<";
            this.btnNavPrev.Name = "btnNavPrev";
            this.btnNavPrev.ScreenTip = "跳转到前一个书签";
            this.btnNavPrev.ShowImage = true;
            this.btnNavPrev.ShowLabel = false;
            this.btnNavPrev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavPrev_Click);
            // 
            // btnNavNext
            // 
            this.btnNavNext.Image = ((System.Drawing.Image)(resources.GetObject("btnNavNext.Image")));
            this.btnNavNext.Label = ">";
            this.btnNavNext.Name = "btnNavNext";
            this.btnNavNext.ScreenTip = "跳转到下一个书签";
            this.btnNavNext.ShowImage = true;
            this.btnNavNext.ShowLabel = false;
            this.btnNavNext.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavNext_Click);
            // 
            // ribBtnJump2Toc
            // 
            this.ribBtnJump2Toc.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnJump2Toc.Image")));
            this.ribBtnJump2Toc.Label = "目录";
            this.ribBtnJump2Toc.Name = "ribBtnJump2Toc";
            this.ribBtnJump2Toc.ShowImage = true;
            this.ribBtnJump2Toc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnJump2Toc_Click);
            // 
            // ribBtnPrevEditPos
            // 
            this.ribBtnPrevEditPos.Label = "前编辑位置";
            this.ribBtnPrevEditPos.Name = "ribBtnPrevEditPos";
            this.ribBtnPrevEditPos.Visible = false;
            this.ribBtnPrevEditPos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnPrevEditPos_Click);
            // 
            // ribBtnNextEditPos
            // 
            this.ribBtnNextEditPos.Label = "后编辑位置";
            this.ribBtnNextEditPos.Name = "ribBtnNextEditPos";
            this.ribBtnNextEditPos.Visible = false;
            this.ribBtnNextEditPos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnNextEditPos_Click);
            // 
            // groupLocalVer
            // 
            this.groupLocalVer.Items.Add(this.btnLocalVerMileStone);
            this.groupLocalVer.Items.Add(this.ribbtnOpenVerDir);
            this.groupLocalVer.Items.Add(this.chkGenLocalVer);
            this.groupLocalVer.Label = "本地版本";
            this.groupLocalVer.Name = "groupLocalVer";
            // 
            // btnLocalVerMileStone
            // 
            this.btnLocalVerMileStone.Image = ((System.Drawing.Image)(resources.GetObject("btnLocalVerMileStone.Image")));
            this.btnLocalVerMileStone.Label = "保存关键版本";
            this.btnLocalVerMileStone.Name = "btnLocalVerMileStone";
            this.btnLocalVerMileStone.ScreenTip = "将当前文档的内容保存一份里程碑版本";
            this.btnLocalVerMileStone.ShowImage = true;
            this.btnLocalVerMileStone.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLocalVerMileStone_Click);
            // 
            // ribbtnOpenVerDir
            // 
            this.ribbtnOpenVerDir.Image = ((System.Drawing.Image)(resources.GetObject("ribbtnOpenVerDir.Image")));
            this.ribbtnOpenVerDir.Label = "打开版本目录";
            this.ribbtnOpenVerDir.Name = "ribbtnOpenVerDir";
            this.ribbtnOpenVerDir.ScreenTip = "打开本地版本文件存放的目录";
            this.ribbtnOpenVerDir.ShowImage = true;
            this.ribbtnOpenVerDir.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbtnOpenVerDir_Click);
            // 
            // chkGenLocalVer
            // 
            this.chkGenLocalVer.Label = "产生本地版本";
            this.chkGenLocalVer.Name = "chkGenLocalVer";
            this.chkGenLocalVer.ScreenTip = "勾选则在保存时对当前文档保存一个版本";
            this.chkGenLocalVer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkGenLocalVer_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnTogglePanePos);
            this.group2.Items.Add(this.toggleTaskWin);
            this.group2.Label = "工作区";
            this.group2.Name = "group2";
            // 
            // btnTogglePanePos
            // 
            this.btnTogglePanePos.Image = ((System.Drawing.Image)(resources.GetObject("btnTogglePanePos.Image")));
            this.btnTogglePanePos.Label = "左右切换";
            this.btnTogglePanePos.Name = "btnTogglePanePos";
            this.btnTogglePanePos.ScreenTip = "切换任务窗居左或居右";
            this.btnTogglePanePos.ShowImage = true;
            this.btnTogglePanePos.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTogglePanePos_Click);
            // 
            // toggleTaskWin
            // 
            this.toggleTaskWin.Image = ((System.Drawing.Image)(resources.GetObject("toggleTaskWin.Image")));
            this.toggleTaskWin.Label = "可见切换";
            this.toggleTaskWin.Name = "toggleTaskWin";
            this.toggleTaskWin.ScreenTip = "切换任务窗居左或居右";
            this.toggleTaskWin.ShowImage = true;
            this.toggleTaskWin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleTaskWin_Click);
            // 
            // grpFuncPages
            // 
            this.grpFuncPages.Items.Add(this.ribBtnCheckUpdate);
            this.grpFuncPages.Items.Add(this.chkAutoCheckUpdate);
            this.grpFuncPages.Label = "版本更新";
            this.grpFuncPages.Name = "grpFuncPages";
            this.grpFuncPages.Visible = false;
            // 
            // ribBtnCheckUpdate
            // 
            this.ribBtnCheckUpdate.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnCheckUpdate.Image")));
            this.ribBtnCheckUpdate.Label = "检查更新";
            this.ribBtnCheckUpdate.Name = "ribBtnCheckUpdate";
            this.ribBtnCheckUpdate.ScreenTip = "手动检查更新";
            this.ribBtnCheckUpdate.ShowImage = true;
            this.ribBtnCheckUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnCheckUpdate_Click);
            // 
            // chkAutoCheckUpdate
            // 
            this.chkAutoCheckUpdate.Label = "自动";
            this.chkAutoCheckUpdate.Name = "chkAutoCheckUpdate";
            this.chkAutoCheckUpdate.ScreenTip = "勾选后则启动时进行自动检查";
            this.chkAutoCheckUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkAutoCheckUpdate_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.ribBtnTutorial);
            this.group1.Items.Add(this.ribBtnHelp);
            this.group1.Items.Add(this.ribbtnAbout);
            this.group1.Items.Add(this.ribLoadSoloLic);
            this.group1.Items.Add(this.RibbtnRegister);
            this.group1.Label = "帮助";
            this.group1.Name = "group1";
            // 
            // ribBtnTutorial
            // 
            this.ribBtnTutorial.Label = "入门";
            this.ribBtnTutorial.Name = "ribBtnTutorial";
            this.ribBtnTutorial.ScreenTip = "入门帮助";
            this.ribBtnTutorial.SuperTip = "将打开PDF文档，请安装PDF阅读器";
            this.ribBtnTutorial.Visible = false;
            this.ribBtnTutorial.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnTutorial_Click);
            // 
            // ribBtnHelp
            // 
            this.ribBtnHelp.Image = ((System.Drawing.Image)(resources.GetObject("ribBtnHelp.Image")));
            this.ribBtnHelp.Label = "帮助";
            this.ribBtnHelp.Name = "ribBtnHelp";
            this.ribBtnHelp.ScreenTip = "详细帮助";
            this.ribBtnHelp.ShowImage = true;
            this.ribBtnHelp.SuperTip = "将打开PDF文档，请安装PDF阅读器";
            this.ribBtnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribBtnHelp_Click);
            // 
            // ribbtnAbout
            // 
            this.ribbtnAbout.Image = ((System.Drawing.Image)(resources.GetObject("ribbtnAbout.Image")));
            this.ribbtnAbout.Label = "关于";
            this.ribbtnAbout.Name = "ribbtnAbout";
            this.ribbtnAbout.ShowImage = true;
            this.ribbtnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribbtnAbout_Click);
            // 
            // ribLoadSoloLic
            // 
            this.ribLoadSoloLic.Image = ((System.Drawing.Image)(resources.GetObject("ribLoadSoloLic.Image")));
            this.ribLoadSoloLic.Label = "本机许可";
            this.ribLoadSoloLic.Name = "ribLoadSoloLic";
            this.ribLoadSoloLic.ScreenTip = "加载单机版许可";
            this.ribLoadSoloLic.ShowImage = true;
            this.ribLoadSoloLic.SuperTip = "用户由此可加载单机版许可文件";
            this.ribLoadSoloLic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ribLoadSoloLic_Click);
            // 
            // RibbtnRegister
            // 
            this.RibbtnRegister.Image = ((System.Drawing.Image)(resources.GetObject("RibbtnRegister.Image")));
            this.RibbtnRegister.Label = "注册";
            this.RibbtnRegister.Name = "RibbtnRegister";
            this.RibbtnRegister.ScreenTip = "个人版注册";
            this.RibbtnRegister.ShowImage = true;
            this.RibbtnRegister.SuperTip = "联网进行个人版许可注册";
            this.RibbtnRegister.Visible = false;
            this.RibbtnRegister.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RibbtnRegister_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Close += new System.EventHandler(this.Ribbon1_Close);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpConfig.ResumeLayout(false);
            this.grpConfig.PerformLayout();
            this.grpComOp.ResumeLayout(false);
            this.grpComOp.PerformLayout();
            this.box8.ResumeLayout(false);
            this.box8.PerformLayout();
            this.box9.ResumeLayout(false);
            this.box9.PerformLayout();
            this.grpAutoNumbering.ResumeLayout(false);
            this.grpAutoNumbering.PerformLayout();
            this.grpOutline.ResumeLayout(false);
            this.grpOutline.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box4.ResumeLayout(false);
            this.box4.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.box5.ResumeLayout(false);
            this.box5.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.grpPane.ResumeLayout(false);
            this.grpPane.PerformLayout();
            this.grpQuickBookmark.ResumeLayout(false);
            this.grpQuickBookmark.PerformLayout();
            this.box6.ResumeLayout(false);
            this.box6.PerformLayout();
            this.box7.ResumeLayout(false);
            this.box7.PerformLayout();
            this.groupLocalVer.ResumeLayout(false);
            this.groupLocalVer.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.grpFuncPages.ResumeLayout(false);
            this.grpFuncPages.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton toggleTaskWin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpQuickBookmark;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpComOp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpFuncPages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavAddBkmk;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavPrev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavNext;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddHeaderLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearBkmk;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpOutline;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutlineSamePrev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutlineLow1Prev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutlineHigh1Prev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutlinePromote;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutlineDemote;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAutoNumbering;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTogglePanePos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveHeaderLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStrictCenter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCenterAllPics;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkBoxUpdTblCntOnSaving;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkBoxUpdTblCntOnClose;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRibCenterTables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCopyHeadingStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPasteHeadingStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRibInsertSeparateTblContent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddFooterLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearFooterLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkOnlyNonTextBodyPara;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkAutoLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbtnSaveCurHeadingStyle2Style;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkHeadingsStylesPersist;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbtnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbtnCopyHeadingsStructure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnTutorial;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnCheckUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkAutoCheckUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbtnOpenCurDocDir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RibbtnRegister;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavFirst;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavLast;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupLocalVer;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkGenLocalVer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbtnOpenVerDir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnRemoveJetNav;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnFillSn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnFillSn2EndRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnFillSelection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevel9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnOutLevelTextBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box4;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box6;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box7;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box8;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box9;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLocalVerMileStone;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbtnUnitedHeaders;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribbtnUnitedFooters;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnPrevEditPos;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnNextEditPos;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator7;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblCurParaOutLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnJump2Toc;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rbBtnCalculate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribLoadSoloLic;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ribBtnViewOutlineLevel;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
