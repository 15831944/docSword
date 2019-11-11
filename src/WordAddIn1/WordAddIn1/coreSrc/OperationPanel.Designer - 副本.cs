using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Collections;

namespace OfficeAssist
{
    partial class OperationPanel
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OperationPanel));
            this.tabCtrl = new System.Windows.Forms.TabControl();
            this.tabPageRel = new System.Windows.Forms.TabPage();
            this.btnRelForceCompute = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.m_tvRel = new System.Windows.Forms.TreeView();
            this.btnFoundNext = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnFoundBack = new System.Windows.Forms.Button();
            this.txtRelKeyword = new System.Windows.Forms.TextBox();
            this.btnRelSearch = new System.Windows.Forms.Button();
            this.btnRefreshRels = new System.Windows.Forms.Button();
            this.btnRelAllTxtOut = new System.Windows.Forms.Button();
            this.btnMove = new System.Windows.Forms.Button();
            this.btnExpEditor = new System.Windows.Forms.Button();
            this.txtRelName = new System.Windows.Forms.TextBox();
            this.chboxOpRulesEnable = new System.Windows.Forms.CheckBox();
            this.txtRelContent = new System.Windows.Forms.TextBox();
            this.txtOpRules = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAddRel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnUpdateRel = new System.Windows.Forms.Button();
            this.btnInsertRel = new System.Windows.Forms.Button();
            this.btnJump2Rel = new System.Windows.Forms.Button();
            this.btnRemoveRel = new System.Windows.Forms.Button();
            this.label36 = new System.Windows.Forms.Label();
            this.label35 = new System.Windows.Forms.Label();
            this.tabPageCheck = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.progbarCheck = new System.Windows.Forms.ProgressBar();
            this.btnCheckSearchNext = new System.Windows.Forms.Button();
            this.btnCheckSearchPrev = new System.Windows.Forms.Button();
            this.btnCheck = new System.Windows.Forms.Button();
            this.btnCheckReset = new System.Windows.Forms.Button();
            this.tvCheckedItems = new System.Windows.Forms.TreeView();
            this.btnCheckSearch = new System.Windows.Forms.Button();
            this.btnCheckIgnore = new System.Windows.Forms.Button();
            this.txtCheckSearchKeyWord = new System.Windows.Forms.TextBox();
            this.tabPageOrganize = new System.Windows.Forms.TabPage();
            this.OrgProgressBar = new System.Windows.Forms.ProgressBar();
            this.chkOrgShowBody = new System.Windows.Forms.CheckBox();
            this.btnOrgCancelProtect = new System.Windows.Forms.Button();
            this.btnOrganProtect = new System.Windows.Forms.Button();
            this.btnOrganNext = new System.Windows.Forms.Button();
            this.btnOrganBack = new System.Windows.Forms.Button();
            this.btnOrganResetSearch = new System.Windows.Forms.Button();
            this.btnOrganSearch = new System.Windows.Forms.Button();
            this.txtOrganKeyWord = new System.Windows.Forms.TextBox();
            this.btnOrganizeRefresh = new System.Windows.Forms.Button();
            this.btnCollapseSel = new System.Windows.Forms.Button();
            this.btnExpandSelChild = new System.Windows.Forms.Button();
            this.btnSelAll = new System.Windows.Forms.Button();
            this.btnSelClear = new System.Windows.Forms.Button();
            this.btnOrgDemote = new System.Windows.Forms.Button();
            this.btnOrgPromote = new System.Windows.Forms.Button();
            this.chkSelCategory = new System.Windows.Forms.CheckedListBox();
            this.m_tvOrganize = new System.Windows.Forms.TreeView();
            this.imageListIcon = new System.Windows.Forms.ImageList(this.components);
            this.label37 = new System.Windows.Forms.Label();
            this.tabPageShare = new System.Windows.Forms.TabPage();
            this.btnShareLibUpdate = new System.Windows.Forms.Button();
            this.btnShareOpen = new System.Windows.Forms.Button();
            this.prgBarLib = new System.Windows.Forms.ProgressBar();
            this.tvShareLib = new System.Windows.Forms.TreeView();
            this.cxtMenuSvr = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuItemApplyStyle = new System.Windows.Forms.ToolStripMenuItem();
            this.menuItemCntReuse = new System.Windows.Forms.ToolStripMenuItem();
            this.btnShareRef = new System.Windows.Forms.Button();
            this.btnShareRefresh = new System.Windows.Forms.Button();
            this.btnShareExpand = new System.Windows.Forms.Button();
            this.btnShareDownload = new System.Windows.Forms.Button();
            this.btnShareCollapse = new System.Windows.Forms.Button();
            this.btnShareSearch = new System.Windows.Forms.Button();
            this.btnShareSearchReset = new System.Windows.Forms.Button();
            this.txtShareKeyWord = new System.Windows.Forms.TextBox();
            this.btnSharePrevSearch = new System.Windows.Forms.Button();
            this.btnShareNextSearch = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.btnShareExternalFile = new System.Windows.Forms.Button();
            this.txtShareName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnShareAdd = new System.Windows.Forms.Button();
            this.btnShareRemove = new System.Windows.Forms.Button();
            this.chkBoxCommonLib = new System.Windows.Forms.CheckBox();
            this.chkBoxCategory = new System.Windows.Forms.CheckBox();
            this.label38 = new System.Windows.Forms.Label();
            this.label39 = new System.Windows.Forms.Label();
            this.tabPageUnitedStyle = new System.Windows.Forms.TabPage();
            this.rchTextBoxUniformStylesPreview = new System.Windows.Forms.RichTextBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.chkIgnoreParaFormat = new System.Windows.Forms.CheckBox();
            this.chkIgnoreFont = new System.Windows.Forms.CheckBox();
            this.chkIgnoreTextBody = new System.Windows.Forms.CheckBox();
            this.chkIgnoreHeadings = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtIgnorePages = new System.Windows.Forms.TextBox();
            this.chkIgnorePages = new System.Windows.Forms.CheckBox();
            this.chkIgnoreTable = new System.Windows.Forms.CheckBox();
            this.chkIgnoreTOC = new System.Windows.Forms.CheckBox();
            this.radioBtnStyleSelection = new System.Windows.Forms.RadioButton();
            this.radioBtnStyleAllDoc = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnUnitFormExitApply = new System.Windows.Forms.Button();
            this.lstUnitedStyleHistoryDoc = new System.Windows.Forms.ListBox();
            this.tblUniformStyleHistoryDocsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.localdbDataSet = new OfficeAssist.localdbDataSet();
            this.btnStyleApply = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.progressBarStyle = new System.Windows.Forms.ProgressBar();
            this.txtBoxStyleFile = new System.Windows.Forms.TextBox();
            this.btnStyleOpenFile = new System.Windows.Forms.Button();
            this.tabPageCompare = new System.Windows.Forms.TabPage();
            this.progBarComp = new System.Windows.Forms.ProgressBar();
            this.btnExecCompare = new System.Windows.Forms.Button();
            this.txtCompResult = new System.Windows.Forms.TextBox();
            this.tvCompCheck = new System.Windows.Forms.TreeView();
            this.tvCompStandard = new System.Windows.Forms.TreeView();
            this.chkCompStrickOrder = new System.Windows.Forms.CheckBox();
            this.chkCompOutline = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btnCompCheckDoc = new System.Windows.Forms.Button();
            this.txtComp2CheckDoc = new System.Windows.Forms.TextBox();
            this.btnCompStandardDoc = new System.Windows.Forms.Button();
            this.txtCompStandardDoc = new System.Windows.Forms.TextBox();
            this.tabPageDataTrans = new System.Windows.Forms.TabPage();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageDocTbls2Excel = new System.Windows.Forms.TabPage();
            this.btnW2XNextSameStructTbl = new System.Windows.Forms.Button();
            this.btnAddTagCol = new System.Windows.Forms.Button();
            this.grpW2XAutoModelScope = new System.Windows.Forms.GroupBox();
            this.chkBoxStrictVerifyTblColName = new System.Windows.Forms.CheckBox();
            this.rdBtnW2XSelScope = new System.Windows.Forms.RadioButton();
            this.rdBtnW2XAllDocScope = new System.Windows.Forms.RadioButton();
            this.trvDataDocTbl2Excel = new System.Windows.Forms.TreeView();
            this.btnClearItems = new System.Windows.Forms.Button();
            this.btnAllProduce = new System.Windows.Forms.Button();
            this.btnPreviewProduce = new System.Windows.Forms.Button();
            this.btnDocTbl2ExcelRemove = new System.Windows.Forms.Button();
            this.btnAddColValue = new System.Windows.Forms.Button();
            this.btnAddColName = new System.Windows.Forms.Button();
            this.tabPageExcel2DocTbl = new System.Windows.Forms.TabPage();
            this.btnDataProduce = new System.Windows.Forms.Button();
            this.btnDataPreviewOne = new System.Windows.Forms.Button();
            this.btnDataTagCombine = new System.Windows.Forms.Button();
            this.btnDataRemoveRel = new System.Windows.Forms.Button();
            this.btnDataInsertData = new System.Windows.Forms.Button();
            this.btnDataInsertName = new System.Windows.Forms.Button();
            this.btnDataDSource = new System.Windows.Forms.Button();
            this.trvData = new System.Windows.Forms.TreeView();
            this.label12 = new System.Windows.Forms.Label();
            this.tabPageFillGather = new System.Windows.Forms.TabPage();
            this.btnFillGatherMoveDown = new System.Windows.Forms.Button();
            this.btnFillGatherMoveUp = new System.Windows.Forms.Button();
            this.btnFillGatherShowRowCol = new System.Windows.Forms.Button();
            this.progBarFG = new System.Windows.Forms.ProgressBar();
            this.btnFillGatherAddUserDefineColName = new System.Windows.Forms.Button();
            this.btnFillGatherAllSelUnSel = new System.Windows.Forms.Button();
            this.btnFillGatherDelFiles = new System.Windows.Forms.Button();
            this.btnFillGatherAddFiles = new System.Windows.Forms.Button();
            this.chkBoxFillGatherStrictMatchColName = new System.Windows.Forms.CheckBox();
            this.btnFillGatherProduce = new System.Windows.Forms.Button();
            this.btnFillGatherPreviewProduce = new System.Windows.Forms.Button();
            this.rdBtnFillGatherCurDoc = new System.Windows.Forms.RadioButton();
            this.rdBtnFillGatherMultiFiles = new System.Windows.Forms.RadioButton();
            this.btnFillGatherViewLog = new System.Windows.Forms.Button();
            this.label42 = new System.Windows.Forms.Label();
            this.chkListBoxTargetFiles = new System.Windows.Forms.CheckedListBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.rdBtnFillGatherSelScope = new System.Windows.Forms.RadioButton();
            this.rdBtnFillGatherAllDocScope = new System.Windows.Forms.RadioButton();
            this.trvFillGatherSchemes = new System.Windows.Forms.TreeView();
            this.txtFillGatherName = new System.Windows.Forms.TextBox();
            this.label43 = new System.Windows.Forms.Label();
            this.btnFillGatherVerifyMatch = new System.Windows.Forms.Button();
            this.btnFillGatherRemoveTblItem = new System.Windows.Forms.Button();
            this.btnFillGatherAddTagNameValue = new System.Windows.Forms.Button();
            this.btnFillGatherAddColValue = new System.Windows.Forms.Button();
            this.btnFillGatherAddColName = new System.Windows.Forms.Button();
            this.btnFillGatherAddTable = new System.Windows.Forms.Button();
            this.btnFillGatherAddScheme = new System.Windows.Forms.Button();
            this.label44 = new System.Windows.Forms.Label();
            this.tabPageCntList = new System.Windows.Forms.TabPage();
            this.btnCntListExpand = new System.Windows.Forms.Button();
            this.btnCntListCollapse = new System.Windows.Forms.Button();
            this.progBarCntList = new System.Windows.Forms.ProgressBar();
            this.trvCntList = new System.Windows.Forms.TreeView();
            this.btnCntListCover = new System.Windows.Forms.Button();
            this.btnCntListRef = new System.Windows.Forms.Button();
            this.btnCntListRemove = new System.Windows.Forms.Button();
            this.btnCntListAddDoc = new System.Windows.Forms.Button();
            this.txtBoxCntListFile = new System.Windows.Forms.TextBox();
            this.tabPageForm = new System.Windows.Forms.TabPage();
            this.label14 = new System.Windows.Forms.Label();
            this.btnFormNextSearch = new System.Windows.Forms.Button();
            this.tblFormLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.btnFormRefresh = new System.Windows.Forms.Button();
            this.btnFormPrevSearch = new System.Windows.Forms.Button();
            this.txtFormKeyWord = new System.Windows.Forms.TextBox();
            this.btnFormSearch = new System.Windows.Forms.Button();
            this.btnFormReset = new System.Windows.Forms.Button();
            this.tabPageInfo = new System.Windows.Forms.TabPage();
            this.btnInfoRefresh = new System.Windows.Forms.Button();
            this.txtInfoBody = new System.Windows.Forms.TextBox();
            this.tabPageNumTrans = new System.Windows.Forms.TabPage();
            this.btnNumTransClear = new System.Windows.Forms.Button();
            this.btnNumTrans = new System.Windows.Forms.Button();
            this.txtMoneySimpBigTbl = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.txtMoneySimpBig = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.txtNumValueSimpBigTbl = new System.Windows.Forms.TextBox();
            this.label29 = new System.Windows.Forms.Label();
            this.txtNumValueSimpBig = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.txtDigitNumSimpBig = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.txtMoneySimpLittleTbl = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.txtMoneySimpLittle = new System.Windows.Forms.TextBox();
            this.txtNumValueSimpLittleTbl = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.txtNumValueSimpLittle = new System.Windows.Forms.TextBox();
            this.txtNumMoney = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtNumValue = new System.Windows.Forms.TextBox();
            this.txtDigitNumSimpLittle = new System.Windows.Forms.TextBox();
            this.txtDigitNum = new System.Windows.Forms.TextBox();
            this.tabPageHeadingSn = new System.Windows.Forms.TabPage();
            this.label33 = new System.Windows.Forms.Label();
            this.btnExitHeadingSnApply = new System.Windows.Forms.Button();
            this.btnHeadingSnPreview = new System.Windows.Forms.Button();
            this.btnHeadingSnReset = new System.Windows.Forms.Button();
            this.chkHeadingSnReserveCurStyle = new System.Windows.Forms.CheckBox();
            this.progBarHeadingSn = new System.Windows.Forms.ProgressBar();
            this.btnHeadingSnNameGen = new System.Windows.Forms.Button();
            this.trvHeadingSnScheme = new System.Windows.Forms.TreeView();
            this.cxtMenuHeadingSn = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.cxtMenuItemPreview = new System.Windows.Forms.ToolStripMenuItem();
            this.btnHeadingSnSchemeApply = new System.Windows.Forms.Button();
            this.btnHeadingSnSchemeGet = new System.Windows.Forms.Button();
            this.btnHeadingSnSchemeUpdate = new System.Windows.Forms.Button();
            this.btnHeadingSnSchemeDel = new System.Windows.Forms.Button();
            this.btnHeadingSnSchemeAdd = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnHeadingSnFont = new System.Windows.Forms.Button();
            this.richTxtHeadingSnPreview = new System.Windows.Forms.RichTextBox();
            this.btnHeadingSnSetDefaultFont = new System.Windows.Forms.Button();
            this.btnHeadingSnFontExtract = new System.Windows.Forms.Button();
            this.btnHeadingSnPos = new System.Windows.Forms.Button();
            this.lstHeadingSnLevel = new System.Windows.Forms.ListBox();
            this.chkHeadingSnLegal = new System.Windows.Forms.CheckBox();
            this.cmbSnShowStyle = new System.Windows.Forms.ComboBox();
            this.label32 = new System.Windows.Forms.Label();
            this.txtSnDefInput = new System.Windows.Forms.TextBox();
            this.label31 = new System.Windows.Forms.Label();
            this.label34 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.txtHeadingSnSchemeName = new System.Windows.Forms.TextBox();
            this.label40 = new System.Windows.Forms.Label();
            this.tabPageHeadingStyles = new System.Windows.Forms.TabPage();
            this.btnHeadingStyleApplyCurSel = new System.Windows.Forms.Button();
            this.btnHeadingStyleExitApply = new System.Windows.Forms.Button();
            this.btnHeadingStyleApplyScope = new System.Windows.Forms.Button();
            this.richHeadingStylePreview = new System.Windows.Forms.RichTextBox();
            this.lstOutlineLevel = new System.Windows.Forms.ListBox();
            this.prgbarHeadingStyleSchemeApply = new System.Windows.Forms.ProgressBar();
            this.txtHeadingStyleSchemeName = new System.Windows.Forms.TextBox();
            this.btnHeadingStyleSchemeApply = new System.Windows.Forms.Button();
            this.btnHeadingStyleSchemeExtract = new System.Windows.Forms.Button();
            this.btnHeadingStyleSchemePreview = new System.Windows.Forms.Button();
            this.btnHeadingStyleSchemeUpdate = new System.Windows.Forms.Button();
            this.btnHeadingStyleSchemeDel = new System.Windows.Forms.Button();
            this.btnHeadingStyleSchemeAdd = new System.Windows.Forms.Button();
            this.trvHeadingStyleScheme = new System.Windows.Forms.TreeView();
            this.label41 = new System.Windows.Forms.Label();
            this.tabPageObjNav = new System.Windows.Forms.TabPage();
            this.groupBox15 = new System.Windows.Forms.GroupBox();
            this.btnONEquationNavLast = new System.Windows.Forms.Button();
            this.btnONObjectNavLast = new System.Windows.Forms.Button();
            this.btnONBookmarkNavLast = new System.Windows.Forms.Button();
            this.btnONEndnoteNavLast = new System.Windows.Forms.Button();
            this.btnONFootnoteNavLast = new System.Windows.Forms.Button();
            this.btnONCommentNavLast = new System.Windows.Forms.Button();
            this.btnONEquationNavPrev = new System.Windows.Forms.Button();
            this.btnONObjectNavPrev = new System.Windows.Forms.Button();
            this.btnONBookmarkNavPrev = new System.Windows.Forms.Button();
            this.btnONEndnoteNavPrev = new System.Windows.Forms.Button();
            this.btnONFootnoteNavPrev = new System.Windows.Forms.Button();
            this.btnONCommentNavPrev = new System.Windows.Forms.Button();
            this.btnONEquationNavFirst = new System.Windows.Forms.Button();
            this.btnONEquationNavNext = new System.Windows.Forms.Button();
            this.btnONObjectNavFirst = new System.Windows.Forms.Button();
            this.btnONObjectNavNext = new System.Windows.Forms.Button();
            this.btnONBookmarkNavFirst = new System.Windows.Forms.Button();
            this.btnONBookmarkNavNext = new System.Windows.Forms.Button();
            this.btnONEndnoteNavFirst = new System.Windows.Forms.Button();
            this.label79 = new System.Windows.Forms.Label();
            this.btnONEndnoteNavNext = new System.Windows.Forms.Button();
            this.label78 = new System.Windows.Forms.Label();
            this.btnONFootnoteNavFirst = new System.Windows.Forms.Button();
            this.label77 = new System.Windows.Forms.Label();
            this.btnONFootnoteNavNext = new System.Windows.Forms.Button();
            this.label69 = new System.Windows.Forms.Label();
            this.btnONCommentNavFirst = new System.Windows.Forms.Button();
            this.label68 = new System.Windows.Forms.Label();
            this.btnONCommentNavNext = new System.Windows.Forms.Button();
            this.label66 = new System.Windows.Forms.Label();
            this.groupBox14 = new System.Windows.Forms.GroupBox();
            this.colorComboBoxNav = new OfficeAssist.ColorComboBox();
            this.btnHighLightGoLast = new System.Windows.Forms.Button();
            this.btnONFieldNavLast = new System.Windows.Forms.Button();
            this.label75 = new System.Windows.Forms.Label();
            this.label67 = new System.Windows.Forms.Label();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.chkListObjNavOutline = new System.Windows.Forms.CheckedListBox();
            this.btnONOutlineAllUnSel = new System.Windows.Forms.Button();
            this.label71 = new System.Windows.Forms.Label();
            this.btnONHeadingNavFirst = new System.Windows.Forms.Button();
            this.btnONOutlineAllSel = new System.Windows.Forms.Button();
            this.btnONHeadingNavPrev = new System.Windows.Forms.Button();
            this.btnONHeadingNavLast = new System.Windows.Forms.Button();
            this.btnONHeadingNavNext = new System.Windows.Forms.Button();
            this.label72 = new System.Windows.Forms.Label();
            this.label70 = new System.Windows.Forms.Label();
            this.btnONSectionNavLast = new System.Windows.Forms.Button();
            this.label76 = new System.Windows.Forms.Label();
            this.btnONPageNavLast = new System.Windows.Forms.Button();
            this.label64 = new System.Windows.Forms.Label();
            this.btnHighLightGoNext = new System.Windows.Forms.Button();
            this.btnONFieldNavNext = new System.Windows.Forms.Button();
            this.btnONGraphicNavLast = new System.Windows.Forms.Button();
            this.btnONSectionNavNext = new System.Windows.Forms.Button();
            this.label63 = new System.Windows.Forms.Label();
            this.btnONPageNavNext = new System.Windows.Forms.Button();
            this.btnHighLightGoFirst = new System.Windows.Forms.Button();
            this.btnONFieldNavFirst = new System.Windows.Forms.Button();
            this.btnONTableNavLast = new System.Windows.Forms.Button();
            this.btnONSectionNavFirst = new System.Windows.Forms.Button();
            this.btnONGraphicNavNext = new System.Windows.Forms.Button();
            this.btnONPageNavFirst = new System.Windows.Forms.Button();
            this.btnHighLightGoPrev = new System.Windows.Forms.Button();
            this.btnONFieldNavPrev = new System.Windows.Forms.Button();
            this.label65 = new System.Windows.Forms.Label();
            this.btnONSectionNavPrev = new System.Windows.Forms.Button();
            this.btnONGraphicNavFirst = new System.Windows.Forms.Button();
            this.btnONPageNavPrev = new System.Windows.Forms.Button();
            this.btnONTableNavNext = new System.Windows.Forms.Button();
            this.btnONGraphicNavPrev = new System.Windows.Forms.Button();
            this.btnONTableNavFirst = new System.Windows.Forms.Button();
            this.btnONTableNavPrev = new System.Windows.Forms.Button();
            this.tabPageMultiSel = new System.Windows.Forms.TabPage();
            this.ExcludeColorComboBox = new OfficeAssist.ColorComboBox();
            this.IncludeColorComboBox = new OfficeAssist.ColorComboBox();
            this.label73 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.chkMultiSelUserDef = new System.Windows.Forms.CheckBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.chkListBoxMultiListSnType = new System.Windows.Forms.CheckedListBox();
            this.groupBox16 = new System.Windows.Forms.GroupBox();
            this.rdBtnMultiSelIgnoreTbls = new System.Windows.Forms.RadioButton();
            this.rdBtnMultiSelOnlyTbls = new System.Windows.Forms.RadioButton();
            this.rdBtnMultiSelIncludeTbls = new System.Windows.Forms.RadioButton();
            this.chkBoxMultiSelIgnoreHeadings = new System.Windows.Forms.CheckBox();
            this.chkBoxMultiSelIgnoreTxtBody = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelSnParas = new System.Windows.Forms.CheckBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.chkWholeTableCells = new System.Windows.Forms.CheckBox();
            this.chkBoxMultiSelLastColumn = new System.Windows.Forms.CheckBox();
            this.chkBoxMulSelTblLastRow = new System.Windows.Forms.CheckBox();
            this.numMultiSelColEnd = new System.Windows.Forms.NumericUpDown();
            this.numMultiSelColStart = new System.Windows.Forms.NumericUpDown();
            this.numMultiSelRowEnd = new System.Windows.Forms.NumericUpDown();
            this.numMultiSelRowStart = new System.Windows.Forms.NumericUpDown();
            this.label46 = new System.Windows.Forms.Label();
            this.label45 = new System.Windows.Forms.Label();
            this.chkBoxMultiSelColumnsScope = new System.Windows.Forms.CheckBox();
            this.chkBoxMultiSelRowsScope = new System.Windows.Forms.CheckBox();
            this.chkBoxMultiSelFirstColumn = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelTables = new System.Windows.Forms.CheckBox();
            this.chkBoxMulSelTblFirstRow = new System.Windows.Forms.CheckBox();
            this.btnMultiSelApplySel = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.rdBtnAfterCurSel = new System.Windows.Forms.RadioButton();
            this.rdBtnBeforeCurSel = new System.Windows.Forms.RadioButton();
            this.radioBtnMultiSelCurSelScope = new System.Windows.Forms.RadioButton();
            this.radioBtnMultiSelWholeStory = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.checkBoxMultiSelHighlight = new System.Windows.Forms.CheckBox();
            this.colorComboBoxHighlight = new OfficeAssist.ColorComboBox();
            this.rdBtnMultiSelObjectPara = new System.Windows.Forms.RadioButton();
            this.rdBtnMultiSelObjectRng = new System.Windows.Forms.RadioButton();
            this.checkBoxMultiSelIndices = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelInlineShapes = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelFields = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelBookMarks = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelCnts = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelComments = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelEndNotes = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelFootNotes = new System.Windows.Forms.CheckBox();
            this.label47 = new System.Windows.Forms.Label();
            this.checkBoxMultiHyperLinks = new System.Windows.Forms.CheckBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.btnMultiSelHeadingAllClear = new System.Windows.Forms.Button();
            this.btnMultiSelHeadingAllSel = new System.Windows.Forms.Button();
            this.checkedListBoxMultiSelHeading = new System.Windows.Forms.CheckedListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBoxMultiSelIgnoreToc = new System.Windows.Forms.CheckBox();
            this.checkBoxMultiSelIgnoreTbl = new System.Windows.Forms.CheckBox();
            this.tabPageMultiTiZhu = new System.Windows.Forms.TabPage();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.label62 = new System.Windows.Forms.Label();
            this.label56 = new System.Windows.Forms.Label();
            this.label52 = new System.Windows.Forms.Label();
            this.btnNavLastField = new System.Windows.Forms.Button();
            this.btnNavNextField = new System.Windows.Forms.Button();
            this.btnNav2LastInShp = new System.Windows.Forms.Button();
            this.btnNav2NextInShp = new System.Windows.Forms.Button();
            this.btnNavPrevField = new System.Windows.Forms.Button();
            this.btnNav2LastTbl = new System.Windows.Forms.Button();
            this.btnNav2PrevInShp = new System.Windows.Forms.Button();
            this.btnNavFirstField = new System.Windows.Forms.Button();
            this.btnNav2NextTbl = new System.Windows.Forms.Button();
            this.btnNav2FirstInShp = new System.Windows.Forms.Button();
            this.btnNav2PrevTbl = new System.Windows.Forms.Button();
            this.btnNav2FirstTbl = new System.Windows.Forms.Button();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.chkInShpCaplblGetFromHeading = new System.Windows.Forms.CheckBox();
            this.chkTblCaplblGetFromHeading = new System.Windows.Forms.CheckBox();
            this.chkSyncUpdateTableOfFigures = new System.Windows.Forms.CheckBox();
            this.txtInShpCapLblPreFix = new System.Windows.Forms.TextBox();
            this.txtInShpCapLblPostFix = new System.Windows.Forms.TextBox();
            this.txtTblCapLblPreFix = new System.Windows.Forms.TextBox();
            this.txtTblCapLblPostFix = new System.Windows.Forms.TextBox();
            this.rdBtnTiZhuAfterCurPos = new System.Windows.Forms.RadioButton();
            this.rdBtnTiZhuBeforeCurPos = new System.Windows.Forms.RadioButton();
            this.rdCapLblScopeSelection = new System.Windows.Forms.RadioButton();
            this.rdCapLblScopeAllDoc = new System.Windows.Forms.RadioButton();
            this.cmbInShpCapLblAlign = new System.Windows.Forms.ComboBox();
            this.cmbInShpCapLblPos = new System.Windows.Forms.ComboBox();
            this.cmbTblCapLblPos = new System.Windows.Forms.ComboBox();
            this.cmbTblCapLblAlign = new System.Windows.Forms.ComboBox();
            this.btnApplyCapLbls2CurDoc = new System.Windows.Forms.Button();
            this.label51 = new System.Windows.Forms.Label();
            this.label55 = new System.Windows.Forms.Label();
            this.label53 = new System.Windows.Forms.Label();
            this.label54 = new System.Windows.Forms.Label();
            this.btnAddSelInShpCapLbl = new System.Windows.Forms.Button();
            this.txtSelectedInShpCapLbl = new System.Windows.Forms.TextBox();
            this.label50 = new System.Windows.Forms.Label();
            this.lstBoxCurSysCapLbls = new System.Windows.Forms.ListBox();
            this.btnRefreshCapsLbl = new System.Windows.Forms.Button();
            this.btnRemoveSelInShpCapLbl = new System.Windows.Forms.Button();
            this.btnSetSysCapLbls = new System.Windows.Forms.Button();
            this.txtSelectedTblCapLbl = new System.Windows.Forms.TextBox();
            this.btnRemoveSelTblCapLbl = new System.Windows.Forms.Button();
            this.label49 = new System.Windows.Forms.Label();
            this.label82 = new System.Windows.Forms.Label();
            this.label81 = new System.Windows.Forms.Label();
            this.label61 = new System.Windows.Forms.Label();
            this.btnAddSelTblCapLbl = new System.Windows.Forms.Button();
            this.label58 = new System.Windows.Forms.Label();
            this.label59 = new System.Windows.Forms.Label();
            this.label57 = new System.Windows.Forms.Label();
            this.label60 = new System.Windows.Forms.Label();
            this.label48 = new System.Windows.Forms.Label();
            this.tabPageTEST = new System.Windows.Forms.TabPage();
            this.button12 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button17 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btn4Test = new System.Windows.Forms.Button();
            this.tblUniformStyleHistoryDocsTableAdapter = new OfficeAssist.localdbDataSetTableAdapters.tblUniformStyleHistoryDocsTableAdapter();
            this.tblUniformStyleHistoryDocsBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.tblUniformStyleHistoryDocsBindingSource2 = new System.Windows.Forms.BindingSource(this.components);
            this.tabCtrl.SuspendLayout();
            this.tabPageRel.SuspendLayout();
            this.tabPageCheck.SuspendLayout();
            this.tabPageOrganize.SuspendLayout();
            this.tabPageShare.SuspendLayout();
            this.cxtMenuSvr.SuspendLayout();
            this.tabPageUnitedStyle.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblUniformStyleHistoryDocsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.localdbDataSet)).BeginInit();
            this.tabPageCompare.SuspendLayout();
            this.tabPageDataTrans.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPageDocTbls2Excel.SuspendLayout();
            this.grpW2XAutoModelScope.SuspendLayout();
            this.tabPageExcel2DocTbl.SuspendLayout();
            this.tabPageFillGather.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.tabPageCntList.SuspendLayout();
            this.tabPageForm.SuspendLayout();
            this.tabPageInfo.SuspendLayout();
            this.tabPageNumTrans.SuspendLayout();
            this.tabPageHeadingSn.SuspendLayout();
            this.cxtMenuHeadingSn.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPageHeadingStyles.SuspendLayout();
            this.tabPageObjNav.SuspendLayout();
            this.groupBox15.SuspendLayout();
            this.groupBox14.SuspendLayout();
            this.groupBox13.SuspendLayout();
            this.tabPageMultiSel.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox16.SuspendLayout();
            this.groupBox9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelColEnd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelColStart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelRowEnd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelRowStart)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPageMultiTiZhu.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.tabPageTEST.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblUniformStyleHistoryDocsBindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblUniformStyleHistoryDocsBindingSource2)).BeginInit();
            this.SuspendLayout();
            // 
            // tabCtrl
            // 
            this.tabCtrl.Controls.Add(this.tabPageRel);
            this.tabCtrl.Controls.Add(this.tabPageCheck);
            this.tabCtrl.Controls.Add(this.tabPageOrganize);
            this.tabCtrl.Controls.Add(this.tabPageShare);
            this.tabCtrl.Controls.Add(this.tabPageUnitedStyle);
            this.tabCtrl.Controls.Add(this.tabPageCompare);
            this.tabCtrl.Controls.Add(this.tabPageDataTrans);
            this.tabCtrl.Controls.Add(this.tabPageFillGather);
            this.tabCtrl.Controls.Add(this.tabPageCntList);
            this.tabCtrl.Controls.Add(this.tabPageForm);
            this.tabCtrl.Controls.Add(this.tabPageInfo);
            this.tabCtrl.Controls.Add(this.tabPageNumTrans);
            this.tabCtrl.Controls.Add(this.tabPageHeadingSn);
            this.tabCtrl.Controls.Add(this.tabPageHeadingStyles);
            this.tabCtrl.Controls.Add(this.tabPageObjNav);
            this.tabCtrl.Controls.Add(this.tabPageMultiSel);
            this.tabCtrl.Controls.Add(this.tabPageMultiTiZhu);
            this.tabCtrl.Controls.Add(this.tabPageTEST);
            resources.ApplyResources(this.tabCtrl, "tabCtrl");
            this.tabCtrl.Multiline = true;
            this.tabCtrl.Name = "tabCtrl";
            this.tabCtrl.SelectedIndex = 0;
            // 
            // tabPageRel
            // 
            resources.ApplyResources(this.tabPageRel, "tabPageRel");
            this.tabPageRel.Controls.Add(this.btnRelForceCompute);
            this.tabPageRel.Controls.Add(this.label1);
            this.tabPageRel.Controls.Add(this.m_tvRel);
            this.tabPageRel.Controls.Add(this.btnFoundNext);
            this.tabPageRel.Controls.Add(this.btnReset);
            this.tabPageRel.Controls.Add(this.btnFoundBack);
            this.tabPageRel.Controls.Add(this.txtRelKeyword);
            this.tabPageRel.Controls.Add(this.btnRelSearch);
            this.tabPageRel.Controls.Add(this.btnRefreshRels);
            this.tabPageRel.Controls.Add(this.btnRelAllTxtOut);
            this.tabPageRel.Controls.Add(this.btnMove);
            this.tabPageRel.Controls.Add(this.btnExpEditor);
            this.tabPageRel.Controls.Add(this.txtRelName);
            this.tabPageRel.Controls.Add(this.chboxOpRulesEnable);
            this.tabPageRel.Controls.Add(this.txtRelContent);
            this.tabPageRel.Controls.Add(this.txtOpRules);
            this.tabPageRel.Controls.Add(this.label2);
            this.tabPageRel.Controls.Add(this.btnAddRel);
            this.tabPageRel.Controls.Add(this.label3);
            this.tabPageRel.Controls.Add(this.btnUpdateRel);
            this.tabPageRel.Controls.Add(this.btnInsertRel);
            this.tabPageRel.Controls.Add(this.btnJump2Rel);
            this.tabPageRel.Controls.Add(this.btnRemoveRel);
            this.tabPageRel.Controls.Add(this.label36);
            this.tabPageRel.Controls.Add(this.label35);
            this.tabPageRel.Name = "tabPageRel";
            this.tabPageRel.UseVisualStyleBackColor = true;
            // 
            // btnRelForceCompute
            // 
            resources.ApplyResources(this.btnRelForceCompute, "btnRelForceCompute");
            this.btnRelForceCompute.Name = "btnRelForceCompute";
            this.btnRelForceCompute.UseVisualStyleBackColor = true;
            this.btnRelForceCompute.Click += new System.EventHandler(this.btnRelForceCompute_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // m_tvRel
            // 
            this.m_tvRel.FullRowSelect = true;
            this.m_tvRel.HideSelection = false;
            this.m_tvRel.HotTracking = true;
            resources.ApplyResources(this.m_tvRel, "m_tvRel");
            this.m_tvRel.Name = "m_tvRel";
            this.m_tvRel.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvRel.Nodes"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvRel.Nodes1")))});
            this.m_tvRel.Tag = "节点树";
            this.m_tvRel.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvRel_AfterSelect);
            this.m_tvRel.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.m_tvRel_NodeMouseClick);
            // 
            // btnFoundNext
            // 
            resources.ApplyResources(this.btnFoundNext, "btnFoundNext");
            this.btnFoundNext.Name = "btnFoundNext";
            this.btnFoundNext.UseVisualStyleBackColor = true;
            this.btnFoundNext.Click += new System.EventHandler(this.btnFoundNext_Click);
            // 
            // btnReset
            // 
            resources.ApplyResources(this.btnReset, "btnReset");
            this.btnReset.Name = "btnReset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnFoundBack
            // 
            resources.ApplyResources(this.btnFoundBack, "btnFoundBack");
            this.btnFoundBack.Name = "btnFoundBack";
            this.btnFoundBack.UseVisualStyleBackColor = true;
            this.btnFoundBack.Click += new System.EventHandler(this.btnFoundBack_Click);
            // 
            // txtRelKeyword
            // 
            resources.ApplyResources(this.txtRelKeyword, "txtRelKeyword");
            this.txtRelKeyword.Name = "txtRelKeyword";
            // 
            // btnRelSearch
            // 
            resources.ApplyResources(this.btnRelSearch, "btnRelSearch");
            this.btnRelSearch.Name = "btnRelSearch";
            this.btnRelSearch.UseVisualStyleBackColor = true;
            this.btnRelSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // btnRefreshRels
            // 
            resources.ApplyResources(this.btnRefreshRels, "btnRefreshRels");
            this.btnRefreshRels.Name = "btnRefreshRels";
            this.btnRefreshRels.UseVisualStyleBackColor = true;
            this.btnRefreshRels.Click += new System.EventHandler(this.btnRefreshRels_Click);
            // 
            // btnRelAllTxtOut
            // 
            resources.ApplyResources(this.btnRelAllTxtOut, "btnRelAllTxtOut");
            this.btnRelAllTxtOut.Name = "btnRelAllTxtOut";
            this.btnRelAllTxtOut.UseVisualStyleBackColor = true;
            this.btnRelAllTxtOut.Click += new System.EventHandler(this.btnRelAllTxtOut_Click);
            // 
            // btnMove
            // 
            resources.ApplyResources(this.btnMove, "btnMove");
            this.btnMove.Name = "btnMove";
            this.btnMove.UseVisualStyleBackColor = true;
            this.btnMove.Click += new System.EventHandler(this.btnMove_Click);
            // 
            // btnExpEditor
            // 
            resources.ApplyResources(this.btnExpEditor, "btnExpEditor");
            this.btnExpEditor.Name = "btnExpEditor";
            this.btnExpEditor.UseVisualStyleBackColor = true;
            // 
            // txtRelName
            // 
            resources.ApplyResources(this.txtRelName, "txtRelName");
            this.txtRelName.Name = "txtRelName";
            // 
            // chboxOpRulesEnable
            // 
            resources.ApplyResources(this.chboxOpRulesEnable, "chboxOpRulesEnable");
            this.chboxOpRulesEnable.Name = "chboxOpRulesEnable";
            this.chboxOpRulesEnable.UseVisualStyleBackColor = true;
            this.chboxOpRulesEnable.CheckedChanged += new System.EventHandler(this.chboxOpRulesEnable_CheckedChanged);
            // 
            // txtRelContent
            // 
            resources.ApplyResources(this.txtRelContent, "txtRelContent");
            this.txtRelContent.Name = "txtRelContent";
            // 
            // txtOpRules
            // 
            resources.ApplyResources(this.txtOpRules, "txtOpRules");
            this.txtOpRules.Name = "txtOpRules";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // btnAddRel
            // 
            resources.ApplyResources(this.btnAddRel, "btnAddRel");
            this.btnAddRel.Name = "btnAddRel";
            this.btnAddRel.UseVisualStyleBackColor = true;
            this.btnAddRel.Click += new System.EventHandler(this.btnAddRel_Click);
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // btnUpdateRel
            // 
            resources.ApplyResources(this.btnUpdateRel, "btnUpdateRel");
            this.btnUpdateRel.Name = "btnUpdateRel";
            this.btnUpdateRel.UseVisualStyleBackColor = true;
            this.btnUpdateRel.Click += new System.EventHandler(this.btnUpdateRel_Click);
            // 
            // btnInsertRel
            // 
            resources.ApplyResources(this.btnInsertRel, "btnInsertRel");
            this.btnInsertRel.Name = "btnInsertRel";
            this.btnInsertRel.UseVisualStyleBackColor = true;
            this.btnInsertRel.Click += new System.EventHandler(this.btnInsertRel_Click);
            // 
            // btnJump2Rel
            // 
            resources.ApplyResources(this.btnJump2Rel, "btnJump2Rel");
            this.btnJump2Rel.Name = "btnJump2Rel";
            this.btnJump2Rel.UseVisualStyleBackColor = true;
            this.btnJump2Rel.Click += new System.EventHandler(this.btnJump2Rel_Click);
            // 
            // btnRemoveRel
            // 
            resources.ApplyResources(this.btnRemoveRel, "btnRemoveRel");
            this.btnRemoveRel.Name = "btnRemoveRel";
            this.btnRemoveRel.UseVisualStyleBackColor = true;
            this.btnRemoveRel.Click += new System.EventHandler(this.btnRemoveRel_Click);
            // 
            // label36
            // 
            resources.ApplyResources(this.label36, "label36");
            this.label36.Name = "label36";
            // 
            // label35
            // 
            resources.ApplyResources(this.label35, "label35");
            this.label35.Name = "label35";
            // 
            // tabPageCheck
            // 
            this.tabPageCheck.Controls.Add(this.label4);
            this.tabPageCheck.Controls.Add(this.progbarCheck);
            this.tabPageCheck.Controls.Add(this.btnCheckSearchNext);
            this.tabPageCheck.Controls.Add(this.btnCheckSearchPrev);
            this.tabPageCheck.Controls.Add(this.btnCheck);
            this.tabPageCheck.Controls.Add(this.btnCheckReset);
            this.tabPageCheck.Controls.Add(this.tvCheckedItems);
            this.tabPageCheck.Controls.Add(this.btnCheckSearch);
            this.tabPageCheck.Controls.Add(this.btnCheckIgnore);
            this.tabPageCheck.Controls.Add(this.txtCheckSearchKeyWord);
            resources.ApplyResources(this.tabPageCheck, "tabPageCheck");
            this.tabPageCheck.Name = "tabPageCheck";
            this.tabPageCheck.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // progbarCheck
            // 
            resources.ApplyResources(this.progbarCheck, "progbarCheck");
            this.progbarCheck.Name = "progbarCheck";
            // 
            // btnCheckSearchNext
            // 
            resources.ApplyResources(this.btnCheckSearchNext, "btnCheckSearchNext");
            this.btnCheckSearchNext.Name = "btnCheckSearchNext";
            this.btnCheckSearchNext.UseVisualStyleBackColor = true;
            this.btnCheckSearchNext.Click += new System.EventHandler(this.btnCheckSearchNext_Click);
            // 
            // btnCheckSearchPrev
            // 
            resources.ApplyResources(this.btnCheckSearchPrev, "btnCheckSearchPrev");
            this.btnCheckSearchPrev.Name = "btnCheckSearchPrev";
            this.btnCheckSearchPrev.UseVisualStyleBackColor = true;
            this.btnCheckSearchPrev.Click += new System.EventHandler(this.btnCheckSearchPrev_Click);
            // 
            // btnCheck
            // 
            resources.ApplyResources(this.btnCheck, "btnCheck");
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.UseVisualStyleBackColor = true;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // btnCheckReset
            // 
            resources.ApplyResources(this.btnCheckReset, "btnCheckReset");
            this.btnCheckReset.Name = "btnCheckReset";
            this.btnCheckReset.UseVisualStyleBackColor = true;
            this.btnCheckReset.Click += new System.EventHandler(this.btnCheckReset_Click);
            // 
            // tvCheckedItems
            // 
            this.tvCheckedItems.FullRowSelect = true;
            this.tvCheckedItems.HideSelection = false;
            this.tvCheckedItems.HotTracking = true;
            resources.ApplyResources(this.tvCheckedItems, "tvCheckedItems");
            this.tvCheckedItems.Name = "tvCheckedItems";
            this.tvCheckedItems.Tag = "检查结果树";
            this.tvCheckedItems.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvCheckedItems_AfterSelect);
            // 
            // btnCheckSearch
            // 
            resources.ApplyResources(this.btnCheckSearch, "btnCheckSearch");
            this.btnCheckSearch.Name = "btnCheckSearch";
            this.btnCheckSearch.UseVisualStyleBackColor = true;
            this.btnCheckSearch.Click += new System.EventHandler(this.btnCheckSearch_Click);
            // 
            // btnCheckIgnore
            // 
            resources.ApplyResources(this.btnCheckIgnore, "btnCheckIgnore");
            this.btnCheckIgnore.Name = "btnCheckIgnore";
            this.btnCheckIgnore.UseVisualStyleBackColor = true;
            this.btnCheckIgnore.Click += new System.EventHandler(this.btnCheckIgnore_Click);
            // 
            // txtCheckSearchKeyWord
            // 
            resources.ApplyResources(this.txtCheckSearchKeyWord, "txtCheckSearchKeyWord");
            this.txtCheckSearchKeyWord.Name = "txtCheckSearchKeyWord";
            // 
            // tabPageOrganize
            // 
            this.tabPageOrganize.Controls.Add(this.OrgProgressBar);
            this.tabPageOrganize.Controls.Add(this.chkOrgShowBody);
            this.tabPageOrganize.Controls.Add(this.btnOrgCancelProtect);
            this.tabPageOrganize.Controls.Add(this.btnOrganProtect);
            this.tabPageOrganize.Controls.Add(this.btnOrganNext);
            this.tabPageOrganize.Controls.Add(this.btnOrganBack);
            this.tabPageOrganize.Controls.Add(this.btnOrganResetSearch);
            this.tabPageOrganize.Controls.Add(this.btnOrganSearch);
            this.tabPageOrganize.Controls.Add(this.txtOrganKeyWord);
            this.tabPageOrganize.Controls.Add(this.btnOrganizeRefresh);
            this.tabPageOrganize.Controls.Add(this.btnCollapseSel);
            this.tabPageOrganize.Controls.Add(this.btnExpandSelChild);
            this.tabPageOrganize.Controls.Add(this.btnSelAll);
            this.tabPageOrganize.Controls.Add(this.btnSelClear);
            this.tabPageOrganize.Controls.Add(this.btnOrgDemote);
            this.tabPageOrganize.Controls.Add(this.btnOrgPromote);
            this.tabPageOrganize.Controls.Add(this.chkSelCategory);
            this.tabPageOrganize.Controls.Add(this.m_tvOrganize);
            this.tabPageOrganize.Controls.Add(this.label37);
            resources.ApplyResources(this.tabPageOrganize, "tabPageOrganize");
            this.tabPageOrganize.Name = "tabPageOrganize";
            this.tabPageOrganize.UseVisualStyleBackColor = true;
            // 
            // OrgProgressBar
            // 
            resources.ApplyResources(this.OrgProgressBar, "OrgProgressBar");
            this.OrgProgressBar.Name = "OrgProgressBar";
            // 
            // chkOrgShowBody
            // 
            resources.ApplyResources(this.chkOrgShowBody, "chkOrgShowBody");
            this.chkOrgShowBody.Name = "chkOrgShowBody";
            this.chkOrgShowBody.UseVisualStyleBackColor = true;
            this.chkOrgShowBody.CheckedChanged += new System.EventHandler(this.chkOrgShowBody_CheckedChanged);
            // 
            // btnOrgCancelProtect
            // 
            resources.ApplyResources(this.btnOrgCancelProtect, "btnOrgCancelProtect");
            this.btnOrgCancelProtect.Name = "btnOrgCancelProtect";
            this.btnOrgCancelProtect.UseVisualStyleBackColor = true;
            this.btnOrgCancelProtect.Click += new System.EventHandler(this.btnOrgCancelProtect_Click);
            // 
            // btnOrganProtect
            // 
            resources.ApplyResources(this.btnOrganProtect, "btnOrganProtect");
            this.btnOrganProtect.Name = "btnOrganProtect";
            this.btnOrganProtect.UseVisualStyleBackColor = true;
            this.btnOrganProtect.Click += new System.EventHandler(this.btnOrganProtect_Click);
            // 
            // btnOrganNext
            // 
            resources.ApplyResources(this.btnOrganNext, "btnOrganNext");
            this.btnOrganNext.Name = "btnOrganNext";
            this.btnOrganNext.UseVisualStyleBackColor = true;
            this.btnOrganNext.Click += new System.EventHandler(this.btnOrganNext_Click);
            // 
            // btnOrganBack
            // 
            resources.ApplyResources(this.btnOrganBack, "btnOrganBack");
            this.btnOrganBack.Name = "btnOrganBack";
            this.btnOrganBack.UseVisualStyleBackColor = true;
            this.btnOrganBack.Click += new System.EventHandler(this.btnOrganBack_Click);
            // 
            // btnOrganResetSearch
            // 
            resources.ApplyResources(this.btnOrganResetSearch, "btnOrganResetSearch");
            this.btnOrganResetSearch.Name = "btnOrganResetSearch";
            this.btnOrganResetSearch.UseVisualStyleBackColor = true;
            this.btnOrganResetSearch.Click += new System.EventHandler(this.btnOrganResetSearch_Click);
            // 
            // btnOrganSearch
            // 
            resources.ApplyResources(this.btnOrganSearch, "btnOrganSearch");
            this.btnOrganSearch.Name = "btnOrganSearch";
            this.btnOrganSearch.UseVisualStyleBackColor = true;
            this.btnOrganSearch.Click += new System.EventHandler(this.btnOrganSearch_Click);
            // 
            // txtOrganKeyWord
            // 
            resources.ApplyResources(this.txtOrganKeyWord, "txtOrganKeyWord");
            this.txtOrganKeyWord.Name = "txtOrganKeyWord";
            // 
            // btnOrganizeRefresh
            // 
            resources.ApplyResources(this.btnOrganizeRefresh, "btnOrganizeRefresh");
            this.btnOrganizeRefresh.Name = "btnOrganizeRefresh";
            this.btnOrganizeRefresh.UseVisualStyleBackColor = true;
            this.btnOrganizeRefresh.Click += new System.EventHandler(this.btnOrganizeRefresh_Click);
            // 
            // btnCollapseSel
            // 
            resources.ApplyResources(this.btnCollapseSel, "btnCollapseSel");
            this.btnCollapseSel.Name = "btnCollapseSel";
            this.btnCollapseSel.UseVisualStyleBackColor = true;
            this.btnCollapseSel.Click += new System.EventHandler(this.btnCollapseSel_Click);
            // 
            // btnExpandSelChild
            // 
            resources.ApplyResources(this.btnExpandSelChild, "btnExpandSelChild");
            this.btnExpandSelChild.Name = "btnExpandSelChild";
            this.btnExpandSelChild.UseVisualStyleBackColor = true;
            this.btnExpandSelChild.Click += new System.EventHandler(this.btnExpandSelChild_Click);
            // 
            // btnSelAll
            // 
            resources.ApplyResources(this.btnSelAll, "btnSelAll");
            this.btnSelAll.Name = "btnSelAll";
            this.btnSelAll.UseVisualStyleBackColor = true;
            this.btnSelAll.Click += new System.EventHandler(this.btnSelAll_Click);
            // 
            // btnSelClear
            // 
            resources.ApplyResources(this.btnSelClear, "btnSelClear");
            this.btnSelClear.Name = "btnSelClear";
            this.btnSelClear.UseVisualStyleBackColor = true;
            this.btnSelClear.Click += new System.EventHandler(this.btnSelClear_Click);
            // 
            // btnOrgDemote
            // 
            resources.ApplyResources(this.btnOrgDemote, "btnOrgDemote");
            this.btnOrgDemote.Name = "btnOrgDemote";
            this.btnOrgDemote.UseVisualStyleBackColor = true;
            this.btnOrgDemote.Click += new System.EventHandler(this.btnOrgDemote_Click);
            // 
            // btnOrgPromote
            // 
            resources.ApplyResources(this.btnOrgPromote, "btnOrgPromote");
            this.btnOrgPromote.Name = "btnOrgPromote";
            this.btnOrgPromote.UseVisualStyleBackColor = true;
            this.btnOrgPromote.Click += new System.EventHandler(this.btnOrgPromote_Click);
            // 
            // chkSelCategory
            // 
            this.chkSelCategory.CheckOnClick = true;
            resources.ApplyResources(this.chkSelCategory, "chkSelCategory");
            this.chkSelCategory.FormattingEnabled = true;
            this.chkSelCategory.Items.AddRange(new object[] {
            resources.GetString("chkSelCategory.Items"),
            resources.GetString("chkSelCategory.Items1"),
            resources.GetString("chkSelCategory.Items2"),
            resources.GetString("chkSelCategory.Items3"),
            resources.GetString("chkSelCategory.Items4"),
            resources.GetString("chkSelCategory.Items5"),
            resources.GetString("chkSelCategory.Items6"),
            resources.GetString("chkSelCategory.Items7"),
            resources.GetString("chkSelCategory.Items8"),
            resources.GetString("chkSelCategory.Items9"),
            resources.GetString("chkSelCategory.Items10")});
            this.chkSelCategory.MultiColumn = true;
            this.chkSelCategory.Name = "chkSelCategory";
            this.chkSelCategory.Tag = "选择区";
            this.chkSelCategory.SelectedIndexChanged += new System.EventHandler(this.chkCategory_SelectedIndexChanged);
            // 
            // m_tvOrganize
            // 
            this.m_tvOrganize.CheckBoxes = true;
            this.m_tvOrganize.FullRowSelect = true;
            this.m_tvOrganize.HideSelection = false;
            this.m_tvOrganize.HotTracking = true;
            resources.ApplyResources(this.m_tvOrganize, "m_tvOrganize");
            this.m_tvOrganize.ImageList = this.imageListIcon;
            this.m_tvOrganize.Name = "m_tvOrganize";
            this.m_tvOrganize.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes1"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes2"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes3"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes4"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes5"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes6"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes7"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("m_tvOrganize.Nodes8")))});
            this.m_tvOrganize.Tag = "节点树";
            this.m_tvOrganize.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.m_tvOrganize_AfterCheck);
            this.m_tvOrganize.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.m_tvOrganize_AfterSelect);
            this.m_tvOrganize.Click += new System.EventHandler(this.m_tvOrganize_Click);
            // 
            // imageListIcon
            // 
            this.imageListIcon.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListIcon.ImageStream")));
            this.imageListIcon.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListIcon.Images.SetKeyName(0, "wordIcon.jpg");
            this.imageListIcon.Images.SetKeyName(1, "number1.jpg");
            this.imageListIcon.Images.SetKeyName(2, "number2.jpg");
            this.imageListIcon.Images.SetKeyName(3, "number3.jpg");
            this.imageListIcon.Images.SetKeyName(4, "number4.jpg");
            this.imageListIcon.Images.SetKeyName(5, "number5.jpg");
            this.imageListIcon.Images.SetKeyName(6, "number6.jpg");
            this.imageListIcon.Images.SetKeyName(7, "number7.jpg");
            this.imageListIcon.Images.SetKeyName(8, "number8.jpg");
            this.imageListIcon.Images.SetKeyName(9, "number9.jpg");
            this.imageListIcon.Images.SetKeyName(10, "正.jpg");
            this.imageListIcon.Images.SetKeyName(11, "upArrow.jpg");
            this.imageListIcon.Images.SetKeyName(12, "downArrow.jpg");
            this.imageListIcon.Images.SetKeyName(13, "folder.jpg");
            this.imageListIcon.Images.SetKeyName(14, "excelIcon.jpg");
            this.imageListIcon.Images.SetKeyName(15, "pptIcon.jpg");
            this.imageListIcon.Images.SetKeyName(16, "pdfIcon.jpg");
            this.imageListIcon.Images.SetKeyName(17, "fileIcon.jpg");
            this.imageListIcon.Images.SetKeyName(18, "cdDriver.jpg");
            this.imageListIcon.Images.SetKeyName(19, "driver.jpg");
            this.imageListIcon.Images.SetKeyName(20, "myComputer.jpg");
            this.imageListIcon.Images.SetKeyName(21, "commonDB.jpg");
            this.imageListIcon.Images.SetKeyName(22, "category.jpg");
            // 
            // label37
            // 
            resources.ApplyResources(this.label37, "label37");
            this.label37.Name = "label37";
            // 
            // tabPageShare
            // 
            this.tabPageShare.Controls.Add(this.btnShareLibUpdate);
            this.tabPageShare.Controls.Add(this.btnShareOpen);
            this.tabPageShare.Controls.Add(this.prgBarLib);
            this.tabPageShare.Controls.Add(this.tvShareLib);
            this.tabPageShare.Controls.Add(this.btnShareRef);
            this.tabPageShare.Controls.Add(this.btnShareRefresh);
            this.tabPageShare.Controls.Add(this.btnShareExpand);
            this.tabPageShare.Controls.Add(this.btnShareDownload);
            this.tabPageShare.Controls.Add(this.btnShareCollapse);
            this.tabPageShare.Controls.Add(this.btnShareSearch);
            this.tabPageShare.Controls.Add(this.btnShareSearchReset);
            this.tabPageShare.Controls.Add(this.txtShareKeyWord);
            this.tabPageShare.Controls.Add(this.btnSharePrevSearch);
            this.tabPageShare.Controls.Add(this.btnShareNextSearch);
            this.tabPageShare.Controls.Add(this.label6);
            this.tabPageShare.Controls.Add(this.btnShareExternalFile);
            this.tabPageShare.Controls.Add(this.txtShareName);
            this.tabPageShare.Controls.Add(this.label5);
            this.tabPageShare.Controls.Add(this.btnShareAdd);
            this.tabPageShare.Controls.Add(this.btnShareRemove);
            this.tabPageShare.Controls.Add(this.chkBoxCommonLib);
            this.tabPageShare.Controls.Add(this.chkBoxCategory);
            this.tabPageShare.Controls.Add(this.label38);
            this.tabPageShare.Controls.Add(this.label39);
            resources.ApplyResources(this.tabPageShare, "tabPageShare");
            this.tabPageShare.Name = "tabPageShare";
            this.tabPageShare.UseVisualStyleBackColor = true;
            // 
            // btnShareLibUpdate
            // 
            resources.ApplyResources(this.btnShareLibUpdate, "btnShareLibUpdate");
            this.btnShareLibUpdate.Name = "btnShareLibUpdate";
            this.btnShareLibUpdate.UseVisualStyleBackColor = true;
            this.btnShareLibUpdate.Click += new System.EventHandler(this.btnShareLibUpdate_Click);
            // 
            // btnShareOpen
            // 
            resources.ApplyResources(this.btnShareOpen, "btnShareOpen");
            this.btnShareOpen.Name = "btnShareOpen";
            this.btnShareOpen.UseVisualStyleBackColor = true;
            this.btnShareOpen.Click += new System.EventHandler(this.btnShareOpen_Click);
            // 
            // prgBarLib
            // 
            resources.ApplyResources(this.prgBarLib, "prgBarLib");
            this.prgBarLib.Name = "prgBarLib";
            // 
            // tvShareLib
            // 
            this.tvShareLib.ContextMenuStrip = this.cxtMenuSvr;
            this.tvShareLib.FullRowSelect = true;
            this.tvShareLib.HotTracking = true;
            resources.ApplyResources(this.tvShareLib, "tvShareLib");
            this.tvShareLib.ImageList = this.imageListIcon;
            this.tvShareLib.Name = "tvShareLib";
            this.tvShareLib.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            ((System.Windows.Forms.TreeNode)(resources.GetObject("tvShareLib.Nodes"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("tvShareLib.Nodes1"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("tvShareLib.Nodes2")))});
            this.tvShareLib.Tag = "文库资源树";
            this.tvShareLib.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvShareLib_BeforeExpand);
            this.tvShareLib.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvShareLib_AfterSelect);
            this.tvShareLib.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvShareLib_NodeMouseDoubleClick);
            // 
            // cxtMenuSvr
            // 
            this.cxtMenuSvr.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuItemApplyStyle,
            this.menuItemCntReuse});
            this.cxtMenuSvr.Name = "cxtMenuSvr";
            resources.ApplyResources(this.cxtMenuSvr, "cxtMenuSvr");
            // 
            // menuItemApplyStyle
            // 
            this.menuItemApplyStyle.Name = "menuItemApplyStyle";
            resources.ApplyResources(this.menuItemApplyStyle, "menuItemApplyStyle");
            this.menuItemApplyStyle.Click += new System.EventHandler(this.menuItemApplyStyle_Click);
            // 
            // menuItemCntReuse
            // 
            this.menuItemCntReuse.Name = "menuItemCntReuse";
            resources.ApplyResources(this.menuItemCntReuse, "menuItemCntReuse");
            this.menuItemCntReuse.Click += new System.EventHandler(this.menuItemCntReuse_Click);
            // 
            // btnShareRef
            // 
            resources.ApplyResources(this.btnShareRef, "btnShareRef");
            this.btnShareRef.Name = "btnShareRef";
            this.btnShareRef.UseVisualStyleBackColor = true;
            this.btnShareRef.Click += new System.EventHandler(this.btnShareRef_Click);
            // 
            // btnShareRefresh
            // 
            resources.ApplyResources(this.btnShareRefresh, "btnShareRefresh");
            this.btnShareRefresh.Name = "btnShareRefresh";
            this.btnShareRefresh.UseVisualStyleBackColor = true;
            this.btnShareRefresh.Click += new System.EventHandler(this.btnShareRefresh_Click);
            // 
            // btnShareExpand
            // 
            resources.ApplyResources(this.btnShareExpand, "btnShareExpand");
            this.btnShareExpand.Name = "btnShareExpand";
            this.btnShareExpand.UseVisualStyleBackColor = true;
            this.btnShareExpand.Click += new System.EventHandler(this.btnShareExpand_Click);
            // 
            // btnShareDownload
            // 
            resources.ApplyResources(this.btnShareDownload, "btnShareDownload");
            this.btnShareDownload.Name = "btnShareDownload";
            this.btnShareDownload.UseVisualStyleBackColor = true;
            this.btnShareDownload.Click += new System.EventHandler(this.btnShareDownload_Click);
            // 
            // btnShareCollapse
            // 
            resources.ApplyResources(this.btnShareCollapse, "btnShareCollapse");
            this.btnShareCollapse.Name = "btnShareCollapse";
            this.btnShareCollapse.UseVisualStyleBackColor = true;
            this.btnShareCollapse.Click += new System.EventHandler(this.btnShareCollapse_Click);
            // 
            // btnShareSearch
            // 
            resources.ApplyResources(this.btnShareSearch, "btnShareSearch");
            this.btnShareSearch.Name = "btnShareSearch";
            this.btnShareSearch.UseVisualStyleBackColor = true;
            this.btnShareSearch.Click += new System.EventHandler(this.btnShareSearch_Click);
            // 
            // btnShareSearchReset
            // 
            resources.ApplyResources(this.btnShareSearchReset, "btnShareSearchReset");
            this.btnShareSearchReset.Name = "btnShareSearchReset";
            this.btnShareSearchReset.UseVisualStyleBackColor = true;
            this.btnShareSearchReset.Click += new System.EventHandler(this.btnShareSearchReset_Click);
            // 
            // txtShareKeyWord
            // 
            resources.ApplyResources(this.txtShareKeyWord, "txtShareKeyWord");
            this.txtShareKeyWord.Name = "txtShareKeyWord";
            // 
            // btnSharePrevSearch
            // 
            resources.ApplyResources(this.btnSharePrevSearch, "btnSharePrevSearch");
            this.btnSharePrevSearch.Name = "btnSharePrevSearch";
            this.btnSharePrevSearch.UseVisualStyleBackColor = true;
            this.btnSharePrevSearch.Click += new System.EventHandler(this.btnSharePrevSearch_Click);
            // 
            // btnShareNextSearch
            // 
            resources.ApplyResources(this.btnShareNextSearch, "btnShareNextSearch");
            this.btnShareNextSearch.Name = "btnShareNextSearch";
            this.btnShareNextSearch.UseVisualStyleBackColor = true;
            this.btnShareNextSearch.Click += new System.EventHandler(this.btnShareNextSearch_Click);
            // 
            // label6
            // 
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            // 
            // btnShareExternalFile
            // 
            resources.ApplyResources(this.btnShareExternalFile, "btnShareExternalFile");
            this.btnShareExternalFile.Name = "btnShareExternalFile";
            this.btnShareExternalFile.UseVisualStyleBackColor = true;
            this.btnShareExternalFile.Click += new System.EventHandler(this.btnShareExternalFile_Click);
            // 
            // txtShareName
            // 
            resources.ApplyResources(this.txtShareName, "txtShareName");
            this.txtShareName.Name = "txtShareName";
            // 
            // label5
            // 
            resources.ApplyResources(this.label5, "label5");
            this.label5.Name = "label5";
            // 
            // btnShareAdd
            // 
            resources.ApplyResources(this.btnShareAdd, "btnShareAdd");
            this.btnShareAdd.Name = "btnShareAdd";
            this.btnShareAdd.UseVisualStyleBackColor = true;
            this.btnShareAdd.Click += new System.EventHandler(this.btnShareAdd_Click);
            // 
            // btnShareRemove
            // 
            resources.ApplyResources(this.btnShareRemove, "btnShareRemove");
            this.btnShareRemove.Name = "btnShareRemove";
            this.btnShareRemove.UseVisualStyleBackColor = true;
            this.btnShareRemove.Click += new System.EventHandler(this.btnShareRemove_Click);
            // 
            // chkBoxCommonLib
            // 
            resources.ApplyResources(this.chkBoxCommonLib, "chkBoxCommonLib");
            this.chkBoxCommonLib.Name = "chkBoxCommonLib";
            this.chkBoxCommonLib.UseVisualStyleBackColor = true;
            // 
            // chkBoxCategory
            // 
            resources.ApplyResources(this.chkBoxCategory, "chkBoxCategory");
            this.chkBoxCategory.Name = "chkBoxCategory";
            this.chkBoxCategory.UseVisualStyleBackColor = true;
            // 
            // label38
            // 
            resources.ApplyResources(this.label38, "label38");
            this.label38.Name = "label38";
            // 
            // label39
            // 
            resources.ApplyResources(this.label39, "label39");
            this.label39.Name = "label39";
            // 
            // tabPageUnitedStyle
            // 
            this.tabPageUnitedStyle.Controls.Add(this.rchTextBoxUniformStylesPreview);
            this.tabPageUnitedStyle.Controls.Add(this.groupBox6);
            this.tabPageUnitedStyle.Controls.Add(this.groupBox4);
            resources.ApplyResources(this.tabPageUnitedStyle, "tabPageUnitedStyle");
            this.tabPageUnitedStyle.Name = "tabPageUnitedStyle";
            this.tabPageUnitedStyle.UseVisualStyleBackColor = true;
            // 
            // rchTextBoxUniformStylesPreview
            // 
            resources.ApplyResources(this.rchTextBoxUniformStylesPreview, "rchTextBoxUniformStylesPreview");
            this.rchTextBoxUniformStylesPreview.Name = "rchTextBoxUniformStylesPreview";
            this.rchTextBoxUniformStylesPreview.ReadOnly = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.chkIgnoreParaFormat);
            this.groupBox6.Controls.Add(this.chkIgnoreFont);
            this.groupBox6.Controls.Add(this.chkIgnoreTextBody);
            this.groupBox6.Controls.Add(this.chkIgnoreHeadings);
            this.groupBox6.Controls.Add(this.label7);
            this.groupBox6.Controls.Add(this.txtIgnorePages);
            this.groupBox6.Controls.Add(this.chkIgnorePages);
            this.groupBox6.Controls.Add(this.chkIgnoreTable);
            this.groupBox6.Controls.Add(this.chkIgnoreTOC);
            this.groupBox6.Controls.Add(this.radioBtnStyleSelection);
            this.groupBox6.Controls.Add(this.radioBtnStyleAllDoc);
            resources.ApplyResources(this.groupBox6, "groupBox6");
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.TabStop = false;
            // 
            // chkIgnoreParaFormat
            // 
            resources.ApplyResources(this.chkIgnoreParaFormat, "chkIgnoreParaFormat");
            this.chkIgnoreParaFormat.Name = "chkIgnoreParaFormat";
            this.chkIgnoreParaFormat.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreFont
            // 
            resources.ApplyResources(this.chkIgnoreFont, "chkIgnoreFont");
            this.chkIgnoreFont.Name = "chkIgnoreFont";
            this.chkIgnoreFont.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreTextBody
            // 
            resources.ApplyResources(this.chkIgnoreTextBody, "chkIgnoreTextBody");
            this.chkIgnoreTextBody.Checked = true;
            this.chkIgnoreTextBody.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreTextBody.Name = "chkIgnoreTextBody";
            this.chkIgnoreTextBody.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreHeadings
            // 
            resources.ApplyResources(this.chkIgnoreHeadings, "chkIgnoreHeadings");
            this.chkIgnoreHeadings.Name = "chkIgnoreHeadings";
            this.chkIgnoreHeadings.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            resources.ApplyResources(this.label7, "label7");
            this.label7.Name = "label7";
            // 
            // txtIgnorePages
            // 
            resources.ApplyResources(this.txtIgnorePages, "txtIgnorePages");
            this.txtIgnorePages.Name = "txtIgnorePages";
            this.txtIgnorePages.TextChanged += new System.EventHandler(this.txtIgnorePages_TextChanged);
            // 
            // chkIgnorePages
            // 
            resources.ApplyResources(this.chkIgnorePages, "chkIgnorePages");
            this.chkIgnorePages.Name = "chkIgnorePages";
            this.chkIgnorePages.UseVisualStyleBackColor = true;
            this.chkIgnorePages.CheckedChanged += new System.EventHandler(this.chkIgnorePages_CheckedChanged);
            // 
            // chkIgnoreTable
            // 
            resources.ApplyResources(this.chkIgnoreTable, "chkIgnoreTable");
            this.chkIgnoreTable.Checked = true;
            this.chkIgnoreTable.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreTable.Name = "chkIgnoreTable";
            this.chkIgnoreTable.UseVisualStyleBackColor = true;
            // 
            // chkIgnoreTOC
            // 
            resources.ApplyResources(this.chkIgnoreTOC, "chkIgnoreTOC");
            this.chkIgnoreTOC.Checked = true;
            this.chkIgnoreTOC.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIgnoreTOC.Name = "chkIgnoreTOC";
            this.chkIgnoreTOC.UseVisualStyleBackColor = true;
            // 
            // radioBtnStyleSelection
            // 
            resources.ApplyResources(this.radioBtnStyleSelection, "radioBtnStyleSelection");
            this.radioBtnStyleSelection.Name = "radioBtnStyleSelection";
            this.radioBtnStyleSelection.UseVisualStyleBackColor = true;
            this.radioBtnStyleSelection.Click += new System.EventHandler(this.radioBtnStyleSelection_Click);
            // 
            // radioBtnStyleAllDoc
            // 
            resources.ApplyResources(this.radioBtnStyleAllDoc, "radioBtnStyleAllDoc");
            this.radioBtnStyleAllDoc.Checked = true;
            this.radioBtnStyleAllDoc.Name = "radioBtnStyleAllDoc";
            this.radioBtnStyleAllDoc.TabStop = true;
            this.radioBtnStyleAllDoc.UseVisualStyleBackColor = true;
            this.radioBtnStyleAllDoc.Click += new System.EventHandler(this.radioBtnStyleAllDoc_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.btnUnitFormExitApply);
            this.groupBox4.Controls.Add(this.lstUnitedStyleHistoryDoc);
            this.groupBox4.Controls.Add(this.btnStyleApply);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.progressBarStyle);
            this.groupBox4.Controls.Add(this.txtBoxStyleFile);
            this.groupBox4.Controls.Add(this.btnStyleOpenFile);
            resources.ApplyResources(this.groupBox4, "groupBox4");
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.TabStop = false;
            // 
            // btnUnitFormExitApply
            // 
            resources.ApplyResources(this.btnUnitFormExitApply, "btnUnitFormExitApply");
            this.btnUnitFormExitApply.Name = "btnUnitFormExitApply";
            this.btnUnitFormExitApply.UseVisualStyleBackColor = true;
            this.btnUnitFormExitApply.Click += new System.EventHandler(this.btnUnitFormExitApply_Click);
            // 
            // lstUnitedStyleHistoryDoc
            // 
            this.lstUnitedStyleHistoryDoc.DataSource = this.tblUniformStyleHistoryDocsBindingSource;
            this.lstUnitedStyleHistoryDoc.DisplayMember = "fullPathDoc";
            this.lstUnitedStyleHistoryDoc.FormattingEnabled = true;
            resources.ApplyResources(this.lstUnitedStyleHistoryDoc, "lstUnitedStyleHistoryDoc");
            this.lstUnitedStyleHistoryDoc.Name = "lstUnitedStyleHistoryDoc";
            this.lstUnitedStyleHistoryDoc.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.lstUnitedStyleHistoryDoc_MouseDoubleClick);
            // 
            // tblUniformStyleHistoryDocsBindingSource
            // 
            this.tblUniformStyleHistoryDocsBindingSource.DataMember = "tblUniformStyleHistoryDocs";
            this.tblUniformStyleHistoryDocsBindingSource.DataSource = this.localdbDataSet;
            // 
            // localdbDataSet
            // 
            this.localdbDataSet.DataSetName = "localdbDataSet";
            this.localdbDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // btnStyleApply
            // 
            resources.ApplyResources(this.btnStyleApply, "btnStyleApply");
            this.btnStyleApply.Name = "btnStyleApply";
            this.btnStyleApply.UseVisualStyleBackColor = true;
            this.btnStyleApply.Click += new System.EventHandler(this.btnStyleApply_Click);
            // 
            // label11
            // 
            resources.ApplyResources(this.label11, "label11");
            this.label11.Name = "label11";
            // 
            // progressBarStyle
            // 
            resources.ApplyResources(this.progressBarStyle, "progressBarStyle");
            this.progressBarStyle.Name = "progressBarStyle";
            // 
            // txtBoxStyleFile
            // 
            resources.ApplyResources(this.txtBoxStyleFile, "txtBoxStyleFile");
            this.txtBoxStyleFile.Name = "txtBoxStyleFile";
            this.txtBoxStyleFile.ReadOnly = true;
            // 
            // btnStyleOpenFile
            // 
            resources.ApplyResources(this.btnStyleOpenFile, "btnStyleOpenFile");
            this.btnStyleOpenFile.Name = "btnStyleOpenFile";
            this.btnStyleOpenFile.UseVisualStyleBackColor = true;
            this.btnStyleOpenFile.Click += new System.EventHandler(this.btnStyleOpenFile_Click);
            // 
            // tabPageCompare
            // 
            this.tabPageCompare.Controls.Add(this.progBarComp);
            this.tabPageCompare.Controls.Add(this.btnExecCompare);
            this.tabPageCompare.Controls.Add(this.txtCompResult);
            this.tabPageCompare.Controls.Add(this.tvCompCheck);
            this.tabPageCompare.Controls.Add(this.tvCompStandard);
            this.tabPageCompare.Controls.Add(this.chkCompStrickOrder);
            this.tabPageCompare.Controls.Add(this.chkCompOutline);
            this.tabPageCompare.Controls.Add(this.label10);
            this.tabPageCompare.Controls.Add(this.label9);
            this.tabPageCompare.Controls.Add(this.btnCompCheckDoc);
            this.tabPageCompare.Controls.Add(this.txtComp2CheckDoc);
            this.tabPageCompare.Controls.Add(this.btnCompStandardDoc);
            this.tabPageCompare.Controls.Add(this.txtCompStandardDoc);
            resources.ApplyResources(this.tabPageCompare, "tabPageCompare");
            this.tabPageCompare.Name = "tabPageCompare";
            this.tabPageCompare.UseVisualStyleBackColor = true;
            // 
            // progBarComp
            // 
            resources.ApplyResources(this.progBarComp, "progBarComp");
            this.progBarComp.Name = "progBarComp";
            // 
            // btnExecCompare
            // 
            resources.ApplyResources(this.btnExecCompare, "btnExecCompare");
            this.btnExecCompare.Name = "btnExecCompare";
            this.btnExecCompare.UseVisualStyleBackColor = true;
            this.btnExecCompare.Click += new System.EventHandler(this.btnExecCompare_Click);
            // 
            // txtCompResult
            // 
            resources.ApplyResources(this.txtCompResult, "txtCompResult");
            this.txtCompResult.Name = "txtCompResult";
            this.txtCompResult.ReadOnly = true;
            // 
            // tvCompCheck
            // 
            this.tvCompCheck.FullRowSelect = true;
            this.tvCompCheck.HideSelection = false;
            this.tvCompCheck.HotTracking = true;
            resources.ApplyResources(this.tvCompCheck, "tvCompCheck");
            this.tvCompCheck.ImageList = this.imageListIcon;
            this.tvCompCheck.Name = "tvCompCheck";
            this.tvCompCheck.ShowNodeToolTips = true;
            this.tvCompCheck.Tag = "检查文档章节树";
            this.tvCompCheck.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvCompCheck_NodeMouseDoubleClick);
            // 
            // tvCompStandard
            // 
            this.tvCompStandard.FullRowSelect = true;
            this.tvCompStandard.HideSelection = false;
            this.tvCompStandard.HotTracking = true;
            resources.ApplyResources(this.tvCompStandard, "tvCompStandard");
            this.tvCompStandard.ImageList = this.imageListIcon;
            this.tvCompStandard.Name = "tvCompStandard";
            this.tvCompStandard.ShowNodeToolTips = true;
            this.tvCompStandard.Tag = "标准文档章节树";
            this.tvCompStandard.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.tvCompStandard_NodeMouseDoubleClick);
            // 
            // chkCompStrickOrder
            // 
            resources.ApplyResources(this.chkCompStrickOrder, "chkCompStrickOrder");
            this.chkCompStrickOrder.Name = "chkCompStrickOrder";
            this.chkCompStrickOrder.UseVisualStyleBackColor = true;
            // 
            // chkCompOutline
            // 
            resources.ApplyResources(this.chkCompOutline, "chkCompOutline");
            this.chkCompOutline.Name = "chkCompOutline";
            this.chkCompOutline.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            resources.ApplyResources(this.label10, "label10");
            this.label10.Name = "label10";
            // 
            // label9
            // 
            resources.ApplyResources(this.label9, "label9");
            this.label9.Name = "label9";
            // 
            // btnCompCheckDoc
            // 
            resources.ApplyResources(this.btnCompCheckDoc, "btnCompCheckDoc");
            this.btnCompCheckDoc.Name = "btnCompCheckDoc";
            this.btnCompCheckDoc.UseVisualStyleBackColor = true;
            this.btnCompCheckDoc.Click += new System.EventHandler(this.btnCompCheckDoc_Click);
            // 
            // txtComp2CheckDoc
            // 
            resources.ApplyResources(this.txtComp2CheckDoc, "txtComp2CheckDoc");
            this.txtComp2CheckDoc.Name = "txtComp2CheckDoc";
            this.txtComp2CheckDoc.ReadOnly = true;
            // 
            // btnCompStandardDoc
            // 
            resources.ApplyResources(this.btnCompStandardDoc, "btnCompStandardDoc");
            this.btnCompStandardDoc.Name = "btnCompStandardDoc";
            this.btnCompStandardDoc.UseVisualStyleBackColor = true;
            this.btnCompStandardDoc.Click += new System.EventHandler(this.btnCompStandardDoc_Click);
            // 
            // txtCompStandardDoc
            // 
            resources.ApplyResources(this.txtCompStandardDoc, "txtCompStandardDoc");
            this.txtCompStandardDoc.Name = "txtCompStandardDoc";
            this.txtCompStandardDoc.ReadOnly = true;
            // 
            // tabPageDataTrans
            // 
            this.tabPageDataTrans.Controls.Add(this.tabControl1);
            resources.ApplyResources(this.tabPageDataTrans, "tabPageDataTrans");
            this.tabPageDataTrans.Name = "tabPageDataTrans";
            this.tabPageDataTrans.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageDocTbls2Excel);
            this.tabControl1.Controls.Add(this.tabPageExcel2DocTbl);
            resources.ApplyResources(this.tabControl1, "tabControl1");
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            // 
            // tabPageDocTbls2Excel
            // 
            this.tabPageDocTbls2Excel.Controls.Add(this.btnW2XNextSameStructTbl);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnAddTagCol);
            this.tabPageDocTbls2Excel.Controls.Add(this.grpW2XAutoModelScope);
            this.tabPageDocTbls2Excel.Controls.Add(this.trvDataDocTbl2Excel);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnClearItems);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnAllProduce);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnPreviewProduce);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnDocTbl2ExcelRemove);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnAddColValue);
            this.tabPageDocTbls2Excel.Controls.Add(this.btnAddColName);
            resources.ApplyResources(this.tabPageDocTbls2Excel, "tabPageDocTbls2Excel");
            this.tabPageDocTbls2Excel.Name = "tabPageDocTbls2Excel";
            this.tabPageDocTbls2Excel.UseVisualStyleBackColor = true;
            // 
            // btnW2XNextSameStructTbl
            // 
            resources.ApplyResources(this.btnW2XNextSameStructTbl, "btnW2XNextSameStructTbl");
            this.btnW2XNextSameStructTbl.Name = "btnW2XNextSameStructTbl";
            this.btnW2XNextSameStructTbl.UseVisualStyleBackColor = true;
            this.btnW2XNextSameStructTbl.Click += new System.EventHandler(this.btnW2XNextSameStructTbl_Click);
            // 
            // btnAddTagCol
            // 
            resources.ApplyResources(this.btnAddTagCol, "btnAddTagCol");
            this.btnAddTagCol.Name = "btnAddTagCol";
            this.btnAddTagCol.UseVisualStyleBackColor = true;
            this.btnAddTagCol.Click += new System.EventHandler(this.btnAddTagCol_Click);
            // 
            // grpW2XAutoModelScope
            // 
            this.grpW2XAutoModelScope.Controls.Add(this.chkBoxStrictVerifyTblColName);
            this.grpW2XAutoModelScope.Controls.Add(this.rdBtnW2XSelScope);
            this.grpW2XAutoModelScope.Controls.Add(this.rdBtnW2XAllDocScope);
            resources.ApplyResources(this.grpW2XAutoModelScope, "grpW2XAutoModelScope");
            this.grpW2XAutoModelScope.Name = "grpW2XAutoModelScope";
            this.grpW2XAutoModelScope.TabStop = false;
            // 
            // chkBoxStrictVerifyTblColName
            // 
            resources.ApplyResources(this.chkBoxStrictVerifyTblColName, "chkBoxStrictVerifyTblColName");
            this.chkBoxStrictVerifyTblColName.Name = "chkBoxStrictVerifyTblColName";
            this.chkBoxStrictVerifyTblColName.UseVisualStyleBackColor = true;
            // 
            // rdBtnW2XSelScope
            // 
            resources.ApplyResources(this.rdBtnW2XSelScope, "rdBtnW2XSelScope");
            this.rdBtnW2XSelScope.Name = "rdBtnW2XSelScope";
            this.rdBtnW2XSelScope.UseVisualStyleBackColor = true;
            // 
            // rdBtnW2XAllDocScope
            // 
            resources.ApplyResources(this.rdBtnW2XAllDocScope, "rdBtnW2XAllDocScope");
            this.rdBtnW2XAllDocScope.Checked = true;
            this.rdBtnW2XAllDocScope.Name = "rdBtnW2XAllDocScope";
            this.rdBtnW2XAllDocScope.TabStop = true;
            this.rdBtnW2XAllDocScope.UseVisualStyleBackColor = true;
            // 
            // trvDataDocTbl2Excel
            // 
            resources.ApplyResources(this.trvDataDocTbl2Excel, "trvDataDocTbl2Excel");
            this.trvDataDocTbl2Excel.Name = "trvDataDocTbl2Excel";
            this.trvDataDocTbl2Excel.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            ((System.Windows.Forms.TreeNode)(resources.GetObject("trvDataDocTbl2Excel.Nodes")))});
            this.trvDataDocTbl2Excel.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.trvDataDocTbl2Excel_NodeMouseClick);
            // 
            // btnClearItems
            // 
            resources.ApplyResources(this.btnClearItems, "btnClearItems");
            this.btnClearItems.Name = "btnClearItems";
            this.btnClearItems.UseVisualStyleBackColor = true;
            this.btnClearItems.Click += new System.EventHandler(this.btnClearItems_Click);
            // 
            // btnAllProduce
            // 
            resources.ApplyResources(this.btnAllProduce, "btnAllProduce");
            this.btnAllProduce.Name = "btnAllProduce";
            this.btnAllProduce.UseVisualStyleBackColor = true;
            this.btnAllProduce.Click += new System.EventHandler(this.btnAllProduce_Click);
            // 
            // btnPreviewProduce
            // 
            resources.ApplyResources(this.btnPreviewProduce, "btnPreviewProduce");
            this.btnPreviewProduce.Name = "btnPreviewProduce";
            this.btnPreviewProduce.UseVisualStyleBackColor = true;
            this.btnPreviewProduce.Click += new System.EventHandler(this.btnPreviewProduce_Click);
            // 
            // btnDocTbl2ExcelRemove
            // 
            resources.ApplyResources(this.btnDocTbl2ExcelRemove, "btnDocTbl2ExcelRemove");
            this.btnDocTbl2ExcelRemove.Name = "btnDocTbl2ExcelRemove";
            this.btnDocTbl2ExcelRemove.UseVisualStyleBackColor = true;
            this.btnDocTbl2ExcelRemove.Click += new System.EventHandler(this.btnDocTbl2ExcelRemove_Click);
            // 
            // btnAddColValue
            // 
            resources.ApplyResources(this.btnAddColValue, "btnAddColValue");
            this.btnAddColValue.Name = "btnAddColValue";
            this.btnAddColValue.UseVisualStyleBackColor = true;
            this.btnAddColValue.Click += new System.EventHandler(this.btnAddColValue_Click);
            // 
            // btnAddColName
            // 
            resources.ApplyResources(this.btnAddColName, "btnAddColName");
            this.btnAddColName.Name = "btnAddColName";
            this.btnAddColName.UseVisualStyleBackColor = true;
            this.btnAddColName.Click += new System.EventHandler(this.btnAddColName_Click);
            // 
            // tabPageExcel2DocTbl
            // 
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataProduce);
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataPreviewOne);
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataTagCombine);
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataRemoveRel);
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataInsertData);
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataInsertName);
            this.tabPageExcel2DocTbl.Controls.Add(this.btnDataDSource);
            this.tabPageExcel2DocTbl.Controls.Add(this.trvData);
            this.tabPageExcel2DocTbl.Controls.Add(this.label12);
            resources.ApplyResources(this.tabPageExcel2DocTbl, "tabPageExcel2DocTbl");
            this.tabPageExcel2DocTbl.Name = "tabPageExcel2DocTbl";
            this.tabPageExcel2DocTbl.UseVisualStyleBackColor = true;
            // 
            // btnDataProduce
            // 
            resources.ApplyResources(this.btnDataProduce, "btnDataProduce");
            this.btnDataProduce.Name = "btnDataProduce";
            this.btnDataProduce.UseVisualStyleBackColor = true;
            this.btnDataProduce.Click += new System.EventHandler(this.btnDataProduce_Click);
            // 
            // btnDataPreviewOne
            // 
            resources.ApplyResources(this.btnDataPreviewOne, "btnDataPreviewOne");
            this.btnDataPreviewOne.Name = "btnDataPreviewOne";
            this.btnDataPreviewOne.UseVisualStyleBackColor = true;
            this.btnDataPreviewOne.Click += new System.EventHandler(this.btnDataPreviewOne_Click);
            // 
            // btnDataTagCombine
            // 
            resources.ApplyResources(this.btnDataTagCombine, "btnDataTagCombine");
            this.btnDataTagCombine.Name = "btnDataTagCombine";
            this.btnDataTagCombine.UseVisualStyleBackColor = true;
            this.btnDataTagCombine.Click += new System.EventHandler(this.btnDataTagCombine_Click);
            // 
            // btnDataRemoveRel
            // 
            resources.ApplyResources(this.btnDataRemoveRel, "btnDataRemoveRel");
            this.btnDataRemoveRel.Name = "btnDataRemoveRel";
            this.btnDataRemoveRel.UseVisualStyleBackColor = true;
            this.btnDataRemoveRel.Click += new System.EventHandler(this.btnDataRemoveRel_Click);
            // 
            // btnDataInsertData
            // 
            resources.ApplyResources(this.btnDataInsertData, "btnDataInsertData");
            this.btnDataInsertData.Name = "btnDataInsertData";
            this.btnDataInsertData.UseVisualStyleBackColor = true;
            this.btnDataInsertData.Click += new System.EventHandler(this.btnDataInsertData_Click);
            // 
            // btnDataInsertName
            // 
            resources.ApplyResources(this.btnDataInsertName, "btnDataInsertName");
            this.btnDataInsertName.Name = "btnDataInsertName";
            this.btnDataInsertName.UseVisualStyleBackColor = true;
            this.btnDataInsertName.Click += new System.EventHandler(this.btnDataInsertName_Click);
            // 
            // btnDataDSource
            // 
            resources.ApplyResources(this.btnDataDSource, "btnDataDSource");
            this.btnDataDSource.Name = "btnDataDSource";
            this.btnDataDSource.UseVisualStyleBackColor = true;
            this.btnDataDSource.Click += new System.EventHandler(this.btnDataDSource_Click);
            // 
            // trvData
            // 
            resources.ApplyResources(this.trvData, "trvData");
            this.trvData.Name = "trvData";
            this.trvData.ShowNodeToolTips = true;
            this.trvData.Tag = "数据字段树";
            this.trvData.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trvData_AfterSelect);
            this.trvData.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.trvData_NodeMouseClick);
            this.trvData.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.trvData_NodeMouseDoubleClick);
            // 
            // label12
            // 
            resources.ApplyResources(this.label12, "label12");
            this.label12.Name = "label12";
            // 
            // tabPageFillGather
            // 
            this.tabPageFillGather.Controls.Add(this.btnFillGatherMoveDown);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherMoveUp);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherShowRowCol);
            this.tabPageFillGather.Controls.Add(this.progBarFG);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddUserDefineColName);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAllSelUnSel);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherDelFiles);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddFiles);
            this.tabPageFillGather.Controls.Add(this.chkBoxFillGatherStrictMatchColName);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherProduce);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherPreviewProduce);
            this.tabPageFillGather.Controls.Add(this.rdBtnFillGatherCurDoc);
            this.tabPageFillGather.Controls.Add(this.rdBtnFillGatherMultiFiles);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherViewLog);
            this.tabPageFillGather.Controls.Add(this.label42);
            this.tabPageFillGather.Controls.Add(this.chkListBoxTargetFiles);
            this.tabPageFillGather.Controls.Add(this.groupBox8);
            this.tabPageFillGather.Controls.Add(this.trvFillGatherSchemes);
            this.tabPageFillGather.Controls.Add(this.txtFillGatherName);
            this.tabPageFillGather.Controls.Add(this.label43);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherVerifyMatch);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherRemoveTblItem);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddTagNameValue);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddColValue);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddColName);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddTable);
            this.tabPageFillGather.Controls.Add(this.btnFillGatherAddScheme);
            this.tabPageFillGather.Controls.Add(this.label44);
            resources.ApplyResources(this.tabPageFillGather, "tabPageFillGather");
            this.tabPageFillGather.Name = "tabPageFillGather";
            this.tabPageFillGather.UseVisualStyleBackColor = true;
            // 
            // btnFillGatherMoveDown
            // 
            resources.ApplyResources(this.btnFillGatherMoveDown, "btnFillGatherMoveDown");
            this.btnFillGatherMoveDown.Name = "btnFillGatherMoveDown";
            this.btnFillGatherMoveDown.UseVisualStyleBackColor = true;
            this.btnFillGatherMoveDown.Click += new System.EventHandler(this.btnFillGatherMoveDown_Click);
            // 
            // btnFillGatherMoveUp
            // 
            resources.ApplyResources(this.btnFillGatherMoveUp, "btnFillGatherMoveUp");
            this.btnFillGatherMoveUp.Name = "btnFillGatherMoveUp";
            this.btnFillGatherMoveUp.UseVisualStyleBackColor = true;
            this.btnFillGatherMoveUp.Click += new System.EventHandler(this.btnFillGatherMoveUp_Click);
            // 
            // btnFillGatherShowRowCol
            // 
            resources.ApplyResources(this.btnFillGatherShowRowCol, "btnFillGatherShowRowCol");
            this.btnFillGatherShowRowCol.Name = "btnFillGatherShowRowCol";
            this.btnFillGatherShowRowCol.UseVisualStyleBackColor = true;
            this.btnFillGatherShowRowCol.Click += new System.EventHandler(this.btnFillGatherShowRowCol_Click);
            // 
            // progBarFG
            // 
            resources.ApplyResources(this.progBarFG, "progBarFG");
            this.progBarFG.Name = "progBarFG";
            // 
            // btnFillGatherAddUserDefineColName
            // 
            resources.ApplyResources(this.btnFillGatherAddUserDefineColName, "btnFillGatherAddUserDefineColName");
            this.btnFillGatherAddUserDefineColName.Name = "btnFillGatherAddUserDefineColName";
            this.btnFillGatherAddUserDefineColName.UseVisualStyleBackColor = true;
            this.btnFillGatherAddUserDefineColName.Click += new System.EventHandler(this.btnFillGatherAddUserDefineColName_Click);
            // 
            // btnFillGatherAllSelUnSel
            // 
            resources.ApplyResources(this.btnFillGatherAllSelUnSel, "btnFillGatherAllSelUnSel");
            this.btnFillGatherAllSelUnSel.Name = "btnFillGatherAllSelUnSel";
            this.btnFillGatherAllSelUnSel.UseVisualStyleBackColor = true;
            this.btnFillGatherAllSelUnSel.Click += new System.EventHandler(this.btnFillGatherAllSelUnSel_Click);
            // 
            // btnFillGatherDelFiles
            // 
            resources.ApplyResources(this.btnFillGatherDelFiles, "btnFillGatherDelFiles");
            this.btnFillGatherDelFiles.Name = "btnFillGatherDelFiles";
            this.btnFillGatherDelFiles.UseVisualStyleBackColor = true;
            this.btnFillGatherDelFiles.Click += new System.EventHandler(this.btnFillGatherDelFiles_Click);
            // 
            // btnFillGatherAddFiles
            // 
            resources.ApplyResources(this.btnFillGatherAddFiles, "btnFillGatherAddFiles");
            this.btnFillGatherAddFiles.Name = "btnFillGatherAddFiles";
            this.btnFillGatherAddFiles.UseVisualStyleBackColor = true;
            this.btnFillGatherAddFiles.Click += new System.EventHandler(this.btnFillGatherAddFiles_Click);
            // 
            // chkBoxFillGatherStrictMatchColName
            // 
            resources.ApplyResources(this.chkBoxFillGatherStrictMatchColName, "chkBoxFillGatherStrictMatchColName");
            this.chkBoxFillGatherStrictMatchColName.Name = "chkBoxFillGatherStrictMatchColName";
            this.chkBoxFillGatherStrictMatchColName.UseVisualStyleBackColor = true;
            // 
            // btnFillGatherProduce
            // 
            resources.ApplyResources(this.btnFillGatherProduce, "btnFillGatherProduce");
            this.btnFillGatherProduce.Name = "btnFillGatherProduce";
            this.btnFillGatherProduce.UseVisualStyleBackColor = true;
            this.btnFillGatherProduce.Click += new System.EventHandler(this.btnFillGatherProduce_Click);
            // 
            // btnFillGatherPreviewProduce
            // 
            resources.ApplyResources(this.btnFillGatherPreviewProduce, "btnFillGatherPreviewProduce");
            this.btnFillGatherPreviewProduce.Name = "btnFillGatherPreviewProduce";
            this.btnFillGatherPreviewProduce.UseVisualStyleBackColor = true;
            this.btnFillGatherPreviewProduce.Click += new System.EventHandler(this.btnFillGatherPreviewProduce_Click);
            // 
            // rdBtnFillGatherCurDoc
            // 
            resources.ApplyResources(this.rdBtnFillGatherCurDoc, "rdBtnFillGatherCurDoc");
            this.rdBtnFillGatherCurDoc.Name = "rdBtnFillGatherCurDoc";
            this.rdBtnFillGatherCurDoc.UseVisualStyleBackColor = true;
            // 
            // rdBtnFillGatherMultiFiles
            // 
            resources.ApplyResources(this.rdBtnFillGatherMultiFiles, "rdBtnFillGatherMultiFiles");
            this.rdBtnFillGatherMultiFiles.Checked = true;
            this.rdBtnFillGatherMultiFiles.Name = "rdBtnFillGatherMultiFiles";
            this.rdBtnFillGatherMultiFiles.TabStop = true;
            this.rdBtnFillGatherMultiFiles.UseVisualStyleBackColor = true;
            this.rdBtnFillGatherMultiFiles.CheckedChanged += new System.EventHandler(this.rdBtnFillGatherMultiFiles_CheckedChanged);
            // 
            // btnFillGatherViewLog
            // 
            resources.ApplyResources(this.btnFillGatherViewLog, "btnFillGatherViewLog");
            this.btnFillGatherViewLog.Name = "btnFillGatherViewLog";
            this.btnFillGatherViewLog.UseVisualStyleBackColor = true;
            this.btnFillGatherViewLog.Click += new System.EventHandler(this.btnFillGatherViewLog_Click);
            // 
            // label42
            // 
            resources.ApplyResources(this.label42, "label42");
            this.label42.Name = "label42";
            // 
            // chkListBoxTargetFiles
            // 
            this.chkListBoxTargetFiles.CheckOnClick = true;
            this.chkListBoxTargetFiles.FormattingEnabled = true;
            this.chkListBoxTargetFiles.Items.AddRange(new object[] {
            resources.GetString("chkListBoxTargetFiles.Items"),
            resources.GetString("chkListBoxTargetFiles.Items1"),
            resources.GetString("chkListBoxTargetFiles.Items2"),
            resources.GetString("chkListBoxTargetFiles.Items3")});
            resources.ApplyResources(this.chkListBoxTargetFiles, "chkListBoxTargetFiles");
            this.chkListBoxTargetFiles.Name = "chkListBoxTargetFiles";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.rdBtnFillGatherSelScope);
            this.groupBox8.Controls.Add(this.rdBtnFillGatherAllDocScope);
            resources.ApplyResources(this.groupBox8, "groupBox8");
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.TabStop = false;
            // 
            // rdBtnFillGatherSelScope
            // 
            resources.ApplyResources(this.rdBtnFillGatherSelScope, "rdBtnFillGatherSelScope");
            this.rdBtnFillGatherSelScope.Name = "rdBtnFillGatherSelScope";
            this.rdBtnFillGatherSelScope.UseVisualStyleBackColor = true;
            // 
            // rdBtnFillGatherAllDocScope
            // 
            resources.ApplyResources(this.rdBtnFillGatherAllDocScope, "rdBtnFillGatherAllDocScope");
            this.rdBtnFillGatherAllDocScope.Checked = true;
            this.rdBtnFillGatherAllDocScope.Name = "rdBtnFillGatherAllDocScope";
            this.rdBtnFillGatherAllDocScope.TabStop = true;
            this.rdBtnFillGatherAllDocScope.UseVisualStyleBackColor = true;
            // 
            // trvFillGatherSchemes
            // 
            resources.ApplyResources(this.trvFillGatherSchemes, "trvFillGatherSchemes");
            this.trvFillGatherSchemes.Name = "trvFillGatherSchemes";
            this.trvFillGatherSchemes.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            ((System.Windows.Forms.TreeNode)(resources.GetObject("trvFillGatherSchemes.Nodes"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("trvFillGatherSchemes.Nodes1")))});
            this.trvFillGatherSchemes.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trvFillGatherSchemes_AfterSelect);
            this.trvFillGatherSchemes.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.trvFillGatherSchemes_NodeMouseClick);
            // 
            // txtFillGatherName
            // 
            resources.ApplyResources(this.txtFillGatherName, "txtFillGatherName");
            this.txtFillGatherName.Name = "txtFillGatherName";
            // 
            // label43
            // 
            resources.ApplyResources(this.label43, "label43");
            this.label43.Name = "label43";
            // 
            // btnFillGatherVerifyMatch
            // 
            resources.ApplyResources(this.btnFillGatherVerifyMatch, "btnFillGatherVerifyMatch");
            this.btnFillGatherVerifyMatch.Name = "btnFillGatherVerifyMatch";
            this.btnFillGatherVerifyMatch.UseVisualStyleBackColor = true;
            this.btnFillGatherVerifyMatch.Click += new System.EventHandler(this.btnFillGatherVerifyMatch_Click);
            // 
            // btnFillGatherRemoveTblItem
            // 
            resources.ApplyResources(this.btnFillGatherRemoveTblItem, "btnFillGatherRemoveTblItem");
            this.btnFillGatherRemoveTblItem.Name = "btnFillGatherRemoveTblItem";
            this.btnFillGatherRemoveTblItem.UseVisualStyleBackColor = true;
            this.btnFillGatherRemoveTblItem.Click += new System.EventHandler(this.btnFillGatherRemoveTblItem_Click);
            // 
            // btnFillGatherAddTagNameValue
            // 
            resources.ApplyResources(this.btnFillGatherAddTagNameValue, "btnFillGatherAddTagNameValue");
            this.btnFillGatherAddTagNameValue.Name = "btnFillGatherAddTagNameValue";
            this.btnFillGatherAddTagNameValue.UseVisualStyleBackColor = true;
            this.btnFillGatherAddTagNameValue.Click += new System.EventHandler(this.btnFillGatherAddTagNameValue_Click);
            // 
            // btnFillGatherAddColValue
            // 
            resources.ApplyResources(this.btnFillGatherAddColValue, "btnFillGatherAddColValue");
            this.btnFillGatherAddColValue.Name = "btnFillGatherAddColValue";
            this.btnFillGatherAddColValue.UseVisualStyleBackColor = true;
            this.btnFillGatherAddColValue.Click += new System.EventHandler(this.btnFillGatherAddColValue_Click);
            // 
            // btnFillGatherAddColName
            // 
            resources.ApplyResources(this.btnFillGatherAddColName, "btnFillGatherAddColName");
            this.btnFillGatherAddColName.Name = "btnFillGatherAddColName";
            this.btnFillGatherAddColName.UseVisualStyleBackColor = true;
            this.btnFillGatherAddColName.Click += new System.EventHandler(this.btnFillGatherAddColName_Click);
            // 
            // btnFillGatherAddTable
            // 
            resources.ApplyResources(this.btnFillGatherAddTable, "btnFillGatherAddTable");
            this.btnFillGatherAddTable.Name = "btnFillGatherAddTable";
            this.btnFillGatherAddTable.UseVisualStyleBackColor = true;
            this.btnFillGatherAddTable.Click += new System.EventHandler(this.btnFillGatherAddTable_Click);
            // 
            // btnFillGatherAddScheme
            // 
            resources.ApplyResources(this.btnFillGatherAddScheme, "btnFillGatherAddScheme");
            this.btnFillGatherAddScheme.Name = "btnFillGatherAddScheme";
            this.btnFillGatherAddScheme.UseVisualStyleBackColor = true;
            this.btnFillGatherAddScheme.Click += new System.EventHandler(this.btnFillGatherAddScheme_Click);
            // 
            // label44
            // 
            resources.ApplyResources(this.label44, "label44");
            this.label44.Name = "label44";
            // 
            // tabPageCntList
            // 
            this.tabPageCntList.Controls.Add(this.btnCntListExpand);
            this.tabPageCntList.Controls.Add(this.btnCntListCollapse);
            this.tabPageCntList.Controls.Add(this.progBarCntList);
            this.tabPageCntList.Controls.Add(this.trvCntList);
            this.tabPageCntList.Controls.Add(this.btnCntListCover);
            this.tabPageCntList.Controls.Add(this.btnCntListRef);
            this.tabPageCntList.Controls.Add(this.btnCntListRemove);
            this.tabPageCntList.Controls.Add(this.btnCntListAddDoc);
            this.tabPageCntList.Controls.Add(this.txtBoxCntListFile);
            resources.ApplyResources(this.tabPageCntList, "tabPageCntList");
            this.tabPageCntList.Name = "tabPageCntList";
            this.tabPageCntList.UseVisualStyleBackColor = true;
            // 
            // btnCntListExpand
            // 
            resources.ApplyResources(this.btnCntListExpand, "btnCntListExpand");
            this.btnCntListExpand.Name = "btnCntListExpand";
            this.btnCntListExpand.UseVisualStyleBackColor = true;
            this.btnCntListExpand.Click += new System.EventHandler(this.btnCntListExpand_Click);
            // 
            // btnCntListCollapse
            // 
            resources.ApplyResources(this.btnCntListCollapse, "btnCntListCollapse");
            this.btnCntListCollapse.Name = "btnCntListCollapse";
            this.btnCntListCollapse.UseVisualStyleBackColor = true;
            this.btnCntListCollapse.Click += new System.EventHandler(this.btnCntListCollapse_Click);
            // 
            // progBarCntList
            // 
            resources.ApplyResources(this.progBarCntList, "progBarCntList");
            this.progBarCntList.Name = "progBarCntList";
            // 
            // trvCntList
            // 
            this.trvCntList.FullRowSelect = true;
            this.trvCntList.HideSelection = false;
            this.trvCntList.HotTracking = true;
            resources.ApplyResources(this.trvCntList, "trvCntList");
            this.trvCntList.ImageList = this.imageListIcon;
            this.trvCntList.Name = "trvCntList";
            this.trvCntList.Tag = "章节及内容块树";
            this.trvCntList.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.trvCntList_NodeMouseDoubleClick);
            // 
            // btnCntListCover
            // 
            resources.ApplyResources(this.btnCntListCover, "btnCntListCover");
            this.btnCntListCover.Name = "btnCntListCover";
            this.btnCntListCover.UseVisualStyleBackColor = true;
            this.btnCntListCover.Click += new System.EventHandler(this.btnCntListCover_Click);
            // 
            // btnCntListRef
            // 
            resources.ApplyResources(this.btnCntListRef, "btnCntListRef");
            this.btnCntListRef.Name = "btnCntListRef";
            this.btnCntListRef.UseVisualStyleBackColor = true;
            this.btnCntListRef.Click += new System.EventHandler(this.btnCntListRef_Click);
            // 
            // btnCntListRemove
            // 
            resources.ApplyResources(this.btnCntListRemove, "btnCntListRemove");
            this.btnCntListRemove.Name = "btnCntListRemove";
            this.btnCntListRemove.UseVisualStyleBackColor = true;
            this.btnCntListRemove.Click += new System.EventHandler(this.btnCntListRemove_Click);
            // 
            // btnCntListAddDoc
            // 
            resources.ApplyResources(this.btnCntListAddDoc, "btnCntListAddDoc");
            this.btnCntListAddDoc.Name = "btnCntListAddDoc";
            this.btnCntListAddDoc.UseVisualStyleBackColor = true;
            this.btnCntListAddDoc.Click += new System.EventHandler(this.btnCntListAddDoc_Click);
            // 
            // txtBoxCntListFile
            // 
            resources.ApplyResources(this.txtBoxCntListFile, "txtBoxCntListFile");
            this.txtBoxCntListFile.Name = "txtBoxCntListFile";
            this.txtBoxCntListFile.ReadOnly = true;
            // 
            // tabPageForm
            // 
            resources.ApplyResources(this.tabPageForm, "tabPageForm");
            this.tabPageForm.Controls.Add(this.label14);
            this.tabPageForm.Controls.Add(this.btnFormNextSearch);
            this.tabPageForm.Controls.Add(this.tblFormLayoutPanel);
            this.tabPageForm.Controls.Add(this.btnFormRefresh);
            this.tabPageForm.Controls.Add(this.btnFormPrevSearch);
            this.tabPageForm.Controls.Add(this.txtFormKeyWord);
            this.tabPageForm.Controls.Add(this.btnFormSearch);
            this.tabPageForm.Controls.Add(this.btnFormReset);
            this.tabPageForm.Name = "tabPageForm";
            this.tabPageForm.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            resources.ApplyResources(this.label14, "label14");
            this.label14.Name = "label14";
            // 
            // btnFormNextSearch
            // 
            resources.ApplyResources(this.btnFormNextSearch, "btnFormNextSearch");
            this.btnFormNextSearch.Name = "btnFormNextSearch";
            this.btnFormNextSearch.UseVisualStyleBackColor = true;
            this.btnFormNextSearch.Click += new System.EventHandler(this.btnFormNextSearch_Click);
            // 
            // tblFormLayoutPanel
            // 
            resources.ApplyResources(this.tblFormLayoutPanel, "tblFormLayoutPanel");
            this.tblFormLayoutPanel.Name = "tblFormLayoutPanel";
            this.tblFormLayoutPanel.Tag = "表单域树";
            // 
            // btnFormRefresh
            // 
            resources.ApplyResources(this.btnFormRefresh, "btnFormRefresh");
            this.btnFormRefresh.Name = "btnFormRefresh";
            this.btnFormRefresh.UseVisualStyleBackColor = true;
            this.btnFormRefresh.Click += new System.EventHandler(this.btnFormRefresh_Click);
            // 
            // btnFormPrevSearch
            // 
            resources.ApplyResources(this.btnFormPrevSearch, "btnFormPrevSearch");
            this.btnFormPrevSearch.Name = "btnFormPrevSearch";
            this.btnFormPrevSearch.UseVisualStyleBackColor = true;
            this.btnFormPrevSearch.Click += new System.EventHandler(this.btnFormPrevSearch_Click);
            // 
            // txtFormKeyWord
            // 
            resources.ApplyResources(this.txtFormKeyWord, "txtFormKeyWord");
            this.txtFormKeyWord.Name = "txtFormKeyWord";
            // 
            // btnFormSearch
            // 
            resources.ApplyResources(this.btnFormSearch, "btnFormSearch");
            this.btnFormSearch.Name = "btnFormSearch";
            this.btnFormSearch.UseVisualStyleBackColor = true;
            this.btnFormSearch.Click += new System.EventHandler(this.btnFormSearch_Click);
            // 
            // btnFormReset
            // 
            resources.ApplyResources(this.btnFormReset, "btnFormReset");
            this.btnFormReset.Name = "btnFormReset";
            this.btnFormReset.UseVisualStyleBackColor = true;
            this.btnFormReset.Click += new System.EventHandler(this.btnFormReset_Click);
            // 
            // tabPageInfo
            // 
            this.tabPageInfo.Controls.Add(this.btnInfoRefresh);
            this.tabPageInfo.Controls.Add(this.txtInfoBody);
            resources.ApplyResources(this.tabPageInfo, "tabPageInfo");
            this.tabPageInfo.Name = "tabPageInfo";
            this.tabPageInfo.UseVisualStyleBackColor = true;
            // 
            // btnInfoRefresh
            // 
            resources.ApplyResources(this.btnInfoRefresh, "btnInfoRefresh");
            this.btnInfoRefresh.Name = "btnInfoRefresh";
            this.btnInfoRefresh.UseVisualStyleBackColor = true;
            this.btnInfoRefresh.Click += new System.EventHandler(this.btnInfoRefresh_Click);
            // 
            // txtInfoBody
            // 
            resources.ApplyResources(this.txtInfoBody, "txtInfoBody");
            this.txtInfoBody.Name = "txtInfoBody";
            this.txtInfoBody.ReadOnly = true;
            this.txtInfoBody.Tag = "基本信息";
            // 
            // tabPageNumTrans
            // 
            this.tabPageNumTrans.Controls.Add(this.btnNumTransClear);
            this.tabPageNumTrans.Controls.Add(this.btnNumTrans);
            this.tabPageNumTrans.Controls.Add(this.txtMoneySimpBigTbl);
            this.tabPageNumTrans.Controls.Add(this.label27);
            this.tabPageNumTrans.Controls.Add(this.txtMoneySimpBig);
            this.tabPageNumTrans.Controls.Add(this.label25);
            this.tabPageNumTrans.Controls.Add(this.txtNumValueSimpBigTbl);
            this.tabPageNumTrans.Controls.Add(this.label29);
            this.tabPageNumTrans.Controls.Add(this.txtNumValueSimpBig);
            this.tabPageNumTrans.Controls.Add(this.label21);
            this.tabPageNumTrans.Controls.Add(this.txtDigitNumSimpBig);
            this.tabPageNumTrans.Controls.Add(this.label24);
            this.tabPageNumTrans.Controls.Add(this.label26);
            this.tabPageNumTrans.Controls.Add(this.label17);
            this.tabPageNumTrans.Controls.Add(this.label23);
            this.tabPageNumTrans.Controls.Add(this.label28);
            this.tabPageNumTrans.Controls.Add(this.label18);
            this.tabPageNumTrans.Controls.Add(this.label20);
            this.tabPageNumTrans.Controls.Add(this.label22);
            this.tabPageNumTrans.Controls.Add(this.label16);
            this.tabPageNumTrans.Controls.Add(this.txtMoneySimpLittleTbl);
            this.tabPageNumTrans.Controls.Add(this.label19);
            this.tabPageNumTrans.Controls.Add(this.txtMoneySimpLittle);
            this.tabPageNumTrans.Controls.Add(this.txtNumValueSimpLittleTbl);
            this.tabPageNumTrans.Controls.Add(this.label15);
            this.tabPageNumTrans.Controls.Add(this.txtNumValueSimpLittle);
            this.tabPageNumTrans.Controls.Add(this.txtNumMoney);
            this.tabPageNumTrans.Controls.Add(this.label8);
            this.tabPageNumTrans.Controls.Add(this.txtNumValue);
            this.tabPageNumTrans.Controls.Add(this.txtDigitNumSimpLittle);
            this.tabPageNumTrans.Controls.Add(this.txtDigitNum);
            resources.ApplyResources(this.tabPageNumTrans, "tabPageNumTrans");
            this.tabPageNumTrans.Name = "tabPageNumTrans";
            this.tabPageNumTrans.UseVisualStyleBackColor = true;
            // 
            // btnNumTransClear
            // 
            resources.ApplyResources(this.btnNumTransClear, "btnNumTransClear");
            this.btnNumTransClear.Name = "btnNumTransClear";
            this.btnNumTransClear.UseVisualStyleBackColor = true;
            this.btnNumTransClear.Click += new System.EventHandler(this.btnNumTransClear_Click);
            // 
            // btnNumTrans
            // 
            resources.ApplyResources(this.btnNumTrans, "btnNumTrans");
            this.btnNumTrans.Name = "btnNumTrans";
            this.btnNumTrans.UseVisualStyleBackColor = true;
            this.btnNumTrans.Click += new System.EventHandler(this.btnNumTrans_Click);
            // 
            // txtMoneySimpBigTbl
            // 
            resources.ApplyResources(this.txtMoneySimpBigTbl, "txtMoneySimpBigTbl");
            this.txtMoneySimpBigTbl.Name = "txtMoneySimpBigTbl";
            this.txtMoneySimpBigTbl.Tag = "中文大写金额（填表）";
            // 
            // label27
            // 
            resources.ApplyResources(this.label27, "label27");
            this.label27.Name = "label27";
            // 
            // txtMoneySimpBig
            // 
            resources.ApplyResources(this.txtMoneySimpBig, "txtMoneySimpBig");
            this.txtMoneySimpBig.Name = "txtMoneySimpBig";
            this.txtMoneySimpBig.Tag = "中文大写金额";
            // 
            // label25
            // 
            resources.ApplyResources(this.label25, "label25");
            this.label25.Name = "label25";
            // 
            // txtNumValueSimpBigTbl
            // 
            resources.ApplyResources(this.txtNumValueSimpBigTbl, "txtNumValueSimpBigTbl");
            this.txtNumValueSimpBigTbl.Name = "txtNumValueSimpBigTbl";
            this.txtNumValueSimpBigTbl.Tag = "中文小写数值（填表）";
            // 
            // label29
            // 
            resources.ApplyResources(this.label29, "label29");
            this.label29.Name = "label29";
            // 
            // txtNumValueSimpBig
            // 
            resources.ApplyResources(this.txtNumValueSimpBig, "txtNumValueSimpBig");
            this.txtNumValueSimpBig.Name = "txtNumValueSimpBig";
            this.txtNumValueSimpBig.Tag = "中文大写数值";
            // 
            // label21
            // 
            resources.ApplyResources(this.label21, "label21");
            this.label21.Name = "label21";
            // 
            // txtDigitNumSimpBig
            // 
            resources.ApplyResources(this.txtDigitNumSimpBig, "txtDigitNumSimpBig");
            this.txtDigitNumSimpBig.Name = "txtDigitNumSimpBig";
            this.txtDigitNumSimpBig.Tag = "中文大写数字";
            // 
            // label24
            // 
            resources.ApplyResources(this.label24, "label24");
            this.label24.Name = "label24";
            // 
            // label26
            // 
            resources.ApplyResources(this.label26, "label26");
            this.label26.Name = "label26";
            // 
            // label17
            // 
            resources.ApplyResources(this.label17, "label17");
            this.label17.Name = "label17";
            // 
            // label23
            // 
            resources.ApplyResources(this.label23, "label23");
            this.label23.Name = "label23";
            // 
            // label28
            // 
            resources.ApplyResources(this.label28, "label28");
            this.label28.Name = "label28";
            // 
            // label18
            // 
            resources.ApplyResources(this.label18, "label18");
            this.label18.Name = "label18";
            // 
            // label20
            // 
            resources.ApplyResources(this.label20, "label20");
            this.label20.Name = "label20";
            // 
            // label22
            // 
            resources.ApplyResources(this.label22, "label22");
            this.label22.Name = "label22";
            // 
            // label16
            // 
            resources.ApplyResources(this.label16, "label16");
            this.label16.Name = "label16";
            // 
            // txtMoneySimpLittleTbl
            // 
            resources.ApplyResources(this.txtMoneySimpLittleTbl, "txtMoneySimpLittleTbl");
            this.txtMoneySimpLittleTbl.Name = "txtMoneySimpLittleTbl";
            this.txtMoneySimpLittleTbl.Tag = "中文小写金额（填表）";
            // 
            // label19
            // 
            resources.ApplyResources(this.label19, "label19");
            this.label19.Name = "label19";
            // 
            // txtMoneySimpLittle
            // 
            resources.ApplyResources(this.txtMoneySimpLittle, "txtMoneySimpLittle");
            this.txtMoneySimpLittle.Name = "txtMoneySimpLittle";
            this.txtMoneySimpLittle.Tag = "中文小写金额";
            // 
            // txtNumValueSimpLittleTbl
            // 
            resources.ApplyResources(this.txtNumValueSimpLittleTbl, "txtNumValueSimpLittleTbl");
            this.txtNumValueSimpLittleTbl.Name = "txtNumValueSimpLittleTbl";
            this.txtNumValueSimpLittleTbl.Tag = "中文小写数值（填表）";
            // 
            // label15
            // 
            resources.ApplyResources(this.label15, "label15");
            this.label15.Name = "label15";
            // 
            // txtNumValueSimpLittle
            // 
            resources.ApplyResources(this.txtNumValueSimpLittle, "txtNumValueSimpLittle");
            this.txtNumValueSimpLittle.Name = "txtNumValueSimpLittle";
            this.txtNumValueSimpLittle.Tag = "中文小写数值";
            // 
            // txtNumMoney
            // 
            resources.ApplyResources(this.txtNumMoney, "txtNumMoney");
            this.txtNumMoney.Name = "txtNumMoney";
            this.txtNumMoney.Tag = "金额数额";
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            // 
            // txtNumValue
            // 
            resources.ApplyResources(this.txtNumValue, "txtNumValue");
            this.txtNumValue.Name = "txtNumValue";
            this.txtNumValue.Tag = "数值";
            // 
            // txtDigitNumSimpLittle
            // 
            resources.ApplyResources(this.txtDigitNumSimpLittle, "txtDigitNumSimpLittle");
            this.txtDigitNumSimpLittle.Name = "txtDigitNumSimpLittle";
            this.txtDigitNumSimpLittle.Tag = "中文小写数字";
            // 
            // txtDigitNum
            // 
            resources.ApplyResources(this.txtDigitNum, "txtDigitNum");
            this.txtDigitNum.Name = "txtDigitNum";
            this.txtDigitNum.Tag = "数字digit";
            // 
            // tabPageHeadingSn
            // 
            this.tabPageHeadingSn.Controls.Add(this.label33);
            this.tabPageHeadingSn.Controls.Add(this.btnExitHeadingSnApply);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnPreview);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnReset);
            this.tabPageHeadingSn.Controls.Add(this.chkHeadingSnReserveCurStyle);
            this.tabPageHeadingSn.Controls.Add(this.progBarHeadingSn);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnNameGen);
            this.tabPageHeadingSn.Controls.Add(this.trvHeadingSnScheme);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnSchemeApply);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnSchemeGet);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnSchemeUpdate);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnSchemeDel);
            this.tabPageHeadingSn.Controls.Add(this.btnHeadingSnSchemeAdd);
            this.tabPageHeadingSn.Controls.Add(this.groupBox1);
            this.tabPageHeadingSn.Controls.Add(this.label30);
            this.tabPageHeadingSn.Controls.Add(this.txtHeadingSnSchemeName);
            this.tabPageHeadingSn.Controls.Add(this.label40);
            resources.ApplyResources(this.tabPageHeadingSn, "tabPageHeadingSn");
            this.tabPageHeadingSn.Name = "tabPageHeadingSn";
            this.tabPageHeadingSn.UseVisualStyleBackColor = true;
            this.tabPageHeadingSn.Enter += new System.EventHandler(this.tabPageHeadingSn_Enter);
            // 
            // label33
            // 
            resources.ApplyResources(this.label33, "label33");
            this.label33.Name = "label33";
            // 
            // btnExitHeadingSnApply
            // 
            resources.ApplyResources(this.btnExitHeadingSnApply, "btnExitHeadingSnApply");
            this.btnExitHeadingSnApply.Name = "btnExitHeadingSnApply";
            this.btnExitHeadingSnApply.UseVisualStyleBackColor = true;
            this.btnExitHeadingSnApply.Click += new System.EventHandler(this.btnExitHeadingSnApply_Click);
            // 
            // btnHeadingSnPreview
            // 
            resources.ApplyResources(this.btnHeadingSnPreview, "btnHeadingSnPreview");
            this.btnHeadingSnPreview.Name = "btnHeadingSnPreview";
            this.btnHeadingSnPreview.UseVisualStyleBackColor = true;
            this.btnHeadingSnPreview.Click += new System.EventHandler(this.btnHeadingSnPreview_Click);
            // 
            // btnHeadingSnReset
            // 
            resources.ApplyResources(this.btnHeadingSnReset, "btnHeadingSnReset");
            this.btnHeadingSnReset.Name = "btnHeadingSnReset";
            this.btnHeadingSnReset.UseVisualStyleBackColor = true;
            this.btnHeadingSnReset.Click += new System.EventHandler(this.btnHeadingSnReset_Click);
            // 
            // chkHeadingSnReserveCurStyle
            // 
            resources.ApplyResources(this.chkHeadingSnReserveCurStyle, "chkHeadingSnReserveCurStyle");
            this.chkHeadingSnReserveCurStyle.Checked = true;
            this.chkHeadingSnReserveCurStyle.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkHeadingSnReserveCurStyle.Name = "chkHeadingSnReserveCurStyle";
            this.chkHeadingSnReserveCurStyle.UseVisualStyleBackColor = true;
            // 
            // progBarHeadingSn
            // 
            resources.ApplyResources(this.progBarHeadingSn, "progBarHeadingSn");
            this.progBarHeadingSn.Name = "progBarHeadingSn";
            // 
            // btnHeadingSnNameGen
            // 
            resources.ApplyResources(this.btnHeadingSnNameGen, "btnHeadingSnNameGen");
            this.btnHeadingSnNameGen.Name = "btnHeadingSnNameGen";
            this.btnHeadingSnNameGen.UseVisualStyleBackColor = true;
            this.btnHeadingSnNameGen.Click += new System.EventHandler(this.btnHeadingSnNameGen_Click);
            // 
            // trvHeadingSnScheme
            // 
            this.trvHeadingSnScheme.ContextMenuStrip = this.cxtMenuHeadingSn;
            resources.ApplyResources(this.trvHeadingSnScheme, "trvHeadingSnScheme");
            this.trvHeadingSnScheme.ImageList = this.imageListIcon;
            this.trvHeadingSnScheme.Name = "trvHeadingSnScheme";
            this.trvHeadingSnScheme.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            ((System.Windows.Forms.TreeNode)(resources.GetObject("trvHeadingSnScheme.Nodes"))),
            ((System.Windows.Forms.TreeNode)(resources.GetObject("trvHeadingSnScheme.Nodes1")))});
            this.trvHeadingSnScheme.ShowNodeToolTips = true;
            this.trvHeadingSnScheme.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trvHeadingSnScheme_AfterSelect);
            this.trvHeadingSnScheme.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.trvHeadingSnScheme_MouseDoubleClick);
            // 
            // cxtMenuHeadingSn
            // 
            this.cxtMenuHeadingSn.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cxtMenuItemPreview});
            this.cxtMenuHeadingSn.Name = "cxtMenuHeadingSn";
            resources.ApplyResources(this.cxtMenuHeadingSn, "cxtMenuHeadingSn");
            // 
            // cxtMenuItemPreview
            // 
            this.cxtMenuItemPreview.Name = "cxtMenuItemPreview";
            resources.ApplyResources(this.cxtMenuItemPreview, "cxtMenuItemPreview");
            this.cxtMenuItemPreview.Click += new System.EventHandler(this.cxtMenuItemPreview_Click);
            // 
            // btnHeadingSnSchemeApply
            // 
            resources.ApplyResources(this.btnHeadingSnSchemeApply, "btnHeadingSnSchemeApply");
            this.btnHeadingSnSchemeApply.Name = "btnHeadingSnSchemeApply";
            this.btnHeadingSnSchemeApply.UseVisualStyleBackColor = true;
            this.btnHeadingSnSchemeApply.Click += new System.EventHandler(this.btnHeadingSnSchemeApply_Click);
            // 
            // btnHeadingSnSchemeGet
            // 
            resources.ApplyResources(this.btnHeadingSnSchemeGet, "btnHeadingSnSchemeGet");
            this.btnHeadingSnSchemeGet.Name = "btnHeadingSnSchemeGet";
            this.btnHeadingSnSchemeGet.UseVisualStyleBackColor = true;
            this.btnHeadingSnSchemeGet.Click += new System.EventHandler(this.btnHeadingSnSchemeGet_Click);
            // 
            // btnHeadingSnSchemeUpdate
            // 
            resources.ApplyResources(this.btnHeadingSnSchemeUpdate, "btnHeadingSnSchemeUpdate");
            this.btnHeadingSnSchemeUpdate.Name = "btnHeadingSnSchemeUpdate";
            this.btnHeadingSnSchemeUpdate.UseVisualStyleBackColor = true;
            this.btnHeadingSnSchemeUpdate.Click += new System.EventHandler(this.btnHeadingSnSchemeUpdate_Click);
            // 
            // btnHeadingSnSchemeDel
            // 
            resources.ApplyResources(this.btnHeadingSnSchemeDel, "btnHeadingSnSchemeDel");
            this.btnHeadingSnSchemeDel.Name = "btnHeadingSnSchemeDel";
            this.btnHeadingSnSchemeDel.UseVisualStyleBackColor = true;
            this.btnHeadingSnSchemeDel.Click += new System.EventHandler(this.btnHeadingSnSchemeDel_Click);
            // 
            // btnHeadingSnSchemeAdd
            // 
            resources.ApplyResources(this.btnHeadingSnSchemeAdd, "btnHeadingSnSchemeAdd");
            this.btnHeadingSnSchemeAdd.Name = "btnHeadingSnSchemeAdd";
            this.btnHeadingSnSchemeAdd.UseVisualStyleBackColor = true;
            this.btnHeadingSnSchemeAdd.Click += new System.EventHandler(this.btnHeadingSnSchemeAdd_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnHeadingSnFont);
            this.groupBox1.Controls.Add(this.richTxtHeadingSnPreview);
            this.groupBox1.Controls.Add(this.btnHeadingSnSetDefaultFont);
            this.groupBox1.Controls.Add(this.btnHeadingSnFontExtract);
            this.groupBox1.Controls.Add(this.btnHeadingSnPos);
            this.groupBox1.Controls.Add(this.lstHeadingSnLevel);
            this.groupBox1.Controls.Add(this.chkHeadingSnLegal);
            this.groupBox1.Controls.Add(this.cmbSnShowStyle);
            this.groupBox1.Controls.Add(this.label32);
            this.groupBox1.Controls.Add(this.txtSnDefInput);
            this.groupBox1.Controls.Add(this.label31);
            this.groupBox1.Controls.Add(this.label34);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // btnHeadingSnFont
            // 
            resources.ApplyResources(this.btnHeadingSnFont, "btnHeadingSnFont");
            this.btnHeadingSnFont.Name = "btnHeadingSnFont";
            this.btnHeadingSnFont.UseVisualStyleBackColor = true;
            this.btnHeadingSnFont.Click += new System.EventHandler(this.btnHeadingSnFont_Click);
            // 
            // richTxtHeadingSnPreview
            // 
            resources.ApplyResources(this.richTxtHeadingSnPreview, "richTxtHeadingSnPreview");
            this.richTxtHeadingSnPreview.Name = "richTxtHeadingSnPreview";
            this.richTxtHeadingSnPreview.ReadOnly = true;
            // 
            // btnHeadingSnSetDefaultFont
            // 
            resources.ApplyResources(this.btnHeadingSnSetDefaultFont, "btnHeadingSnSetDefaultFont");
            this.btnHeadingSnSetDefaultFont.Name = "btnHeadingSnSetDefaultFont";
            this.btnHeadingSnSetDefaultFont.UseVisualStyleBackColor = true;
            this.btnHeadingSnSetDefaultFont.Click += new System.EventHandler(this.btnHeadingSnSetDefaultFont_Click);
            // 
            // btnHeadingSnFontExtract
            // 
            resources.ApplyResources(this.btnHeadingSnFontExtract, "btnHeadingSnFontExtract");
            this.btnHeadingSnFontExtract.Name = "btnHeadingSnFontExtract";
            this.btnHeadingSnFontExtract.UseVisualStyleBackColor = true;
            this.btnHeadingSnFontExtract.Click += new System.EventHandler(this.btnHeadingSnFontExtract_Click);
            // 
            // btnHeadingSnPos
            // 
            resources.ApplyResources(this.btnHeadingSnPos, "btnHeadingSnPos");
            this.btnHeadingSnPos.Name = "btnHeadingSnPos";
            this.btnHeadingSnPos.UseVisualStyleBackColor = true;
            this.btnHeadingSnPos.Click += new System.EventHandler(this.btnHeadingSnPos_Click);
            // 
            // lstHeadingSnLevel
            // 
            this.lstHeadingSnLevel.BackColor = System.Drawing.SystemColors.Window;
            this.lstHeadingSnLevel.FormattingEnabled = true;
            resources.ApplyResources(this.lstHeadingSnLevel, "lstHeadingSnLevel");
            this.lstHeadingSnLevel.Items.AddRange(new object[] {
            resources.GetString("lstHeadingSnLevel.Items"),
            resources.GetString("lstHeadingSnLevel.Items1"),
            resources.GetString("lstHeadingSnLevel.Items2"),
            resources.GetString("lstHeadingSnLevel.Items3"),
            resources.GetString("lstHeadingSnLevel.Items4"),
            resources.GetString("lstHeadingSnLevel.Items5"),
            resources.GetString("lstHeadingSnLevel.Items6"),
            resources.GetString("lstHeadingSnLevel.Items7"),
            resources.GetString("lstHeadingSnLevel.Items8")});
            this.lstHeadingSnLevel.Name = "lstHeadingSnLevel";
            this.lstHeadingSnLevel.Tag = "章节序号大纲级别";
            this.lstHeadingSnLevel.SelectedIndexChanged += new System.EventHandler(this.lstHeadingSnLevel_SelectedIndexChanged);
            // 
            // chkHeadingSnLegal
            // 
            resources.ApplyResources(this.chkHeadingSnLegal, "chkHeadingSnLegal");
            this.chkHeadingSnLegal.Name = "chkHeadingSnLegal";
            this.chkHeadingSnLegal.UseVisualStyleBackColor = true;
            this.chkHeadingSnLegal.CheckedChanged += new System.EventHandler(this.chkHeadingSnLegal_CheckedChanged);
            // 
            // cmbSnShowStyle
            // 
            this.cmbSnShowStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSnShowStyle.FormattingEnabled = true;
            this.cmbSnShowStyle.Items.AddRange(new object[] {
            resources.GetString("cmbSnShowStyle.Items"),
            resources.GetString("cmbSnShowStyle.Items1"),
            resources.GetString("cmbSnShowStyle.Items2"),
            resources.GetString("cmbSnShowStyle.Items3"),
            resources.GetString("cmbSnShowStyle.Items4"),
            resources.GetString("cmbSnShowStyle.Items5"),
            resources.GetString("cmbSnShowStyle.Items6"),
            resources.GetString("cmbSnShowStyle.Items7"),
            resources.GetString("cmbSnShowStyle.Items8"),
            resources.GetString("cmbSnShowStyle.Items9"),
            resources.GetString("cmbSnShowStyle.Items10"),
            resources.GetString("cmbSnShowStyle.Items11"),
            resources.GetString("cmbSnShowStyle.Items12"),
            resources.GetString("cmbSnShowStyle.Items13")});
            resources.ApplyResources(this.cmbSnShowStyle, "cmbSnShowStyle");
            this.cmbSnShowStyle.Name = "cmbSnShowStyle";
            this.cmbSnShowStyle.Tag = "章节序号显示样式";
            this.cmbSnShowStyle.SelectedIndexChanged += new System.EventHandler(this.cmbSnShowStyle_SelectedIndexChanged);
            // 
            // label32
            // 
            resources.ApplyResources(this.label32, "label32");
            this.label32.Name = "label32";
            // 
            // txtSnDefInput
            // 
            resources.ApplyResources(this.txtSnDefInput, "txtSnDefInput");
            this.txtSnDefInput.Name = "txtSnDefInput";
            this.txtSnDefInput.Tag = "章节序号格式输入框";
            this.txtSnDefInput.Leave += new System.EventHandler(this.txtSnDefInput_Leave);
            // 
            // label31
            // 
            resources.ApplyResources(this.label31, "label31");
            this.label31.Name = "label31";
            // 
            // label34
            // 
            resources.ApplyResources(this.label34, "label34");
            this.label34.Name = "label34";
            // 
            // label30
            // 
            resources.ApplyResources(this.label30, "label30");
            this.label30.Name = "label30";
            // 
            // txtHeadingSnSchemeName
            // 
            resources.ApplyResources(this.txtHeadingSnSchemeName, "txtHeadingSnSchemeName");
            this.txtHeadingSnSchemeName.Name = "txtHeadingSnSchemeName";
            this.txtHeadingSnSchemeName.Tag = "章节序号方案名称";
            // 
            // label40
            // 
            resources.ApplyResources(this.label40, "label40");
            this.label40.Name = "label40";
            // 
            // tabPageHeadingStyles
            // 
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleApplyCurSel);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleExitApply);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleApplyScope);
            this.tabPageHeadingStyles.Controls.Add(this.richHeadingStylePreview);
            this.tabPageHeadingStyles.Controls.Add(this.lstOutlineLevel);
            this.tabPageHeadingStyles.Controls.Add(this.prgbarHeadingStyleSchemeApply);
            this.tabPageHeadingStyles.Controls.Add(this.txtHeadingStyleSchemeName);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleSchemeApply);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleSchemeExtract);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleSchemePreview);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleSchemeUpdate);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleSchemeDel);
            this.tabPageHeadingStyles.Controls.Add(this.btnHeadingStyleSchemeAdd);
            this.tabPageHeadingStyles.Controls.Add(this.trvHeadingStyleScheme);
            this.tabPageHeadingStyles.Controls.Add(this.label41);
            resources.ApplyResources(this.tabPageHeadingStyles, "tabPageHeadingStyles");
            this.tabPageHeadingStyles.Name = "tabPageHeadingStyles";
            this.tabPageHeadingStyles.UseVisualStyleBackColor = true;
            this.tabPageHeadingStyles.Enter += new System.EventHandler(this.tabPageHeadingStyles_Enter);
            // 
            // btnHeadingStyleApplyCurSel
            // 
            resources.ApplyResources(this.btnHeadingStyleApplyCurSel, "btnHeadingStyleApplyCurSel");
            this.btnHeadingStyleApplyCurSel.Name = "btnHeadingStyleApplyCurSel";
            this.btnHeadingStyleApplyCurSel.UseVisualStyleBackColor = true;
            this.btnHeadingStyleApplyCurSel.Click += new System.EventHandler(this.btnHeadingStyleApplyCurSel_Click);
            // 
            // btnHeadingStyleExitApply
            // 
            resources.ApplyResources(this.btnHeadingStyleExitApply, "btnHeadingStyleExitApply");
            this.btnHeadingStyleExitApply.Name = "btnHeadingStyleExitApply";
            this.btnHeadingStyleExitApply.UseVisualStyleBackColor = true;
            this.btnHeadingStyleExitApply.Click += new System.EventHandler(this.btnHeadingStyleExitApply_Click);
            // 
            // btnHeadingStyleApplyScope
            // 
            resources.ApplyResources(this.btnHeadingStyleApplyScope, "btnHeadingStyleApplyScope");
            this.btnHeadingStyleApplyScope.Name = "btnHeadingStyleApplyScope";
            this.btnHeadingStyleApplyScope.UseVisualStyleBackColor = true;
            this.btnHeadingStyleApplyScope.Click += new System.EventHandler(this.btnHeadingStyleApplyScope_Click);
            // 
            // richHeadingStylePreview
            // 
            resources.ApplyResources(this.richHeadingStylePreview, "richHeadingStylePreview");
            this.richHeadingStylePreview.Name = "richHeadingStylePreview";
            this.richHeadingStylePreview.ReadOnly = true;
            // 
            // lstOutlineLevel
            // 
            this.lstOutlineLevel.FormattingEnabled = true;
            resources.ApplyResources(this.lstOutlineLevel, "lstOutlineLevel");
            this.lstOutlineLevel.Items.AddRange(new object[] {
            resources.GetString("lstOutlineLevel.Items"),
            resources.GetString("lstOutlineLevel.Items1"),
            resources.GetString("lstOutlineLevel.Items2"),
            resources.GetString("lstOutlineLevel.Items3"),
            resources.GetString("lstOutlineLevel.Items4"),
            resources.GetString("lstOutlineLevel.Items5"),
            resources.GetString("lstOutlineLevel.Items6"),
            resources.GetString("lstOutlineLevel.Items7"),
            resources.GetString("lstOutlineLevel.Items8"),
            resources.GetString("lstOutlineLevel.Items9")});
            this.lstOutlineLevel.Name = "lstOutlineLevel";
            this.lstOutlineLevel.SelectedIndexChanged += new System.EventHandler(this.lstOutlineLevel_SelectedIndexChanged);
            // 
            // prgbarHeadingStyleSchemeApply
            // 
            resources.ApplyResources(this.prgbarHeadingStyleSchemeApply, "prgbarHeadingStyleSchemeApply");
            this.prgbarHeadingStyleSchemeApply.Name = "prgbarHeadingStyleSchemeApply";
            // 
            // txtHeadingStyleSchemeName
            // 
            resources.ApplyResources(this.txtHeadingStyleSchemeName, "txtHeadingStyleSchemeName");
            this.txtHeadingStyleSchemeName.Name = "txtHeadingStyleSchemeName";
            // 
            // btnHeadingStyleSchemeApply
            // 
            resources.ApplyResources(this.btnHeadingStyleSchemeApply, "btnHeadingStyleSchemeApply");
            this.btnHeadingStyleSchemeApply.Name = "btnHeadingStyleSchemeApply";
            this.btnHeadingStyleSchemeApply.UseVisualStyleBackColor = true;
            this.btnHeadingStyleSchemeApply.Click += new System.EventHandler(this.btnHeadingStyleSchemeApply_Click);
            // 
            // btnHeadingStyleSchemeExtract
            // 
            resources.ApplyResources(this.btnHeadingStyleSchemeExtract, "btnHeadingStyleSchemeExtract");
            this.btnHeadingStyleSchemeExtract.Name = "btnHeadingStyleSchemeExtract";
            this.btnHeadingStyleSchemeExtract.UseVisualStyleBackColor = true;
            this.btnHeadingStyleSchemeExtract.Click += new System.EventHandler(this.btnHeadingStyleSchemeExtract_Click);
            // 
            // btnHeadingStyleSchemePreview
            // 
            resources.ApplyResources(this.btnHeadingStyleSchemePreview, "btnHeadingStyleSchemePreview");
            this.btnHeadingStyleSchemePreview.Name = "btnHeadingStyleSchemePreview";
            this.btnHeadingStyleSchemePreview.UseVisualStyleBackColor = true;
            this.btnHeadingStyleSchemePreview.Click += new System.EventHandler(this.btnHeadingStyleSchemePreview_Click);
            // 
            // btnHeadingStyleSchemeUpdate
            // 
            resources.ApplyResources(this.btnHeadingStyleSchemeUpdate, "btnHeadingStyleSchemeUpdate");
            this.btnHeadingStyleSchemeUpdate.Name = "btnHeadingStyleSchemeUpdate";
            this.btnHeadingStyleSchemeUpdate.UseVisualStyleBackColor = true;
            this.btnHeadingStyleSchemeUpdate.Click += new System.EventHandler(this.btnHeadingStyleSchemeUpdate_Click);
            // 
            // btnHeadingStyleSchemeDel
            // 
            resources.ApplyResources(this.btnHeadingStyleSchemeDel, "btnHeadingStyleSchemeDel");
            this.btnHeadingStyleSchemeDel.Name = "btnHeadingStyleSchemeDel";
            this.btnHeadingStyleSchemeDel.UseVisualStyleBackColor = true;
            this.btnHeadingStyleSchemeDel.Click += new System.EventHandler(this.btnHeadingStyleSchemeDel_Click);
            // 
            // btnHeadingStyleSchemeAdd
            // 
            resources.ApplyResources(this.btnHeadingStyleSchemeAdd, "btnHeadingStyleSchemeAdd");
            this.btnHeadingStyleSchemeAdd.Name = "btnHeadingStyleSchemeAdd";
            this.btnHeadingStyleSchemeAdd.UseVisualStyleBackColor = true;
            this.btnHeadingStyleSchemeAdd.Click += new System.EventHandler(this.btnHeadingStyleSchemeAdd_Click);
            // 
            // trvHeadingStyleScheme
            // 
            resources.ApplyResources(this.trvHeadingStyleScheme, "trvHeadingStyleScheme");
            this.trvHeadingStyleScheme.ImageList = this.imageListIcon;
            this.trvHeadingStyleScheme.Name = "trvHeadingStyleScheme";
            this.trvHeadingStyleScheme.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.trvHeadingStyleScheme_AfterSelect);
            this.trvHeadingStyleScheme.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.trvHeadingStyleScheme_MouseDoubleClick);
            // 
            // label41
            // 
            resources.ApplyResources(this.label41, "label41");
            this.label41.Name = "label41";
            // 
            // tabPageObjNav
            // 
            this.tabPageObjNav.Controls.Add(this.groupBox15);
            this.tabPageObjNav.Controls.Add(this.groupBox14);
            resources.ApplyResources(this.tabPageObjNav, "tabPageObjNav");
            this.tabPageObjNav.Name = "tabPageObjNav";
            this.tabPageObjNav.UseVisualStyleBackColor = true;
            // 
            // groupBox15
            // 
            this.groupBox15.Controls.Add(this.btnONEquationNavLast);
            this.groupBox15.Controls.Add(this.btnONObjectNavLast);
            this.groupBox15.Controls.Add(this.btnONBookmarkNavLast);
            this.groupBox15.Controls.Add(this.btnONEndnoteNavLast);
            this.groupBox15.Controls.Add(this.btnONFootnoteNavLast);
            this.groupBox15.Controls.Add(this.btnONCommentNavLast);
            this.groupBox15.Controls.Add(this.btnONEquationNavPrev);
            this.groupBox15.Controls.Add(this.btnONObjectNavPrev);
            this.groupBox15.Controls.Add(this.btnONBookmarkNavPrev);
            this.groupBox15.Controls.Add(this.btnONEndnoteNavPrev);
            this.groupBox15.Controls.Add(this.btnONFootnoteNavPrev);
            this.groupBox15.Controls.Add(this.btnONCommentNavPrev);
            this.groupBox15.Controls.Add(this.btnONEquationNavFirst);
            this.groupBox15.Controls.Add(this.btnONEquationNavNext);
            this.groupBox15.Controls.Add(this.btnONObjectNavFirst);
            this.groupBox15.Controls.Add(this.btnONObjectNavNext);
            this.groupBox15.Controls.Add(this.btnONBookmarkNavFirst);
            this.groupBox15.Controls.Add(this.btnONBookmarkNavNext);
            this.groupBox15.Controls.Add(this.btnONEndnoteNavFirst);
            this.groupBox15.Controls.Add(this.label79);
            this.groupBox15.Controls.Add(this.btnONEndnoteNavNext);
            this.groupBox15.Controls.Add(this.label78);
            this.groupBox15.Controls.Add(this.btnONFootnoteNavFirst);
            this.groupBox15.Controls.Add(this.label77);
            this.groupBox15.Controls.Add(this.btnONFootnoteNavNext);
            this.groupBox15.Controls.Add(this.label69);
            this.groupBox15.Controls.Add(this.btnONCommentNavFirst);
            this.groupBox15.Controls.Add(this.label68);
            this.groupBox15.Controls.Add(this.btnONCommentNavNext);
            this.groupBox15.Controls.Add(this.label66);
            resources.ApplyResources(this.groupBox15, "groupBox15");
            this.groupBox15.Name = "groupBox15";
            this.groupBox15.TabStop = false;
            // 
            // btnONEquationNavLast
            // 
            resources.ApplyResources(this.btnONEquationNavLast, "btnONEquationNavLast");
            this.btnONEquationNavLast.Name = "btnONEquationNavLast";
            this.btnONEquationNavLast.UseVisualStyleBackColor = true;
            this.btnONEquationNavLast.Click += new System.EventHandler(this.btnONEquationNavLast_Click);
            // 
            // btnONObjectNavLast
            // 
            resources.ApplyResources(this.btnONObjectNavLast, "btnONObjectNavLast");
            this.btnONObjectNavLast.Name = "btnONObjectNavLast";
            this.btnONObjectNavLast.UseVisualStyleBackColor = true;
            this.btnONObjectNavLast.Click += new System.EventHandler(this.btnONObjectNavLast_Click);
            // 
            // btnONBookmarkNavLast
            // 
            resources.ApplyResources(this.btnONBookmarkNavLast, "btnONBookmarkNavLast");
            this.btnONBookmarkNavLast.Name = "btnONBookmarkNavLast";
            this.btnONBookmarkNavLast.UseVisualStyleBackColor = true;
            this.btnONBookmarkNavLast.Click += new System.EventHandler(this.btnONBookmarkNavLast_Click);
            // 
            // btnONEndnoteNavLast
            // 
            resources.ApplyResources(this.btnONEndnoteNavLast, "btnONEndnoteNavLast");
            this.btnONEndnoteNavLast.Name = "btnONEndnoteNavLast";
            this.btnONEndnoteNavLast.UseVisualStyleBackColor = true;
            this.btnONEndnoteNavLast.Click += new System.EventHandler(this.btnONEndnoteNavLast_Click);
            // 
            // btnONFootnoteNavLast
            // 
            resources.ApplyResources(this.btnONFootnoteNavLast, "btnONFootnoteNavLast");
            this.btnONFootnoteNavLast.Name = "btnONFootnoteNavLast";
            this.btnONFootnoteNavLast.UseVisualStyleBackColor = true;
            this.btnONFootnoteNavLast.Click += new System.EventHandler(this.btnONFootnoteNavLast_Click);
            // 
            // btnONCommentNavLast
            // 
            resources.ApplyResources(this.btnONCommentNavLast, "btnONCommentNavLast");
            this.btnONCommentNavLast.Name = "btnONCommentNavLast";
            this.btnONCommentNavLast.UseVisualStyleBackColor = true;
            this.btnONCommentNavLast.Click += new System.EventHandler(this.btnONCommentNavLast_Click);
            // 
            // btnONEquationNavPrev
            // 
            resources.ApplyResources(this.btnONEquationNavPrev, "btnONEquationNavPrev");
            this.btnONEquationNavPrev.Name = "btnONEquationNavPrev";
            this.btnONEquationNavPrev.UseVisualStyleBackColor = true;
            this.btnONEquationNavPrev.Click += new System.EventHandler(this.btnONEquationNavPrev_Click);
            // 
            // btnONObjectNavPrev
            // 
            resources.ApplyResources(this.btnONObjectNavPrev, "btnONObjectNavPrev");
            this.btnONObjectNavPrev.Name = "btnONObjectNavPrev";
            this.btnONObjectNavPrev.UseVisualStyleBackColor = true;
            this.btnONObjectNavPrev.Click += new System.EventHandler(this.btnONObjectNavPrev_Click);
            // 
            // btnONBookmarkNavPrev
            // 
            resources.ApplyResources(this.btnONBookmarkNavPrev, "btnONBookmarkNavPrev");
            this.btnONBookmarkNavPrev.Name = "btnONBookmarkNavPrev";
            this.btnONBookmarkNavPrev.UseVisualStyleBackColor = true;
            this.btnONBookmarkNavPrev.Click += new System.EventHandler(this.btnONBookmarkNavPrev_Click);
            // 
            // btnONEndnoteNavPrev
            // 
            resources.ApplyResources(this.btnONEndnoteNavPrev, "btnONEndnoteNavPrev");
            this.btnONEndnoteNavPrev.Name = "btnONEndnoteNavPrev";
            this.btnONEndnoteNavPrev.UseVisualStyleBackColor = true;
            this.btnONEndnoteNavPrev.Click += new System.EventHandler(this.btnONEndnoteNavPrev_Click);
            // 
            // btnONFootnoteNavPrev
            // 
            resources.ApplyResources(this.btnONFootnoteNavPrev, "btnONFootnoteNavPrev");
            this.btnONFootnoteNavPrev.Name = "btnONFootnoteNavPrev";
            this.btnONFootnoteNavPrev.UseVisualStyleBackColor = true;
            this.btnONFootnoteNavPrev.Click += new System.EventHandler(this.btnONFootnoteNavPrev_Click);
            // 
            // btnONCommentNavPrev
            // 
            resources.ApplyResources(this.btnONCommentNavPrev, "btnONCommentNavPrev");
            this.btnONCommentNavPrev.Name = "btnONCommentNavPrev";
            this.btnONCommentNavPrev.UseVisualStyleBackColor = true;
            this.btnONCommentNavPrev.Click += new System.EventHandler(this.btnONCommentNavPrev_Click);
            // 
            // btnONEquationNavFirst
            // 
            resources.ApplyResources(this.btnONEquationNavFirst, "btnONEquationNavFirst");
            this.btnONEquationNavFirst.Name = "btnONEquationNavFirst";
            this.btnONEquationNavFirst.UseVisualStyleBackColor = true;
            this.btnONEquationNavFirst.Click += new System.EventHandler(this.btnONEquationNavFirst_Click);
            // 
            // btnONEquationNavNext
            // 
            resources.ApplyResources(this.btnONEquationNavNext, "btnONEquationNavNext");
            this.btnONEquationNavNext.Name = "btnONEquationNavNext";
            this.btnONEquationNavNext.UseVisualStyleBackColor = true;
            this.btnONEquationNavNext.Click += new System.EventHandler(this.btnONEquationNavNext_Click);
            // 
            // btnONObjectNavFirst
            // 
            resources.ApplyResources(this.btnONObjectNavFirst, "btnONObjectNavFirst");
            this.btnONObjectNavFirst.Name = "btnONObjectNavFirst";
            this.btnONObjectNavFirst.UseVisualStyleBackColor = true;
            this.btnONObjectNavFirst.Click += new System.EventHandler(this.btnONObjectNavFirst_Click);
            // 
            // btnONObjectNavNext
            // 
            resources.ApplyResources(this.btnONObjectNavNext, "btnONObjectNavNext");
            this.btnONObjectNavNext.Name = "btnONObjectNavNext";
            this.btnONObjectNavNext.UseVisualStyleBackColor = true;
            this.btnONObjectNavNext.Click += new System.EventHandler(this.btnONObjectNavNext_Click);
            // 
            // btnONBookmarkNavFirst
            // 
            resources.ApplyResources(this.btnONBookmarkNavFirst, "btnONBookmarkNavFirst");
            this.btnONBookmarkNavFirst.Name = "btnONBookmarkNavFirst";
            this.btnONBookmarkNavFirst.UseVisualStyleBackColor = true;
            this.btnONBookmarkNavFirst.Click += new System.EventHandler(this.btnONBookmarkNavFirst_Click);
            // 
            // btnONBookmarkNavNext
            // 
            resources.ApplyResources(this.btnONBookmarkNavNext, "btnONBookmarkNavNext");
            this.btnONBookmarkNavNext.Name = "btnONBookmarkNavNext";
            this.btnONBookmarkNavNext.UseVisualStyleBackColor = true;
            this.btnONBookmarkNavNext.Click += new System.EventHandler(this.btnONBookmarkNavNext_Click);
            // 
            // btnONEndnoteNavFirst
            // 
            resources.ApplyResources(this.btnONEndnoteNavFirst, "btnONEndnoteNavFirst");
            this.btnONEndnoteNavFirst.Name = "btnONEndnoteNavFirst";
            this.btnONEndnoteNavFirst.UseVisualStyleBackColor = true;
            this.btnONEndnoteNavFirst.Click += new System.EventHandler(this.btnONEndnoteNavFirst_Click);
            // 
            // label79
            // 
            resources.ApplyResources(this.label79, "label79");
            this.label79.Name = "label79";
            // 
            // btnONEndnoteNavNext
            // 
            resources.ApplyResources(this.btnONEndnoteNavNext, "btnONEndnoteNavNext");
            this.btnONEndnoteNavNext.Name = "btnONEndnoteNavNext";
            this.btnONEndnoteNavNext.UseVisualStyleBackColor = true;
            this.btnONEndnoteNavNext.Click += new System.EventHandler(this.btnONEndnoteNavNext_Click);
            // 
            // label78
            // 
            resources.ApplyResources(this.label78, "label78");
            this.label78.Name = "label78";
            // 
            // btnONFootnoteNavFirst
            // 
            resources.ApplyResources(this.btnONFootnoteNavFirst, "btnONFootnoteNavFirst");
            this.btnONFootnoteNavFirst.Name = "btnONFootnoteNavFirst";
            this.btnONFootnoteNavFirst.UseVisualStyleBackColor = true;
            this.btnONFootnoteNavFirst.Click += new System.EventHandler(this.btnONFootnoteNavFirst_Click);
            // 
            // label77
            // 
            resources.ApplyResources(this.label77, "label77");
            this.label77.Name = "label77";
            // 
            // btnONFootnoteNavNext
            // 
            resources.ApplyResources(this.btnONFootnoteNavNext, "btnONFootnoteNavNext");
            this.btnONFootnoteNavNext.Name = "btnONFootnoteNavNext";
            this.btnONFootnoteNavNext.UseVisualStyleBackColor = true;
            this.btnONFootnoteNavNext.Click += new System.EventHandler(this.btnONFootnoteNavNext_Click);
            // 
            // label69
            // 
            resources.ApplyResources(this.label69, "label69");
            this.label69.Name = "label69";
            // 
            // btnONCommentNavFirst
            // 
            resources.ApplyResources(this.btnONCommentNavFirst, "btnONCommentNavFirst");
            this.btnONCommentNavFirst.Name = "btnONCommentNavFirst";
            this.btnONCommentNavFirst.UseVisualStyleBackColor = true;
            this.btnONCommentNavFirst.Click += new System.EventHandler(this.btnONCommentNavFirst_Click);
            // 
            // label68
            // 
            resources.ApplyResources(this.label68, "label68");
            this.label68.Name = "label68";
            // 
            // btnONCommentNavNext
            // 
            resources.ApplyResources(this.btnONCommentNavNext, "btnONCommentNavNext");
            this.btnONCommentNavNext.Name = "btnONCommentNavNext";
            this.btnONCommentNavNext.UseVisualStyleBackColor = true;
            this.btnONCommentNavNext.Click += new System.EventHandler(this.btnONCommentNavNext_Click);
            // 
            // label66
            // 
            resources.ApplyResources(this.label66, "label66");
            this.label66.Name = "label66";
            // 
            // groupBox14
            // 
            this.groupBox14.Controls.Add(this.colorComboBoxNav);
            this.groupBox14.Controls.Add(this.btnHighLightGoLast);
            this.groupBox14.Controls.Add(this.btnONFieldNavLast);
            this.groupBox14.Controls.Add(this.label75);
            this.groupBox14.Controls.Add(this.label67);
            this.groupBox14.Controls.Add(this.groupBox13);
            this.groupBox14.Controls.Add(this.btnONSectionNavLast);
            this.groupBox14.Controls.Add(this.label76);
            this.groupBox14.Controls.Add(this.btnONPageNavLast);
            this.groupBox14.Controls.Add(this.label64);
            this.groupBox14.Controls.Add(this.btnHighLightGoNext);
            this.groupBox14.Controls.Add(this.btnONFieldNavNext);
            this.groupBox14.Controls.Add(this.btnONGraphicNavLast);
            this.groupBox14.Controls.Add(this.btnONSectionNavNext);
            this.groupBox14.Controls.Add(this.label63);
            this.groupBox14.Controls.Add(this.btnONPageNavNext);
            this.groupBox14.Controls.Add(this.btnHighLightGoFirst);
            this.groupBox14.Controls.Add(this.btnONFieldNavFirst);
            this.groupBox14.Controls.Add(this.btnONTableNavLast);
            this.groupBox14.Controls.Add(this.btnONSectionNavFirst);
            this.groupBox14.Controls.Add(this.btnONGraphicNavNext);
            this.groupBox14.Controls.Add(this.btnONPageNavFirst);
            this.groupBox14.Controls.Add(this.btnHighLightGoPrev);
            this.groupBox14.Controls.Add(this.btnONFieldNavPrev);
            this.groupBox14.Controls.Add(this.label65);
            this.groupBox14.Controls.Add(this.btnONSectionNavPrev);
            this.groupBox14.Controls.Add(this.btnONGraphicNavFirst);
            this.groupBox14.Controls.Add(this.btnONPageNavPrev);
            this.groupBox14.Controls.Add(this.btnONTableNavNext);
            this.groupBox14.Controls.Add(this.btnONGraphicNavPrev);
            this.groupBox14.Controls.Add(this.btnONTableNavFirst);
            this.groupBox14.Controls.Add(this.btnONTableNavPrev);
            resources.ApplyResources(this.groupBox14, "groupBox14");
            this.groupBox14.Name = "groupBox14";
            this.groupBox14.TabStop = false;
            // 
            // colorComboBoxNav
            // 
            this.colorComboBoxNav.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.colorComboBoxNav.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.colorComboBoxNav.FormattingEnabled = true;
            this.colorComboBoxNav.Items.AddRange(new object[] {
            resources.GetString("colorComboBoxNav.Items"),
            resources.GetString("colorComboBoxNav.Items1"),
            resources.GetString("colorComboBoxNav.Items2"),
            resources.GetString("colorComboBoxNav.Items3"),
            resources.GetString("colorComboBoxNav.Items4"),
            resources.GetString("colorComboBoxNav.Items5"),
            resources.GetString("colorComboBoxNav.Items6"),
            resources.GetString("colorComboBoxNav.Items7"),
            resources.GetString("colorComboBoxNav.Items8"),
            resources.GetString("colorComboBoxNav.Items9"),
            resources.GetString("colorComboBoxNav.Items10"),
            resources.GetString("colorComboBoxNav.Items11"),
            resources.GetString("colorComboBoxNav.Items12"),
            resources.GetString("colorComboBoxNav.Items13"),
            resources.GetString("colorComboBoxNav.Items14"),
            resources.GetString("colorComboBoxNav.Items15"),
            resources.GetString("colorComboBoxNav.Items16"),
            resources.GetString("colorComboBoxNav.Items17"),
            resources.GetString("colorComboBoxNav.Items18"),
            resources.GetString("colorComboBoxNav.Items19"),
            resources.GetString("colorComboBoxNav.Items20"),
            resources.GetString("colorComboBoxNav.Items21"),
            resources.GetString("colorComboBoxNav.Items22"),
            resources.GetString("colorComboBoxNav.Items23"),
            resources.GetString("colorComboBoxNav.Items24"),
            resources.GetString("colorComboBoxNav.Items25"),
            resources.GetString("colorComboBoxNav.Items26"),
            resources.GetString("colorComboBoxNav.Items27"),
            resources.GetString("colorComboBoxNav.Items28"),
            resources.GetString("colorComboBoxNav.Items29"),
            resources.GetString("colorComboBoxNav.Items30"),
            resources.GetString("colorComboBoxNav.Items31"),
            resources.GetString("colorComboBoxNav.Items32"),
            resources.GetString("colorComboBoxNav.Items33"),
            resources.GetString("colorComboBoxNav.Items34"),
            resources.GetString("colorComboBoxNav.Items35"),
            resources.GetString("colorComboBoxNav.Items36"),
            resources.GetString("colorComboBoxNav.Items37"),
            resources.GetString("colorComboBoxNav.Items38"),
            resources.GetString("colorComboBoxNav.Items39"),
            resources.GetString("colorComboBoxNav.Items40"),
            resources.GetString("colorComboBoxNav.Items41"),
            resources.GetString("colorComboBoxNav.Items42"),
            resources.GetString("colorComboBoxNav.Items43"),
            resources.GetString("colorComboBoxNav.Items44"),
            resources.GetString("colorComboBoxNav.Items45"),
            resources.GetString("colorComboBoxNav.Items46"),
            resources.GetString("colorComboBoxNav.Items47"),
            resources.GetString("colorComboBoxNav.Items48"),
            resources.GetString("colorComboBoxNav.Items49"),
            resources.GetString("colorComboBoxNav.Items50"),
            resources.GetString("colorComboBoxNav.Items51"),
            resources.GetString("colorComboBoxNav.Items52"),
            resources.GetString("colorComboBoxNav.Items53"),
            resources.GetString("colorComboBoxNav.Items54"),
            resources.GetString("colorComboBoxNav.Items55"),
            resources.GetString("colorComboBoxNav.Items56"),
            resources.GetString("colorComboBoxNav.Items57"),
            resources.GetString("colorComboBoxNav.Items58"),
            resources.GetString("colorComboBoxNav.Items59"),
            resources.GetString("colorComboBoxNav.Items60"),
            resources.GetString("colorComboBoxNav.Items61"),
            resources.GetString("colorComboBoxNav.Items62"),
            resources.GetString("colorComboBoxNav.Items63"),
            resources.GetString("colorComboBoxNav.Items64"),
            resources.GetString("colorComboBoxNav.Items65"),
            resources.GetString("colorComboBoxNav.Items66"),
            resources.GetString("colorComboBoxNav.Items67"),
            resources.GetString("colorComboBoxNav.Items68"),
            resources.GetString("colorComboBoxNav.Items69"),
            resources.GetString("colorComboBoxNav.Items70"),
            resources.GetString("colorComboBoxNav.Items71"),
            resources.GetString("colorComboBoxNav.Items72"),
            resources.GetString("colorComboBoxNav.Items73"),
            resources.GetString("colorComboBoxNav.Items74"),
            resources.GetString("colorComboBoxNav.Items75"),
            resources.GetString("colorComboBoxNav.Items76"),
            resources.GetString("colorComboBoxNav.Items77"),
            resources.GetString("colorComboBoxNav.Items78"),
            resources.GetString("colorComboBoxNav.Items79"),
            resources.GetString("colorComboBoxNav.Items80"),
            resources.GetString("colorComboBoxNav.Items81"),
            resources.GetString("colorComboBoxNav.Items82"),
            resources.GetString("colorComboBoxNav.Items83"),
            resources.GetString("colorComboBoxNav.Items84"),
            resources.GetString("colorComboBoxNav.Items85"),
            resources.GetString("colorComboBoxNav.Items86"),
            resources.GetString("colorComboBoxNav.Items87"),
            resources.GetString("colorComboBoxNav.Items88"),
            resources.GetString("colorComboBoxNav.Items89"),
            resources.GetString("colorComboBoxNav.Items90"),
            resources.GetString("colorComboBoxNav.Items91"),
            resources.GetString("colorComboBoxNav.Items92"),
            resources.GetString("colorComboBoxNav.Items93"),
            resources.GetString("colorComboBoxNav.Items94"),
            resources.GetString("colorComboBoxNav.Items95"),
            resources.GetString("colorComboBoxNav.Items96"),
            resources.GetString("colorComboBoxNav.Items97"),
            resources.GetString("colorComboBoxNav.Items98"),
            resources.GetString("colorComboBoxNav.Items99"),
            resources.GetString("colorComboBoxNav.Items100"),
            resources.GetString("colorComboBoxNav.Items101"),
            resources.GetString("colorComboBoxNav.Items102"),
            resources.GetString("colorComboBoxNav.Items103"),
            resources.GetString("colorComboBoxNav.Items104"),
            resources.GetString("colorComboBoxNav.Items105"),
            resources.GetString("colorComboBoxNav.Items106"),
            resources.GetString("colorComboBoxNav.Items107"),
            resources.GetString("colorComboBoxNav.Items108"),
            resources.GetString("colorComboBoxNav.Items109"),
            resources.GetString("colorComboBoxNav.Items110"),
            resources.GetString("colorComboBoxNav.Items111"),
            resources.GetString("colorComboBoxNav.Items112"),
            resources.GetString("colorComboBoxNav.Items113"),
            resources.GetString("colorComboBoxNav.Items114"),
            resources.GetString("colorComboBoxNav.Items115"),
            resources.GetString("colorComboBoxNav.Items116"),
            resources.GetString("colorComboBoxNav.Items117"),
            resources.GetString("colorComboBoxNav.Items118"),
            resources.GetString("colorComboBoxNav.Items119"),
            resources.GetString("colorComboBoxNav.Items120"),
            resources.GetString("colorComboBoxNav.Items121"),
            resources.GetString("colorComboBoxNav.Items122"),
            resources.GetString("colorComboBoxNav.Items123"),
            resources.GetString("colorComboBoxNav.Items124"),
            resources.GetString("colorComboBoxNav.Items125"),
            resources.GetString("colorComboBoxNav.Items126"),
            resources.GetString("colorComboBoxNav.Items127"),
            resources.GetString("colorComboBoxNav.Items128"),
            resources.GetString("colorComboBoxNav.Items129"),
            resources.GetString("colorComboBoxNav.Items130"),
            resources.GetString("colorComboBoxNav.Items131"),
            resources.GetString("colorComboBoxNav.Items132"),
            resources.GetString("colorComboBoxNav.Items133"),
            resources.GetString("colorComboBoxNav.Items134"),
            resources.GetString("colorComboBoxNav.Items135"),
            resources.GetString("colorComboBoxNav.Items136"),
            resources.GetString("colorComboBoxNav.Items137"),
            resources.GetString("colorComboBoxNav.Items138"),
            resources.GetString("colorComboBoxNav.Items139"),
            resources.GetString("colorComboBoxNav.Items140"),
            resources.GetString("colorComboBoxNav.Items141"),
            resources.GetString("colorComboBoxNav.Items142"),
            resources.GetString("colorComboBoxNav.Items143"),
            resources.GetString("colorComboBoxNav.Items144"),
            resources.GetString("colorComboBoxNav.Items145"),
            resources.GetString("colorComboBoxNav.Items146"),
            resources.GetString("colorComboBoxNav.Items147"),
            resources.GetString("colorComboBoxNav.Items148"),
            resources.GetString("colorComboBoxNav.Items149"),
            resources.GetString("colorComboBoxNav.Items150"),
            resources.GetString("colorComboBoxNav.Items151"),
            resources.GetString("colorComboBoxNav.Items152"),
            resources.GetString("colorComboBoxNav.Items153"),
            resources.GetString("colorComboBoxNav.Items154"),
            resources.GetString("colorComboBoxNav.Items155"),
            resources.GetString("colorComboBoxNav.Items156"),
            resources.GetString("colorComboBoxNav.Items157"),
            resources.GetString("colorComboBoxNav.Items158"),
            resources.GetString("colorComboBoxNav.Items159"),
            resources.GetString("colorComboBoxNav.Items160"),
            resources.GetString("colorComboBoxNav.Items161"),
            resources.GetString("colorComboBoxNav.Items162"),
            resources.GetString("colorComboBoxNav.Items163"),
            resources.GetString("colorComboBoxNav.Items164"),
            resources.GetString("colorComboBoxNav.Items165"),
            resources.GetString("colorComboBoxNav.Items166"),
            resources.GetString("colorComboBoxNav.Items167"),
            resources.GetString("colorComboBoxNav.Items168"),
            resources.GetString("colorComboBoxNav.Items169"),
            resources.GetString("colorComboBoxNav.Items170"),
            resources.GetString("colorComboBoxNav.Items171"),
            resources.GetString("colorComboBoxNav.Items172"),
            resources.GetString("colorComboBoxNav.Items173"),
            resources.GetString("colorComboBoxNav.Items174"),
            resources.GetString("colorComboBoxNav.Items175"),
            resources.GetString("colorComboBoxNav.Items176"),
            resources.GetString("colorComboBoxNav.Items177"),
            resources.GetString("colorComboBoxNav.Items178"),
            resources.GetString("colorComboBoxNav.Items179"),
            resources.GetString("colorComboBoxNav.Items180"),
            resources.GetString("colorComboBoxNav.Items181"),
            resources.GetString("colorComboBoxNav.Items182"),
            resources.GetString("colorComboBoxNav.Items183"),
            resources.GetString("colorComboBoxNav.Items184"),
            resources.GetString("colorComboBoxNav.Items185"),
            resources.GetString("colorComboBoxNav.Items186"),
            resources.GetString("colorComboBoxNav.Items187"),
            resources.GetString("colorComboBoxNav.Items188"),
            resources.GetString("colorComboBoxNav.Items189"),
            resources.GetString("colorComboBoxNav.Items190"),
            resources.GetString("colorComboBoxNav.Items191"),
            resources.GetString("colorComboBoxNav.Items192"),
            resources.GetString("colorComboBoxNav.Items193"),
            resources.GetString("colorComboBoxNav.Items194"),
            resources.GetString("colorComboBoxNav.Items195"),
            resources.GetString("colorComboBoxNav.Items196"),
            resources.GetString("colorComboBoxNav.Items197"),
            resources.GetString("colorComboBoxNav.Items198"),
            resources.GetString("colorComboBoxNav.Items199"),
            resources.GetString("colorComboBoxNav.Items200"),
            resources.GetString("colorComboBoxNav.Items201"),
            resources.GetString("colorComboBoxNav.Items202"),
            resources.GetString("colorComboBoxNav.Items203"),
            resources.GetString("colorComboBoxNav.Items204"),
            resources.GetString("colorComboBoxNav.Items205"),
            resources.GetString("colorComboBoxNav.Items206"),
            resources.GetString("colorComboBoxNav.Items207"),
            resources.GetString("colorComboBoxNav.Items208"),
            resources.GetString("colorComboBoxNav.Items209"),
            resources.GetString("colorComboBoxNav.Items210"),
            resources.GetString("colorComboBoxNav.Items211"),
            resources.GetString("colorComboBoxNav.Items212"),
            resources.GetString("colorComboBoxNav.Items213"),
            resources.GetString("colorComboBoxNav.Items214"),
            resources.GetString("colorComboBoxNav.Items215"),
            resources.GetString("colorComboBoxNav.Items216"),
            resources.GetString("colorComboBoxNav.Items217"),
            resources.GetString("colorComboBoxNav.Items218"),
            resources.GetString("colorComboBoxNav.Items219"),
            resources.GetString("colorComboBoxNav.Items220"),
            resources.GetString("colorComboBoxNav.Items221"),
            resources.GetString("colorComboBoxNav.Items222"),
            resources.GetString("colorComboBoxNav.Items223"),
            resources.GetString("colorComboBoxNav.Items224"),
            resources.GetString("colorComboBoxNav.Items225"),
            resources.GetString("colorComboBoxNav.Items226"),
            resources.GetString("colorComboBoxNav.Items227"),
            resources.GetString("colorComboBoxNav.Items228"),
            resources.GetString("colorComboBoxNav.Items229"),
            resources.GetString("colorComboBoxNav.Items230"),
            resources.GetString("colorComboBoxNav.Items231"),
            resources.GetString("colorComboBoxNav.Items232"),
            resources.GetString("colorComboBoxNav.Items233"),
            resources.GetString("colorComboBoxNav.Items234"),
            resources.GetString("colorComboBoxNav.Items235"),
            resources.GetString("colorComboBoxNav.Items236"),
            resources.GetString("colorComboBoxNav.Items237"),
            resources.GetString("colorComboBoxNav.Items238"),
            resources.GetString("colorComboBoxNav.Items239"),
            resources.GetString("colorComboBoxNav.Items240"),
            resources.GetString("colorComboBoxNav.Items241"),
            resources.GetString("colorComboBoxNav.Items242"),
            resources.GetString("colorComboBoxNav.Items243"),
            resources.GetString("colorComboBoxNav.Items244"),
            resources.GetString("colorComboBoxNav.Items245"),
            resources.GetString("colorComboBoxNav.Items246"),
            resources.GetString("colorComboBoxNav.Items247"),
            resources.GetString("colorComboBoxNav.Items248"),
            resources.GetString("colorComboBoxNav.Items249"),
            resources.GetString("colorComboBoxNav.Items250"),
            resources.GetString("colorComboBoxNav.Items251"),
            resources.GetString("colorComboBoxNav.Items252"),
            resources.GetString("colorComboBoxNav.Items253"),
            resources.GetString("colorComboBoxNav.Items254"),
            resources.GetString("colorComboBoxNav.Items255"),
            resources.GetString("colorComboBoxNav.Items256"),
            resources.GetString("colorComboBoxNav.Items257"),
            resources.GetString("colorComboBoxNav.Items258"),
            resources.GetString("colorComboBoxNav.Items259"),
            resources.GetString("colorComboBoxNav.Items260"),
            resources.GetString("colorComboBoxNav.Items261"),
            resources.GetString("colorComboBoxNav.Items262"),
            resources.GetString("colorComboBoxNav.Items263"),
            resources.GetString("colorComboBoxNav.Items264"),
            resources.GetString("colorComboBoxNav.Items265"),
            resources.GetString("colorComboBoxNav.Items266"),
            resources.GetString("colorComboBoxNav.Items267"),
            resources.GetString("colorComboBoxNav.Items268"),
            resources.GetString("colorComboBoxNav.Items269"),
            resources.GetString("colorComboBoxNav.Items270"),
            resources.GetString("colorComboBoxNav.Items271"),
            resources.GetString("colorComboBoxNav.Items272"),
            resources.GetString("colorComboBoxNav.Items273"),
            resources.GetString("colorComboBoxNav.Items274"),
            resources.GetString("colorComboBoxNav.Items275"),
            resources.GetString("colorComboBoxNav.Items276"),
            resources.GetString("colorComboBoxNav.Items277"),
            resources.GetString("colorComboBoxNav.Items278"),
            resources.GetString("colorComboBoxNav.Items279"),
            resources.GetString("colorComboBoxNav.Items280"),
            resources.GetString("colorComboBoxNav.Items281"),
            resources.GetString("colorComboBoxNav.Items282"),
            resources.GetString("colorComboBoxNav.Items283"),
            resources.GetString("colorComboBoxNav.Items284"),
            resources.GetString("colorComboBoxNav.Items285"),
            resources.GetString("colorComboBoxNav.Items286"),
            resources.GetString("colorComboBoxNav.Items287"),
            resources.GetString("colorComboBoxNav.Items288"),
            resources.GetString("colorComboBoxNav.Items289"),
            resources.GetString("colorComboBoxNav.Items290"),
            resources.GetString("colorComboBoxNav.Items291"),
            resources.GetString("colorComboBoxNav.Items292"),
            resources.GetString("colorComboBoxNav.Items293"),
            resources.GetString("colorComboBoxNav.Items294"),
            resources.GetString("colorComboBoxNav.Items295"),
            resources.GetString("colorComboBoxNav.Items296"),
            resources.GetString("colorComboBoxNav.Items297"),
            resources.GetString("colorComboBoxNav.Items298"),
            resources.GetString("colorComboBoxNav.Items299"),
            resources.GetString("colorComboBoxNav.Items300"),
            resources.GetString("colorComboBoxNav.Items301"),
            resources.GetString("colorComboBoxNav.Items302"),
            resources.GetString("colorComboBoxNav.Items303"),
            resources.GetString("colorComboBoxNav.Items304"),
            resources.GetString("colorComboBoxNav.Items305"),
            resources.GetString("colorComboBoxNav.Items306"),
            resources.GetString("colorComboBoxNav.Items307"),
            resources.GetString("colorComboBoxNav.Items308"),
            resources.GetString("colorComboBoxNav.Items309"),
            resources.GetString("colorComboBoxNav.Items310"),
            resources.GetString("colorComboBoxNav.Items311"),
            resources.GetString("colorComboBoxNav.Items312"),
            resources.GetString("colorComboBoxNav.Items313"),
            resources.GetString("colorComboBoxNav.Items314"),
            resources.GetString("colorComboBoxNav.Items315"),
            resources.GetString("colorComboBoxNav.Items316"),
            resources.GetString("colorComboBoxNav.Items317"),
            resources.GetString("colorComboBoxNav.Items318"),
            resources.GetString("colorComboBoxNav.Items319"),
            resources.GetString("colorComboBoxNav.Items320"),
            resources.GetString("colorComboBoxNav.Items321"),
            resources.GetString("colorComboBoxNav.Items322"),
            resources.GetString("colorComboBoxNav.Items323"),
            resources.GetString("colorComboBoxNav.Items324"),
            resources.GetString("colorComboBoxNav.Items325"),
            resources.GetString("colorComboBoxNav.Items326"),
            resources.GetString("colorComboBoxNav.Items327"),
            resources.GetString("colorComboBoxNav.Items328"),
            resources.GetString("colorComboBoxNav.Items329"),
            resources.GetString("colorComboBoxNav.Items330"),
            resources.GetString("colorComboBoxNav.Items331"),
            resources.GetString("colorComboBoxNav.Items332"),
            resources.GetString("colorComboBoxNav.Items333"),
            resources.GetString("colorComboBoxNav.Items334"),
            resources.GetString("colorComboBoxNav.Items335"),
            resources.GetString("colorComboBoxNav.Items336"),
            resources.GetString("colorComboBoxNav.Items337"),
            resources.GetString("colorComboBoxNav.Items338"),
            resources.GetString("colorComboBoxNav.Items339"),
            resources.GetString("colorComboBoxNav.Items340"),
            resources.GetString("colorComboBoxNav.Items341"),
            resources.GetString("colorComboBoxNav.Items342"),
            resources.GetString("colorComboBoxNav.Items343"),
            resources.GetString("colorComboBoxNav.Items344"),
            resources.GetString("colorComboBoxNav.Items345"),
            resources.GetString("colorComboBoxNav.Items346"),
            resources.GetString("colorComboBoxNav.Items347"),
            resources.GetString("colorComboBoxNav.Items348"),
            resources.GetString("colorComboBoxNav.Items349"),
            resources.GetString("colorComboBoxNav.Items350"),
            resources.GetString("colorComboBoxNav.Items351"),
            resources.GetString("colorComboBoxNav.Items352"),
            resources.GetString("colorComboBoxNav.Items353"),
            resources.GetString("colorComboBoxNav.Items354"),
            resources.GetString("colorComboBoxNav.Items355"),
            resources.GetString("colorComboBoxNav.Items356"),
            resources.GetString("colorComboBoxNav.Items357"),
            resources.GetString("colorComboBoxNav.Items358"),
            resources.GetString("colorComboBoxNav.Items359"),
            resources.GetString("colorComboBoxNav.Items360"),
            resources.GetString("colorComboBoxNav.Items361"),
            resources.GetString("colorComboBoxNav.Items362"),
            resources.GetString("colorComboBoxNav.Items363"),
            resources.GetString("colorComboBoxNav.Items364"),
            resources.GetString("colorComboBoxNav.Items365"),
            resources.GetString("colorComboBoxNav.Items366"),
            resources.GetString("colorComboBoxNav.Items367"),
            resources.GetString("colorComboBoxNav.Items368"),
            resources.GetString("colorComboBoxNav.Items369"),
            resources.GetString("colorComboBoxNav.Items370"),
            resources.GetString("colorComboBoxNav.Items371"),
            resources.GetString("colorComboBoxNav.Items372"),
            resources.GetString("colorComboBoxNav.Items373"),
            resources.GetString("colorComboBoxNav.Items374"),
            resources.GetString("colorComboBoxNav.Items375"),
            resources.GetString("colorComboBoxNav.Items376"),
            resources.GetString("colorComboBoxNav.Items377"),
            resources.GetString("colorComboBoxNav.Items378"),
            resources.GetString("colorComboBoxNav.Items379"),
            resources.GetString("colorComboBoxNav.Items380"),
            resources.GetString("colorComboBoxNav.Items381"),
            resources.GetString("colorComboBoxNav.Items382"),
            resources.GetString("colorComboBoxNav.Items383"),
            resources.GetString("colorComboBoxNav.Items384"),
            resources.GetString("colorComboBoxNav.Items385"),
            resources.GetString("colorComboBoxNav.Items386"),
            resources.GetString("colorComboBoxNav.Items387"),
            resources.GetString("colorComboBoxNav.Items388"),
            resources.GetString("colorComboBoxNav.Items389"),
            resources.GetString("colorComboBoxNav.Items390"),
            resources.GetString("colorComboBoxNav.Items391"),
            resources.GetString("colorComboBoxNav.Items392"),
            resources.GetString("colorComboBoxNav.Items393"),
            resources.GetString("colorComboBoxNav.Items394"),
            resources.GetString("colorComboBoxNav.Items395"),
            resources.GetString("colorComboBoxNav.Items396"),
            resources.GetString("colorComboBoxNav.Items397"),
            resources.GetString("colorComboBoxNav.Items398"),
            resources.GetString("colorComboBoxNav.Items399"),
            resources.GetString("colorComboBoxNav.Items400"),
            resources.GetString("colorComboBoxNav.Items401"),
            resources.GetString("colorComboBoxNav.Items402"),
            resources.GetString("colorComboBoxNav.Items403"),
            resources.GetString("colorComboBoxNav.Items404"),
            resources.GetString("colorComboBoxNav.Items405"),
            resources.GetString("colorComboBoxNav.Items406"),
            resources.GetString("colorComboBoxNav.Items407"),
            resources.GetString("colorComboBoxNav.Items408"),
            resources.GetString("colorComboBoxNav.Items409"),
            resources.GetString("colorComboBoxNav.Items410"),
            resources.GetString("colorComboBoxNav.Items411"),
            resources.GetString("colorComboBoxNav.Items412"),
            resources.GetString("colorComboBoxNav.Items413"),
            resources.GetString("colorComboBoxNav.Items414"),
            resources.GetString("colorComboBoxNav.Items415"),
            resources.GetString("colorComboBoxNav.Items416"),
            resources.GetString("colorComboBoxNav.Items417"),
            resources.GetString("colorComboBoxNav.Items418"),
            resources.GetString("colorComboBoxNav.Items419"),
            resources.GetString("colorComboBoxNav.Items420"),
            resources.GetString("colorComboBoxNav.Items421"),
            resources.GetString("colorComboBoxNav.Items422"),
            resources.GetString("colorComboBoxNav.Items423"),
            resources.GetString("colorComboBoxNav.Items424"),
            resources.GetString("colorComboBoxNav.Items425"),
            resources.GetString("colorComboBoxNav.Items426"),
            resources.GetString("colorComboBoxNav.Items427"),
            resources.GetString("colorComboBoxNav.Items428"),
            resources.GetString("colorComboBoxNav.Items429"),
            resources.GetString("colorComboBoxNav.Items430"),
            resources.GetString("colorComboBoxNav.Items431"),
            resources.GetString("colorComboBoxNav.Items432"),
            resources.GetString("colorComboBoxNav.Items433"),
            resources.GetString("colorComboBoxNav.Items434"),
            resources.GetString("colorComboBoxNav.Items435"),
            resources.GetString("colorComboBoxNav.Items436"),
            resources.GetString("colorComboBoxNav.Items437"),
            resources.GetString("colorComboBoxNav.Items438"),
            resources.GetString("colorComboBoxNav.Items439"),
            resources.GetString("colorComboBoxNav.Items440"),
            resources.GetString("colorComboBoxNav.Items441"),
            resources.GetString("colorComboBoxNav.Items442"),
            resources.GetString("colorComboBoxNav.Items443"),
            resources.GetString("colorComboBoxNav.Items444"),
            resources.GetString("colorComboBoxNav.Items445"),
            resources.GetString("colorComboBoxNav.Items446"),
            resources.GetString("colorComboBoxNav.Items447"),
            resources.GetString("colorComboBoxNav.Items448"),
            resources.GetString("colorComboBoxNav.Items449"),
            resources.GetString("colorComboBoxNav.Items450"),
            resources.GetString("colorComboBoxNav.Items451"),
            resources.GetString("colorComboBoxNav.Items452"),
            resources.GetString("colorComboBoxNav.Items453"),
            resources.GetString("colorComboBoxNav.Items454"),
            resources.GetString("colorComboBoxNav.Items455"),
            resources.GetString("colorComboBoxNav.Items456"),
            resources.GetString("colorComboBoxNav.Items457"),
            resources.GetString("colorComboBoxNav.Items458"),
            resources.GetString("colorComboBoxNav.Items459"),
            resources.GetString("colorComboBoxNav.Items460"),
            resources.GetString("colorComboBoxNav.Items461"),
            resources.GetString("colorComboBoxNav.Items462"),
            resources.GetString("colorComboBoxNav.Items463"),
            resources.GetString("colorComboBoxNav.Items464"),
            resources.GetString("colorComboBoxNav.Items465"),
            resources.GetString("colorComboBoxNav.Items466"),
            resources.GetString("colorComboBoxNav.Items467"),
            resources.GetString("colorComboBoxNav.Items468"),
            resources.GetString("colorComboBoxNav.Items469"),
            resources.GetString("colorComboBoxNav.Items470"),
            resources.GetString("colorComboBoxNav.Items471"),
            resources.GetString("colorComboBoxNav.Items472"),
            resources.GetString("colorComboBoxNav.Items473"),
            resources.GetString("colorComboBoxNav.Items474"),
            resources.GetString("colorComboBoxNav.Items475"),
            resources.GetString("colorComboBoxNav.Items476"),
            resources.GetString("colorComboBoxNav.Items477"),
            resources.GetString("colorComboBoxNav.Items478"),
            resources.GetString("colorComboBoxNav.Items479"),
            resources.GetString("colorComboBoxNav.Items480"),
            resources.GetString("colorComboBoxNav.Items481"),
            resources.GetString("colorComboBoxNav.Items482"),
            resources.GetString("colorComboBoxNav.Items483"),
            resources.GetString("colorComboBoxNav.Items484"),
            resources.GetString("colorComboBoxNav.Items485"),
            resources.GetString("colorComboBoxNav.Items486"),
            resources.GetString("colorComboBoxNav.Items487"),
            resources.GetString("colorComboBoxNav.Items488"),
            resources.GetString("colorComboBoxNav.Items489"),
            resources.GetString("colorComboBoxNav.Items490"),
            resources.GetString("colorComboBoxNav.Items491"),
            resources.GetString("colorComboBoxNav.Items492"),
            resources.GetString("colorComboBoxNav.Items493"),
            resources.GetString("colorComboBoxNav.Items494"),
            resources.GetString("colorComboBoxNav.Items495"),
            resources.GetString("colorComboBoxNav.Items496"),
            resources.GetString("colorComboBoxNav.Items497"),
            resources.GetString("colorComboBoxNav.Items498"),
            resources.GetString("colorComboBoxNav.Items499"),
            resources.GetString("colorComboBoxNav.Items500"),
            resources.GetString("colorComboBoxNav.Items501"),
            resources.GetString("colorComboBoxNav.Items502"),
            resources.GetString("colorComboBoxNav.Items503"),
            resources.GetString("colorComboBoxNav.Items504"),
            resources.GetString("colorComboBoxNav.Items505"),
            resources.GetString("colorComboBoxNav.Items506"),
            resources.GetString("colorComboBoxNav.Items507"),
            resources.GetString("colorComboBoxNav.Items508"),
            resources.GetString("colorComboBoxNav.Items509"),
            resources.GetString("colorComboBoxNav.Items510"),
            resources.GetString("colorComboBoxNav.Items511"),
            resources.GetString("colorComboBoxNav.Items512"),
            resources.GetString("colorComboBoxNav.Items513"),
            resources.GetString("colorComboBoxNav.Items514"),
            resources.GetString("colorComboBoxNav.Items515"),
            resources.GetString("colorComboBoxNav.Items516"),
            resources.GetString("colorComboBoxNav.Items517"),
            resources.GetString("colorComboBoxNav.Items518"),
            resources.GetString("colorComboBoxNav.Items519"),
            resources.GetString("colorComboBoxNav.Items520"),
            resources.GetString("colorComboBoxNav.Items521"),
            resources.GetString("colorComboBoxNav.Items522"),
            resources.GetString("colorComboBoxNav.Items523"),
            resources.GetString("colorComboBoxNav.Items524"),
            resources.GetString("colorComboBoxNav.Items525"),
            resources.GetString("colorComboBoxNav.Items526"),
            resources.GetString("colorComboBoxNav.Items527"),
            resources.GetString("colorComboBoxNav.Items528"),
            resources.GetString("colorComboBoxNav.Items529"),
            resources.GetString("colorComboBoxNav.Items530"),
            resources.GetString("colorComboBoxNav.Items531"),
            resources.GetString("colorComboBoxNav.Items532"),
            resources.GetString("colorComboBoxNav.Items533"),
            resources.GetString("colorComboBoxNav.Items534"),
            resources.GetString("colorComboBoxNav.Items535"),
            resources.GetString("colorComboBoxNav.Items536"),
            resources.GetString("colorComboBoxNav.Items537"),
            resources.GetString("colorComboBoxNav.Items538"),
            resources.GetString("colorComboBoxNav.Items539"),
            resources.GetString("colorComboBoxNav.Items540"),
            resources.GetString("colorComboBoxNav.Items541"),
            resources.GetString("colorComboBoxNav.Items542"),
            resources.GetString("colorComboBoxNav.Items543"),
            resources.GetString("colorComboBoxNav.Items544"),
            resources.GetString("colorComboBoxNav.Items545"),
            resources.GetString("colorComboBoxNav.Items546"),
            resources.GetString("colorComboBoxNav.Items547"),
            resources.GetString("colorComboBoxNav.Items548"),
            resources.GetString("colorComboBoxNav.Items549"),
            resources.GetString("colorComboBoxNav.Items550"),
            resources.GetString("colorComboBoxNav.Items551"),
            resources.GetString("colorComboBoxNav.Items552"),
            resources.GetString("colorComboBoxNav.Items553"),
            resources.GetString("colorComboBoxNav.Items554"),
            resources.GetString("colorComboBoxNav.Items555"),
            resources.GetString("colorComboBoxNav.Items556"),
            resources.GetString("colorComboBoxNav.Items557"),
            resources.GetString("colorComboBoxNav.Items558"),
            resources.GetString("colorComboBoxNav.Items559"),
            resources.GetString("colorComboBoxNav.Items560"),
            resources.GetString("colorComboBoxNav.Items561"),
            resources.GetString("colorComboBoxNav.Items562"),
            resources.GetString("colorComboBoxNav.Items563"),
            resources.GetString("colorComboBoxNav.Items564"),
            resources.GetString("colorComboBoxNav.Items565"),
            resources.GetString("colorComboBoxNav.Items566"),
            resources.GetString("colorComboBoxNav.Items567"),
            resources.GetString("colorComboBoxNav.Items568"),
            resources.GetString("colorComboBoxNav.Items569"),
            resources.GetString("colorComboBoxNav.Items570"),
            resources.GetString("colorComboBoxNav.Items571"),
            resources.GetString("colorComboBoxNav.Items572"),
            resources.GetString("colorComboBoxNav.Items573"),
            resources.GetString("colorComboBoxNav.Items574"),
            resources.GetString("colorComboBoxNav.Items575"),
            resources.GetString("colorComboBoxNav.Items576"),
            resources.GetString("colorComboBoxNav.Items577"),
            resources.GetString("colorComboBoxNav.Items578"),
            resources.GetString("colorComboBoxNav.Items579"),
            resources.GetString("colorComboBoxNav.Items580"),
            resources.GetString("colorComboBoxNav.Items581"),
            resources.GetString("colorComboBoxNav.Items582"),
            resources.GetString("colorComboBoxNav.Items583"),
            resources.GetString("colorComboBoxNav.Items584"),
            resources.GetString("colorComboBoxNav.Items585"),
            resources.GetString("colorComboBoxNav.Items586"),
            resources.GetString("colorComboBoxNav.Items587"),
            resources.GetString("colorComboBoxNav.Items588"),
            resources.GetString("colorComboBoxNav.Items589"),
            resources.GetString("colorComboBoxNav.Items590"),
            resources.GetString("colorComboBoxNav.Items591"),
            resources.GetString("colorComboBoxNav.Items592"),
            resources.GetString("colorComboBoxNav.Items593"),
            resources.GetString("colorComboBoxNav.Items594"),
            resources.GetString("colorComboBoxNav.Items595"),
            resources.GetString("colorComboBoxNav.Items596"),
            resources.GetString("colorComboBoxNav.Items597"),
            resources.GetString("colorComboBoxNav.Items598"),
            resources.GetString("colorComboBoxNav.Items599"),
            resources.GetString("colorComboBoxNav.Items600"),
            resources.GetString("colorComboBoxNav.Items601"),
            resources.GetString("colorComboBoxNav.Items602"),
            resources.GetString("colorComboBoxNav.Items603"),
            resources.GetString("colorComboBoxNav.Items604"),
            resources.GetString("colorComboBoxNav.Items605"),
            resources.GetString("colorComboBoxNav.Items606"),
            resources.GetString("colorComboBoxNav.Items607"),
            resources.GetString("colorComboBoxNav.Items608"),
            resources.GetString("colorComboBoxNav.Items609"),
            resources.GetString("colorComboBoxNav.Items610"),
            resources.GetString("colorComboBoxNav.Items611"),
            resources.GetString("colorComboBoxNav.Items612"),
            resources.GetString("colorComboBoxNav.Items613"),
            resources.GetString("colorComboBoxNav.Items614"),
            resources.GetString("colorComboBoxNav.Items615"),
            resources.GetString("colorComboBoxNav.Items616"),
            resources.GetString("colorComboBoxNav.Items617"),
            resources.GetString("colorComboBoxNav.Items618"),
            resources.GetString("colorComboBoxNav.Items619"),
            resources.GetString("colorComboBoxNav.Items620"),
            resources.GetString("colorComboBoxNav.Items621"),
            resources.GetString("colorComboBoxNav.Items622"),
            resources.GetString("colorComboBoxNav.Items623"),
            resources.GetString("colorComboBoxNav.Items624"),
            resources.GetString("colorComboBoxNav.Items625"),
            resources.GetString("colorComboBoxNav.Items626"),
            resources.GetString("colorComboBoxNav.Items627"),
            resources.GetString("colorComboBoxNav.Items628"),
            resources.GetString("colorComboBoxNav.Items629"),
            resources.GetString("colorComboBoxNav.Items630"),
            resources.GetString("colorComboBoxNav.Items631"),
            resources.GetString("colorComboBoxNav.Items632"),
            resources.GetString("colorComboBoxNav.Items633"),
            resources.GetString("colorComboBoxNav.Items634"),
            resources.GetString("colorComboBoxNav.Items635"),
            resources.GetString("colorComboBoxNav.Items636"),
            resources.GetString("colorComboBoxNav.Items637"),
            resources.GetString("colorComboBoxNav.Items638"),
            resources.GetString("colorComboBoxNav.Items639"),
            resources.GetString("colorComboBoxNav.Items640"),
            resources.GetString("colorComboBoxNav.Items641"),
            resources.GetString("colorComboBoxNav.Items642"),
            resources.GetString("colorComboBoxNav.Items643"),
            resources.GetString("colorComboBoxNav.Items644"),
            resources.GetString("colorComboBoxNav.Items645"),
            resources.GetString("colorComboBoxNav.Items646"),
            resources.GetString("colorComboBoxNav.Items647"),
            resources.GetString("colorComboBoxNav.Items648"),
            resources.GetString("colorComboBoxNav.Items649"),
            resources.GetString("colorComboBoxNav.Items650"),
            resources.GetString("colorComboBoxNav.Items651"),
            resources.GetString("colorComboBoxNav.Items652"),
            resources.GetString("colorComboBoxNav.Items653"),
            resources.GetString("colorComboBoxNav.Items654"),
            resources.GetString("colorComboBoxNav.Items655"),
            resources.GetString("colorComboBoxNav.Items656"),
            resources.GetString("colorComboBoxNav.Items657"),
            resources.GetString("colorComboBoxNav.Items658"),
            resources.GetString("colorComboBoxNav.Items659"),
            resources.GetString("colorComboBoxNav.Items660"),
            resources.GetString("colorComboBoxNav.Items661"),
            resources.GetString("colorComboBoxNav.Items662"),
            resources.GetString("colorComboBoxNav.Items663"),
            resources.GetString("colorComboBoxNav.Items664"),
            resources.GetString("colorComboBoxNav.Items665"),
            resources.GetString("colorComboBoxNav.Items666"),
            resources.GetString("colorComboBoxNav.Items667"),
            resources.GetString("colorComboBoxNav.Items668"),
            resources.GetString("colorComboBoxNav.Items669"),
            resources.GetString("colorComboBoxNav.Items670"),
            resources.GetString("colorComboBoxNav.Items671"),
            resources.GetString("colorComboBoxNav.Items672"),
            resources.GetString("colorComboBoxNav.Items673"),
            resources.GetString("colorComboBoxNav.Items674"),
            resources.GetString("colorComboBoxNav.Items675"),
            resources.GetString("colorComboBoxNav.Items676"),
            resources.GetString("colorComboBoxNav.Items677"),
            resources.GetString("colorComboBoxNav.Items678"),
            resources.GetString("colorComboBoxNav.Items679"),
            resources.GetString("colorComboBoxNav.Items680"),
            resources.GetString("colorComboBoxNav.Items681"),
            resources.GetString("colorComboBoxNav.Items682"),
            resources.GetString("colorComboBoxNav.Items683"),
            resources.GetString("colorComboBoxNav.Items684"),
            resources.GetString("colorComboBoxNav.Items685"),
            resources.GetString("colorComboBoxNav.Items686"),
            resources.GetString("colorComboBoxNav.Items687"),
            resources.GetString("colorComboBoxNav.Items688"),
            resources.GetString("colorComboBoxNav.Items689"),
            resources.GetString("colorComboBoxNav.Items690"),
            resources.GetString("colorComboBoxNav.Items691"),
            resources.GetString("colorComboBoxNav.Items692"),
            resources.GetString("colorComboBoxNav.Items693"),
            resources.GetString("colorComboBoxNav.Items694"),
            resources.GetString("colorComboBoxNav.Items695"),
            resources.GetString("colorComboBoxNav.Items696"),
            resources.GetString("colorComboBoxNav.Items697"),
            resources.GetString("colorComboBoxNav.Items698"),
            resources.GetString("colorComboBoxNav.Items699"),
            resources.GetString("colorComboBoxNav.Items700"),
            resources.GetString("colorComboBoxNav.Items701"),
            resources.GetString("colorComboBoxNav.Items702"),
            resources.GetString("colorComboBoxNav.Items703"),
            resources.GetString("colorComboBoxNav.Items704"),
            resources.GetString("colorComboBoxNav.Items705"),
            resources.GetString("colorComboBoxNav.Items706"),
            resources.GetString("colorComboBoxNav.Items707"),
            resources.GetString("colorComboBoxNav.Items708"),
            resources.GetString("colorComboBoxNav.Items709"),
            resources.GetString("colorComboBoxNav.Items710"),
            resources.GetString("colorComboBoxNav.Items711"),
            resources.GetString("colorComboBoxNav.Items712"),
            resources.GetString("colorComboBoxNav.Items713"),
            resources.GetString("colorComboBoxNav.Items714"),
            resources.GetString("colorComboBoxNav.Items715"),
            resources.GetString("colorComboBoxNav.Items716"),
            resources.GetString("colorComboBoxNav.Items717"),
            resources.GetString("colorComboBoxNav.Items718"),
            resources.GetString("colorComboBoxNav.Items719"),
            resources.GetString("colorComboBoxNav.Items720"),
            resources.GetString("colorComboBoxNav.Items721"),
            resources.GetString("colorComboBoxNav.Items722"),
            resources.GetString("colorComboBoxNav.Items723"),
            resources.GetString("colorComboBoxNav.Items724"),
            resources.GetString("colorComboBoxNav.Items725"),
            resources.GetString("colorComboBoxNav.Items726"),
            resources.GetString("colorComboBoxNav.Items727"),
            resources.GetString("colorComboBoxNav.Items728"),
            resources.GetString("colorComboBoxNav.Items729"),
            resources.GetString("colorComboBoxNav.Items730"),
            resources.GetString("colorComboBoxNav.Items731"),
            resources.GetString("colorComboBoxNav.Items732"),
            resources.GetString("colorComboBoxNav.Items733"),
            resources.GetString("colorComboBoxNav.Items734"),
            resources.GetString("colorComboBoxNav.Items735"),
            resources.GetString("colorComboBoxNav.Items736"),
            resources.GetString("colorComboBoxNav.Items737"),
            resources.GetString("colorComboBoxNav.Items738"),
            resources.GetString("colorComboBoxNav.Items739"),
            resources.GetString("colorComboBoxNav.Items740"),
            resources.GetString("colorComboBoxNav.Items741"),
            resources.GetString("colorComboBoxNav.Items742"),
            resources.GetString("colorComboBoxNav.Items743"),
            resources.GetString("colorComboBoxNav.Items744"),
            resources.GetString("colorComboBoxNav.Items745"),
            resources.GetString("colorComboBoxNav.Items746"),
            resources.GetString("colorComboBoxNav.Items747"),
            resources.GetString("colorComboBoxNav.Items748"),
            resources.GetString("colorComboBoxNav.Items749"),
            resources.GetString("colorComboBoxNav.Items750"),
            resources.GetString("colorComboBoxNav.Items751"),
            resources.GetString("colorComboBoxNav.Items752"),
            resources.GetString("colorComboBoxNav.Items753"),
            resources.GetString("colorComboBoxNav.Items754"),
            resources.GetString("colorComboBoxNav.Items755"),
            resources.GetString("colorComboBoxNav.Items756"),
            resources.GetString("colorComboBoxNav.Items757"),
            resources.GetString("colorComboBoxNav.Items758"),
            resources.GetString("colorComboBoxNav.Items759"),
            resources.GetString("colorComboBoxNav.Items760"),
            resources.GetString("colorComboBoxNav.Items761"),
            resources.GetString("colorComboBoxNav.Items762"),
            resources.GetString("colorComboBoxNav.Items763"),
            resources.GetString("colorComboBoxNav.Items764"),
            resources.GetString("colorComboBoxNav.Items765"),
            resources.GetString("colorComboBoxNav.Items766"),
            resources.GetString("colorComboBoxNav.Items767"),
            resources.GetString("colorComboBoxNav.Items768"),
            resources.GetString("colorComboBoxNav.Items769"),
            resources.GetString("colorComboBoxNav.Items770"),
            resources.GetString("colorComboBoxNav.Items771"),
            resources.GetString("colorComboBoxNav.Items772"),
            resources.GetString("colorComboBoxNav.Items773"),
            resources.GetString("colorComboBoxNav.Items774"),
            resources.GetString("colorComboBoxNav.Items775"),
            resources.GetString("colorComboBoxNav.Items776"),
            resources.GetString("colorComboBoxNav.Items777"),
            resources.GetString("colorComboBoxNav.Items778"),
            resources.GetString("colorComboBoxNav.Items779"),
            resources.GetString("colorComboBoxNav.Items780"),
            resources.GetString("colorComboBoxNav.Items781"),
            resources.GetString("colorComboBoxNav.Items782"),
            resources.GetString("colorComboBoxNav.Items783"),
            resources.GetString("colorComboBoxNav.Items784"),
            resources.GetString("colorComboBoxNav.Items785"),
            resources.GetString("colorComboBoxNav.Items786"),
            resources.GetString("colorComboBoxNav.Items787"),
            resources.GetString("colorComboBoxNav.Items788"),
            resources.GetString("colorComboBoxNav.Items789"),
            resources.GetString("colorComboBoxNav.Items790"),
            resources.GetString("colorComboBoxNav.Items791"),
            resources.GetString("colorComboBoxNav.Items792"),
            resources.GetString("colorComboBoxNav.Items793"),
            resources.GetString("colorComboBoxNav.Items794"),
            resources.GetString("colorComboBoxNav.Items795"),
            resources.GetString("colorComboBoxNav.Items796"),
            resources.GetString("colorComboBoxNav.Items797"),
            resources.GetString("colorComboBoxNav.Items798"),
            resources.GetString("colorComboBoxNav.Items799"),
            resources.GetString("colorComboBoxNav.Items800"),
            resources.GetString("colorComboBoxNav.Items801"),
            resources.GetString("colorComboBoxNav.Items802"),
            resources.GetString("colorComboBoxNav.Items803"),
            resources.GetString("colorComboBoxNav.Items804"),
            resources.GetString("colorComboBoxNav.Items805"),
            resources.GetString("colorComboBoxNav.Items806"),
            resources.GetString("colorComboBoxNav.Items807"),
            resources.GetString("colorComboBoxNav.Items808"),
            resources.GetString("colorComboBoxNav.Items809"),
            resources.GetString("colorComboBoxNav.Items810"),
            resources.GetString("colorComboBoxNav.Items811"),
            resources.GetString("colorComboBoxNav.Items812"),
            resources.GetString("colorComboBoxNav.Items813"),
            resources.GetString("colorComboBoxNav.Items814"),
            resources.GetString("colorComboBoxNav.Items815"),
            resources.GetString("colorComboBoxNav.Items816"),
            resources.GetString("colorComboBoxNav.Items817"),
            resources.GetString("colorComboBoxNav.Items818"),
            resources.GetString("colorComboBoxNav.Items819"),
            resources.GetString("colorComboBoxNav.Items820"),
            resources.GetString("colorComboBoxNav.Items821"),
            resources.GetString("colorComboBoxNav.Items822"),
            resources.GetString("colorComboBoxNav.Items823"),
            resources.GetString("colorComboBoxNav.Items824"),
            resources.GetString("colorComboBoxNav.Items825"),
            resources.GetString("colorComboBoxNav.Items826"),
            resources.GetString("colorComboBoxNav.Items827"),
            resources.GetString("colorComboBoxNav.Items828"),
            resources.GetString("colorComboBoxNav.Items829"),
            resources.GetString("colorComboBoxNav.Items830"),
            resources.GetString("colorComboBoxNav.Items831"),
            resources.GetString("colorComboBoxNav.Items832"),
            resources.GetString("colorComboBoxNav.Items833"),
            resources.GetString("colorComboBoxNav.Items834"),
            resources.GetString("colorComboBoxNav.Items835"),
            resources.GetString("colorComboBoxNav.Items836"),
            resources.GetString("colorComboBoxNav.Items837"),
            resources.GetString("colorComboBoxNav.Items838"),
            resources.GetString("colorComboBoxNav.Items839"),
            resources.GetString("colorComboBoxNav.Items840"),
            resources.GetString("colorComboBoxNav.Items841"),
            resources.GetString("colorComboBoxNav.Items842"),
            resources.GetString("colorComboBoxNav.Items843"),
            resources.GetString("colorComboBoxNav.Items844"),
            resources.GetString("colorComboBoxNav.Items845"),
            resources.GetString("colorComboBoxNav.Items846"),
            resources.GetString("colorComboBoxNav.Items847"),
            resources.GetString("colorComboBoxNav.Items848"),
            resources.GetString("colorComboBoxNav.Items849"),
            resources.GetString("colorComboBoxNav.Items850"),
            resources.GetString("colorComboBoxNav.Items851"),
            resources.GetString("colorComboBoxNav.Items852"),
            resources.GetString("colorComboBoxNav.Items853"),
            resources.GetString("colorComboBoxNav.Items854"),
            resources.GetString("colorComboBoxNav.Items855"),
            resources.GetString("colorComboBoxNav.Items856"),
            resources.GetString("colorComboBoxNav.Items857"),
            resources.GetString("colorComboBoxNav.Items858"),
            resources.GetString("colorComboBoxNav.Items859"),
            resources.GetString("colorComboBoxNav.Items860"),
            resources.GetString("colorComboBoxNav.Items861"),
            resources.GetString("colorComboBoxNav.Items862"),
            resources.GetString("colorComboBoxNav.Items863"),
            resources.GetString("colorComboBoxNav.Items864"),
            resources.GetString("colorComboBoxNav.Items865"),
            resources.GetString("colorComboBoxNav.Items866"),
            resources.GetString("colorComboBoxNav.Items867"),
            resources.GetString("colorComboBoxNav.Items868"),
            resources.GetString("colorComboBoxNav.Items869"),
            resources.GetString("colorComboBoxNav.Items870"),
            resources.GetString("colorComboBoxNav.Items871"),
            resources.GetString("colorComboBoxNav.Items872"),
            resources.GetString("colorComboBoxNav.Items873"),
            resources.GetString("colorComboBoxNav.Items874"),
            resources.GetString("colorComboBoxNav.Items875"),
            resources.GetString("colorComboBoxNav.Items876"),
            resources.GetString("colorComboBoxNav.Items877"),
            resources.GetString("colorComboBoxNav.Items878"),
            resources.GetString("colorComboBoxNav.Items879"),
            resources.GetString("colorComboBoxNav.Items880"),
            resources.GetString("colorComboBoxNav.Items881"),
            resources.GetString("colorComboBoxNav.Items882"),
            resources.GetString("colorComboBoxNav.Items883"),
            resources.GetString("colorComboBoxNav.Items884"),
            resources.GetString("colorComboBoxNav.Items885"),
            resources.GetString("colorComboBoxNav.Items886"),
            resources.GetString("colorComboBoxNav.Items887"),
            resources.GetString("colorComboBoxNav.Items888"),
            resources.GetString("colorComboBoxNav.Items889"),
            resources.GetString("colorComboBoxNav.Items890"),
            resources.GetString("colorComboBoxNav.Items891"),
            resources.GetString("colorComboBoxNav.Items892"),
            resources.GetString("colorComboBoxNav.Items893"),
            resources.GetString("colorComboBoxNav.Items894"),
            resources.GetString("colorComboBoxNav.Items895"),
            resources.GetString("colorComboBoxNav.Items896"),
            resources.GetString("colorComboBoxNav.Items897"),
            resources.GetString("colorComboBoxNav.Items898"),
            resources.GetString("colorComboBoxNav.Items899"),
            resources.GetString("colorComboBoxNav.Items900"),
            resources.GetString("colorComboBoxNav.Items901"),
            resources.GetString("colorComboBoxNav.Items902"),
            resources.GetString("colorComboBoxNav.Items903"),
            resources.GetString("colorComboBoxNav.Items904"),
            resources.GetString("colorComboBoxNav.Items905"),
            resources.GetString("colorComboBoxNav.Items906"),
            resources.GetString("colorComboBoxNav.Items907"),
            resources.GetString("colorComboBoxNav.Items908"),
            resources.GetString("colorComboBoxNav.Items909"),
            resources.GetString("colorComboBoxNav.Items910"),
            resources.GetString("colorComboBoxNav.Items911"),
            resources.GetString("colorComboBoxNav.Items912"),
            resources.GetString("colorComboBoxNav.Items913"),
            resources.GetString("colorComboBoxNav.Items914"),
            resources.GetString("colorComboBoxNav.Items915"),
            resources.GetString("colorComboBoxNav.Items916"),
            resources.GetString("colorComboBoxNav.Items917"),
            resources.GetString("colorComboBoxNav.Items918"),
            resources.GetString("colorComboBoxNav.Items919"),
            resources.GetString("colorComboBoxNav.Items920"),
            resources.GetString("colorComboBoxNav.Items921"),
            resources.GetString("colorComboBoxNav.Items922"),
            resources.GetString("colorComboBoxNav.Items923"),
            resources.GetString("colorComboBoxNav.Items924"),
            resources.GetString("colorComboBoxNav.Items925"),
            resources.GetString("colorComboBoxNav.Items926"),
            resources.GetString("colorComboBoxNav.Items927"),
            resources.GetString("colorComboBoxNav.Items928"),
            resources.GetString("colorComboBoxNav.Items929"),
            resources.GetString("colorComboBoxNav.Items930"),
            resources.GetString("colorComboBoxNav.Items931"),
            resources.GetString("colorComboBoxNav.Items932"),
            resources.GetString("colorComboBoxNav.Items933"),
            resources.GetString("colorComboBoxNav.Items934"),
            resources.GetString("colorComboBoxNav.Items935"),
            resources.GetString("colorComboBoxNav.Items936"),
            resources.GetString("colorComboBoxNav.Items937"),
            resources.GetString("colorComboBoxNav.Items938"),
            resources.GetString("colorComboBoxNav.Items939"),
            resources.GetString("colorComboBoxNav.Items940"),
            resources.GetString("colorComboBoxNav.Items941"),
            resources.GetString("colorComboBoxNav.Items942"),
            resources.GetString("colorComboBoxNav.Items943"),
            resources.GetString("colorComboBoxNav.Items944"),
            resources.GetString("colorComboBoxNav.Items945"),
            resources.GetString("colorComboBoxNav.Items946"),
            resources.GetString("colorComboBoxNav.Items947"),
            resources.GetString("colorComboBoxNav.Items948"),
            resources.GetString("colorComboBoxNav.Items949"),
            resources.GetString("colorComboBoxNav.Items950"),
            resources.GetString("colorComboBoxNav.Items951"),
            resources.GetString("colorComboBoxNav.Items952"),
            resources.GetString("colorComboBoxNav.Items953"),
            resources.GetString("colorComboBoxNav.Items954"),
            resources.GetString("colorComboBoxNav.Items955"),
            resources.GetString("colorComboBoxNav.Items956"),
            resources.GetString("colorComboBoxNav.Items957"),
            resources.GetString("colorComboBoxNav.Items958"),
            resources.GetString("colorComboBoxNav.Items959"),
            resources.GetString("colorComboBoxNav.Items960"),
            resources.GetString("colorComboBoxNav.Items961"),
            resources.GetString("colorComboBoxNav.Items962"),
            resources.GetString("colorComboBoxNav.Items963"),
            resources.GetString("colorComboBoxNav.Items964"),
            resources.GetString("colorComboBoxNav.Items965"),
            resources.GetString("colorComboBoxNav.Items966"),
            resources.GetString("colorComboBoxNav.Items967"),
            resources.GetString("colorComboBoxNav.Items968"),
            resources.GetString("colorComboBoxNav.Items969"),
            resources.GetString("colorComboBoxNav.Items970"),
            resources.GetString("colorComboBoxNav.Items971"),
            resources.GetString("colorComboBoxNav.Items972"),
            resources.GetString("colorComboBoxNav.Items973"),
            resources.GetString("colorComboBoxNav.Items974"),
            resources.GetString("colorComboBoxNav.Items975"),
            resources.GetString("colorComboBoxNav.Items976"),
            resources.GetString("colorComboBoxNav.Items977"),
            resources.GetString("colorComboBoxNav.Items978"),
            resources.GetString("colorComboBoxNav.Items979"),
            resources.GetString("colorComboBoxNav.Items980"),
            resources.GetString("colorComboBoxNav.Items981"),
            resources.GetString("colorComboBoxNav.Items982"),
            resources.GetString("colorComboBoxNav.Items983"),
            resources.GetString("colorComboBoxNav.Items984"),
            resources.GetString("colorComboBoxNav.Items985"),
            resources.GetString("colorComboBoxNav.Items986"),
            resources.GetString("colorComboBoxNav.Items987"),
            resources.GetString("colorComboBoxNav.Items988"),
            resources.GetString("colorComboBoxNav.Items989"),
            resources.GetString("colorComboBoxNav.Items990"),
            resources.GetString("colorComboBoxNav.Items991"),
            resources.GetString("colorComboBoxNav.Items992"),
            resources.GetString("colorComboBoxNav.Items993"),
            resources.GetString("colorComboBoxNav.Items994"),
            resources.GetString("colorComboBoxNav.Items995"),
            resources.GetString("colorComboBoxNav.Items996"),
            resources.GetString("colorComboBoxNav.Items997"),
            resources.GetString("colorComboBoxNav.Items998"),
            resources.GetString("colorComboBoxNav.Items999"),
            resources.GetString("colorComboBoxNav.Items1000"),
            resources.GetString("colorComboBoxNav.Items1001"),
            resources.GetString("colorComboBoxNav.Items1002"),
            resources.GetString("colorComboBoxNav.Items1003"),
            resources.GetString("colorComboBoxNav.Items1004"),
            resources.GetString("colorComboBoxNav.Items1005"),
            resources.GetString("colorComboBoxNav.Items1006"),
            resources.GetString("colorComboBoxNav.Items1007"),
            resources.GetString("colorComboBoxNav.Items1008"),
            resources.GetString("colorComboBoxNav.Items1009"),
            resources.GetString("colorComboBoxNav.Items1010"),
            resources.GetString("colorComboBoxNav.Items1011"),
            resources.GetString("colorComboBoxNav.Items1012"),
            resources.GetString("colorComboBoxNav.Items1013"),
            resources.GetString("colorComboBoxNav.Items1014"),
            resources.GetString("colorComboBoxNav.Items1015"),
            resources.GetString("colorComboBoxNav.Items1016"),
            resources.GetString("colorComboBoxNav.Items1017"),
            resources.GetString("colorComboBoxNav.Items1018"),
            resources.GetString("colorComboBoxNav.Items1019"),
            resources.GetString("colorComboBoxNav.Items1020"),
            resources.GetString("colorComboBoxNav.Items1021"),
            resources.GetString("colorComboBoxNav.Items1022"),
            resources.GetString("colorComboBoxNav.Items1023"),
            resources.GetString("colorComboBoxNav.Items1024"),
            resources.GetString("colorComboBoxNav.Items1025"),
            resources.GetString("colorComboBoxNav.Items1026"),
            resources.GetString("colorComboBoxNav.Items1027"),
            resources.GetString("colorComboBoxNav.Items1028"),
            resources.GetString("colorComboBoxNav.Items1029"),
            resources.GetString("colorComboBoxNav.Items1030"),
            resources.GetString("colorComboBoxNav.Items1031"),
            resources.GetString("colorComboBoxNav.Items1032"),
            resources.GetString("colorComboBoxNav.Items1033"),
            resources.GetString("colorComboBoxNav.Items1034"),
            resources.GetString("colorComboBoxNav.Items1035"),
            resources.GetString("colorComboBoxNav.Items1036"),
            resources.GetString("colorComboBoxNav.Items1037"),
            resources.GetString("colorComboBoxNav.Items1038"),
            resources.GetString("colorComboBoxNav.Items1039"),
            resources.GetString("colorComboBoxNav.Items1040"),
            resources.GetString("colorComboBoxNav.Items1041"),
            resources.GetString("colorComboBoxNav.Items1042"),
            resources.GetString("colorComboBoxNav.Items1043"),
            resources.GetString("colorComboBoxNav.Items1044"),
            resources.GetString("colorComboBoxNav.Items1045"),
            resources.GetString("colorComboBoxNav.Items1046"),
            resources.GetString("colorComboBoxNav.Items1047"),
            resources.GetString("colorComboBoxNav.Items1048"),
            resources.GetString("colorComboBoxNav.Items1049"),
            resources.GetString("colorComboBoxNav.Items1050"),
            resources.GetString("colorComboBoxNav.Items1051"),
            resources.GetString("colorComboBoxNav.Items1052"),
            resources.GetString("colorComboBoxNav.Items1053"),
            resources.GetString("colorComboBoxNav.Items1054"),
            resources.GetString("colorComboBoxNav.Items1055"),
            resources.GetString("colorComboBoxNav.Items1056"),
            resources.GetString("colorComboBoxNav.Items1057"),
            resources.GetString("colorComboBoxNav.Items1058"),
            resources.GetString("colorComboBoxNav.Items1059"),
            resources.GetString("colorComboBoxNav.Items1060"),
            resources.GetString("colorComboBoxNav.Items1061"),
            resources.GetString("colorComboBoxNav.Items1062"),
            resources.GetString("colorComboBoxNav.Items1063"),
            resources.GetString("colorComboBoxNav.Items1064"),
            resources.GetString("colorComboBoxNav.Items1065"),
            resources.GetString("colorComboBoxNav.Items1066"),
            resources.GetString("colorComboBoxNav.Items1067"),
            resources.GetString("colorComboBoxNav.Items1068"),
            resources.GetString("colorComboBoxNav.Items1069"),
            resources.GetString("colorComboBoxNav.Items1070"),
            resources.GetString("colorComboBoxNav.Items1071"),
            resources.GetString("colorComboBoxNav.Items1072"),
            resources.GetString("colorComboBoxNav.Items1073"),
            resources.GetString("colorComboBoxNav.Items1074"),
            resources.GetString("colorComboBoxNav.Items1075"),
            resources.GetString("colorComboBoxNav.Items1076"),
            resources.GetString("colorComboBoxNav.Items1077"),
            resources.GetString("colorComboBoxNav.Items1078"),
            resources.GetString("colorComboBoxNav.Items1079"),
            resources.GetString("colorComboBoxNav.Items1080"),
            resources.GetString("colorComboBoxNav.Items1081"),
            resources.GetString("colorComboBoxNav.Items1082"),
            resources.GetString("colorComboBoxNav.Items1083"),
            resources.GetString("colorComboBoxNav.Items1084"),
            resources.GetString("colorComboBoxNav.Items1085"),
            resources.GetString("colorComboBoxNav.Items1086"),
            resources.GetString("colorComboBoxNav.Items1087"),
            resources.GetString("colorComboBoxNav.Items1088"),
            resources.GetString("colorComboBoxNav.Items1089"),
            resources.GetString("colorComboBoxNav.Items1090"),
            resources.GetString("colorComboBoxNav.Items1091"),
            resources.GetString("colorComboBoxNav.Items1092"),
            resources.GetString("colorComboBoxNav.Items1093"),
            resources.GetString("colorComboBoxNav.Items1094"),
            resources.GetString("colorComboBoxNav.Items1095"),
            resources.GetString("colorComboBoxNav.Items1096"),
            resources.GetString("colorComboBoxNav.Items1097"),
            resources.GetString("colorComboBoxNav.Items1098"),
            resources.GetString("colorComboBoxNav.Items1099"),
            resources.GetString("colorComboBoxNav.Items1100"),
            resources.GetString("colorComboBoxNav.Items1101"),
            resources.GetString("colorComboBoxNav.Items1102"),
            resources.GetString("colorComboBoxNav.Items1103"),
            resources.GetString("colorComboBoxNav.Items1104"),
            resources.GetString("colorComboBoxNav.Items1105"),
            resources.GetString("colorComboBoxNav.Items1106"),
            resources.GetString("colorComboBoxNav.Items1107"),
            resources.GetString("colorComboBoxNav.Items1108"),
            resources.GetString("colorComboBoxNav.Items1109"),
            resources.GetString("colorComboBoxNav.Items1110"),
            resources.GetString("colorComboBoxNav.Items1111"),
            resources.GetString("colorComboBoxNav.Items1112"),
            resources.GetString("colorComboBoxNav.Items1113"),
            resources.GetString("colorComboBoxNav.Items1114"),
            resources.GetString("colorComboBoxNav.Items1115"),
            resources.GetString("colorComboBoxNav.Items1116"),
            resources.GetString("colorComboBoxNav.Items1117"),
            resources.GetString("colorComboBoxNav.Items1118"),
            resources.GetString("colorComboBoxNav.Items1119"),
            resources.GetString("colorComboBoxNav.Items1120"),
            resources.GetString("colorComboBoxNav.Items1121"),
            resources.GetString("colorComboBoxNav.Items1122"),
            resources.GetString("colorComboBoxNav.Items1123"),
            resources.GetString("colorComboBoxNav.Items1124"),
            resources.GetString("colorComboBoxNav.Items1125"),
            resources.GetString("colorComboBoxNav.Items1126"),
            resources.GetString("colorComboBoxNav.Items1127")});
            resources.ApplyResources(this.colorComboBoxNav, "colorComboBoxNav");
            this.colorComboBoxNav.Name = "colorComboBoxNav";
            // 
            // btnHighLightGoLast
            // 
            resources.ApplyResources(this.btnHighLightGoLast, "btnHighLightGoLast");
            this.btnHighLightGoLast.Name = "btnHighLightGoLast";
            this.btnHighLightGoLast.UseVisualStyleBackColor = true;
            this.btnHighLightGoLast.Click += new System.EventHandler(this.btnHighLightGoLast_Click);
            // 
            // btnONFieldNavLast
            // 
            resources.ApplyResources(this.btnONFieldNavLast, "btnONFieldNavLast");
            this.btnONFieldNavLast.Name = "btnONFieldNavLast";
            this.btnONFieldNavLast.UseVisualStyleBackColor = true;
            this.btnONFieldNavLast.Click += new System.EventHandler(this.btnONFieldNavLast_Click);
            // 
            // label75
            // 
            resources.ApplyResources(this.label75, "label75");
            this.label75.Name = "label75";
            // 
            // label67
            // 
            resources.ApplyResources(this.label67, "label67");
            this.label67.Name = "label67";
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.chkListObjNavOutline);
            this.groupBox13.Controls.Add(this.btnONOutlineAllUnSel);
            this.groupBox13.Controls.Add(this.label71);
            this.groupBox13.Controls.Add(this.btnONHeadingNavFirst);
            this.groupBox13.Controls.Add(this.btnONOutlineAllSel);
            this.groupBox13.Controls.Add(this.btnONHeadingNavPrev);
            this.groupBox13.Controls.Add(this.btnONHeadingNavLast);
            this.groupBox13.Controls.Add(this.btnONHeadingNavNext);
            this.groupBox13.Controls.Add(this.label72);
            this.groupBox13.Controls.Add(this.label70);
            resources.ApplyResources(this.groupBox13, "groupBox13");
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.TabStop = false;
            // 
            // chkListObjNavOutline
            // 
            this.chkListObjNavOutline.CheckOnClick = true;
            resources.ApplyResources(this.chkListObjNavOutline, "chkListObjNavOutline");
            this.chkListObjNavOutline.FormattingEnabled = true;
            this.chkListObjNavOutline.Items.AddRange(new object[] {
            resources.GetString("chkListObjNavOutline.Items"),
            resources.GetString("chkListObjNavOutline.Items1"),
            resources.GetString("chkListObjNavOutline.Items2"),
            resources.GetString("chkListObjNavOutline.Items3"),
            resources.GetString("chkListObjNavOutline.Items4"),
            resources.GetString("chkListObjNavOutline.Items5"),
            resources.GetString("chkListObjNavOutline.Items6"),
            resources.GetString("chkListObjNavOutline.Items7"),
            resources.GetString("chkListObjNavOutline.Items8")});
            this.chkListObjNavOutline.MultiColumn = true;
            this.chkListObjNavOutline.Name = "chkListObjNavOutline";
            // 
            // btnONOutlineAllUnSel
            // 
            resources.ApplyResources(this.btnONOutlineAllUnSel, "btnONOutlineAllUnSel");
            this.btnONOutlineAllUnSel.Name = "btnONOutlineAllUnSel";
            this.btnONOutlineAllUnSel.UseVisualStyleBackColor = true;
            this.btnONOutlineAllUnSel.Click += new System.EventHandler(this.btnONOutlineAllUnSel_Click);
            // 
            // label71
            // 
            resources.ApplyResources(this.label71, "label71");
            this.label71.Name = "label71";
            // 
            // btnONHeadingNavFirst
            // 
            resources.ApplyResources(this.btnONHeadingNavFirst, "btnONHeadingNavFirst");
            this.btnONHeadingNavFirst.Name = "btnONHeadingNavFirst";
            this.btnONHeadingNavFirst.UseVisualStyleBackColor = true;
            this.btnONHeadingNavFirst.Click += new System.EventHandler(this.btnONHeadingNavFirst_Click);
            // 
            // btnONOutlineAllSel
            // 
            resources.ApplyResources(this.btnONOutlineAllSel, "btnONOutlineAllSel");
            this.btnONOutlineAllSel.Name = "btnONOutlineAllSel";
            this.btnONOutlineAllSel.UseVisualStyleBackColor = true;
            this.btnONOutlineAllSel.Click += new System.EventHandler(this.btnONOutlineAllSel_Click);
            // 
            // btnONHeadingNavPrev
            // 
            resources.ApplyResources(this.btnONHeadingNavPrev, "btnONHeadingNavPrev");
            this.btnONHeadingNavPrev.Name = "btnONHeadingNavPrev";
            this.btnONHeadingNavPrev.UseVisualStyleBackColor = true;
            this.btnONHeadingNavPrev.Click += new System.EventHandler(this.btnONHeadingNavPrev_Click);
            // 
            // btnONHeadingNavLast
            // 
            resources.ApplyResources(this.btnONHeadingNavLast, "btnONHeadingNavLast");
            this.btnONHeadingNavLast.Name = "btnONHeadingNavLast";
            this.btnONHeadingNavLast.UseVisualStyleBackColor = true;
            this.btnONHeadingNavLast.Click += new System.EventHandler(this.btnONHeadingNavLast_Click);
            // 
            // btnONHeadingNavNext
            // 
            resources.ApplyResources(this.btnONHeadingNavNext, "btnONHeadingNavNext");
            this.btnONHeadingNavNext.Name = "btnONHeadingNavNext";
            this.btnONHeadingNavNext.UseVisualStyleBackColor = true;
            this.btnONHeadingNavNext.Click += new System.EventHandler(this.btnONHeadingNavNext_Click);
            // 
            // label72
            // 
            resources.ApplyResources(this.label72, "label72");
            this.label72.Name = "label72";
            // 
            // label70
            // 
            resources.ApplyResources(this.label70, "label70");
            this.label70.Name = "label70";
            // 
            // btnONSectionNavLast
            // 
            resources.ApplyResources(this.btnONSectionNavLast, "btnONSectionNavLast");
            this.btnONSectionNavLast.Name = "btnONSectionNavLast";
            this.btnONSectionNavLast.UseVisualStyleBackColor = true;
            this.btnONSectionNavLast.Click += new System.EventHandler(this.btnONSectionNavLast_Click);
            // 
            // label76
            // 
            resources.ApplyResources(this.label76, "label76");
            this.label76.Name = "label76";
            // 
            // btnONPageNavLast
            // 
            resources.ApplyResources(this.btnONPageNavLast, "btnONPageNavLast");
            this.btnONPageNavLast.Name = "btnONPageNavLast";
            this.btnONPageNavLast.UseVisualStyleBackColor = true;
            this.btnONPageNavLast.Click += new System.EventHandler(this.btnONPageNavLast_Click);
            // 
            // label64
            // 
            resources.ApplyResources(this.label64, "label64");
            this.label64.Name = "label64";
            // 
            // btnHighLightGoNext
            // 
            resources.ApplyResources(this.btnHighLightGoNext, "btnHighLightGoNext");
            this.btnHighLightGoNext.Name = "btnHighLightGoNext";
            this.btnHighLightGoNext.UseVisualStyleBackColor = true;
            this.btnHighLightGoNext.Click += new System.EventHandler(this.btnHighLightGoNext_Click);
            // 
            // btnONFieldNavNext
            // 
            resources.ApplyResources(this.btnONFieldNavNext, "btnONFieldNavNext");
            this.btnONFieldNavNext.Name = "btnONFieldNavNext";
            this.btnONFieldNavNext.UseVisualStyleBackColor = true;
            this.btnONFieldNavNext.Click += new System.EventHandler(this.btnONFieldNavNext_Click);
            // 
            // btnONGraphicNavLast
            // 
            resources.ApplyResources(this.btnONGraphicNavLast, "btnONGraphicNavLast");
            this.btnONGraphicNavLast.Name = "btnONGraphicNavLast";
            this.btnONGraphicNavLast.UseVisualStyleBackColor = true;
            this.btnONGraphicNavLast.Click += new System.EventHandler(this.btnONGraphicNavLast_Click);
            // 
            // btnONSectionNavNext
            // 
            resources.ApplyResources(this.btnONSectionNavNext, "btnONSectionNavNext");
            this.btnONSectionNavNext.Name = "btnONSectionNavNext";
            this.btnONSectionNavNext.UseVisualStyleBackColor = true;
            this.btnONSectionNavNext.Click += new System.EventHandler(this.btnONSectionNavNext_Click);
            // 
            // label63
            // 
            resources.ApplyResources(this.label63, "label63");
            this.label63.Name = "label63";
            // 
            // btnONPageNavNext
            // 
            resources.ApplyResources(this.btnONPageNavNext, "btnONPageNavNext");
            this.btnONPageNavNext.Name = "btnONPageNavNext";
            this.btnONPageNavNext.UseVisualStyleBackColor = true;
            this.btnONPageNavNext.Click += new System.EventHandler(this.btnONPageNavNext_Click);
            // 
            // btnHighLightGoFirst
            // 
            resources.ApplyResources(this.btnHighLightGoFirst, "btnHighLightGoFirst");
            this.btnHighLightGoFirst.Name = "btnHighLightGoFirst";
            this.btnHighLightGoFirst.UseVisualStyleBackColor = true;
            this.btnHighLightGoFirst.Click += new System.EventHandler(this.btnHighLightGoFirst_Click);
            // 
            // btnONFieldNavFirst
            // 
            resources.ApplyResources(this.btnONFieldNavFirst, "btnONFieldNavFirst");
            this.btnONFieldNavFirst.Name = "btnONFieldNavFirst";
            this.btnONFieldNavFirst.UseVisualStyleBackColor = true;
            this.btnONFieldNavFirst.Click += new System.EventHandler(this.btnONFieldNavFirst_Click);
            // 
            // btnONTableNavLast
            // 
            resources.ApplyResources(this.btnONTableNavLast, "btnONTableNavLast");
            this.btnONTableNavLast.Name = "btnONTableNavLast";
            this.btnONTableNavLast.UseVisualStyleBackColor = true;
            this.btnONTableNavLast.Click += new System.EventHandler(this.btnONTableNavLast_Click);
            // 
            // btnONSectionNavFirst
            // 
            resources.ApplyResources(this.btnONSectionNavFirst, "btnONSectionNavFirst");
            this.btnONSectionNavFirst.Name = "btnONSectionNavFirst";
            this.btnONSectionNavFirst.UseVisualStyleBackColor = true;
            this.btnONSectionNavFirst.Click += new System.EventHandler(this.btnONSectionNavFirst_Click);
            // 
            // btnONGraphicNavNext
            // 
            resources.ApplyResources(this.btnONGraphicNavNext, "btnONGraphicNavNext");
            this.btnONGraphicNavNext.Name = "btnONGraphicNavNext";
            this.btnONGraphicNavNext.UseVisualStyleBackColor = true;
            this.btnONGraphicNavNext.Click += new System.EventHandler(this.btnONGraphicNavNext_Click);
            // 
            // btnONPageNavFirst
            // 
            resources.ApplyResources(this.btnONPageNavFirst, "btnONPageNavFirst");
            this.btnONPageNavFirst.Name = "btnONPageNavFirst";
            this.btnONPageNavFirst.UseVisualStyleBackColor = true;
            this.btnONPageNavFirst.Click += new System.EventHandler(this.btnONPageNavFirst_Click);
            // 
            // btnHighLightGoPrev
            // 
            resources.ApplyResources(this.btnHighLightGoPrev, "btnHighLightGoPrev");
            this.btnHighLightGoPrev.Name = "btnHighLightGoPrev";
            this.btnHighLightGoPrev.UseVisualStyleBackColor = true;
            this.btnHighLightGoPrev.Click += new System.EventHandler(this.btnHighLightGoPrev_Click);
            // 
            // btnONFieldNavPrev
            // 
            resources.ApplyResources(this.btnONFieldNavPrev, "btnONFieldNavPrev");
            this.btnONFieldNavPrev.Name = "btnONFieldNavPrev";
            this.btnONFieldNavPrev.UseVisualStyleBackColor = true;
            this.btnONFieldNavPrev.Click += new System.EventHandler(this.btnONFieldNavPrev_Click);
            // 
            // label65
            // 
            resources.ApplyResources(this.label65, "label65");
            this.label65.Name = "label65";
            // 
            // btnONSectionNavPrev
            // 
            resources.ApplyResources(this.btnONSectionNavPrev, "btnONSectionNavPrev");
            this.btnONSectionNavPrev.Name = "btnONSectionNavPrev";
            this.btnONSectionNavPrev.UseVisualStyleBackColor = true;
            this.btnONSectionNavPrev.Click += new System.EventHandler(this.btnONSectionNavPrev_Click);
            // 
            // btnONGraphicNavFirst
            // 
            resources.ApplyResources(this.btnONGraphicNavFirst, "btnONGraphicNavFirst");
            this.btnONGraphicNavFirst.Name = "btnONGraphicNavFirst";
            this.btnONGraphicNavFirst.UseVisualStyleBackColor = true;
            this.btnONGraphicNavFirst.Click += new System.EventHandler(this.btnONGraphicNavFirst_Click);
            // 
            // btnONPageNavPrev
            // 
            resources.ApplyResources(this.btnONPageNavPrev, "btnONPageNavPrev");
            this.btnONPageNavPrev.Name = "btnONPageNavPrev";
            this.btnONPageNavPrev.UseVisualStyleBackColor = true;
            this.btnONPageNavPrev.Click += new System.EventHandler(this.btnONPageNavPrev_Click);
            // 
            // btnONTableNavNext
            // 
            resources.ApplyResources(this.btnONTableNavNext, "btnONTableNavNext");
            this.btnONTableNavNext.Name = "btnONTableNavNext";
            this.btnONTableNavNext.UseVisualStyleBackColor = true;
            this.btnONTableNavNext.Click += new System.EventHandler(this.btnONTableNavNext_Click);
            // 
            // btnONGraphicNavPrev
            // 
            resources.ApplyResources(this.btnONGraphicNavPrev, "btnONGraphicNavPrev");
            this.btnONGraphicNavPrev.Name = "btnONGraphicNavPrev";
            this.btnONGraphicNavPrev.UseVisualStyleBackColor = true;
            this.btnONGraphicNavPrev.Click += new System.EventHandler(this.btnONGraphicNavPrev_Click);
            // 
            // btnONTableNavFirst
            // 
            resources.ApplyResources(this.btnONTableNavFirst, "btnONTableNavFirst");
            this.btnONTableNavFirst.Name = "btnONTableNavFirst";
            this.btnONTableNavFirst.UseVisualStyleBackColor = true;
            this.btnONTableNavFirst.Click += new System.EventHandler(this.btnONTableNavFirst_Click);
            // 
            // btnONTableNavPrev
            // 
            resources.ApplyResources(this.btnONTableNavPrev, "btnONTableNavPrev");
            this.btnONTableNavPrev.Name = "btnONTableNavPrev";
            this.btnONTableNavPrev.UseVisualStyleBackColor = true;
            this.btnONTableNavPrev.Click += new System.EventHandler(this.btnONTableNavPrev_Click);
            // 
            // tabPageMultiSel
            // 
            this.tabPageMultiSel.Controls.Add(this.ExcludeColorComboBox);
            this.tabPageMultiSel.Controls.Add(this.IncludeColorComboBox);
            this.tabPageMultiSel.Controls.Add(this.label73);
            this.tabPageMultiSel.Controls.Add(this.label13);
            this.tabPageMultiSel.Controls.Add(this.chkMultiSelUserDef);
            this.tabPageMultiSel.Controls.Add(this.groupBox10);
            this.tabPageMultiSel.Controls.Add(this.groupBox9);
            this.tabPageMultiSel.Controls.Add(this.btnMultiSelApplySel);
            this.tabPageMultiSel.Controls.Add(this.groupBox5);
            this.tabPageMultiSel.Controls.Add(this.groupBox3);
            this.tabPageMultiSel.Controls.Add(this.groupBox7);
            resources.ApplyResources(this.tabPageMultiSel, "tabPageMultiSel");
            this.tabPageMultiSel.Name = "tabPageMultiSel";
            this.tabPageMultiSel.UseVisualStyleBackColor = true;
            // 
            // ExcludeColorComboBox
            // 
            this.ExcludeColorComboBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.ExcludeColorComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ExcludeColorComboBox.FormattingEnabled = true;
            resources.ApplyResources(this.ExcludeColorComboBox, "ExcludeColorComboBox");
            this.ExcludeColorComboBox.Items.AddRange(new object[] {
            resources.GetString("ExcludeColorComboBox.Items"),
            resources.GetString("ExcludeColorComboBox.Items1")});
            this.ExcludeColorComboBox.Name = "ExcludeColorComboBox";
            // 
            // IncludeColorComboBox
            // 
            this.IncludeColorComboBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.IncludeColorComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.IncludeColorComboBox.FormattingEnabled = true;
            resources.ApplyResources(this.IncludeColorComboBox, "IncludeColorComboBox");
            this.IncludeColorComboBox.Items.AddRange(new object[] {
            resources.GetString("IncludeColorComboBox.Items"),
            resources.GetString("IncludeColorComboBox.Items1")});
            this.IncludeColorComboBox.Name = "IncludeColorComboBox";
            // 
            // label73
            // 
            resources.ApplyResources(this.label73, "label73");
            this.label73.Name = "label73";
            // 
            // label13
            // 
            resources.ApplyResources(this.label13, "label13");
            this.label13.Name = "label13";
            // 
            // chkMultiSelUserDef
            // 
            resources.ApplyResources(this.chkMultiSelUserDef, "chkMultiSelUserDef");
            this.chkMultiSelUserDef.Name = "chkMultiSelUserDef";
            this.chkMultiSelUserDef.UseVisualStyleBackColor = true;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.chkListBoxMultiListSnType);
            this.groupBox10.Controls.Add(this.groupBox16);
            this.groupBox10.Controls.Add(this.checkBoxMultiSelSnParas);
            resources.ApplyResources(this.groupBox10, "groupBox10");
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.TabStop = false;
            // 
            // chkListBoxMultiListSnType
            // 
            this.chkListBoxMultiListSnType.CheckOnClick = true;
            this.chkListBoxMultiListSnType.FormattingEnabled = true;
            this.chkListBoxMultiListSnType.Items.AddRange(new object[] {
            resources.GetString("chkListBoxMultiListSnType.Items"),
            resources.GetString("chkListBoxMultiListSnType.Items1"),
            resources.GetString("chkListBoxMultiListSnType.Items2"),
            resources.GetString("chkListBoxMultiListSnType.Items3"),
            resources.GetString("chkListBoxMultiListSnType.Items4"),
            resources.GetString("chkListBoxMultiListSnType.Items5"),
            resources.GetString("chkListBoxMultiListSnType.Items6"),
            resources.GetString("chkListBoxMultiListSnType.Items7"),
            resources.GetString("chkListBoxMultiListSnType.Items8"),
            resources.GetString("chkListBoxMultiListSnType.Items9"),
            resources.GetString("chkListBoxMultiListSnType.Items10"),
            resources.GetString("chkListBoxMultiListSnType.Items11"),
            resources.GetString("chkListBoxMultiListSnType.Items12")});
            resources.ApplyResources(this.chkListBoxMultiListSnType, "chkListBoxMultiListSnType");
            this.chkListBoxMultiListSnType.Name = "chkListBoxMultiListSnType";
            // 
            // groupBox16
            // 
            this.groupBox16.Controls.Add(this.rdBtnMultiSelIgnoreTbls);
            this.groupBox16.Controls.Add(this.rdBtnMultiSelOnlyTbls);
            this.groupBox16.Controls.Add(this.rdBtnMultiSelIncludeTbls);
            this.groupBox16.Controls.Add(this.chkBoxMultiSelIgnoreHeadings);
            this.groupBox16.Controls.Add(this.chkBoxMultiSelIgnoreTxtBody);
            resources.ApplyResources(this.groupBox16, "groupBox16");
            this.groupBox16.Name = "groupBox16";
            this.groupBox16.TabStop = false;
            // 
            // rdBtnMultiSelIgnoreTbls
            // 
            resources.ApplyResources(this.rdBtnMultiSelIgnoreTbls, "rdBtnMultiSelIgnoreTbls");
            this.rdBtnMultiSelIgnoreTbls.Checked = true;
            this.rdBtnMultiSelIgnoreTbls.Name = "rdBtnMultiSelIgnoreTbls";
            this.rdBtnMultiSelIgnoreTbls.TabStop = true;
            this.rdBtnMultiSelIgnoreTbls.UseVisualStyleBackColor = true;
            // 
            // rdBtnMultiSelOnlyTbls
            // 
            resources.ApplyResources(this.rdBtnMultiSelOnlyTbls, "rdBtnMultiSelOnlyTbls");
            this.rdBtnMultiSelOnlyTbls.Name = "rdBtnMultiSelOnlyTbls";
            this.rdBtnMultiSelOnlyTbls.UseVisualStyleBackColor = true;
            // 
            // rdBtnMultiSelIncludeTbls
            // 
            resources.ApplyResources(this.rdBtnMultiSelIncludeTbls, "rdBtnMultiSelIncludeTbls");
            this.rdBtnMultiSelIncludeTbls.Name = "rdBtnMultiSelIncludeTbls";
            this.rdBtnMultiSelIncludeTbls.UseVisualStyleBackColor = true;
            // 
            // chkBoxMultiSelIgnoreHeadings
            // 
            resources.ApplyResources(this.chkBoxMultiSelIgnoreHeadings, "chkBoxMultiSelIgnoreHeadings");
            this.chkBoxMultiSelIgnoreHeadings.Checked = true;
            this.chkBoxMultiSelIgnoreHeadings.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBoxMultiSelIgnoreHeadings.Name = "chkBoxMultiSelIgnoreHeadings";
            this.chkBoxMultiSelIgnoreHeadings.UseVisualStyleBackColor = true;
            // 
            // chkBoxMultiSelIgnoreTxtBody
            // 
            resources.ApplyResources(this.chkBoxMultiSelIgnoreTxtBody, "chkBoxMultiSelIgnoreTxtBody");
            this.chkBoxMultiSelIgnoreTxtBody.Name = "chkBoxMultiSelIgnoreTxtBody";
            this.chkBoxMultiSelIgnoreTxtBody.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelSnParas
            // 
            resources.ApplyResources(this.checkBoxMultiSelSnParas, "checkBoxMultiSelSnParas");
            this.checkBoxMultiSelSnParas.Name = "checkBoxMultiSelSnParas";
            this.checkBoxMultiSelSnParas.UseVisualStyleBackColor = true;
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.chkWholeTableCells);
            this.groupBox9.Controls.Add(this.chkBoxMultiSelLastColumn);
            this.groupBox9.Controls.Add(this.chkBoxMulSelTblLastRow);
            this.groupBox9.Controls.Add(this.numMultiSelColEnd);
            this.groupBox9.Controls.Add(this.numMultiSelColStart);
            this.groupBox9.Controls.Add(this.numMultiSelRowEnd);
            this.groupBox9.Controls.Add(this.numMultiSelRowStart);
            this.groupBox9.Controls.Add(this.label46);
            this.groupBox9.Controls.Add(this.label45);
            this.groupBox9.Controls.Add(this.chkBoxMultiSelColumnsScope);
            this.groupBox9.Controls.Add(this.chkBoxMultiSelRowsScope);
            this.groupBox9.Controls.Add(this.chkBoxMultiSelFirstColumn);
            this.groupBox9.Controls.Add(this.checkBoxMultiSelTables);
            this.groupBox9.Controls.Add(this.chkBoxMulSelTblFirstRow);
            resources.ApplyResources(this.groupBox9, "groupBox9");
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.TabStop = false;
            // 
            // chkWholeTableCells
            // 
            resources.ApplyResources(this.chkWholeTableCells, "chkWholeTableCells");
            this.chkWholeTableCells.Name = "chkWholeTableCells";
            this.chkWholeTableCells.UseVisualStyleBackColor = true;
            // 
            // chkBoxMultiSelLastColumn
            // 
            resources.ApplyResources(this.chkBoxMultiSelLastColumn, "chkBoxMultiSelLastColumn");
            this.chkBoxMultiSelLastColumn.Name = "chkBoxMultiSelLastColumn";
            this.chkBoxMultiSelLastColumn.UseVisualStyleBackColor = true;
            // 
            // chkBoxMulSelTblLastRow
            // 
            resources.ApplyResources(this.chkBoxMulSelTblLastRow, "chkBoxMulSelTblLastRow");
            this.chkBoxMulSelTblLastRow.Name = "chkBoxMulSelTblLastRow";
            this.chkBoxMulSelTblLastRow.UseVisualStyleBackColor = true;
            // 
            // numMultiSelColEnd
            // 
            resources.ApplyResources(this.numMultiSelColEnd, "numMultiSelColEnd");
            this.numMultiSelColEnd.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numMultiSelColEnd.Name = "numMultiSelColEnd";
            this.numMultiSelColEnd.ReadOnly = true;
            // 
            // numMultiSelColStart
            // 
            resources.ApplyResources(this.numMultiSelColStart, "numMultiSelColStart");
            this.numMultiSelColStart.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numMultiSelColStart.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numMultiSelColStart.Name = "numMultiSelColStart";
            this.numMultiSelColStart.ReadOnly = true;
            this.numMultiSelColStart.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // numMultiSelRowEnd
            // 
            resources.ApplyResources(this.numMultiSelRowEnd, "numMultiSelRowEnd");
            this.numMultiSelRowEnd.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numMultiSelRowEnd.Name = "numMultiSelRowEnd";
            this.numMultiSelRowEnd.ReadOnly = true;
            // 
            // numMultiSelRowStart
            // 
            resources.ApplyResources(this.numMultiSelRowStart, "numMultiSelRowStart");
            this.numMultiSelRowStart.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numMultiSelRowStart.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numMultiSelRowStart.Name = "numMultiSelRowStart";
            this.numMultiSelRowStart.ReadOnly = true;
            this.numMultiSelRowStart.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label46
            // 
            resources.ApplyResources(this.label46, "label46");
            this.label46.Name = "label46";
            // 
            // label45
            // 
            resources.ApplyResources(this.label45, "label45");
            this.label45.Name = "label45";
            // 
            // chkBoxMultiSelColumnsScope
            // 
            resources.ApplyResources(this.chkBoxMultiSelColumnsScope, "chkBoxMultiSelColumnsScope");
            this.chkBoxMultiSelColumnsScope.Name = "chkBoxMultiSelColumnsScope";
            this.chkBoxMultiSelColumnsScope.UseVisualStyleBackColor = true;
            this.chkBoxMultiSelColumnsScope.CheckedChanged += new System.EventHandler(this.chkBoxMultiSelColumnsScope_CheckedChanged);
            // 
            // chkBoxMultiSelRowsScope
            // 
            resources.ApplyResources(this.chkBoxMultiSelRowsScope, "chkBoxMultiSelRowsScope");
            this.chkBoxMultiSelRowsScope.Name = "chkBoxMultiSelRowsScope";
            this.chkBoxMultiSelRowsScope.UseVisualStyleBackColor = true;
            this.chkBoxMultiSelRowsScope.CheckedChanged += new System.EventHandler(this.chkBoxMultiSelRowsScope_CheckedChanged);
            // 
            // chkBoxMultiSelFirstColumn
            // 
            resources.ApplyResources(this.chkBoxMultiSelFirstColumn, "chkBoxMultiSelFirstColumn");
            this.chkBoxMultiSelFirstColumn.Name = "chkBoxMultiSelFirstColumn";
            this.chkBoxMultiSelFirstColumn.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelTables
            // 
            resources.ApplyResources(this.checkBoxMultiSelTables, "checkBoxMultiSelTables");
            this.checkBoxMultiSelTables.Name = "checkBoxMultiSelTables";
            this.checkBoxMultiSelTables.UseVisualStyleBackColor = true;
            this.checkBoxMultiSelTables.CheckedChanged += new System.EventHandler(this.checkBoxMultiSelTables_CheckedChanged);
            // 
            // chkBoxMulSelTblFirstRow
            // 
            resources.ApplyResources(this.chkBoxMulSelTblFirstRow, "chkBoxMulSelTblFirstRow");
            this.chkBoxMulSelTblFirstRow.Name = "chkBoxMulSelTblFirstRow";
            this.chkBoxMulSelTblFirstRow.UseVisualStyleBackColor = true;
            // 
            // btnMultiSelApplySel
            // 
            resources.ApplyResources(this.btnMultiSelApplySel, "btnMultiSelApplySel");
            this.btnMultiSelApplySel.Name = "btnMultiSelApplySel";
            this.btnMultiSelApplySel.UseVisualStyleBackColor = true;
            this.btnMultiSelApplySel.Click += new System.EventHandler(this.btnMultiSelApplySel_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.rdBtnAfterCurSel);
            this.groupBox5.Controls.Add(this.rdBtnBeforeCurSel);
            this.groupBox5.Controls.Add(this.radioBtnMultiSelCurSelScope);
            this.groupBox5.Controls.Add(this.radioBtnMultiSelWholeStory);
            resources.ApplyResources(this.groupBox5, "groupBox5");
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.TabStop = false;
            // 
            // rdBtnAfterCurSel
            // 
            resources.ApplyResources(this.rdBtnAfterCurSel, "rdBtnAfterCurSel");
            this.rdBtnAfterCurSel.Name = "rdBtnAfterCurSel";
            this.rdBtnAfterCurSel.UseVisualStyleBackColor = true;
            // 
            // rdBtnBeforeCurSel
            // 
            resources.ApplyResources(this.rdBtnBeforeCurSel, "rdBtnBeforeCurSel");
            this.rdBtnBeforeCurSel.Name = "rdBtnBeforeCurSel";
            this.rdBtnBeforeCurSel.UseVisualStyleBackColor = true;
            // 
            // radioBtnMultiSelCurSelScope
            // 
            resources.ApplyResources(this.radioBtnMultiSelCurSelScope, "radioBtnMultiSelCurSelScope");
            this.radioBtnMultiSelCurSelScope.Name = "radioBtnMultiSelCurSelScope";
            this.radioBtnMultiSelCurSelScope.UseVisualStyleBackColor = true;
            // 
            // radioBtnMultiSelWholeStory
            // 
            resources.ApplyResources(this.radioBtnMultiSelWholeStory, "radioBtnMultiSelWholeStory");
            this.radioBtnMultiSelWholeStory.Checked = true;
            this.radioBtnMultiSelWholeStory.Name = "radioBtnMultiSelWholeStory";
            this.radioBtnMultiSelWholeStory.TabStop = true;
            this.radioBtnMultiSelWholeStory.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBoxMultiSelHighlight);
            this.groupBox3.Controls.Add(this.colorComboBoxHighlight);
            this.groupBox3.Controls.Add(this.rdBtnMultiSelObjectPara);
            this.groupBox3.Controls.Add(this.rdBtnMultiSelObjectRng);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelIndices);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelInlineShapes);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelFields);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelBookMarks);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelCnts);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelComments);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelEndNotes);
            this.groupBox3.Controls.Add(this.checkBoxMultiSelFootNotes);
            this.groupBox3.Controls.Add(this.label47);
            this.groupBox3.Controls.Add(this.checkBoxMultiHyperLinks);
            resources.ApplyResources(this.groupBox3, "groupBox3");
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.TabStop = false;
            // 
            // checkBoxMultiSelHighlight
            // 
            resources.ApplyResources(this.checkBoxMultiSelHighlight, "checkBoxMultiSelHighlight");
            this.checkBoxMultiSelHighlight.Name = "checkBoxMultiSelHighlight";
            this.checkBoxMultiSelHighlight.UseVisualStyleBackColor = true;
            // 
            // colorComboBoxHighlight
            // 
            this.colorComboBoxHighlight.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.colorComboBoxHighlight.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.colorComboBoxHighlight.FormattingEnabled = true;
            resources.ApplyResources(this.colorComboBoxHighlight, "colorComboBoxHighlight");
            this.colorComboBoxHighlight.Items.AddRange(new object[] {
            resources.GetString("colorComboBoxHighlight.Items"),
            resources.GetString("colorComboBoxHighlight.Items1")});
            this.colorComboBoxHighlight.Name = "colorComboBoxHighlight";
            // 
            // rdBtnMultiSelObjectPara
            // 
            resources.ApplyResources(this.rdBtnMultiSelObjectPara, "rdBtnMultiSelObjectPara");
            this.rdBtnMultiSelObjectPara.Checked = true;
            this.rdBtnMultiSelObjectPara.Name = "rdBtnMultiSelObjectPara";
            this.rdBtnMultiSelObjectPara.TabStop = true;
            this.rdBtnMultiSelObjectPara.UseVisualStyleBackColor = true;
            // 
            // rdBtnMultiSelObjectRng
            // 
            resources.ApplyResources(this.rdBtnMultiSelObjectRng, "rdBtnMultiSelObjectRng");
            this.rdBtnMultiSelObjectRng.Name = "rdBtnMultiSelObjectRng";
            this.rdBtnMultiSelObjectRng.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelIndices
            // 
            resources.ApplyResources(this.checkBoxMultiSelIndices, "checkBoxMultiSelIndices");
            this.checkBoxMultiSelIndices.Name = "checkBoxMultiSelIndices";
            this.checkBoxMultiSelIndices.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelInlineShapes
            // 
            resources.ApplyResources(this.checkBoxMultiSelInlineShapes, "checkBoxMultiSelInlineShapes");
            this.checkBoxMultiSelInlineShapes.Name = "checkBoxMultiSelInlineShapes";
            this.checkBoxMultiSelInlineShapes.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelFields
            // 
            resources.ApplyResources(this.checkBoxMultiSelFields, "checkBoxMultiSelFields");
            this.checkBoxMultiSelFields.Name = "checkBoxMultiSelFields";
            this.checkBoxMultiSelFields.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelBookMarks
            // 
            resources.ApplyResources(this.checkBoxMultiSelBookMarks, "checkBoxMultiSelBookMarks");
            this.checkBoxMultiSelBookMarks.Name = "checkBoxMultiSelBookMarks";
            this.checkBoxMultiSelBookMarks.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelCnts
            // 
            resources.ApplyResources(this.checkBoxMultiSelCnts, "checkBoxMultiSelCnts");
            this.checkBoxMultiSelCnts.Name = "checkBoxMultiSelCnts";
            this.checkBoxMultiSelCnts.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelComments
            // 
            resources.ApplyResources(this.checkBoxMultiSelComments, "checkBoxMultiSelComments");
            this.checkBoxMultiSelComments.Name = "checkBoxMultiSelComments";
            this.checkBoxMultiSelComments.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelEndNotes
            // 
            resources.ApplyResources(this.checkBoxMultiSelEndNotes, "checkBoxMultiSelEndNotes");
            this.checkBoxMultiSelEndNotes.Name = "checkBoxMultiSelEndNotes";
            this.checkBoxMultiSelEndNotes.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelFootNotes
            // 
            resources.ApplyResources(this.checkBoxMultiSelFootNotes, "checkBoxMultiSelFootNotes");
            this.checkBoxMultiSelFootNotes.Name = "checkBoxMultiSelFootNotes";
            this.checkBoxMultiSelFootNotes.UseVisualStyleBackColor = true;
            // 
            // label47
            // 
            resources.ApplyResources(this.label47, "label47");
            this.label47.Name = "label47";
            // 
            // checkBoxMultiHyperLinks
            // 
            resources.ApplyResources(this.checkBoxMultiHyperLinks, "checkBoxMultiHyperLinks");
            this.checkBoxMultiHyperLinks.Name = "checkBoxMultiHyperLinks";
            this.checkBoxMultiHyperLinks.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.btnMultiSelHeadingAllClear);
            this.groupBox7.Controls.Add(this.btnMultiSelHeadingAllSel);
            this.groupBox7.Controls.Add(this.checkedListBoxMultiSelHeading);
            this.groupBox7.Controls.Add(this.groupBox2);
            resources.ApplyResources(this.groupBox7, "groupBox7");
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.TabStop = false;
            // 
            // btnMultiSelHeadingAllClear
            // 
            resources.ApplyResources(this.btnMultiSelHeadingAllClear, "btnMultiSelHeadingAllClear");
            this.btnMultiSelHeadingAllClear.Name = "btnMultiSelHeadingAllClear";
            this.btnMultiSelHeadingAllClear.UseVisualStyleBackColor = true;
            this.btnMultiSelHeadingAllClear.Click += new System.EventHandler(this.btnMultiSelHeadingAllClear_Click);
            // 
            // btnMultiSelHeadingAllSel
            // 
            resources.ApplyResources(this.btnMultiSelHeadingAllSel, "btnMultiSelHeadingAllSel");
            this.btnMultiSelHeadingAllSel.Name = "btnMultiSelHeadingAllSel";
            this.btnMultiSelHeadingAllSel.UseVisualStyleBackColor = true;
            this.btnMultiSelHeadingAllSel.Click += new System.EventHandler(this.btnMultiSelHeadingAllSel_Click);
            // 
            // checkedListBoxMultiSelHeading
            // 
            this.checkedListBoxMultiSelHeading.CheckOnClick = true;
            resources.ApplyResources(this.checkedListBoxMultiSelHeading, "checkedListBoxMultiSelHeading");
            this.checkedListBoxMultiSelHeading.FormattingEnabled = true;
            this.checkedListBoxMultiSelHeading.Items.AddRange(new object[] {
            resources.GetString("checkedListBoxMultiSelHeading.Items"),
            resources.GetString("checkedListBoxMultiSelHeading.Items1"),
            resources.GetString("checkedListBoxMultiSelHeading.Items2"),
            resources.GetString("checkedListBoxMultiSelHeading.Items3"),
            resources.GetString("checkedListBoxMultiSelHeading.Items4"),
            resources.GetString("checkedListBoxMultiSelHeading.Items5"),
            resources.GetString("checkedListBoxMultiSelHeading.Items6"),
            resources.GetString("checkedListBoxMultiSelHeading.Items7"),
            resources.GetString("checkedListBoxMultiSelHeading.Items8"),
            resources.GetString("checkedListBoxMultiSelHeading.Items9")});
            this.checkedListBoxMultiSelHeading.MultiColumn = true;
            this.checkedListBoxMultiSelHeading.Name = "checkedListBoxMultiSelHeading";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkBoxMultiSelIgnoreToc);
            this.groupBox2.Controls.Add(this.checkBoxMultiSelIgnoreTbl);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // checkBoxMultiSelIgnoreToc
            // 
            resources.ApplyResources(this.checkBoxMultiSelIgnoreToc, "checkBoxMultiSelIgnoreToc");
            this.checkBoxMultiSelIgnoreToc.Checked = true;
            this.checkBoxMultiSelIgnoreToc.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxMultiSelIgnoreToc.Name = "checkBoxMultiSelIgnoreToc";
            this.checkBoxMultiSelIgnoreToc.UseVisualStyleBackColor = true;
            // 
            // checkBoxMultiSelIgnoreTbl
            // 
            resources.ApplyResources(this.checkBoxMultiSelIgnoreTbl, "checkBoxMultiSelIgnoreTbl");
            this.checkBoxMultiSelIgnoreTbl.Checked = true;
            this.checkBoxMultiSelIgnoreTbl.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxMultiSelIgnoreTbl.Name = "checkBoxMultiSelIgnoreTbl";
            this.checkBoxMultiSelIgnoreTbl.UseVisualStyleBackColor = true;
            // 
            // tabPageMultiTiZhu
            // 
            this.tabPageMultiTiZhu.Controls.Add(this.groupBox12);
            this.tabPageMultiTiZhu.Controls.Add(this.groupBox11);
            resources.ApplyResources(this.tabPageMultiTiZhu, "tabPageMultiTiZhu");
            this.tabPageMultiTiZhu.Name = "tabPageMultiTiZhu";
            this.tabPageMultiTiZhu.UseVisualStyleBackColor = true;
            this.tabPageMultiTiZhu.Enter += new System.EventHandler(this.tabPageMultiTiZhu_Enter);
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.label62);
            this.groupBox12.Controls.Add(this.label56);
            this.groupBox12.Controls.Add(this.label52);
            this.groupBox12.Controls.Add(this.btnNavLastField);
            this.groupBox12.Controls.Add(this.btnNavNextField);
            this.groupBox12.Controls.Add(this.btnNav2LastInShp);
            this.groupBox12.Controls.Add(this.btnNav2NextInShp);
            this.groupBox12.Controls.Add(this.btnNavPrevField);
            this.groupBox12.Controls.Add(this.btnNav2LastTbl);
            this.groupBox12.Controls.Add(this.btnNav2PrevInShp);
            this.groupBox12.Controls.Add(this.btnNavFirstField);
            this.groupBox12.Controls.Add(this.btnNav2NextTbl);
            this.groupBox12.Controls.Add(this.btnNav2FirstInShp);
            this.groupBox12.Controls.Add(this.btnNav2PrevTbl);
            this.groupBox12.Controls.Add(this.btnNav2FirstTbl);
            resources.ApplyResources(this.groupBox12, "groupBox12");
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.TabStop = false;
            // 
            // label62
            // 
            resources.ApplyResources(this.label62, "label62");
            this.label62.Name = "label62";
            // 
            // label56
            // 
            resources.ApplyResources(this.label56, "label56");
            this.label56.Name = "label56";
            // 
            // label52
            // 
            resources.ApplyResources(this.label52, "label52");
            this.label52.Name = "label52";
            // 
            // btnNavLastField
            // 
            resources.ApplyResources(this.btnNavLastField, "btnNavLastField");
            this.btnNavLastField.Name = "btnNavLastField";
            this.btnNavLastField.UseVisualStyleBackColor = true;
            this.btnNavLastField.Click += new System.EventHandler(this.btnNavLastField_Click);
            // 
            // btnNavNextField
            // 
            resources.ApplyResources(this.btnNavNextField, "btnNavNextField");
            this.btnNavNextField.Name = "btnNavNextField";
            this.btnNavNextField.UseVisualStyleBackColor = true;
            this.btnNavNextField.Click += new System.EventHandler(this.btnNavNextField_Click);
            // 
            // btnNav2LastInShp
            // 
            resources.ApplyResources(this.btnNav2LastInShp, "btnNav2LastInShp");
            this.btnNav2LastInShp.Name = "btnNav2LastInShp";
            this.btnNav2LastInShp.UseVisualStyleBackColor = true;
            this.btnNav2LastInShp.Click += new System.EventHandler(this.btnNav2LastInShp_Click);
            // 
            // btnNav2NextInShp
            // 
            resources.ApplyResources(this.btnNav2NextInShp, "btnNav2NextInShp");
            this.btnNav2NextInShp.Name = "btnNav2NextInShp";
            this.btnNav2NextInShp.UseVisualStyleBackColor = true;
            this.btnNav2NextInShp.Click += new System.EventHandler(this.btnNav2NextInShp_Click);
            // 
            // btnNavPrevField
            // 
            resources.ApplyResources(this.btnNavPrevField, "btnNavPrevField");
            this.btnNavPrevField.Name = "btnNavPrevField";
            this.btnNavPrevField.UseVisualStyleBackColor = true;
            this.btnNavPrevField.Click += new System.EventHandler(this.btnNavPrevField_Click);
            // 
            // btnNav2LastTbl
            // 
            resources.ApplyResources(this.btnNav2LastTbl, "btnNav2LastTbl");
            this.btnNav2LastTbl.Name = "btnNav2LastTbl";
            this.btnNav2LastTbl.UseVisualStyleBackColor = true;
            this.btnNav2LastTbl.Click += new System.EventHandler(this.btnNav2LastTbl_Click);
            // 
            // btnNav2PrevInShp
            // 
            resources.ApplyResources(this.btnNav2PrevInShp, "btnNav2PrevInShp");
            this.btnNav2PrevInShp.Name = "btnNav2PrevInShp";
            this.btnNav2PrevInShp.UseVisualStyleBackColor = true;
            this.btnNav2PrevInShp.Click += new System.EventHandler(this.btnNav2PrevInShp_Click);
            // 
            // btnNavFirstField
            // 
            resources.ApplyResources(this.btnNavFirstField, "btnNavFirstField");
            this.btnNavFirstField.Name = "btnNavFirstField";
            this.btnNavFirstField.UseVisualStyleBackColor = true;
            this.btnNavFirstField.Click += new System.EventHandler(this.btnNavFirstField_Click);
            // 
            // btnNav2NextTbl
            // 
            resources.ApplyResources(this.btnNav2NextTbl, "btnNav2NextTbl");
            this.btnNav2NextTbl.Name = "btnNav2NextTbl";
            this.btnNav2NextTbl.UseVisualStyleBackColor = true;
            this.btnNav2NextTbl.Click += new System.EventHandler(this.btnNav2NextTbl_Click);
            // 
            // btnNav2FirstInShp
            // 
            resources.ApplyResources(this.btnNav2FirstInShp, "btnNav2FirstInShp");
            this.btnNav2FirstInShp.Name = "btnNav2FirstInShp";
            this.btnNav2FirstInShp.UseVisualStyleBackColor = true;
            this.btnNav2FirstInShp.Click += new System.EventHandler(this.btnNav2FirstInShp_Click);
            // 
            // btnNav2PrevTbl
            // 
            resources.ApplyResources(this.btnNav2PrevTbl, "btnNav2PrevTbl");
            this.btnNav2PrevTbl.Name = "btnNav2PrevTbl";
            this.btnNav2PrevTbl.UseVisualStyleBackColor = true;
            this.btnNav2PrevTbl.Click += new System.EventHandler(this.btnNav2PrevTbl_Click);
            // 
            // btnNav2FirstTbl
            // 
            resources.ApplyResources(this.btnNav2FirstTbl, "btnNav2FirstTbl");
            this.btnNav2FirstTbl.Name = "btnNav2FirstTbl";
            this.btnNav2FirstTbl.UseVisualStyleBackColor = true;
            this.btnNav2FirstTbl.Click += new System.EventHandler(this.btnNav2FirstTbl_Click);
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.chkInShpCaplblGetFromHeading);
            this.groupBox11.Controls.Add(this.chkTblCaplblGetFromHeading);
            this.groupBox11.Controls.Add(this.chkSyncUpdateTableOfFigures);
            this.groupBox11.Controls.Add(this.txtInShpCapLblPreFix);
            this.groupBox11.Controls.Add(this.txtInShpCapLblPostFix);
            this.groupBox11.Controls.Add(this.txtTblCapLblPreFix);
            this.groupBox11.Controls.Add(this.txtTblCapLblPostFix);
            this.groupBox11.Controls.Add(this.rdBtnTiZhuAfterCurPos);
            this.groupBox11.Controls.Add(this.rdBtnTiZhuBeforeCurPos);
            this.groupBox11.Controls.Add(this.rdCapLblScopeSelection);
            this.groupBox11.Controls.Add(this.rdCapLblScopeAllDoc);
            this.groupBox11.Controls.Add(this.cmbInShpCapLblAlign);
            this.groupBox11.Controls.Add(this.cmbInShpCapLblPos);
            this.groupBox11.Controls.Add(this.cmbTblCapLblPos);
            this.groupBox11.Controls.Add(this.cmbTblCapLblAlign);
            this.groupBox11.Controls.Add(this.btnApplyCapLbls2CurDoc);
            this.groupBox11.Controls.Add(this.label51);
            this.groupBox11.Controls.Add(this.label55);
            this.groupBox11.Controls.Add(this.label53);
            this.groupBox11.Controls.Add(this.label54);
            this.groupBox11.Controls.Add(this.btnAddSelInShpCapLbl);
            this.groupBox11.Controls.Add(this.txtSelectedInShpCapLbl);
            this.groupBox11.Controls.Add(this.label50);
            this.groupBox11.Controls.Add(this.lstBoxCurSysCapLbls);
            this.groupBox11.Controls.Add(this.btnRefreshCapsLbl);
            this.groupBox11.Controls.Add(this.btnRemoveSelInShpCapLbl);
            this.groupBox11.Controls.Add(this.btnSetSysCapLbls);
            this.groupBox11.Controls.Add(this.txtSelectedTblCapLbl);
            this.groupBox11.Controls.Add(this.btnRemoveSelTblCapLbl);
            this.groupBox11.Controls.Add(this.label49);
            this.groupBox11.Controls.Add(this.label82);
            this.groupBox11.Controls.Add(this.label81);
            this.groupBox11.Controls.Add(this.label61);
            this.groupBox11.Controls.Add(this.btnAddSelTblCapLbl);
            this.groupBox11.Controls.Add(this.label58);
            this.groupBox11.Controls.Add(this.label59);
            this.groupBox11.Controls.Add(this.label57);
            this.groupBox11.Controls.Add(this.label60);
            this.groupBox11.Controls.Add(this.label48);
            resources.ApplyResources(this.groupBox11, "groupBox11");
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.TabStop = false;
            // 
            // chkInShpCaplblGetFromHeading
            // 
            resources.ApplyResources(this.chkInShpCaplblGetFromHeading, "chkInShpCaplblGetFromHeading");
            this.chkInShpCaplblGetFromHeading.Name = "chkInShpCaplblGetFromHeading";
            this.chkInShpCaplblGetFromHeading.UseVisualStyleBackColor = true;
            // 
            // chkTblCaplblGetFromHeading
            // 
            resources.ApplyResources(this.chkTblCaplblGetFromHeading, "chkTblCaplblGetFromHeading");
            this.chkTblCaplblGetFromHeading.Name = "chkTblCaplblGetFromHeading";
            this.chkTblCaplblGetFromHeading.UseVisualStyleBackColor = true;
            // 
            // chkSyncUpdateTableOfFigures
            // 
            resources.ApplyResources(this.chkSyncUpdateTableOfFigures, "chkSyncUpdateTableOfFigures");
            this.chkSyncUpdateTableOfFigures.Name = "chkSyncUpdateTableOfFigures";
            this.chkSyncUpdateTableOfFigures.UseVisualStyleBackColor = true;
            // 
            // txtInShpCapLblPreFix
            // 
            resources.ApplyResources(this.txtInShpCapLblPreFix, "txtInShpCapLblPreFix");
            this.txtInShpCapLblPreFix.Name = "txtInShpCapLblPreFix";
            // 
            // txtInShpCapLblPostFix
            // 
            resources.ApplyResources(this.txtInShpCapLblPostFix, "txtInShpCapLblPostFix");
            this.txtInShpCapLblPostFix.Name = "txtInShpCapLblPostFix";
            // 
            // txtTblCapLblPreFix
            // 
            resources.ApplyResources(this.txtTblCapLblPreFix, "txtTblCapLblPreFix");
            this.txtTblCapLblPreFix.Name = "txtTblCapLblPreFix";
            // 
            // txtTblCapLblPostFix
            // 
            resources.ApplyResources(this.txtTblCapLblPostFix, "txtTblCapLblPostFix");
            this.txtTblCapLblPostFix.Name = "txtTblCapLblPostFix";
            // 
            // rdBtnTiZhuAfterCurPos
            // 
            resources.ApplyResources(this.rdBtnTiZhuAfterCurPos, "rdBtnTiZhuAfterCurPos");
            this.rdBtnTiZhuAfterCurPos.Name = "rdBtnTiZhuAfterCurPos";
            this.rdBtnTiZhuAfterCurPos.UseVisualStyleBackColor = true;
            // 
            // rdBtnTiZhuBeforeCurPos
            // 
            resources.ApplyResources(this.rdBtnTiZhuBeforeCurPos, "rdBtnTiZhuBeforeCurPos");
            this.rdBtnTiZhuBeforeCurPos.Name = "rdBtnTiZhuBeforeCurPos";
            this.rdBtnTiZhuBeforeCurPos.UseVisualStyleBackColor = true;
            // 
            // rdCapLblScopeSelection
            // 
            resources.ApplyResources(this.rdCapLblScopeSelection, "rdCapLblScopeSelection");
            this.rdCapLblScopeSelection.Name = "rdCapLblScopeSelection";
            this.rdCapLblScopeSelection.UseVisualStyleBackColor = true;
            // 
            // rdCapLblScopeAllDoc
            // 
            resources.ApplyResources(this.rdCapLblScopeAllDoc, "rdCapLblScopeAllDoc");
            this.rdCapLblScopeAllDoc.Checked = true;
            this.rdCapLblScopeAllDoc.Name = "rdCapLblScopeAllDoc";
            this.rdCapLblScopeAllDoc.TabStop = true;
            this.rdCapLblScopeAllDoc.UseVisualStyleBackColor = true;
            // 
            // cmbInShpCapLblAlign
            // 
            this.cmbInShpCapLblAlign.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbInShpCapLblAlign.FormattingEnabled = true;
            this.cmbInShpCapLblAlign.Items.AddRange(new object[] {
            resources.GetString("cmbInShpCapLblAlign.Items"),
            resources.GetString("cmbInShpCapLblAlign.Items1"),
            resources.GetString("cmbInShpCapLblAlign.Items2")});
            resources.ApplyResources(this.cmbInShpCapLblAlign, "cmbInShpCapLblAlign");
            this.cmbInShpCapLblAlign.Name = "cmbInShpCapLblAlign";
            // 
            // cmbInShpCapLblPos
            // 
            this.cmbInShpCapLblPos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbInShpCapLblPos.FormattingEnabled = true;
            this.cmbInShpCapLblPos.Items.AddRange(new object[] {
            resources.GetString("cmbInShpCapLblPos.Items"),
            resources.GetString("cmbInShpCapLblPos.Items1")});
            resources.ApplyResources(this.cmbInShpCapLblPos, "cmbInShpCapLblPos");
            this.cmbInShpCapLblPos.Name = "cmbInShpCapLblPos";
            // 
            // cmbTblCapLblPos
            // 
            this.cmbTblCapLblPos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTblCapLblPos.FormattingEnabled = true;
            this.cmbTblCapLblPos.Items.AddRange(new object[] {
            resources.GetString("cmbTblCapLblPos.Items"),
            resources.GetString("cmbTblCapLblPos.Items1")});
            resources.ApplyResources(this.cmbTblCapLblPos, "cmbTblCapLblPos");
            this.cmbTblCapLblPos.Name = "cmbTblCapLblPos";
            // 
            // cmbTblCapLblAlign
            // 
            this.cmbTblCapLblAlign.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTblCapLblAlign.FormattingEnabled = true;
            this.cmbTblCapLblAlign.Items.AddRange(new object[] {
            resources.GetString("cmbTblCapLblAlign.Items"),
            resources.GetString("cmbTblCapLblAlign.Items1"),
            resources.GetString("cmbTblCapLblAlign.Items2")});
            resources.ApplyResources(this.cmbTblCapLblAlign, "cmbTblCapLblAlign");
            this.cmbTblCapLblAlign.Name = "cmbTblCapLblAlign";
            // 
            // btnApplyCapLbls2CurDoc
            // 
            resources.ApplyResources(this.btnApplyCapLbls2CurDoc, "btnApplyCapLbls2CurDoc");
            this.btnApplyCapLbls2CurDoc.Name = "btnApplyCapLbls2CurDoc";
            this.btnApplyCapLbls2CurDoc.UseVisualStyleBackColor = true;
            this.btnApplyCapLbls2CurDoc.Click += new System.EventHandler(this.btnApplyCapLbls2CurDoc_Click);
            // 
            // label51
            // 
            resources.ApplyResources(this.label51, "label51");
            this.label51.Name = "label51";
            // 
            // label55
            // 
            resources.ApplyResources(this.label55, "label55");
            this.label55.Name = "label55";
            // 
            // label53
            // 
            resources.ApplyResources(this.label53, "label53");
            this.label53.Name = "label53";
            // 
            // label54
            // 
            resources.ApplyResources(this.label54, "label54");
            this.label54.Name = "label54";
            // 
            // btnAddSelInShpCapLbl
            // 
            resources.ApplyResources(this.btnAddSelInShpCapLbl, "btnAddSelInShpCapLbl");
            this.btnAddSelInShpCapLbl.Name = "btnAddSelInShpCapLbl";
            this.btnAddSelInShpCapLbl.UseVisualStyleBackColor = true;
            this.btnAddSelInShpCapLbl.Click += new System.EventHandler(this.btnAddSelInShpCapLbl_Click);
            // 
            // txtSelectedInShpCapLbl
            // 
            resources.ApplyResources(this.txtSelectedInShpCapLbl, "txtSelectedInShpCapLbl");
            this.txtSelectedInShpCapLbl.Name = "txtSelectedInShpCapLbl";
            this.txtSelectedInShpCapLbl.ReadOnly = true;
            // 
            // label50
            // 
            resources.ApplyResources(this.label50, "label50");
            this.label50.Name = "label50";
            // 
            // lstBoxCurSysCapLbls
            // 
            this.lstBoxCurSysCapLbls.FormattingEnabled = true;
            resources.ApplyResources(this.lstBoxCurSysCapLbls, "lstBoxCurSysCapLbls");
            this.lstBoxCurSysCapLbls.Items.AddRange(new object[] {
            resources.GetString("lstBoxCurSysCapLbls.Items"),
            resources.GetString("lstBoxCurSysCapLbls.Items1"),
            resources.GetString("lstBoxCurSysCapLbls.Items2"),
            resources.GetString("lstBoxCurSysCapLbls.Items3"),
            resources.GetString("lstBoxCurSysCapLbls.Items4")});
            this.lstBoxCurSysCapLbls.Name = "lstBoxCurSysCapLbls";
            // 
            // btnRefreshCapsLbl
            // 
            resources.ApplyResources(this.btnRefreshCapsLbl, "btnRefreshCapsLbl");
            this.btnRefreshCapsLbl.Name = "btnRefreshCapsLbl";
            this.btnRefreshCapsLbl.UseVisualStyleBackColor = true;
            this.btnRefreshCapsLbl.Click += new System.EventHandler(this.btnRefreshCapsLbl_Click);
            // 
            // btnRemoveSelInShpCapLbl
            // 
            resources.ApplyResources(this.btnRemoveSelInShpCapLbl, "btnRemoveSelInShpCapLbl");
            this.btnRemoveSelInShpCapLbl.Name = "btnRemoveSelInShpCapLbl";
            this.btnRemoveSelInShpCapLbl.UseVisualStyleBackColor = true;
            this.btnRemoveSelInShpCapLbl.Click += new System.EventHandler(this.btnRemoveSelInShpCapLbl_Click);
            // 
            // btnSetSysCapLbls
            // 
            resources.ApplyResources(this.btnSetSysCapLbls, "btnSetSysCapLbls");
            this.btnSetSysCapLbls.Name = "btnSetSysCapLbls";
            this.btnSetSysCapLbls.UseVisualStyleBackColor = true;
            this.btnSetSysCapLbls.Click += new System.EventHandler(this.btnSetSysCapLbls_Click);
            // 
            // txtSelectedTblCapLbl
            // 
            resources.ApplyResources(this.txtSelectedTblCapLbl, "txtSelectedTblCapLbl");
            this.txtSelectedTblCapLbl.Name = "txtSelectedTblCapLbl";
            this.txtSelectedTblCapLbl.ReadOnly = true;
            // 
            // btnRemoveSelTblCapLbl
            // 
            resources.ApplyResources(this.btnRemoveSelTblCapLbl, "btnRemoveSelTblCapLbl");
            this.btnRemoveSelTblCapLbl.Name = "btnRemoveSelTblCapLbl";
            this.btnRemoveSelTblCapLbl.UseVisualStyleBackColor = true;
            this.btnRemoveSelTblCapLbl.Click += new System.EventHandler(this.btnRemoveSelTblCapLbl_Click);
            // 
            // label49
            // 
            resources.ApplyResources(this.label49, "label49");
            this.label49.Name = "label49";
            // 
            // label82
            // 
            resources.ApplyResources(this.label82, "label82");
            this.label82.Name = "label82";
            // 
            // label81
            // 
            resources.ApplyResources(this.label81, "label81");
            this.label81.Name = "label81";
            // 
            // label61
            // 
            resources.ApplyResources(this.label61, "label61");
            this.label61.Name = "label61";
            // 
            // btnAddSelTblCapLbl
            // 
            resources.ApplyResources(this.btnAddSelTblCapLbl, "btnAddSelTblCapLbl");
            this.btnAddSelTblCapLbl.Name = "btnAddSelTblCapLbl";
            this.btnAddSelTblCapLbl.UseVisualStyleBackColor = true;
            this.btnAddSelTblCapLbl.Click += new System.EventHandler(this.btnAddSelTblCapLbl_Click);
            // 
            // label58
            // 
            resources.ApplyResources(this.label58, "label58");
            this.label58.Name = "label58";
            // 
            // label59
            // 
            resources.ApplyResources(this.label59, "label59");
            this.label59.Name = "label59";
            // 
            // label57
            // 
            resources.ApplyResources(this.label57, "label57");
            this.label57.Name = "label57";
            // 
            // label60
            // 
            resources.ApplyResources(this.label60, "label60");
            this.label60.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label60.Name = "label60";
            // 
            // label48
            // 
            resources.ApplyResources(this.label48, "label48");
            this.label48.ForeColor = System.Drawing.SystemColors.ButtonShadow;
            this.label48.Name = "label48";
            // 
            // tabPageTEST
            // 
            this.tabPageTEST.Controls.Add(this.button12);
            this.tabPageTEST.Controls.Add(this.button11);
            this.tabPageTEST.Controls.Add(this.button10);
            this.tabPageTEST.Controls.Add(this.button9);
            this.tabPageTEST.Controls.Add(this.button17);
            this.tabPageTEST.Controls.Add(this.textBox1);
            this.tabPageTEST.Controls.Add(this.button5);
            this.tabPageTEST.Controls.Add(this.button4);
            this.tabPageTEST.Controls.Add(this.button3);
            this.tabPageTEST.Controls.Add(this.button2);
            this.tabPageTEST.Controls.Add(this.button1);
            this.tabPageTEST.Controls.Add(this.btn4Test);
            resources.ApplyResources(this.tabPageTEST, "tabPageTEST");
            this.tabPageTEST.Name = "tabPageTEST";
            this.tabPageTEST.UseVisualStyleBackColor = true;
            // 
            // button12
            // 
            resources.ApplyResources(this.button12, "button12");
            this.button12.Name = "button12";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click_1);
            // 
            // button11
            // 
            resources.ApplyResources(this.button11, "button11");
            this.button11.Name = "button11";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click_1);
            // 
            // button10
            // 
            resources.ApplyResources(this.button10, "button10");
            this.button10.Name = "button10";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click_1);
            // 
            // button9
            // 
            resources.ApplyResources(this.button9, "button9");
            this.button9.Name = "button9";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click_1);
            // 
            // button17
            // 
            resources.ApplyResources(this.button17, "button17");
            this.button17.Name = "button17";
            this.button17.UseVisualStyleBackColor = true;
            this.button17.Click += new System.EventHandler(this.button17_Click);
            // 
            // textBox1
            // 
            resources.ApplyResources(this.textBox1, "textBox1");
            this.textBox1.Name = "textBox1";
            // 
            // button5
            // 
            resources.ApplyResources(this.button5, "button5");
            this.button5.Name = "button5";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            resources.ApplyResources(this.button4, "button4");
            this.button4.Name = "button4";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            resources.ApplyResources(this.button3, "button3");
            this.button3.Name = "button3";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            resources.ApplyResources(this.button2, "button2");
            this.button2.Name = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn4Test
            // 
            resources.ApplyResources(this.btn4Test, "btn4Test");
            this.btn4Test.Name = "btn4Test";
            this.btn4Test.UseVisualStyleBackColor = true;
            this.btn4Test.Click += new System.EventHandler(this.btn4Test_Click);
            // 
            // tblUniformStyleHistoryDocsTableAdapter
            // 
            this.tblUniformStyleHistoryDocsTableAdapter.ClearBeforeFill = true;
            // 
            // tblUniformStyleHistoryDocsBindingSource1
            // 
            this.tblUniformStyleHistoryDocsBindingSource1.DataMember = "tblUniformStyleHistoryDocs";
            this.tblUniformStyleHistoryDocsBindingSource1.DataSource = this.localdbDataSet;
            // 
            // tblUniformStyleHistoryDocsBindingSource2
            // 
            this.tblUniformStyleHistoryDocsBindingSource2.DataMember = "tblUniformStyleHistoryDocs";
            this.tblUniformStyleHistoryDocsBindingSource2.DataSource = this.localdbDataSet;
            // 
            // OperationPanel
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.tabCtrl);
            this.Name = "OperationPanel";
            this.Resize += new System.EventHandler(this.UserControl1_Resize);
            this.tabCtrl.ResumeLayout(false);
            this.tabPageRel.ResumeLayout(false);
            this.tabPageRel.PerformLayout();
            this.tabPageCheck.ResumeLayout(false);
            this.tabPageCheck.PerformLayout();
            this.tabPageOrganize.ResumeLayout(false);
            this.tabPageOrganize.PerformLayout();
            this.tabPageShare.ResumeLayout(false);
            this.tabPageShare.PerformLayout();
            this.cxtMenuSvr.ResumeLayout(false);
            this.tabPageUnitedStyle.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblUniformStyleHistoryDocsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.localdbDataSet)).EndInit();
            this.tabPageCompare.ResumeLayout(false);
            this.tabPageCompare.PerformLayout();
            this.tabPageDataTrans.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPageDocTbls2Excel.ResumeLayout(false);
            this.grpW2XAutoModelScope.ResumeLayout(false);
            this.grpW2XAutoModelScope.PerformLayout();
            this.tabPageExcel2DocTbl.ResumeLayout(false);
            this.tabPageExcel2DocTbl.PerformLayout();
            this.tabPageFillGather.ResumeLayout(false);
            this.tabPageFillGather.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.tabPageCntList.ResumeLayout(false);
            this.tabPageCntList.PerformLayout();
            this.tabPageForm.ResumeLayout(false);
            this.tabPageForm.PerformLayout();
            this.tabPageInfo.ResumeLayout(false);
            this.tabPageInfo.PerformLayout();
            this.tabPageNumTrans.ResumeLayout(false);
            this.tabPageNumTrans.PerformLayout();
            this.tabPageHeadingSn.ResumeLayout(false);
            this.tabPageHeadingSn.PerformLayout();
            this.cxtMenuHeadingSn.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPageHeadingStyles.ResumeLayout(false);
            this.tabPageHeadingStyles.PerformLayout();
            this.tabPageObjNav.ResumeLayout(false);
            this.groupBox15.ResumeLayout(false);
            this.groupBox15.PerformLayout();
            this.groupBox14.ResumeLayout(false);
            this.groupBox14.PerformLayout();
            this.groupBox13.ResumeLayout(false);
            this.groupBox13.PerformLayout();
            this.tabPageMultiSel.ResumeLayout(false);
            this.tabPageMultiSel.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.groupBox16.ResumeLayout(false);
            this.groupBox16.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelColEnd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelColStart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelRowEnd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMultiSelRowStart)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPageMultiTiZhu.ResumeLayout(false);
            this.groupBox12.ResumeLayout(false);
            this.groupBox12.PerformLayout();
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            this.tabPageTEST.ResumeLayout(false);
            this.tabPageTEST.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tblUniformStyleHistoryDocsBindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tblUniformStyleHistoryDocsBindingSource2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

 
        public TabControl tabCtrl;
        public TabPage tabPageRel;
        public TabPage tabPageCheck;
        private TreeView m_tvRel;
        private TextBox txtOpRules;
        private Label label3;
        private Button btnInsertRel;
        private Button btnRemoveRel;
        private Button btnJump2Rel;
        private Button btnUpdateRel;
        private Button btnAddRel;
        private Label label2;
        private TextBox txtRelContent;
        private TextBox txtRelName;
        private Button btnReset;
        private Button btnRelSearch;
        private TextBox txtRelKeyword;
        private CheckBox chboxOpRulesEnable;
        private Button btnExpEditor;
        private Button btnFoundNext;
        private Button btnFoundBack;
        private Button btnMove;
        private TabPage tabPageOrganize;
        private CheckedListBox chkSelCategory;
        private TreeView m_tvOrganize;
        private Button btnSelAll;
        private Button btnSelClear;
        private Button btnCheck;
        private Button btnCollapseSel;
        private Button btnExpandSelChild;
        private Button btnOrganizeRefresh;
        private TabPage tabPageShare;
        private Button btnShareSearch;
        private TextBox txtShareKeyWord;
        private Button btnOrganSearch;
        private TextBox txtOrganKeyWord;
        private Button btnOrganResetSearch;
        private Button btnOrganNext;
        private Button btnOrganBack;
        private Button btnShareSearchReset;
        private Button btnShareRef;
        private Label label5;
        private TextBox txtShareName;
        private Button btnShareRemove;
        private Button btnShareAdd;
        private TabPage tabPageForm;
        private Button btnFormRefresh;
        private TabPage tabPageInfo;
        private Button btnInfoRefresh;
        private TextBox txtInfoBody;
        private Button btnRelAllTxtOut;
        private Button btnShareNextSearch;
        private Button btnSharePrevSearch;
        private Button btnOrganProtect;
        private TreeView tvCheckedItems;
        private Button btnCheckIgnore;
        private Button btnCheckReset;
        private Button btnCheckSearch;
        private TextBox txtCheckSearchKeyWord;
        private Button btnCheckSearchNext;
        private Button btnCheckSearchPrev;
        private Label label6;
        private CheckBox chkBoxCommonLib;
        private CheckBox chkBoxCategory;
        private Button btnOrgCancelProtect;
        private ImageList imageListIcon;
        private Button btnShareExternalFile;
        private Button btnShareRefresh;
        private Button btnShareCollapse;
        private Button btnShareExpand;
        private Button btnOrgPromote;
        private Button btnOrgDemote;
        private CheckBox chkOrgShowBody;
        private ProgressBar OrgProgressBar;
        private Button btnShareDownload;
        private Button btnShareOpen;
        private TabPage tabPageCompare;
        private Label label10;
        private Label label9;
        private Button btnCompCheckDoc;
        private TextBox txtComp2CheckDoc;
        private Button btnCompStandardDoc;
        private TextBox txtCompStandardDoc;
        private TextBox txtCompResult;
        private TreeView tvCompCheck;
        private TreeView tvCompStandard;
        private CheckBox chkCompStrickOrder;
        private CheckBox chkCompOutline;
        private Button btnExecCompare;
        private ProgressBar progBarComp;
        private ProgressBar progbarCheck;
        private TabPage tabPageUnitedStyle;
        private Button btnStyleOpenFile;
        private TextBox txtBoxStyleFile;
        private Label label11;
        private Button btnStyleApply;
        private ProgressBar progressBarStyle;
        private ProgressBar prgBarLib;
        private GroupBox groupBox6;
        private RadioButton radioBtnStyleSelection;
        private RadioButton radioBtnStyleAllDoc;
        private GroupBox groupBox4;
        private Button btnFormNextSearch;
        private TableLayoutPanel tblFormLayoutPanel;
        private Button btnFormPrevSearch;
        private TextBox txtFormKeyWord;
        private Button btnFormReset;
        private Button btnFormSearch;
        private Label label14;
        private TabPage tabPageCntList;
        private TreeView trvCntList;
        private Button btnCntListCover;
        private Button btnCntListRef;
        private Button btnCntListRemove;
        private Button btnCntListAddDoc;
        private TextBox txtBoxCntListFile;
        private ProgressBar progBarCntList;
        private Button btnCntListExpand;
        private Button btnCntListCollapse;
        private ContextMenuStrip cxtMenuSvr;
        private ToolStripMenuItem menuItemApplyStyle;
        private ToolStripMenuItem menuItemCntReuse;
        private Button btnRefreshRels;
        private Label label1;
        private Label label4;
        private Button btnRelForceCompute;
        private TabPage tabPageTEST;
        private Button btn4Test;
        private Button button1;
        private CheckBox chkIgnoreTable;
        private CheckBox chkIgnoreTOC;
        private Label label7;
        private TextBox txtIgnorePages;
        private CheckBox chkIgnorePages;
        private Button button3;
        private Button button2;
        private TabPage tabPageNumTrans;
        private Label label8;
        private TextBox txtDigitNumSimpLittle;
        private TextBox txtDigitNum;
        private TextBox txtDigitNumSimpBig;
        private Label label17;
        private Label label16;
        private Label label15;
        private Label label18;
        private TextBox txtNumValueSimpBig;
        private Label label21;
        private Label label20;
        private Label label19;
        private TextBox txtNumValueSimpLittle;
        private TextBox txtNumValue;
        private TextBox txtMoneySimpBig;
        private Label label25;
        private Label label24;
        private Label label23;
        private Label label22;
        private TextBox txtMoneySimpLittle;
        private TextBox txtNumMoney;
        private TextBox txtMoneySimpBigTbl;
        private Label label27;
        private Label label26;
        private TextBox txtMoneySimpLittleTbl;
        private Button btnNumTrans;
        private TextBox txtNumValueSimpBigTbl;
        private Label label29;
        private Label label28;
        private TextBox txtNumValueSimpLittleTbl;
        private Button btnNumTransClear;
        private CheckBox chkIgnoreTextBody;
        private CheckBox chkIgnoreHeadings;
        private CheckBox chkIgnoreParaFormat;
        private CheckBox chkIgnoreFont;
        private Button button4;
        private TabPage tabPageHeadingSn;
        private Button button5;
        private ListBox lstHeadingSnLevel;
        private Button btnHeadingSnFont;
        private CheckBox chkHeadingSnLegal;
        private ComboBox cmbSnShowStyle;
        private TextBox txtSnDefInput;
        private Label label31;
        private Label label32;
        private Label label30;
        private TextBox txtHeadingSnSchemeName;
        private GroupBox groupBox1;
        private Label label33;
        private TreeView trvHeadingSnScheme;
        private Button btnHeadingSnSchemeApply;
        private Button btnHeadingSnSchemeGet;
        private Button btnHeadingSnSchemeUpdate;
        private Button btnHeadingSnSchemeDel;
        private Button btnHeadingSnSchemeAdd;
        private Button btnHeadingSnPos;
        private ListBox lstUnitedStyleHistoryDoc;
        private BindingSource tblUniformStyleHistoryDocsBindingSource;
        private localdbDataSet localdbDataSet;
        private localdbDataSetTableAdapters.tblUniformStyleHistoryDocsTableAdapter tblUniformStyleHistoryDocsTableAdapter;
        private Button btnHeadingSnNameGen;
        private ProgressBar progBarHeadingSn;
        private CheckBox chkHeadingSnReserveCurStyle;
        private TabPage tabPageHeadingStyles;
        private Button btnHeadingSnReset;
        private ContextMenuStrip cxtMenuHeadingSn;
        private ToolStripMenuItem cxtMenuItemPreview;
        private Button btnHeadingSnPreview;
        private Button btnHeadingStyleSchemePreview;
        private Button btnHeadingStyleSchemeUpdate;
        private Button btnHeadingStyleSchemeDel;
        private Button btnHeadingStyleSchemeAdd;
        private TreeView trvHeadingStyleScheme;
        private Button btnHeadingStyleSchemeApply;
        private Button btnHeadingStyleSchemeExtract;
        private TextBox txtHeadingStyleSchemeName;
        private ProgressBar prgbarHeadingStyleSchemeApply;
        private ListBox lstOutlineLevel;
        private RichTextBox richHeadingStylePreview;
        private Button btnHeadingStyleApplyScope;
        private Button btnHeadingStyleExitApply;
        private Button btnUnitFormExitApply;
        private Button btnExitHeadingSnApply;
        private Button btnHeadingStyleApplyCurSel;
        private RichTextBox richTxtHeadingSnPreview;
        private Button btnHeadingSnSetDefaultFont;
        private Button btnHeadingSnFontExtract;
        private TextBox textBox1;
        private Label label34;
        private Label label35;
        private Label label36;
        private Label label37;
        private Label label38;
        private Label label39;
        private Label label40;
        private Label label41;
        private BindingSource tblUniformStyleHistoryDocsBindingSource1;
        private BindingSource tblUniformStyleHistoryDocsBindingSource2;
        public TreeView tvShareLib;
        private Button btnShareLibUpdate;
        private TabPage tabPageMultiSel;
        private TabPage tabPageDataTrans;
        private TabPage tabPageDocTbls2Excel;
        private Button btnAddColName;
        private TabPage tabPageExcel2DocTbl;
        private Button btnAddColValue;
        private Button btnClearItems;
        private Button btnAllProduce;
        private Button btnPreviewProduce;
        private Button btnDocTbl2ExcelRemove;
        private TreeView trvDataDocTbl2Excel;
        private TreeView trvData;
        private Label label12;
        private Button btnDataInsertData;
        private Button btnDataInsertName;
        private Button btnDataDSource;
        private Button btnDataProduce;
        private Button btnDataPreviewOne;
        private Button btnDataTagCombine;
        private Button btnDataRemoveRel;
        private CheckedListBox checkedListBoxMultiSelHeading;
        private GroupBox groupBox2;
        private CheckBox checkBoxMultiSelIgnoreToc;
        private CheckBox checkBoxMultiSelIgnoreTbl;
        private CheckBox checkBoxMultiSelComments;
        private CheckBox checkBoxMultiSelBookMarks;
        private CheckBox checkBoxMultiSelInlineShapes;
        private CheckBox checkBoxMultiSelTables;
        private CheckBox checkBoxMultiSelEndNotes;
        private CheckBox checkBoxMultiSelFootNotes;
        private CheckBox checkBoxMultiSelIndices;
        private CheckBox checkBoxMultiHyperLinks;
        private CheckBox checkBoxMultiSelFields;
        private CheckBox checkBoxMultiSelCnts;
        private GroupBox groupBox5;
        private RadioButton radioBtnMultiSelCurSelScope;
        private RadioButton radioBtnMultiSelWholeStory;
        private GroupBox groupBox3;
        private GroupBox groupBox7;
        private Button btnMultiSelApplySel;
        private Button btnMultiSelHeadingAllClear;
        private Button btnMultiSelHeadingAllSel;
        private GroupBox grpW2XAutoModelScope;
        private RadioButton rdBtnW2XSelScope;
        private RadioButton rdBtnW2XAllDocScope;
        private Button btnAddTagCol;
        private Button btnW2XNextSameStructTbl;
        private TabPage tabPageFillGather;
        private CheckBox chkBoxStrictVerifyTblColName;
        private CheckedListBox chkListBoxTargetFiles;
        private Label label42;
        private Button btnFillGatherViewLog;
        private TreeView trvFillGatherSchemes;
        private RadioButton rdBtnFillGatherCurDoc;
        private RadioButton rdBtnFillGatherMultiFiles;
        private Button btnFillGatherProduce;
        private Button btnFillGatherPreviewProduce;
        private Label label43;
        private TextBox txtFillGatherName;
        private Button btnFillGatherAddTagNameValue;
        private Button btnFillGatherAddColValue;
        private Button btnFillGatherAddColName;
        private Button btnFillGatherAddTable;
        private Button btnFillGatherAddScheme;
        private Button btnFillGatherRemoveTblItem;
        private CheckBox chkBoxFillGatherStrictMatchColName;
        private GroupBox groupBox8;
        private RadioButton rdBtnFillGatherSelScope;
        private RadioButton rdBtnFillGatherAllDocScope;
        private Button btnFillGatherVerifyMatch;
        private Button btnFillGatherDelFiles;
        private Button btnFillGatherAddFiles;
        private Label label44;
        private Button btnFillGatherAllSelUnSel;
        private Button btnFillGatherAddUserDefineColName;
        private ProgressBar progBarFG;
        private Button btnFillGatherShowRowCol;
        private Button btnFillGatherMoveDown;
        private Button btnFillGatherMoveUp;
        private CheckBox chkBoxMulSelTblFirstRow;
        private CheckBox chkBoxMultiSelFirstColumn;
        private GroupBox groupBox9;
        private CheckBox chkBoxMultiSelColumnsScope;
        private CheckBox chkBoxMultiSelRowsScope;
        private Label label46;
        private Label label45;
        private NumericUpDown numMultiSelColEnd;
        private NumericUpDown numMultiSelColStart;
        private NumericUpDown numMultiSelRowEnd;
        private NumericUpDown numMultiSelRowStart;
        private CheckBox chkBoxMultiSelLastColumn;
        private CheckBox chkBoxMulSelTblLastRow;
        private Label label47;
        private RadioButton rdBtnMultiSelObjectPara;
        private RadioButton rdBtnMultiSelObjectRng;
        private TabPage tabPageMultiTiZhu;
        private TextBox txtSelectedInShpCapLbl;
        private Label label50;
        private Button btnAddSelInShpCapLbl;
        private Button btnRemoveSelInShpCapLbl;
        private TextBox txtSelectedTblCapLbl;
        private Label label49;
        private Button btnAddSelTblCapLbl;
        private Button btnRemoveSelTblCapLbl;
        private Button btnSetSysCapLbls;
        private Button btnRefreshCapsLbl;
        private ListBox lstBoxCurSysCapLbls;
        private GroupBox groupBox11;
        private Button btnApplyCapLbls2CurDoc;
        private ComboBox cmbTblCapLblPos;
        private ComboBox cmbTblCapLblAlign;
        private Label label55;
        private Label label54;
        private ComboBox cmbInShpCapLblAlign;
        private ComboBox cmbInShpCapLblPos;
        private Label label51;
        private Label label53;
        private RadioButton rdCapLblScopeSelection;
        private RadioButton rdCapLblScopeAllDoc;
        private Label label48;
        private GroupBox groupBox12;
        private Label label56;
        private Label label52;
        private Button btnNav2LastInShp;
        private Button btnNav2NextInShp;
        private Button btnNav2LastTbl;
        private Button btnNav2PrevInShp;
        private Button btnNav2NextTbl;
        private Button btnNav2FirstInShp;
        private Button btnNav2PrevTbl;
        private Button btnNav2FirstTbl;
        private TextBox txtInShpCapLblPostFix;
        private TextBox txtTblCapLblPostFix;
        private Label label58;
        private Label label57;
        private Label label62;
        private Button btnNavLastField;
        private Button btnNavNextField;
        private Button btnNavPrevField;
        private Button btnNavFirstField;
        private CheckBox chkSyncUpdateTableOfFigures;
        private CheckBox chkWholeTableCells;
        private TabPage tabPageObjNav;
        private Label label65;
        private Button btnONHeadingNavLast;
        private Button btnONHeadingNavNext;
        private Button btnONHeadingNavPrev;
        private Button btnONHeadingNavFirst;
        private Button btnONOutlineAllUnSel;
        private Button btnONOutlineAllSel;
        private CheckedListBox chkListObjNavOutline;
        private GroupBox groupBox13;
        private GroupBox groupBox15;
        private Button btnONEquationNavLast;
        private Button btnONObjectNavLast;
        private Button btnONBookmarkNavLast;
        private Button btnONEndnoteNavLast;
        private Button btnONFootnoteNavLast;
        private Button btnONCommentNavLast;
        private Button btnONEquationNavPrev;
        private Button btnONObjectNavPrev;
        private Button btnONBookmarkNavPrev;
        private Button btnONEndnoteNavPrev;
        private Button btnONFootnoteNavPrev;
        private Button btnONCommentNavPrev;
        private Button btnONEquationNavFirst;
        private Button btnONEquationNavNext;
        private Button btnONObjectNavFirst;
        private Button btnONObjectNavNext;
        private Button btnONBookmarkNavFirst;
        private Button btnONBookmarkNavNext;
        private Button btnONEndnoteNavFirst;
        private Label label79;
        private Button btnONEndnoteNavNext;
        private Label label78;
        private Button btnONFootnoteNavFirst;
        private Label label77;
        private Button btnONFootnoteNavNext;
        private Label label69;
        private Button btnONCommentNavFirst;
        private Label label68;
        private Button btnONCommentNavNext;
        private Label label66;
        private GroupBox groupBox14;
        private Button btnONFieldNavLast;
        private Label label67;
        private Button btnONSectionNavLast;
        private Label label76;
        private Button btnONPageNavLast;
        private Label label64;
        private Button btnONFieldNavNext;
        private Button btnONGraphicNavLast;
        private Button btnONSectionNavNext;
        private Label label63;
        private Button btnONPageNavNext;
        private Button btnONFieldNavFirst;
        private Button btnONTableNavLast;
        private Button btnONSectionNavFirst;
        private Button btnONGraphicNavNext;
        private Button btnONPageNavFirst;
        private Button btnONFieldNavPrev;
        private Button btnONSectionNavPrev;
        private Button btnONGraphicNavFirst;
        private Button btnONPageNavPrev;
        private Button btnONTableNavNext;
        private Button btnONGraphicNavPrev;
        private Button btnONTableNavFirst;
        private Button btnONTableNavPrev;
        private Label label70;
        private Label label71;
        private Label label72;
        private CheckBox chkTblCaplblGetFromHeading;
        private TextBox txtTblCapLblPreFix;
        private Label label59;
        private CheckBox chkInShpCaplblGetFromHeading;
        private TextBox txtInShpCapLblPreFix;
        private Label label61;
        private Label label60;
        private RichTextBox rchTextBoxUniformStylesPreview;
        private Button button17;
        public TabControl tabControl1;
        private Button btnHighLightGoLast;
        private Label label75;
        private Button btnHighLightGoNext;
        private Button btnHighLightGoFirst;
        private Button btnHighLightGoPrev;
        private RadioButton rdBtnAfterCurSel;
        private RadioButton rdBtnBeforeCurSel;
        private RadioButton rdBtnTiZhuAfterCurPos;
        private RadioButton rdBtnTiZhuBeforeCurPos;
        private Label label81;
        private Label label82;
        private CheckBox checkBoxMultiSelHighlight;
        private CheckBox checkBoxMultiSelSnParas;
        private GroupBox groupBox10;
        private GroupBox groupBox16;
        private CheckBox chkBoxMultiSelIgnoreHeadings;
        private CheckBox chkBoxMultiSelIgnoreTxtBody;
        private RadioButton rdBtnMultiSelIgnoreTbls;
        private RadioButton rdBtnMultiSelOnlyTbls;
        private RadioButton rdBtnMultiSelIncludeTbls;
        private CheckedListBox chkListBoxMultiListSnType;
        private Label label73;
        private Label label13;
        private CheckBox chkMultiSelUserDef;
        private ColorComboBox IncludeColorComboBox;
        private ColorComboBox ExcludeColorComboBox;
        private ColorComboBox colorComboBoxNav;
        private ColorComboBox colorComboBoxHighlight;
        private Button button12;
        private Button button11;
        private Button button10;
        private Button button9;

    }
}
