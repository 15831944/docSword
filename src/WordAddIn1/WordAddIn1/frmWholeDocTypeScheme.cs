using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeTools.Common;

using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using System.Collections.Specialized;

namespace OfficeAssist
{
    public partial class frmWholeDocTypeScheme : Form
    {
        Word.Application app = null;
        Word.Document doc = null;

        public String m_strNoChange = "(保持原值不变)";

        Boolean bStopUI2Data = false;
        TreeView trvheadingStyle = null;
        TreeView trvheadingSn = null;

        ClassWholeDocType inWholeDocType = new ClassWholeDocType();
        ClassWholeDocType oWholeDocType = null;

        // 封面
        ClassWholeDocType.sepPart curSepPart1stPage = null;
        ClassWholeDocType.sepPart defaultSepPart1stPage = new ClassWholeDocType.sepPart();

        ClassOfficeCommon cmnTools = Globals.ThisAddIn.m_commTools;

        // 章节目录
        int nCurHeadingTocLevel = -1;

        // 图文目录
        // 

        // 节和页眉页脚
        ClassWholeDocType.sepPart curSepPartSection = null;
        ClassWholeDocType.sepPart defaultSepPartSection = new ClassWholeDocType.sepPart();

        public frmWholeDocTypeScheme()
        {
            InitializeComponent();
        }

        public void setFeatures(Word.Application oApp, TreeView oTrvHeadingStyle, TreeView oTrvHeadingSn,
                                ClassWholeDocType outWholeDocType = null)
        {
            app = oApp;
            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                doc = null;
            }
            finally
            {

            }

            trvheadingStyle = oTrvHeadingStyle;
            trvheadingSn = oTrvHeadingSn;

            oWholeDocType = outWholeDocType;
            // TODO, copy to inWholeDocType

            // Word.Selection sel = app.ActiveWindow.Selection;

            return;
        }


        ///////////////////
        private void loadListData()
        {
            //
            cmb1stPageParts.Items.Add("其余部分");
            for (int i = 1; i < 11; i++)
            {
                cmb1stPageParts.Items.Add("第" + i + "部分");
            }


            cmbSections.Items.Add("其余节");
            for (int i = 1; i < 11; i++)
            {
                cmbSections.Items.Add("第" + i + "节");
            }

            ArrayList arrFontNames = new ArrayList();
            String strNoChange = m_strNoChange;
            //
            arrFontNames.Add(strNoChange);

            if (app != null)
            {
                foreach (String strFntNameItem in app.FontNames)
                {
                    arrFontNames.Add(strFntNameItem);
                }
            }

            foreach (String strFntNameItem in arrFontNames)
            {
                cmb1stPageChineseFontName.Items.Add(strFntNameItem);
                cmb1stPageAsciiFontName.Items.Add(strFntNameItem);

                cmbHeadingTOCTotalChineseFontName.Items.Add(strFntNameItem);
                cmbHeadingTOCTotalAsciiFontName.Items.Add(strFntNameItem);

                cmbHeadingTocLevelChineseFontName.Items.Add(strFntNameItem);
                cmbHeadingTocLevelAsciiFontName.Items.Add(strFntNameItem);

                cmbTuWenTocChineseFontName.Items.Add(strFntNameItem);
                cmbTuWenTocAsciiFontName.Items.Add(strFntNameItem);

                cmbTbleChineseFontName.Items.Add(strFntNameItem);
                cmbTbleAsciiFontName.Items.Add(strFntNameItem);
                cmbTiZhuChineseFontName.Items.Add(strFntNameItem);
                cmbTiZhuAsciiFontName.Items.Add(strFntNameItem);

                cmbTextBodyZoneChineseFontName.Items.Add(strFntNameItem);
                cmbTextBodyZoneAsciiFontName.Items.Add(strFntNameItem);
                cmbSectionChineseFontName.Items.Add(strFntNameItem);
                cmbSectionAsciiFontName.Items.Add(strFntNameItem);

            }

            ArrayList arrWordFontSize = new ArrayList();

            arrWordFontSize.Add(strNoChange);

            ArrayList arrFontSizes = Globals.ThisAddIn.getFontSizes();

            foreach (String strSize in arrFontSizes)
            {
                arrWordFontSize.Add(strSize);
            }

            foreach (String strFntSize in arrWordFontSize)
            {
                cmb1stPageFontSize.Items.Add(strFntSize);
                cmbHeadingTOCTotalFontSize.Items.Add(strFntSize);
                cmbHeadingTocLevelFontSize.Items.Add(strFntSize);
                cmbTuWenTocFontSize.Items.Add(strFntSize);
                cmbTbleFontSize.Items.Add(strFntSize);
                cmbTiZhuFontSize.Items.Add(strFntSize);
                cmbTextBodyZoneFontSize.Items.Add(strFntSize);
                cmbSectionFontSize.Items.Add(strFntSize);
            }

            // 段落，居左中右，行距RULE
            cmb1stPageParaAlignStyle.Items.Add(m_strNoChange);
            cmbTiZhuParaAlignStyle.Items.Add(m_strNoChange);
            cmbSectionParaAlignStyle.Items.Add(m_strNoChange);

            foreach (String strAlignStyleItem in Globals.ThisAddIn.m_arrStrAlignStyle)
            {
                cmb1stPageParaAlignStyle.Items.Add(strAlignStyleItem);
                cmbTiZhuParaAlignStyle.Items.Add(strAlignStyleItem);
                cmbSectionParaAlignStyle.Items.Add(strAlignStyleItem);
            }


            // String[] arrParaLineSpaceRule = { strNoChange, "单倍行距", "1.5 倍行距", "2 倍行距", "最小值", "固定值", "多倍行距" };

            cmb1stPageParaLineSpacingRule.Items.Add(m_strNoChange);
            cmbHeadingTOCTotalLineSpaceRule.Items.Add(m_strNoChange);
            cmbHeadingTocLevelParaLineSpaceRule.Items.Add(m_strNoChange);
            cmbTuWenTocParaLineSpaceRule.Items.Add(m_strNoChange);
            cmbTblParaLineSpaceRule.Items.Add(m_strNoChange);
            cmbTiZhuParaLineSpaceRule.Items.Add(m_strNoChange);
            cmbTextBodyZoneParaLineSpaceRule.Items.Add(m_strNoChange);
            cmbSectionParaLineSpaceRule.Items.Add(m_strNoChange);

            foreach (String strLineSpaceItem in Globals.ThisAddIn.m_arrParaLineSpaceRule)
            {
                cmb1stPageParaLineSpacingRule.Items.Add(strLineSpaceItem);
                cmbHeadingTOCTotalLineSpaceRule.Items.Add(strLineSpaceItem);
                cmbHeadingTocLevelParaLineSpaceRule.Items.Add(strLineSpaceItem);
                cmbTuWenTocParaLineSpaceRule.Items.Add(strLineSpaceItem);
                cmbTblParaLineSpaceRule.Items.Add(strLineSpaceItem);
                cmbTiZhuParaLineSpaceRule.Items.Add(strLineSpaceItem);
                cmbTextBodyZoneParaLineSpaceRule.Items.Add(strLineSpaceItem);
                cmbSectionParaLineSpaceRule.Items.Add(strLineSpaceItem);
            }

            // String[] arrParaLineSpaceUnit = { "磅", "厘米", "毫米", "英寸" };

            foreach (String strItem in Globals.ThisAddIn.m_arrSpaceUnit)
            {
                cmb1stParaLineSpaceUnit.Items.Add(strItem);
                cmbHeadingTOCTotalLineSpaceUnit.Items.Add(strItem);
                cmbHeadingTocLevelLineSpaceUnit.Items.Add(strItem);
                cmbTuWenTocLineSpaceUnit.Items.Add(strItem);
                cmbTableLineSpaceUnit.Items.Add(strItem);
                cmbTiZhuLineSpaceUnit.Items.Add(strItem);
                cmbTextBodyZoneLineSpaceUnit.Items.Add(strItem);
                cmbSectionLineSpaceUnit.Items.Add(strItem);
            }

            // load headingStyle / headingSn
            if (trvheadingStyle != null)
            {
                foreach (TreeNode rootNodes in trvheadingStyle.Nodes)
                {
                    trvHeadingStyleSchemes.Nodes.Add((TreeNode)rootNodes.Clone());
                }
            }

            if (trvheadingSn != null)
            {
                foreach (TreeNode rootNodes in trvheadingSn.Nodes)
                {
                    trvHeadingSnSchemes.Nodes.Add((TreeNode)rootNodes.Clone());
                }
            }

            return;
        }


        private void toggleEnable()
        {
            // change enable

            chk1stPageChosen.Checked = inWholeDocType.b1stPageEnable;
                grp1stBody.Enabled = chk1stPageChosen.Checked;

            chkHeadingTOCChosen.Checked = inWholeDocType.bHeadingTocEnable;
                rdHeadingTocTotal.Enabled = chkHeadingTOCChosen.Checked;
                rdHeadingTocLevel.Enabled = chkHeadingTOCChosen.Checked;
                grpHeadingTocTotal.Enabled = (chkHeadingTOCChosen.Checked && rdHeadingTocTotal.Checked);
                grpHeadingTocLevel.Enabled = (chkHeadingTOCChosen.Checked && rdHeadingTocLevel.Checked);

            chkTuWenTocChosen.Checked = inWholeDocType.bTuWenTocEnable;
            grpTuWenTocBody.Enabled = chkTuWenTocChosen.Checked;

            chkHeadingChosen.Checked = inWholeDocType.bHeadingEnable;
            grpHeadingBody.Enabled = chkHeadingChosen.Checked;

            chkTableChosen.Checked = inWholeDocType.bTableEnable;
            grpTableBody.Enabled = chkTableChosen.Checked;

            chkTiZuChosen.Checked = inWholeDocType.bTizhuEnable;
            grpTiZhuBody.Enabled = chkTiZuChosen.Checked;

            chkTextBodyZoneChosen.Checked = inWholeDocType.bTextBodyZoneEnable;
            grpTextBodyZoneBody.Enabled = chkTextBodyZoneChosen.Checked;

            chkSectionChosen.Checked = inWholeDocType.bSectionEnable;
            grpSectionBody.Enabled = chkSectionChosen.Checked;

            return;
        }


        private void frmWholeDocTypeScheme_Load(object sender, EventArgs e)
        {

            tblCtrlSchemeSetting.TabPages.Clear();

            tblCtrlSchemeSetting.TabPages.Add(tabPageStyle);

            curSepPart1stPage = defaultSepPart1stPage;
            curSepPartSection = defaultSepPartSection;

            bStopUI2Data = true;

            loadListData();
            toggleEnable();

            Data2UI_1stPage(inWholeDocType);
            Data2UI_headingTOC(inWholeDocType);
            Data2UI_TuWenToc(inWholeDocType);
            Data2UI_Table(inWholeDocType);
            Data2UI_TiZhu(inWholeDocType);
            Data2UI_TextBodyZone(inWholeDocType);
            Data2UI_Section(inWholeDocType);
            bStopUI2Data = false;

            return;
        }


        private void btn1stPageFonts_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();

            fntDialog.setFeatures(app, curSepPart1stPage.cFont);

            DialogResult res = fntDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassFont cFont = fntDialog.getFont();

            cFont.SelCopy2(curSepPart1stPage.cFont);
            Data2UI_Font_1stPage(cFont);

            return;
        }

        private void btn1stPageParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            paraFmtDialog.setFeatures(app,curSepPart1stPage.cParaFmt);

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassParagraphFormat paraFmt = paraFmtDialog.getParaFmt();

            paraFmt.SelCopy2(curSepPart1stPage.cParaFmt);
            Data2UI_ParaFmt_1stPage(paraFmt);

            return;
        }

        private void btnHeadingTOCTotalFonts_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();

            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            return;
        }



        private void btnHeadingTOCTotalParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            ClassParagraphFormat totalParaFmt = inWholeDocType.arrsHeadingTocParaFmt[0];

            paraFmtDialog.setFeatures(app, totalParaFmt);

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            ClassParagraphFormat dlgParaFmt = paraFmtDialog.getParaFmt();

            dlgParaFmt.SelCopy2(totalParaFmt);

            //TODO
            Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt,0);

            return;
        }

        private void btnHeadingTOCLevelFonts_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();
            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            return;
        }

        private void btnHeadingTOCLevelParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            ClassParagraphFormat curParaFmt = inWholeDocType.arrsHeadingTocParaFmt[nCurHeadingTocLevel];

            paraFmtDialog.setFeatures(app, curParaFmt);

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            ClassParagraphFormat dlgParaFmt = paraFmtDialog.getParaFmt();

            dlgParaFmt.SelCopy2(curParaFmt);

            //TODO
            Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt,nCurHeadingTocLevel);

            return;
        }

        private void btnTuWenTOCFonts_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();
            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            return;
        }

        private void btnTuWenTOCParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassParagraphFormat oParaFmt = paraFmtDialog.getParaFmt();
            oParaFmt.SelCopy2(inWholeDocType.tuWenTocTotalParaFmt);

            Data2UI_ParaFmt_TuWenToc(inWholeDocType.tuWenTocTotalParaFmt);

            return;
        }

        private void btnTableTotalFonts_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();
            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }




            return;
        }

        private void btnTableTotalParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassParagraphFormat oParaFmt = paraFmtDialog.getParaFmt();
            oParaFmt.SelCopy2(inWholeDocType.tableTotalParaFmt);

            Data2UI_ParaFmt_Table(inWholeDocType.tableTotalParaFmt);

            return;
        }

        private void btnTiZhuFonts_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();
            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            return;
        }


        private void btnTiZhuParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassParagraphFormat oParaFmt = paraFmtDialog.getParaFmt();
            oParaFmt.SelCopy2(inWholeDocType.tizhuParaFmt);

            Data2UI_ParaFmt_TiZhu(inWholeDocType.tizhuParaFmt);

            return;
        }

        private void btnTextBodyZoneFont_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();
            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            return;
        }

        private void btnTextBodyZoneParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassParagraphFormat oParaFmt = paraFmtDialog.getParaFmt();
            oParaFmt.SelCopy2(inWholeDocType.textbodyZoneParaFmt);

            Data2UI_ParaFmt_TextBodyZone(inWholeDocType.textbodyZoneParaFmt);

            return;
        }

        private void btnSectionFont_Click(object sender, EventArgs e)
        {
            frmFontSetting fntDialog = new frmFontSetting();
            fntDialog.setFeatures(app);

            DialogResult res = fntDialog.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            return;
        }

        private void btnSectionParagraph_Click(object sender, EventArgs e)
        {
            frmParagraphFormat paraFmtDialog = new frmParagraphFormat();

            DialogResult res = paraFmtDialog.ShowDialog();

            if (res != DialogResult.OK)
            {
                return;
            }

            ClassParagraphFormat paraFmt = paraFmtDialog.getParaFmt();

            paraFmt.SelCopy2(curSepPartSection.cParaFmt);
            Data2UI_ParaFmt_Section(paraFmt);

            return;
        }


        private void chk1stPageChosen_CheckedChanged(object sender, EventArgs e)
        {
            grp1stBody.Enabled = chk1stPageChosen.Checked;
            inWholeDocType.b1stPageEnable = chk1stPageChosen.Checked;

            return;
        }


        private void chkHeadingTOCChosen_CheckedChanged(object sender, EventArgs e)
        {
            rdHeadingTocTotal.Enabled = chkHeadingTOCChosen.Checked;
            rdHeadingTocLevel.Enabled = chkHeadingTOCChosen.Checked;

            grpHeadingTocTotal.Enabled = (chkHeadingTOCChosen.Checked && rdHeadingTocTotal.Checked);
            grpHeadingTocLevel.Enabled = (chkHeadingTOCChosen.Checked && rdHeadingTocLevel.Checked);

            inWholeDocType.bHeadingTocEnable = chkHeadingTOCChosen.Checked;

            if (!chkHeadingTOCChosen.Checked)
            {
                return;
            }

            if (rdHeadingTocTotal.Checked)
            {
                Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt, 0);
                Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt,0);
            }
            else
            {
                Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt, nCurHeadingTocLevel);
                Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt,nCurHeadingTocLevel);
            }

            return;
        }


        private void chkTuWenTocChosen_CheckedChanged(object sender, EventArgs e)
        {
            grpTuWenTocBody.Enabled = chkTuWenTocChosen.Checked;

            inWholeDocType.bTuWenTocEnable = chkTuWenTocChosen.Checked;

            if (!chkTuWenTocChosen.Checked)
            {
                return;
            }

            Data2UI_Font_TuWenToc(inWholeDocType.tuWenTocTotalFnt);
            Data2UI_ParaFmt_TuWenToc(inWholeDocType.tuWenTocTotalParaFmt);

            return;
        }

        private void chkHeadingChosen_CheckedChanged(object sender, EventArgs e)
        {
            grpHeadingBody.Enabled = chkHeadingChosen.Checked;

            return;
        }

        private void chkTableChosen_CheckedChanged(object sender, EventArgs e)
        {
            grpTableBody.Enabled = chkTableChosen.Checked;
            inWholeDocType.bTableEnable = chkTableChosen.Checked;

            if (!chkTableChosen.Checked)
            {
                return;
            }

            Data2UI_Font_Table(inWholeDocType.tableTotalFont);
            Data2UI_ParaFmt_Table(inWholeDocType.tableTotalParaFmt);

            return;
        }

        private void chkTiZuChosen_CheckedChanged(object sender, EventArgs e)
        {
            grpTiZhuBody.Enabled = chkTiZuChosen.Checked;
            inWholeDocType.bTizhuEnable = chkTiZuChosen.Checked;

            if (!chkTiZuChosen.Checked)
            {
                return;
            }

            Data2UI_Font_TiZhu(inWholeDocType.tizhuFont);
            Data2UI_ParaFmt_TiZhu(inWholeDocType.tizhuParaFmt);

            return;
        }

        private void chkTextBodyZoneChosen_CheckedChanged(object sender, EventArgs e)
        {
            grpTextBodyZoneBody.Enabled = chkTextBodyZoneChosen.Checked;
            inWholeDocType.bTextBodyZoneEnable = chkTextBodyZoneChosen.Checked;

            if (!chkTextBodyZoneChosen.Checked)
            {
                return;
            }

            Data2UI_Font_TextBodyZone(inWholeDocType.textbodyZoneFont);
            Data2UI_ParaFmt_TextBodyZone(inWholeDocType.textbodyZoneParaFmt);

            return;
        }

        private void chkSectionChosen_CheckedChanged(object sender, EventArgs e)
        {
            grpSectionBody.Enabled = chkSectionChosen.Checked;

            return;
        }

        private void trvHeadingStyleSchemes_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode selNode = trvHeadingStyleSchemes.SelectedNode;
            String strSchemeName = "";

            if (selNode != null)
            {
                switch (selNode.Level)
                {
                    case 1:
                        strSchemeName = selNode.Text;
                        txtChosenHeadingStyleScheme.Text = strSchemeName;
                        break;

                    case 2:
                        strSchemeName = selNode.Parent.Text;
                        txtChosenHeadingStyleScheme.Text = strSchemeName;
                        break;

                    case 0:
                    default:
                        break;
                }
            }
            
            return;
        }


        private void trvHeadingSnSchemes_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode selNode = trvHeadingSnSchemes.SelectedNode;
            String strSchemeName = "";

            if (selNode != null)
            {
                switch (selNode.Level)
                {
                    case 1:
                        strSchemeName = selNode.Text;
                        txtChosenHeadingSnScheme.Text = strSchemeName;
                        break;

                    case 2:
                        strSchemeName = selNode.Parent.Text;
                        txtChosenHeadingSnScheme.Text = strSchemeName;
                        break;

                    case 0:
                    default:
                        break;
                }
            }

            return;
        }

        private void btnHeadingClearStyleChosen_Click(object sender, EventArgs e)
        {
            txtChosenHeadingStyleScheme.Text = "";

            return;
        }

        private void btnHeadingClearSnChosen_Click(object sender, EventArgs e)
        {
            txtChosenHeadingSnScheme.Text = "";

            return;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //if(cWholeDocType == null)
            //{
            
            inWholeDocType = new ClassWholeDocType();
            
            //}

            // enable
            inWholeDocType.b1stPageEnable = chk1stPageChosen.Checked;

            // save 2 
            int nSelIndex = lstChosenParts.SelectedIndex;
            if (nSelIndex != -1)
            {
                //String strItem = (String)lstChosenParts.Items[nSelIndex];

                //ClassWholeDocType.sepPart sep = (ClassWholeDocType.sepPart)cWholeDocType.hsh1stPagePart[strItem];

                //if (sep != null)
                //{
                //    sep.clone(curSepPart1stPage);
                //}
            }
            
            
            inWholeDocType.bHeadingTocEnable = chkHeadingTOCChosen.Checked;
            inWholeDocType.bTuWenTocEnable = chkTuWenTocChosen.Checked;
            inWholeDocType.bHeadingEnable = chkHeadingChosen.Checked;
            inWholeDocType.bTableEnable = chkTableChosen.Checked;
            inWholeDocType.bTizhuEnable = chkTiZuChosen.Checked;
            inWholeDocType.bTextBodyZoneEnable = chkTextBodyZoneChosen.Checked;
            inWholeDocType.bSectionEnable = chkSectionChosen.Checked;

            // 



            return;
        }

        private void btn1stPageAddPart_Click(object sender, EventArgs e)
        {
            String strText = cmb1stPageParts.Text;
            // 判重
            // 

            if (String.IsNullOrWhiteSpace(strText))
            {
                MessageBox.Show("不能添加空选择");
                return;
            }

            if (lstChosenParts.Items.Contains(strText))
            {
                MessageBox.Show("不能添加重复项"); // lstChosenParts
                return;
            }
            
            
            ClassWholeDocType.sepPart sep = new ClassWholeDocType.sepPart();

            curSepPart1stPage.cFont.SelCopy2(sep.cFont);
            curSepPart1stPage.cParaFmt.SelCopy2(sep.cParaFmt);

            inWholeDocType.hsh1stPagePart.Add(strText, sep);

            curSepPart1stPage = sep;

            lstChosenParts.Items.Add(strText);
            lstChosenParts.SelectedItem = strText;

            return;
        }

        private void btn1stPageRemovePart_Click(object sender, EventArgs e)
        {
            int nSelIndex = lstChosenParts.SelectedIndex;
            int nCount = lstChosenParts.Items.Count;

            if (nCount == 0 || nSelIndex == -1)
            {
                MessageBox.Show("请选中一项后进行此操作");
            }

            DialogResult res = MessageBox.Show("确认删除选中项？","确认",MessageBoxButtons.YesNo);

            if (res == DialogResult.No)
            {
                return;
            }

            int nNextIndex = -1;

            if (nCount == 1)
            {
                nNextIndex = -1; // 
            }
            else if (nSelIndex == nCount - 1)
            {
                nNextIndex = nCount - 2;
            }
            else
            {
                nNextIndex = nSelIndex; // 
            }

            String strItem = (String)lstChosenParts.Items[lstChosenParts.SelectedIndex];

            lstChosenParts.Items.RemoveAt(lstChosenParts.SelectedIndex);

            inWholeDocType.hsh1stPagePart.Remove(strItem);

            if (lstChosenParts.Items.Count == 0)
            {
                curSepPart1stPage = defaultSepPart1stPage;
            }
            else if(nNextIndex != -1)
            {
                lstChosenParts.SelectedIndex = nNextIndex;
                strItem = (String)lstChosenParts.Items[nNextIndex];
                // lstChosenParts.SelectedItem = strItem;

                ClassWholeDocType.sepPart sep = (ClassWholeDocType.sepPart)inWholeDocType.hsh1stPagePart[strItem];

                if (sep != null)
                {
                    curSepPart1stPage = sep;
                }
                // curSepPart1stPage.
            }

            Data2UI_Font_1stPage(curSepPart1stPage.cFont);
            Data2UI_ParaFmt_1stPage(curSepPart1stPage.cParaFmt);

            return;
        }

        private void btn1stPageUpdatePart_Click(object sender, EventArgs e)
        {
            int nSelIndex = lstChosenParts.SelectedIndex;
            int nCount = lstChosenParts.Items.Count;

            if (lstChosenParts.Items.Count == 0 || lstChosenParts.SelectedIndex == -1)
            {
                MessageBox.Show("请选中一项后进行此操作");
            }

            String strText = cmb1stPageParts.Text;

            // 判重
            // 
            if (lstChosenParts.Items.Contains(strText))
            {
                MessageBox.Show("不能添加重复项"); // lstChosenParts
                return;
            }

            String strOldItem = (String)lstChosenParts.Items[lstChosenParts.SelectedIndex];

            inWholeDocType.hsh1stPagePart.Remove(strOldItem);

            lstChosenParts.Items.RemoveAt(lstChosenParts.SelectedIndex);

            int nIndex = lstChosenParts.Items.Add(strText);

            // get current font / paragraph format

            inWholeDocType.hsh1stPagePart.Add(strText, curSepPart1stPage);

            lstChosenParts.SelectedIndex = nIndex;

            // update data
            Data2UI_Font_1stPage(curSepPart1stPage.cFont);
            Data2UI_ParaFmt_1stPage(curSepPart1stPage.cParaFmt);
            // 

            return;
        }

        private void lstChosenParts_SelectedIndexChanged(object sender, EventArgs e)
        {
            int nSelIndex = lstChosenParts.SelectedIndex;

            if (nSelIndex == -1)
            {
                return;
            }

            ClassWholeDocType.sepPart sep = null;

            String strItem = (String)lstChosenParts.Items[nSelIndex];

            sep = (ClassWholeDocType.sepPart)inWholeDocType.hsh1stPagePart[strItem];

            bStopUI2Data = true;
            // load current font and paragraph
            if (sep != null)
            {
                curSepPart1stPage = sep;

                Data2UI_Font_1stPage(curSepPart1stPage.cFont);
                Data2UI_ParaFmt_1stPage(curSepPart1stPage.cParaFmt);
            }

            bStopUI2Data = false;

            return;
        }

        private void cmb1stPageChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_1stPage(ref curSepPart1stPage.cFont);
            }

            return;
        }

        private void cmb1stPageAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_1stPage(ref curSepPart1stPage.cFont);
            }

            return;
        }


        private void cmb1stPageFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_1stPage(ref curSepPart1stPage.cFont);
            }
            return;
        }

        private void cmb1stPageParaAlignStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_1stPage(ref curSepPart1stPage.cParaFmt);
            }
            return;
        }


        private void cmb1stPageParaLineSpacingRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmb1stPageParaLineSpacingRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                num1stPageParaLineValue.Enabled = true;
                cmb1stParaLineSpaceUnit.Enabled = true;
                cmb1stParaLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                num1stPageParaLineValue.Enabled = true;
                cmb1stParaLineSpaceUnit.Enabled = false;
                cmb1stParaLineSpaceUnit.Text = "行";
            }
            else
            {
                num1stPageParaLineValue.Enabled = false;
                cmb1stParaLineSpaceUnit.Enabled = false;
                cmb1stParaLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_1stPage(ref curSepPart1stPage.cParaFmt);
            return;
        }


        private void cmb1stParaLineSpaceUnit_TextChanged(object sender, EventArgs e)
        {
            if (!cmb1stParaLineSpaceUnit.Enabled)
            {
                return;
            }

            // 
            if (curSepPart1stPage == null)
            {
                return;
            }

            String strLineSpacingRule = cmb1stPageParaLineSpacingRule.Text;

            if (String.IsNullOrWhiteSpace(strLineSpacingRule))
            {
                return;
            }

            float fLineSpaceValue = (float)num1stPageParaLineValue.Value;
            float fPonds = 0.0f;

            if (strLineSpacingRule.Equals(m_strNoChange))
            {
                //curSepPart1stPage.cParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else if (strLineSpacingRule.Equals("单倍行距"))
            {
                //curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                //curSepPart1stPage.cParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                //num1stPageParaLineValue.Enabled = false;
                //cmb1stParaLineSpaceUnit.Enabled = false;

            }
            else if (strLineSpacingRule.Equals("1.5 倍行距"))
            {
                //curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                //curSepPart1stPage.cParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;

                //num1stPageParaLineValue.Enabled = false;
                //cmb1stParaLineSpaceUnit.Enabled = false;

            }
            else if (strLineSpacingRule.Equals("2 倍行距"))
            {
                //curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                //curSepPart1stPage.cParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;

                //num1stPageParaLineValue.Enabled = false;
                //cmb1stParaLineSpaceUnit.Enabled = false;
            }
            else if (strLineSpacingRule.Equals("最小值"))
            {
                //curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                //curSepPart1stPage.cParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                String strText = cmb1stParaLineSpaceUnit.Text;

                if (strText.Equals("磅"))
                {
                    curSepPart1stPage.cParaFmt.LineSpacing = fLineSpaceValue;
                }
                else if (strText.Equals("厘米"))
                {
                    curSepPart1stPage.cParaFmt.LineSpacing = fLineSpaceValue * 28.35f;
                }
                else if (strText.Equals("毫米"))
                {
                    curSepPart1stPage.cParaFmt.LineSpacing = fLineSpaceValue * 2.835f;
                }
                else if (strText.Equals("英寸"))
                {
                    curSepPart1stPage.cParaFmt.LineSpacing = fLineSpaceValue * 2.54f * 28.35f;
                }

            }
            else if (strLineSpacingRule.Equals("固定值"))
            {
                String strText = cmb1stParaLineSpaceUnit.Text;

                if (strText.Equals("磅"))
                {
                    fPonds = fLineSpaceValue;

                    if (fPonds < 0.7f)
                    {
                        MessageBox.Show("行间距固定值最少0.7磅");
                        num1stPageParaLineValue.Value = 1;
                        return;
                    }

                    curSepPart1stPage.cParaFmt.LineSpacing = fLineSpaceValue;
                    curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                }
                else if (strText.Equals("厘米"))
                {
                    fPonds = fLineSpaceValue * 28.35f;

                    if (fPonds < 0.7f)
                    {
                        MessageBox.Show("行间距固定值最少0.025厘米");
                        num1stPageParaLineValue.Value = 0.025m;
                        return;
                    }

                    curSepPart1stPage.cParaFmt.LineSpacing = fPonds;
                    curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                }
                else if (strText.Equals("毫米"))
                {
                    fPonds = fLineSpaceValue * 2.835f;

                    if (fPonds < 0.7f)
                    {
                        MessageBox.Show("行间距固定值最少0.25毫米");
                        num1stPageParaLineValue.Value = 0.25m;
                        return;
                    }

                    curSepPart1stPage.cParaFmt.LineSpacing = fPonds;
                    curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                }
                else if (strText.Equals("英寸"))
                {
                    fPonds = fLineSpaceValue * 2.54f * 28.35f;

                    if (fPonds < 0.7f)
                    {
                        MessageBox.Show("行间距固定值最少0.01英寸");
                        num1stPageParaLineValue.Value = 0.01m;
                        return;
                    }

                    curSepPart1stPage.cParaFmt.LineSpacing = fPonds;
                    curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                }
            }
            else if (strLineSpacingRule.Equals("多倍行距"))
            {
                if (fLineSpaceValue < 0.25f)
                {
                    MessageBox.Show("行间距固定值最少0.25");
                    num1stPageParaLineValue.Value = 0.25m;
                    return;
                }

                curSepPart1stPage.cParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                curSepPart1stPage.cParaFmt.LineSpacing = app.LinesToPoints((float)num1stPageParaLineValue.Value);
            }


            return;
        }

        private void num1stPageParaLineValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_1stPage(ref curSepPart1stPage.cParaFmt);
            }
            return;
        }

        private void cmb1stParaLineSpaceUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_1stPage(ref curSepPart1stPage.cParaFmt);
            }
            return;
        }

        // 
        private void Data2UI_1stPage(ClassWholeDocType oWdt = null)
        {
            cmb1stPageParts.Text = "第1部分";

            if(oWdt == null)
            {
                Data2UI_Font_1stPage(null);
                Data2UI_ParaFmt_1stPage(null);

                return;
            }

            // oWdt.hsh1stPagePart
            String strItem = "";
            foreach (DictionaryEntry ent in oWdt.hsh1stPagePart)
            {
                strItem = (String)ent.Key;

                lstChosenParts.Items.Add(strItem);
            }

            if (lstChosenParts.Items.Count > 0)
            {
                lstChosenParts.SelectedIndex = 0;
            }
            else
            {
                Data2UI_Font_1stPage(null);
                Data2UI_ParaFmt_1stPage(null);
            }

            return;
        }



        private void Data2UI_Font_1stPage(ClassFont oFont = null)
        {
            if(oFont == null)
            {
                cmb1stPageChineseFontName.Text = m_strNoChange;
                cmb1stPageAsciiFontName.Text = m_strNoChange;
                cmb1stPageFontSize.Text = m_strNoChange;

                chk1stPageFontBold.CheckState = CheckState.Indeterminate;
                chk1stPageFontItalic.CheckState = CheckState.Indeterminate;

                return;
            }

            if (oFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmb1stPageChineseFontName.Text = oFont.NameFarEast;
            }
            else
            {
                cmb1stPageChineseFontName.Text = m_strNoChange;
            }


            if (oFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmb1stPageAsciiFontName.Text = oFont.NameAscii;
            }
            else
            {
                cmb1stPageAsciiFontName.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Size))
            {
                cmb1stPageFontSize.Text = "" + oFont.Size;
            }
            else
            {
                cmb1stPageFontSize.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Bold))
            {
                if (oFont.Bold != 0)
                {
                    chk1stPageFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chk1stPageFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chk1stPageFontBold.CheckState = CheckState.Indeterminate;
            }


            if (oFont.isSet(ClassFont.euMembers.Italic))
            {
                if (oFont.Italic != 0)
                {
                    chk1stPageFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chk1stPageFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chk1stPageFontItalic.CheckState = CheckState.Indeterminate;
            }

            return;
        }


        private void Data2UI_ParaFmt_1stPage(ClassParagraphFormat oParaFmt = null)
        {
            if (oParaFmt == null)
            {
                cmb1stPageParaAlignStyle.Text = m_strNoChange;
                cmb1stPageParaLineSpacingRule.Text = m_strNoChange;

                num1stPageParaLineValue.Value = 0.0m;
                cmb1stParaLineSpaceUnit.Text = "磅";
                return;
            }


            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.Alignment))
            {
                cmb1stPageParaAlignStyle.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.Alignment)
                {
                    case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                        cmb1stPageParaAlignStyle.Text = "居左";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphRight:
                        cmb1stPageParaAlignStyle.Text = "居右";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                        cmb1stPageParaAlignStyle.Text = "居中";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                        cmb1stPageParaAlignStyle.Text = "分散对齐";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                        cmb1stPageParaAlignStyle.Text = "两端对齐";
                        break;

                    default:
                        break;
                }
            }

            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
            {
                cmb1stPageParaLineSpacingRule.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmb1stPageParaLineSpacingRule.Text = "单倍行距";
                        cmb1stParaLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmb1stPageParaLineSpacingRule.Text = "1.5 倍行距";
                        cmb1stParaLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmb1stPageParaLineSpacingRule.Text = "2 倍行距";
                        cmb1stParaLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmb1stPageParaLineSpacingRule.Text = "最小值";

                        num1stPageParaLineValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmb1stParaLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmb1stPageParaLineSpacingRule.Text = "固定值";
                        num1stPageParaLineValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmb1stParaLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmb1stPageParaLineSpacingRule.Text = "多倍行距";
                        num1stPageParaLineValue.Value = (decimal)app.PointsToLines(oParaFmt.LineSpacing);
                        cmb1stParaLineSpaceUnit.Text = "行";
                        break;

                    default:
                        break;
                }
            }

            return;
        }



        // 
        private void UI2Data_1stPage()
        {
            return;
        }

        private void UI2Data_Font_1stPage(ref ClassFont oFont)
        {
            String strItem = cmb1stPageChineseFontName.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                oFont.NameFarEast = strItem;
            }

            strItem = cmb1stPageAsciiFontName.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                oFont.NameAscii = strItem;
            }

            String strFntSize = cmb1stPageFontSize.Text;

            if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    oFont.AddSelMember((int)ClassFont.euMembers.Size);
                    oFont.Size = fSize;
                }
                else
                {
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }


            switch (chk1stPageFontBold.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            switch (chk1stPageFontItalic.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }


            return;
        }


        private void UI2Data_ParaFmt_1stPage(ref ClassParagraphFormat oParaFmt)
        {
            String strItem = cmb1stPageParaAlignStyle.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.Alignment);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.Alignment);
                if (strItem.Equals("居左"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                else if (strItem.Equals("居中"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (strItem.Equals("居右"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }
                else if (strItem.Equals("两端对齐"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (strItem.Equals("分散对齐"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                }

            }

            strItem = cmb1stPageParaLineSpacingRule.Text;
            String strUnit = cmb1stParaLineSpaceUnit.Text;
            float fValue = (float)num1stPageParaLineValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    oParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            return;
        }

        private void chk1stPageFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_1stPage(ref curSepPart1stPage.cFont);
            }
            return;
        }

        private void chk1stPageFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_1stPage(ref curSepPart1stPage.cFont);
            }
            return;
        }


        // Heading TOC, UI2Data
        private void UI2Data_headingTOC()
        {
            return;
        }

        private void UI2Data_Font_headingTOC(ref ClassFont[] oFont)
		{
            Boolean bTotal = rdHeadingTocTotal.Checked;

            ClassFont dstFont = null;
            String strItem = "";
            String strFntSize = "";

            if (bTotal)
            {
                dstFont = oFont[0];

                strItem = cmbHeadingTOCTotalChineseFontName.Text;

                if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
                {
                    dstFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
                }
                else
                {
                    dstFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                    dstFont.NameFarEast = strItem;
                }

                strItem = cmbHeadingTOCTotalAsciiFontName.Text;
                if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
                {
                    dstFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
                }
                else
                {
                    dstFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                    dstFont.NameAscii = strItem;
                }

                strFntSize = cmbHeadingTOCTotalFontSize.Text;

                if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
                {
                    dstFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
                else
                {
                    float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                    if (!float.IsNaN(fSize) && fSize > 0.0f)
                    {
                        dstFont.AddSelMember((int)ClassFont.euMembers.Size);
                        dstFont.Size = fSize;
                    }
                    else
                    {
                        dstFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                    }
                }


                switch (chkHeadingTOCTotalFontBold.CheckState)
                {
                    case CheckState.Checked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Bold);
                        dstFont.Bold = -1;
                        break;

                    case CheckState.Unchecked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Bold);
                        dstFont.Bold = 0;
                        break;

                    case CheckState.Indeterminate:
                        dstFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                        break;
                }


                switch (chkHeadingTOCTotalFontItalic.CheckState)
                {
                    case CheckState.Checked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Italic);
                        dstFont.Italic = -1;
                        break;

                    case CheckState.Unchecked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Italic);
                        dstFont.Italic = 0;
                        break;

                    case CheckState.Indeterminate:
                        dstFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                        break;
                }


                for (int nIndex = 1; nIndex <= 9; nIndex++)
                {
                    dstFont.SelCopy2(oFont[nIndex]);
                }

                // 

            }
            else
            {
                if (!(nCurHeadingTocLevel >= 1 && nCurHeadingTocLevel <= 9))
                {
                    return;
                }

                dstFont = oFont[nCurHeadingTocLevel];

                strItem = cmbHeadingTocLevelChineseFontName.Text;

                if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
                {
                    dstFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
                }
                else
                {
                    dstFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                    dstFont.NameFarEast = strItem;
                }

                strItem = cmbHeadingTocLevelAsciiFontName.Text;
                if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
                {
                    dstFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
                }
                else
                {
                    dstFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                    dstFont.NameAscii = strItem;
                }

                strFntSize = cmbHeadingTocLevelFontSize.Text;

                if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
                {
                    dstFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
                else
                {
                    float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                    if (!float.IsNaN(fSize) && fSize > 0.0f)
                    {
                        dstFont.AddSelMember((int)ClassFont.euMembers.Size);
                        dstFont.Size = fSize;
                    }
                    else
                    {
                        dstFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                    }
                }


                switch (chkHeadingTocLevelFontBold.CheckState)
                {
                    case CheckState.Checked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Bold);
                        dstFont.Bold = -1;
                        break;

                    case CheckState.Unchecked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Bold);
                        dstFont.Bold = 0;
                        break;

                    case CheckState.Indeterminate:
                        dstFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                        break;
                }


                switch (chkHeadingTocLevelFontItalic.CheckState)
                {
                    case CheckState.Checked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Italic);
                        dstFont.Italic = -1;
                        break;

                    case CheckState.Unchecked:
                        dstFont.AddSelMember((int)ClassFont.euMembers.Italic);
                        dstFont.Italic = 0;
                        break;

                    case CheckState.Indeterminate:
                        dstFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                        break;
                }
            }

            return;
        }


        private void UI2Data_ParaFmt_headingTOC(ref ClassParagraphFormat[] oParaFmt)
		{
            ClassParagraphFormat dstParaFmt = oParaFmt[0];

            String strItem = cmbHeadingTOCTotalLineSpaceRule.Text;
            String strUnit = cmbHeadingTOCTotalLineSpaceUnit.Text;
            float fValue = (float)numHeadingTOCTotalLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                dstParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    dstParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    dstParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    dstParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            if (!(nCurHeadingTocLevel >= 1 && nCurHeadingTocLevel <= 9))
            {
                return;
            }

            dstParaFmt = oParaFmt[nCurHeadingTocLevel];

            strItem = cmbHeadingTocLevelParaLineSpaceRule.Text;
            strUnit = cmbHeadingTocLevelLineSpaceUnit.Text;
            fValue = (float)numHeadingTocLevelParaLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                dstParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    dstParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    dstParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    dstParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    dstParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    dstParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            return;
        }

        // Data2UI
		
        private void Data2UI_headingTOC(ClassWholeDocType oWdt = null)
        {
            nCurHeadingTocLevel = 1;

            Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt, 0);
            Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt, 0);

            Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt, nCurHeadingTocLevel);
            Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt, nCurHeadingTocLevel);

            return;
        }

        private void Data2UI_Font_headingTOC(ClassFont[] oFont, int nIndex)
        {
            ClassFont dstFont = oFont[nIndex];

            if (nIndex == 0)
            {
                if (dstFont.isSet(ClassFont.euMembers.NameFarEast))
                {
                    cmbHeadingTOCTotalChineseFontName.Text = dstFont.NameFarEast;
                }
                else
                {
                    cmbHeadingTOCTotalChineseFontName.Text = m_strNoChange;
                }


                if (dstFont.isSet(ClassFont.euMembers.NameAscii))
                {
                    cmbHeadingTOCTotalAsciiFontName.Text = dstFont.NameAscii;
                }
                else
                {
                    cmbHeadingTOCTotalAsciiFontName.Text = m_strNoChange;
                }

                if (dstFont.isSet(ClassFont.euMembers.Size))
                {
                    cmbHeadingTOCTotalFontSize.Text = "" + dstFont.Size;
                }
                else
                {
                    cmbHeadingTOCTotalFontSize.Text = m_strNoChange;
                }

                if (dstFont.isSet(ClassFont.euMembers.Bold))
                {
                    if (dstFont.Bold != 0)
                    {
                        chkHeadingTOCTotalFontBold.CheckState = CheckState.Checked;
                    }
                    else
                    {
                        chkHeadingTOCTotalFontBold.CheckState = CheckState.Unchecked;
                    }
                }
                else
                {
                    chkHeadingTOCTotalFontBold.CheckState = CheckState.Indeterminate;
                }


                if (dstFont.isSet(ClassFont.euMembers.Italic))
                {
                    if (dstFont.Italic != 0)
                    {
                        chkHeadingTOCTotalFontItalic.CheckState = CheckState.Checked;
                    }
                    else
                    {
                        chkHeadingTOCTotalFontItalic.CheckState = CheckState.Unchecked;
                    }
                }
                else
                {
                    chkHeadingTOCTotalFontItalic.CheckState = CheckState.Indeterminate;
                }
            }
            else
            {
                // lstHeadingLevel.SelectedIndex = nIndex - 1;
                if (dstFont.isSet(ClassFont.euMembers.NameFarEast))
                {
                    cmbHeadingTocLevelChineseFontName.Text = dstFont.NameFarEast;
                }
                else
                {
                    cmbHeadingTocLevelChineseFontName.Text = m_strNoChange;
                }


                if (dstFont.isSet(ClassFont.euMembers.NameAscii))
                {
                    cmbHeadingTocLevelAsciiFontName.Text = dstFont.NameAscii;
                }
                else
                {
                    cmbHeadingTocLevelAsciiFontName.Text = m_strNoChange;
                }

                if (dstFont.isSet(ClassFont.euMembers.Size))
                {
                    cmbHeadingTocLevelFontSize.Text = "" + dstFont.Size;
                }
                else
                {
                    cmbHeadingTocLevelFontSize.Text = m_strNoChange;
                }

                if (dstFont.isSet(ClassFont.euMembers.Bold))
                {
                    if (dstFont.Bold != 0)
                    {
                        chkHeadingTocLevelFontBold.CheckState = CheckState.Checked;
                    }
                    else
                    {
                        chkHeadingTocLevelFontBold.CheckState = CheckState.Unchecked;
                    }
                }
                else
                {
                    chkHeadingTocLevelFontBold.CheckState = CheckState.Indeterminate;
                }


                if (dstFont.isSet(ClassFont.euMembers.Italic))
                {
                    if (dstFont.Italic != 0)
                    {
                        chkHeadingTocLevelFontItalic.CheckState = CheckState.Checked;
                    }
                    else
                    {
                        chkHeadingTocLevelFontItalic.CheckState = CheckState.Unchecked;
                    }
                }
                else
                {
                    chkHeadingTocLevelFontItalic.CheckState = CheckState.Indeterminate;
                }
            }
            

            return;
        }

        private void Data2UI_ParaFmt_headingTOC(ClassParagraphFormat[] oParaFmt, int nIndex)
        {
            ClassParagraphFormat dstParaFmt = oParaFmt[nIndex];

            if (nIndex == 0)
            {
                if (!dstParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
                {
                    cmbHeadingTOCTotalLineSpaceRule.Text = m_strNoChange;
                    cmbHeadingTOCTotalLineSpaceUnit.Text = "磅";
                }
                else
                {
                    switch (dstParaFmt.LineSpacingRule)
                    {
                        case Word.WdLineSpacing.wdLineSpaceSingle:
                            cmbHeadingTOCTotalLineSpaceRule.Text = "单倍行距";
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        case Word.WdLineSpacing.wdLineSpace1pt5:
                            cmbHeadingTOCTotalLineSpaceRule.Text = "1.5 倍行距";
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceDouble:
                            cmbHeadingTOCTotalLineSpaceRule.Text = "2 倍行距";
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceAtLeast:
                            cmbHeadingTOCTotalLineSpaceRule.Text = "最小值";

                            numHeadingTOCTotalLineSpaceValue.Value = (decimal)dstParaFmt.LineSpacing;
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "磅";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceExactly:
                            cmbHeadingTOCTotalLineSpaceRule.Text = "固定值";
                            numHeadingTOCTotalLineSpaceValue.Value = (decimal)dstParaFmt.LineSpacing;
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "磅";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceMultiple:
                            cmbHeadingTOCTotalLineSpaceRule.Text = "多倍行距";
                            numHeadingTOCTotalLineSpaceValue.Value = (decimal)app.PointsToLines(dstParaFmt.LineSpacing);
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        default:
                            break;
                    }
                }
            }
            else
            {
                if (!dstParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
                {
                    cmbHeadingTocLevelParaLineSpaceRule.Text = m_strNoChange;
                    cmbHeadingTOCTotalLineSpaceUnit.Text = "磅";
                }
                else
                {
                    switch (dstParaFmt.LineSpacingRule)
                    {
                        case Word.WdLineSpacing.wdLineSpaceSingle:
                            cmbHeadingTocLevelParaLineSpaceRule.Text = "单倍行距";
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        case Word.WdLineSpacing.wdLineSpace1pt5:
                            cmbHeadingTocLevelParaLineSpaceRule.Text = "1.5 倍行距";
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceDouble:
                            cmbHeadingTocLevelParaLineSpaceRule.Text = "2 倍行距";
                            cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceAtLeast:
                            cmbHeadingTocLevelParaLineSpaceRule.Text = "最小值";

                            numHeadingTocLevelParaLineSpaceValue.Value = (decimal)dstParaFmt.LineSpacing;
                            cmbHeadingTocLevelLineSpaceUnit.Text = "磅";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceExactly:
                            cmbHeadingTocLevelParaLineSpaceRule.Text = "固定值";
                            numHeadingTocLevelParaLineSpaceValue.Value = (decimal)dstParaFmt.LineSpacing;
                            cmbHeadingTocLevelLineSpaceUnit.Text = "磅";
                            break;

                        case Word.WdLineSpacing.wdLineSpaceMultiple:
                            cmbHeadingTocLevelParaLineSpaceRule.Text = "多倍行距";
                            numHeadingTocLevelParaLineSpaceValue.Value = (decimal)app.PointsToLines(dstParaFmt.LineSpacing);
                            cmbHeadingTocLevelLineSpaceUnit.Text = "行";
                            break;

                        default:
                            break;
                    }
                }
            }


            return;
        }


        private void lstHeadingLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            bStopUI2Data = true;

            nCurHeadingTocLevel = lstHeadingLevel.SelectedIndex + 1;

            Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt,nCurHeadingTocLevel);
            Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt, nCurHeadingTocLevel);

            bStopUI2Data = false;

            return;
        }

        private void rdHeadingTocTotal_CheckedChanged(object sender, EventArgs e)
        {
            grpHeadingTocTotal.Enabled = rdHeadingTocTotal.Checked;

            if (rdHeadingTocTotal.Checked)
            {
                Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt, 0);
                Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt, 0);
            }

            return;
        }


        private void rdHeadingTocLevel_CheckedChanged(object sender, EventArgs e)
        {
            grpHeadingTocLevel.Enabled = rdHeadingTocLevel.Checked;

            if (rdHeadingTocLevel.Checked)
            {
                Data2UI_Font_headingTOC(inWholeDocType.arrsHeadingTocFnt, nCurHeadingTocLevel);
                Data2UI_ParaFmt_headingTOC(inWholeDocType.arrsHeadingTocParaFmt, nCurHeadingTocLevel);

                lstHeadingLevel.SelectedIndex = nCurHeadingTocLevel - 1;
            }

            return;
        }


        private void cmbHeadingTOCTotalFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }
            return;
        }


        private void cmbHeadingTOCTotalChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }
            return;
        }


        private void cmbHeadingTOCTotalAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }
            return;
        }

        private void chkHeadingTOCTotalFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }
            return;
        }

        private void chkHeadingTOCTotalFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }
            return;
        }

        private void cmbHeadingTOCTotalLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmbHeadingTOCTotalLineSpaceRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                numHeadingTOCTotalLineSpaceValue.Enabled = true;
                cmbHeadingTOCTotalLineSpaceUnit.Enabled = true;
                cmbHeadingTOCTotalLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                numHeadingTOCTotalLineSpaceValue.Enabled = true;
                cmbHeadingTOCTotalLineSpaceUnit.Enabled = false;
                cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
            }
            else
            {
                numHeadingTOCTotalLineSpaceValue.Enabled = false;
                cmbHeadingTOCTotalLineSpaceUnit.Enabled = false;
                cmbHeadingTOCTotalLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_headingTOC(ref inWholeDocType.arrsHeadingTocParaFmt);
            return;
        }

        private void cmbHeadingTOCTotalLineSpaceUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_headingTOC(ref inWholeDocType.arrsHeadingTocParaFmt);
            }
            return;
        }

        private void numHeadingTOCTotalLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_headingTOC(ref inWholeDocType.arrsHeadingTocParaFmt);
            }

            return;
        }

        private void cmbHeadingTocLevelChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }

            return;
        }

        private void cmbHeadingTocLevelAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }

            return;
        }

        private void cmbHeadingTocLevelFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }

            return;
        }

        private void chkHeadingTocLevelFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }

            return;
        }

        private void chkHeadingTocLevelFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_headingTOC(ref inWholeDocType.arrsHeadingTocFnt);
            }

            return;
        }        

        private void cmbHeadingTocLevelParaLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmbHeadingTocLevelParaLineSpaceRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                numHeadingTocLevelParaLineSpaceValue.Enabled = true;
                cmbHeadingTocLevelLineSpaceUnit.Enabled = true;
                cmbHeadingTocLevelLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                numHeadingTocLevelParaLineSpaceValue.Enabled = true;
                cmbHeadingTocLevelLineSpaceUnit.Enabled = false;
                cmbHeadingTocLevelLineSpaceUnit.Text = "行";
            }
            else
            {
                numHeadingTocLevelParaLineSpaceValue.Enabled = false;
                cmbHeadingTocLevelLineSpaceUnit.Enabled = false;
                cmbHeadingTocLevelLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_headingTOC(ref inWholeDocType.arrsHeadingTocParaFmt);
            return;
        }

        private void cmbHeadingTocLevelLineSpaceUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_headingTOC(ref inWholeDocType.arrsHeadingTocParaFmt);
            }

            return;
        }

        private void numHeadingTocLevelParaLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_headingTOC(ref inWholeDocType.arrsHeadingTocParaFmt);
            }

            return;
        }

        // 
        // UI2Data
        private void UI2Data_TuWenToc()
        {
            return;
        }

        private void UI2Data_Font_TuWenToc(ref ClassFont oFont)
        {
            String strItem = cmbTuWenTocChineseFontName.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                oFont.NameFarEast = strItem;
            }

            strItem = cmbTuWenTocAsciiFontName.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                oFont.NameAscii = strItem;
            }

            String strFntSize = cmbTuWenTocFontSize.Text;

            if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    oFont.AddSelMember((int)ClassFont.euMembers.Size);
                    oFont.Size = fSize;
                }
                else
                {
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }


            switch (chkTuWenTocFontBold.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            switch (chkTuWenTocFontItalic.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }

            return;
        }

        private void UI2Data_ParaFmt_TuWenToc(ref ClassParagraphFormat oParaFmt)
        {
            String strItem = cmbTuWenTocParaLineSpaceRule.Text;
            String strUnit = cmbTuWenTocLineSpaceUnit.Text;
            float fValue = (float)numTuWenTocParaLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    oParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            return;
        }
		

        // Data2UI
        private void Data2UI_TuWenToc(ClassWholeDocType oWdt = null)
        {
            if (oWdt == null)
            {
                Data2UI_Font_TuWenToc();
                Data2UI_ParaFmt_TuWenToc();
            }
            else
            {
                Data2UI_Font_TuWenToc(oWdt.tuWenTocTotalFnt);
                Data2UI_ParaFmt_TuWenToc(oWdt.tuWenTocTotalParaFmt);
            }

            return;
        }

        private void Data2UI_Font_TuWenToc(ClassFont oFont = null)
        {
            if (oFont == null)
            {
                cmbTuWenTocChineseFontName.Text = m_strNoChange;
                cmbTuWenTocAsciiFontName.Text = m_strNoChange;
                cmbTuWenTocFontSize.Text = m_strNoChange;

                chkTuWenTocFontBold.CheckState = CheckState.Indeterminate;
                chkTuWenTocFontItalic.CheckState = CheckState.Indeterminate;

                return;
            }

            if (oFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmbTuWenTocChineseFontName.Text = oFont.NameFarEast;
            }
            else
            {
                cmbTuWenTocChineseFontName.Text = m_strNoChange;
            }


            if (oFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmbTuWenTocAsciiFontName.Text = oFont.NameAscii;
            }
            else
            {
                cmbTuWenTocAsciiFontName.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Size))
            {
                cmbTuWenTocFontSize.Text = "" + oFont.Size;
            }
            else
            {
                cmbTuWenTocFontSize.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Bold))
            {
                if (oFont.Bold != 0)
                {
                    chkTuWenTocFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTuWenTocFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTuWenTocFontBold.CheckState = CheckState.Indeterminate;
            }


            if (oFont.isSet(ClassFont.euMembers.Italic))
            {
                if (oFont.Italic != 0)
                {
                    chkTuWenTocFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTuWenTocFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTuWenTocFontItalic.CheckState = CheckState.Indeterminate;
            }

            return;
        }

        private void Data2UI_ParaFmt_TuWenToc(ClassParagraphFormat oParaFmt = null)
        {
            if (oParaFmt == null)
            {
                cmbTuWenTocParaLineSpaceRule.Text = m_strNoChange;

                numTuWenTocParaLineSpaceValue.Value = 0.0m;
                cmbTuWenTocLineSpaceUnit.Text = "磅";
                return;
            }

            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
            {
                cmbTuWenTocParaLineSpaceRule.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmbTuWenTocParaLineSpaceRule.Text = "单倍行距";
                        cmbTuWenTocLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmbTuWenTocParaLineSpaceRule.Text = "1.5 倍行距";
                        cmbTuWenTocLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmbTuWenTocParaLineSpaceRule.Text = "2 倍行距";
                        cmbTuWenTocLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmbTuWenTocParaLineSpaceRule.Text = "最小值";

                        numTuWenTocParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTuWenTocLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmbTuWenTocParaLineSpaceRule.Text = "固定值";
                        numTuWenTocParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTuWenTocLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmbTuWenTocParaLineSpaceRule.Text = "多倍行距";
                        numTuWenTocParaLineSpaceValue.Value = (decimal)app.PointsToLines(oParaFmt.LineSpacing);
                        cmbTuWenTocLineSpaceUnit.Text = "行";
                        break;

                    default:
                        break;
                }
            }

            return;
        }

        private void cmbTuWenTocChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TuWenToc(ref inWholeDocType.tuWenTocTotalFnt);
            }
            return;
        }

        private void cmbTuWenTocAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TuWenToc(ref inWholeDocType.tuWenTocTotalFnt);
            }
            return;
        }

        private void cmbTuWenTocFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TuWenToc(ref inWholeDocType.tuWenTocTotalFnt);
            }
            return;
        }

        private void chkTuWenTocFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TuWenToc(ref inWholeDocType.tuWenTocTotalFnt);
            }
            return;
        }

        private void chkTuWenTocFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TuWenToc(ref inWholeDocType.tuWenTocTotalFnt);
            }
            return;
        }

        private void numTuWenTocParaLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TuWenToc(ref inWholeDocType.tuWenTocTotalParaFmt);
            }
            return;
        }

        private void cmbTuWenTocLineSpaceUnit_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TuWenToc(ref inWholeDocType.tuWenTocTotalParaFmt);
            }
            return;
        }


        // UI2Data

        private void UI2Data_Table()
        {
            return;
        }

        private void UI2Data_Font_Table(ref ClassFont oFont)
        {
            String strItem = cmbTbleChineseFontName.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                oFont.NameFarEast = strItem;
            }

            strItem = cmbTbleAsciiFontName.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                oFont.NameAscii = strItem;
            }

            String strFntSize = cmbTbleFontSize.Text;

            if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    oFont.AddSelMember((int)ClassFont.euMembers.Size);
                    oFont.Size = fSize;
                }
                else
                {
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }


            switch (chkTbleFontBold.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            switch (chkTbleFontItalic.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }

            return;
        }
		
        private void UI2Data_ParaFmt_Table(ref ClassParagraphFormat oParaFmt)
        {
            String strItem = cmbTblParaLineSpaceRule.Text;
            String strUnit = cmbTableLineSpaceUnit.Text;
            float fValue = (float)nmTblParaLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    oParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            return;
        }
		

        // Data2UI
        private void Data2UI_Table(ClassWholeDocType oWdt = null)
        {
            if (oWdt == null)
            {
                Data2UI_Font_Table();
                Data2UI_ParaFmt_Table();
            }
            else
            {
                Data2UI_Font_Table(oWdt.tableTotalFont);
                Data2UI_ParaFmt_Table(oWdt.tableTotalParaFmt);
            }

            return;
        }

        private void Data2UI_Font_Table(ClassFont oFont = null)
        {
            if (oFont == null)
            {
                cmbTbleChineseFontName.Text = m_strNoChange;
                cmbTbleAsciiFontName.Text = m_strNoChange;
                cmbTbleFontSize.Text = m_strNoChange;

                chkTbleFontBold.CheckState = CheckState.Indeterminate;
                chkTbleFontItalic.CheckState = CheckState.Indeterminate;

                return;
            }

            if (oFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmbTbleChineseFontName.Text = oFont.NameFarEast;
            }
            else
            {
                cmbTbleChineseFontName.Text = m_strNoChange;
            }


            if (oFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmbTbleAsciiFontName.Text = oFont.NameAscii;
            }
            else
            {
                cmbTbleAsciiFontName.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Size))
            {
                cmbTbleFontSize.Text = "" + oFont.Size;
            }
            else
            {
                cmbTbleFontSize.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Bold))
            {
                if (oFont.Bold != 0)
                {
                    chkTbleFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTbleFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTbleFontBold.CheckState = CheckState.Indeterminate;
            }


            if (oFont.isSet(ClassFont.euMembers.Italic))
            {
                if (oFont.Italic != 0)
                {
                    chkTbleFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTbleFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTbleFontItalic.CheckState = CheckState.Indeterminate;
            }
            return;
        }

        private void Data2UI_ParaFmt_Table(ClassParagraphFormat oParaFmt = null)
        {
            if (oParaFmt == null)
            {
                cmbTblParaLineSpaceRule.Text = m_strNoChange;

                nmTblParaLineSpaceValue.Value = 0.0m;
                cmbTableLineSpaceUnit.Text = "磅";
                return;
            }

            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
            {
                cmbTblParaLineSpaceRule.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmbTblParaLineSpaceRule.Text = "单倍行距";
                        cmbTableLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmbTblParaLineSpaceRule.Text = "1.5 倍行距";
                        cmbTableLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmbTblParaLineSpaceRule.Text = "2 倍行距";
                        cmbTableLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmbTblParaLineSpaceRule.Text = "最小值";

                        nmTblParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTableLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmbTblParaLineSpaceRule.Text = "固定值";
                        nmTblParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTableLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmbTblParaLineSpaceRule.Text = "多倍行距";
                        nmTblParaLineSpaceValue.Value = (decimal)app.PointsToLines(oParaFmt.LineSpacing);
                        cmbTableLineSpaceUnit.Text = "行";
                        break;

                    default:
                        break;
                }
            }
            return;
        }


        private void chkTblClearIndent_CheckedChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                inWholeDocType.bClearIndent = chkTblClearIndent.Checked;
            }
            return;
        }

        private void cmbTbleChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Table(ref inWholeDocType.tableTotalFont);
            }
            return;
        }

        private void cmbTbleAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Table(ref inWholeDocType.tableTotalFont);
            }
            return;
        }

        private void cmbTbleFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Table(ref inWholeDocType.tableTotalFont);
            }
            return;
        }

        private void chkTbleFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Table(ref inWholeDocType.tableTotalFont);
            }
            return;
        }

        private void chkTbleFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Table(ref inWholeDocType.tableTotalFont);
            }
            return;
        }

        private void nmTblParaLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_Table(ref inWholeDocType.tableTotalParaFmt);
            }
            return;
        }

        private void cmbTableLineSpaceUnit_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_Table(ref inWholeDocType.tableTotalParaFmt);
            }
            return;
        }

        private void cmbTuWenTocParaLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmbTuWenTocParaLineSpaceRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                numTuWenTocParaLineSpaceValue.Enabled = true;
                cmbTuWenTocLineSpaceUnit.Enabled = true;
                cmbTuWenTocLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                numTuWenTocParaLineSpaceValue.Enabled = true;
                cmbTuWenTocLineSpaceUnit.Enabled = false;
                cmbTuWenTocLineSpaceUnit.Text = "行";
            }
            else
            {
                numTuWenTocParaLineSpaceValue.Enabled = false;
                cmbTuWenTocLineSpaceUnit.Enabled = false;
                cmbTuWenTocLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_TuWenToc(ref inWholeDocType.tuWenTocTotalParaFmt);
            return;
        }

        private void cmbTblParaLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmbTblParaLineSpaceRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                nmTblParaLineSpaceValue.Enabled = true;
                cmbTableLineSpaceUnit.Enabled = true;
                cmbTableLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                nmTblParaLineSpaceValue.Enabled = true;
                cmbTableLineSpaceUnit.Enabled = false;
                cmbTableLineSpaceUnit.Text = "行";
            }
            else
            {
                nmTblParaLineSpaceValue.Enabled = false;
                cmbTableLineSpaceUnit.Enabled = false;
                cmbTableLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_Table(ref inWholeDocType.tableTotalParaFmt);
            return;
        }

        private void Data2UI_TiZhu(ClassWholeDocType oWdt = null)
        {
            if (oWdt == null)
            {
                Data2UI_Font_TiZhu(null);
                Data2UI_ParaFmt_TiZhu(null);
            }
            else
            {
                Data2UI_Font_TiZhu(oWdt.tableTotalFont);
                Data2UI_ParaFmt_TiZhu(oWdt.tableTotalParaFmt);
            }

            return;
        }

        private void Data2UI_Font_TiZhu(ClassFont oFont = null)
        {
            if (oFont == null)
            {
                cmbTiZhuChineseFontName.Text = m_strNoChange;
                cmbTiZhuAsciiFontName.Text = m_strNoChange;
                cmbTiZhuFontSize.Text = m_strNoChange;

                chkTiZhuFontBold.CheckState = CheckState.Indeterminate;
                chkTiZhuFontItalic.CheckState = CheckState.Indeterminate;

                return;
            }

            if (oFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmbTiZhuChineseFontName.Text = oFont.NameFarEast;
            }
            else
            {
                cmbTiZhuChineseFontName.Text = m_strNoChange;
            }


            if (oFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmbTiZhuAsciiFontName.Text = oFont.NameAscii;
            }
            else
            {
                cmbTiZhuAsciiFontName.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Size))
            {
                cmbTiZhuFontSize.Text = "" + oFont.Size;
            }
            else
            {
                cmbTiZhuFontSize.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Bold))
            {
                if (oFont.Bold != 0)
                {
                    chkTiZhuFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTiZhuFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTiZhuFontBold.CheckState = CheckState.Indeterminate;
            }


            if (oFont.isSet(ClassFont.euMembers.Italic))
            {
                if (oFont.Italic != 0)
                {
                    chkTiZhuFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTiZhuFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTiZhuFontItalic.CheckState = CheckState.Indeterminate;
            }

            return;
        }

        private void Data2UI_ParaFmt_TiZhu(ClassParagraphFormat oParaFmt = null)
        {
            if (oParaFmt == null)
            {
                cmbTiZhuParaAlignStyle.Text = m_strNoChange;
                cmbTiZhuParaLineSpaceRule.Text = m_strNoChange;

                numTiZhuParaLineSpaceValue.Value = 0.0m;
                cmbTiZhuLineSpaceUnit.Text = "磅";
                return;
            }


            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.Alignment))
            {
                cmbTiZhuParaAlignStyle.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.Alignment)
                {
                    case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbTiZhuParaAlignStyle.Text = "居左";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphRight:
                        cmbTiZhuParaAlignStyle.Text = "居右";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbTiZhuParaAlignStyle.Text = "居中";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                        cmbTiZhuParaAlignStyle.Text = "分散对齐";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbTiZhuParaAlignStyle.Text = "两端对齐";
                        break;

                    default:
                        break;
                }
            }

            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
            {
                cmbTiZhuParaLineSpaceRule.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmbTiZhuParaLineSpaceRule.Text = "单倍行距";
                        cmbTiZhuLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmbTiZhuParaLineSpaceRule.Text = "1.5 倍行距";
                        cmbTiZhuLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmbTiZhuParaLineSpaceRule.Text = "2 倍行距";
                        cmbTiZhuLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmbTiZhuParaLineSpaceRule.Text = "最小值";

                        numTiZhuParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTiZhuLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmbTiZhuParaLineSpaceRule.Text = "固定值";
                        numTiZhuParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTiZhuLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmbTiZhuParaLineSpaceRule.Text = "多倍行距";
                        numTiZhuParaLineSpaceValue.Value = (decimal)app.PointsToLines(oParaFmt.LineSpacing);
                        cmbTiZhuLineSpaceUnit.Text = "行";
                        break;

                    default:
                        break;
                }
            }

            return;
        }



        // 
        private void UI2Data_TiZhu()
        {
            return;
        }

        private void UI2Data_Font_TiZhu(ref ClassFont oFont)
        {
            String strItem = cmbTiZhuChineseFontName.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                oFont.NameFarEast = strItem;
            }

            strItem = cmbTiZhuAsciiFontName.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                oFont.NameAscii = strItem;
            }

            String strFntSize = cmbTiZhuFontSize.Text;

            if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    oFont.AddSelMember((int)ClassFont.euMembers.Size);
                    oFont.Size = fSize;
                }
                else
                {
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }


            switch (chkTiZhuFontBold.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            switch (chkTiZhuFontItalic.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }

            return;
        }


        private void UI2Data_ParaFmt_TiZhu(ref ClassParagraphFormat oParaFmt)
        {
            String strItem = cmbTiZhuParaAlignStyle.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.Alignment);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.Alignment);
                if (strItem.Equals("居左"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                else if (strItem.Equals("居中"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (strItem.Equals("居右"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }
                else if (strItem.Equals("两端对齐"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (strItem.Equals("分散对齐"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                }

            }

            strItem = cmbTiZhuParaLineSpaceRule.Text;
            String strUnit = cmbTiZhuLineSpaceUnit.Text;
            float fValue = (float)numTiZhuParaLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    oParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }
            return;
        }


        private void cmbTiZhuParaLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmbTiZhuParaLineSpaceRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                numTiZhuParaLineSpaceValue.Enabled = true;
                cmbTiZhuLineSpaceUnit.Enabled = true;
                cmbTiZhuLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                numTiZhuParaLineSpaceValue.Enabled = true;
                cmbTiZhuLineSpaceUnit.Enabled = false;
                cmbTiZhuLineSpaceUnit.Text = "行";
            }
            else
            {
                numTiZhuParaLineSpaceValue.Enabled = false;
                cmbTiZhuLineSpaceUnit.Enabled = false;
                cmbTiZhuLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_TiZhu(ref inWholeDocType.tizhuParaFmt);
            return;
        }

        private void cmbTiZhuChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TiZhu(ref inWholeDocType.tizhuFont);
            }
            return;
        }

        private void cmbTiZhuAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TiZhu(ref inWholeDocType.tizhuFont);
            }
            return;
        }

        private void cmbTiZhuFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TiZhu(ref inWholeDocType.tizhuFont);
            }
            return;
        }

        private void chkTiZhuFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TiZhu(ref inWholeDocType.tizhuFont);
            }
            return;
        }

        private void chkTiZhuFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TiZhu(ref inWholeDocType.tizhuFont);
            }
            return;
        }

        private void cmbTiZhuParaAlignStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TiZhu(ref inWholeDocType.tizhuParaFmt);
            }
            return;
        }

        private void numTiZhuParaLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TiZhu(ref inWholeDocType.tizhuParaFmt);
            }
            return;
        }

        private void cmbTiZhuLineSpaceUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TiZhu(ref inWholeDocType.tizhuParaFmt);
            }
            return;
        }


        private void Data2UI_TextBodyZone(ClassWholeDocType oWdt = null)
        {
            if (oWdt == null)
            {
                Data2UI_Font_TextBodyZone(null);
                Data2UI_ParaFmt_TextBodyZone(null);

                return;
            }
            else
            {
                Data2UI_Font_TextBodyZone(inWholeDocType.textbodyZoneFont);
                Data2UI_ParaFmt_TextBodyZone(inWholeDocType.textbodyZoneParaFmt);
            }
            return;
        }

        private void Data2UI_Font_TextBodyZone(ClassFont oFont = null)
        {
            if (oFont == null)
            {
                cmbTextBodyZoneChineseFontName.Text = m_strNoChange;
                cmbTextBodyZoneAsciiFontName.Text = m_strNoChange;
                cmbTextBodyZoneFontSize.Text = m_strNoChange;

                chkTextBodyZoneFontBold.CheckState = CheckState.Indeterminate;
                chkTextBodyZoneFontItalic.CheckState = CheckState.Indeterminate;

                return;
            }

            if (oFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmbTextBodyZoneChineseFontName.Text = oFont.NameFarEast;
            }
            else
            {
                cmbTextBodyZoneChineseFontName.Text = m_strNoChange;
            }


            if (oFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmbTextBodyZoneAsciiFontName.Text = oFont.NameAscii;
            }
            else
            {
                cmbTextBodyZoneAsciiFontName.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Size))
            {
                cmbTextBodyZoneFontSize.Text = "" + oFont.Size;
            }
            else
            {
                cmbTextBodyZoneFontSize.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Bold))
            {
                if (oFont.Bold != 0)
                {
                    chkTextBodyZoneFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTextBodyZoneFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTextBodyZoneFontBold.CheckState = CheckState.Indeterminate;
            }


            if (oFont.isSet(ClassFont.euMembers.Italic))
            {
                if (oFont.Italic != 0)
                {
                    chkTextBodyZoneFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chkTextBodyZoneFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkTextBodyZoneFontItalic.CheckState = CheckState.Indeterminate;
            }

            return;
        }

        private void Data2UI_ParaFmt_TextBodyZone(ClassParagraphFormat oParaFmt = null)
        {
            if (oParaFmt == null)
            {
                cmbTextBodyZoneParaLineSpaceRule.Text = m_strNoChange;

                numTextBodyZoneParaLineSpaceValue.Value = 0.0m;
                cmbTextBodyZoneLineSpaceUnit.Text = "磅";
                return;
            }

            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
            {
                cmbTextBodyZoneParaLineSpaceRule.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmbTextBodyZoneParaLineSpaceRule.Text = "单倍行距";
                        cmbTextBodyZoneLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmbTextBodyZoneParaLineSpaceRule.Text = "1.5 倍行距";
                        cmbTextBodyZoneLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmbTextBodyZoneParaLineSpaceRule.Text = "2 倍行距";
                        cmbTextBodyZoneLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmbTextBodyZoneParaLineSpaceRule.Text = "最小值";

                        numTextBodyZoneParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTextBodyZoneLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmbTextBodyZoneParaLineSpaceRule.Text = "固定值";
                        numTextBodyZoneParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbTextBodyZoneLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmbTextBodyZoneParaLineSpaceRule.Text = "多倍行距";
                        numTextBodyZoneParaLineSpaceValue.Value = (decimal)app.PointsToLines(oParaFmt.LineSpacing);
                        cmbTextBodyZoneLineSpaceUnit.Text = "行";
                        break;

                    default:
                        break;
                }
            }

            return;
        }



        // 
        private void UI2Data_TextBodyZone()
        {
            return;
        }

        private void UI2Data_Font_TextBodyZone(ref ClassFont oFont)
        {
            String strItem = cmbTextBodyZoneChineseFontName.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                oFont.NameFarEast = strItem;
            }

            strItem = cmbTextBodyZoneAsciiFontName.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                oFont.NameAscii = strItem;
            }

            String strFntSize = cmbTextBodyZoneFontSize.Text;

            if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    oFont.AddSelMember((int)ClassFont.euMembers.Size);
                    oFont.Size = fSize;
                }
                else
                {
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }


            switch (chkTextBodyZoneFontBold.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            switch (chkTextBodyZoneFontItalic.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }

            return;
        }


        private void UI2Data_ParaFmt_TextBodyZone(ref ClassParagraphFormat oParaFmt)
        {
            String strItem = cmbTextBodyZoneParaLineSpaceRule.Text;
            String strUnit = cmbTextBodyZoneLineSpaceUnit.Text;
            float fValue = (float)numTextBodyZoneParaLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    oParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            return;
        }

        private void cmbTextBodyZoneChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TextBodyZone(ref inWholeDocType.textbodyZoneFont);
            }
            return;
        }

        private void cmbTextBodyZoneAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TextBodyZone(ref inWholeDocType.textbodyZoneFont);
            }
            return;
        }

        private void cmbTextBodyZoneFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TextBodyZone(ref inWholeDocType.textbodyZoneFont);
            }
            return;
        }

        private void chkTextBodyZoneFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TextBodyZone(ref inWholeDocType.textbodyZoneFont);
            }
            return;
        }

        private void chkTextBodyZoneFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_TextBodyZone(ref inWholeDocType.textbodyZoneFont);
            }
            return;
        }

        private void cmbTextBodyZoneParaLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bStopUI2Data)
            {
                return;
            }

            String strText = cmbTextBodyZoneParaLineSpaceRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                numTextBodyZoneParaLineSpaceValue.Enabled = true;
                cmbTextBodyZoneLineSpaceUnit.Enabled = true;
                cmbTextBodyZoneLineSpaceUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                numTextBodyZoneParaLineSpaceValue.Enabled = true;
                cmbTextBodyZoneLineSpaceUnit.Enabled = false;
                cmbTextBodyZoneLineSpaceUnit.Text = "行";
            }
            else
            {
                numTextBodyZoneParaLineSpaceValue.Enabled = false;
                cmbTextBodyZoneLineSpaceUnit.Enabled = false;
                cmbTextBodyZoneLineSpaceUnit.Text = "行";
            }

            UI2Data_ParaFmt_TextBodyZone(ref inWholeDocType.textbodyZoneParaFmt);
            return;
        }

        private void numTextBodyZoneParaLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TextBodyZone(ref inWholeDocType.textbodyZoneParaFmt);
            }
            return;
        }

        private void cmbTextBodyZoneLineSpaceUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_TextBodyZone(ref inWholeDocType.textbodyZoneParaFmt);
            }
            return;
        }


        /// <summary>
        /// //////////////
        /// </summary>
        private void Data2UI_Section(ClassWholeDocType oWdt = null)
        {
            cmbSections.Text = "第1节";

            if(oWdt == null)
            {
                Data2UI_Font_Section(null);
                Data2UI_ParaFmt_Section(null);

                return;
            }

            ////////
            String strItem = "";
            foreach (DictionaryEntry ent in oWdt.hshSectionPart)
            {
                strItem = (String)ent.Key;

                lstSectionsChosen.Items.Add(strItem);
            }

            if (lstSectionsChosen.Items.Count > 0)
            {
                lstSectionsChosen.SelectedIndex = 0;
            }
            else
            {
                Data2UI_Font_Section(null);
                Data2UI_ParaFmt_Section(null);
            }

            return;
        }


        private void Data2UI_Section(ClassWholeDocType.sepPart sep = null)
        {
            if (sep == null)
            {
                chkSectionHeader.Checked = false;
                chkSectionFooter.Checked = false;
                return;
            }

            chkSectionHeader.Checked = sep.bHeader;
            chkSectionFooter.Checked = sep.bFooter;

            return;
        }


        private void Data2UI_Font_Section(ClassFont oFont = null)
        {
            if (oFont == null)
            {
                cmbSectionChineseFontName.Text = m_strNoChange;
                cmbSectionAsciiFontName.Text = m_strNoChange;
                cmbSectionFontSize.Text = m_strNoChange;

                chkSectionFontBold.CheckState = CheckState.Indeterminate;
                chkSectionFontItalic.CheckState = CheckState.Indeterminate;

                return;
            }

            if (oFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmbSectionChineseFontName.Text = oFont.NameFarEast;
            }
            else
            {
                cmbSectionChineseFontName.Text = m_strNoChange;
            }


            if (oFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmbSectionAsciiFontName.Text = oFont.NameAscii;
            }
            else
            {
                cmbSectionAsciiFontName.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Size))
            {
                cmbSectionFontSize.Text = "" + oFont.Size;
            }
            else
            {
                cmbSectionFontSize.Text = m_strNoChange;
            }

            if (oFont.isSet(ClassFont.euMembers.Bold))
            {
                if (oFont.Bold != 0)
                {
                    chkSectionFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chkSectionFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkSectionFontBold.CheckState = CheckState.Indeterminate;
            }


            if (oFont.isSet(ClassFont.euMembers.Italic))
            {
                if (oFont.Italic != 0)
                {
                    chkSectionFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chkSectionFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkSectionFontItalic.CheckState = CheckState.Indeterminate;
            }

            return;
        }

        private void Data2UI_ParaFmt_Section(ClassParagraphFormat oParaFmt = null)
        {
            if (oParaFmt == null)
            {
                cmbSectionParaAlignStyle.Text = m_strNoChange;
                cmbSectionParaLineSpaceRule.Text = m_strNoChange;

                numSectionParaLineSpaceValue.Value = 0.0m;
                cmbSectionLineSpaceUnit.Text = "磅";
                return;
            }


            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.Alignment))
            {
                cmbSectionParaAlignStyle.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.Alignment)
                {
                    case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbSectionParaAlignStyle.Text = "居左";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphRight:
                        cmbSectionParaAlignStyle.Text = "居右";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbSectionParaAlignStyle.Text = "居中";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                        cmbSectionParaAlignStyle.Text = "分散对齐";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbSectionParaAlignStyle.Text = "两端对齐";
                        break;

                    default:
                        break;
                }
            }

            if (!oParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule))
            {
                cmbSectionParaLineSpaceRule.Text = m_strNoChange;
            }
            else
            {
                switch (oParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmbSectionParaLineSpaceRule.Text = "单倍行距";
                        cmbSectionLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmbSectionParaLineSpaceRule.Text = "1.5 倍行距";
                        cmbSectionLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmbSectionParaLineSpaceRule.Text = "2 倍行距";
                        cmbSectionLineSpaceUnit.Text = "行";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmbSectionParaLineSpaceRule.Text = "最小值";

                        numSectionParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbSectionLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmbSectionParaLineSpaceRule.Text = "固定值";
                        numSectionParaLineSpaceValue.Value = (decimal)oParaFmt.LineSpacing;
                        cmbSectionLineSpaceUnit.Text = "磅";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmbSectionParaLineSpaceRule.Text = "多倍行距";
                        numSectionParaLineSpaceValue.Value = (decimal)app.PointsToLines(oParaFmt.LineSpacing);
                        cmbSectionLineSpaceUnit.Text = "行";
                        break;

                    default:
                        break;
                }
            }

            return;
        }

        // 
        private void UI2Data_Section(ref ClassWholeDocType oWdt)
        {
            return;
        }

        private void UI2Data_Section(ref ClassWholeDocType.sepPart sep)
        {
            sep.bHeader = chkSectionHeader.Checked;
            sep.bFooter = chkSectionFooter.Checked;

            return;
        }


        private void UI2Data_Font_Section(ref ClassFont oFont)
        {
            String strItem = cmbSectionChineseFontName.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                oFont.NameFarEast = strItem;
            }

            strItem = cmbSectionAsciiFontName.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                oFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                oFont.NameAscii = strItem;
            }

            String strFntSize = cmbSectionFontSize.Text;

            if (String.IsNullOrWhiteSpace(strFntSize) || strFntSize.Equals(m_strNoChange))
            {
                oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strFntSize);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    oFont.AddSelMember((int)ClassFont.euMembers.Size);
                    oFont.Size = fSize;
                }
                else
                {
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }


            switch (chkSectionFontBold.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    oFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            switch (chkSectionFontItalic.CheckState)
            {
                case CheckState.Checked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    oFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    oFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                    oFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }


            return;
        }

        private void UI2Data_ParaFmt_Section(ref ClassParagraphFormat oParaFmt)
        {
            String strItem = cmbSectionParaAlignStyle.Text;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.Alignment);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.Alignment);
                if (strItem.Equals("居左"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                else if (strItem.Equals("居中"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (strItem.Equals("居右"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }
                else if (strItem.Equals("两端对齐"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                }
                else if (strItem.Equals("分散对齐"))
                {
                    oParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                }

            }

            strItem = cmbSectionParaLineSpaceRule.Text;
            String strUnit = cmbSectionLineSpaceUnit.Text;
            float fValue = (float)numSectionParaLineSpaceValue.Value;

            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                oParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
            }
            else
            {
                oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);

                if (strItem.Equals("单倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("固定值"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    oParaFmt.LineSpacing = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                }
                else if (strItem.Equals("多倍行距"))
                {
                    oParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    oParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    oParaFmt.LineSpacing = app.LinesToPoints(fLines);
                }
            }

            return;
        }

        private void btnSectionAdd_Click(object sender, EventArgs e)
        {
            String strText = cmbSections.Text;
            // 判重
            // 

            if (String.IsNullOrWhiteSpace(strText))
            {
                MessageBox.Show("不能添加空选择");
                return;
            }

            if (lstSectionsChosen.Items.Contains(strText))
            {
                MessageBox.Show("不能添加重复项"); // lstChosenParts
                return;
            }

            Boolean bHeader = chkSectionHeader.Checked;
            Boolean bFooter = chkSectionFooter.Checked;

            ClassWholeDocType.sepPart sep = new ClassWholeDocType.sepPart();

            sep.bHeader = bHeader;
            sep.bFooter = bFooter;

            curSepPartSection.cFont.SelCopy2(sep.cFont);
            curSepPartSection.cParaFmt.SelCopy2(sep.cParaFmt);

            inWholeDocType.hshSectionPart.Add(strText, sep);

            curSepPartSection = sep;

            lstSectionsChosen.Items.Add(strText);
            lstSectionsChosen.SelectedItem = strText;

            return;
        }

        private void btnSectionDel_Click(object sender, EventArgs e)
        {
            int nSelIndex = lstSectionsChosen.SelectedIndex;
            int nCount = lstSectionsChosen.Items.Count;

            if (nCount == 0 || nSelIndex == -1)
            {
                MessageBox.Show("请选中一项后进行此操作");
            }

            DialogResult res = MessageBox.Show("确认删除选中项？", "确认", MessageBoxButtons.YesNo);

            if (res == DialogResult.No)
            {
                return;
            }

            int nNextIndex = -1;

            if (nCount == 1)
            {
                nNextIndex = -1; // 
            }
            else if (nSelIndex == nCount - 1)
            {
                nNextIndex = nCount - 2;
            }
            else
            {
                nNextIndex = nSelIndex; // 
            }

            String strItem = (String)lstSectionsChosen.Items[lstSectionsChosen.SelectedIndex];

            lstSectionsChosen.Items.RemoveAt(lstSectionsChosen.SelectedIndex);

            inWholeDocType.hshSectionPart.Remove(strItem);

            if (lstSectionsChosen.Items.Count == 0)
            {
                curSepPartSection = defaultSepPartSection;
            }
            else if (nNextIndex != -1)
            {
                lstSectionsChosen.SelectedIndex = nNextIndex;
                strItem = (String)lstSectionsChosen.Items[nNextIndex];
                // lstChosenParts.SelectedItem = strItem;

                ClassWholeDocType.sepPart sep = (ClassWholeDocType.sepPart)inWholeDocType.hshSectionPart[strItem];

                if (sep != null)
                {
                    curSepPartSection = sep;
                }
            }

            Data2UI_Section(curSepPartSection);
            Data2UI_Font_Section(curSepPartSection.cFont);
            Data2UI_ParaFmt_Section(curSepPartSection.cParaFmt);
            return;
        }

        private void btnSectionUpdate_Click(object sender, EventArgs e)
        {
            int nSelIndex = lstSectionsChosen.SelectedIndex;
            int nCount = lstSectionsChosen.Items.Count;

            if (lstSectionsChosen.Items.Count == 0 || lstSectionsChosen.SelectedIndex == -1)
            {
                MessageBox.Show("请选中一项后进行此操作");
            }

            String strText = cmbSections.Text;

            // 判重
            // 
            if (lstSectionsChosen.Items.Contains(strText))
            {
                MessageBox.Show("不能添加重复项"); // lstChosenParts
                return;
            }

            String strOldItem = (String)lstSectionsChosen.Items[lstSectionsChosen.SelectedIndex];

            inWholeDocType.hshSectionPart.Remove(strOldItem);

            lstSectionsChosen.Items.RemoveAt(lstSectionsChosen.SelectedIndex);

            int nIndex = lstSectionsChosen.Items.Add(strText);

            // get current font / paragraph format

            inWholeDocType.hshSectionPart.Add(strText, curSepPartSection);

            lstSectionsChosen.SelectedIndex = nIndex;

            // update data
            Data2UI_Section(curSepPartSection);
            Data2UI_Font_Section(curSepPartSection.cFont);
            Data2UI_ParaFmt_Section(curSepPartSection.cParaFmt);
            // 
            return;
        }

        private void lstSectionsChosen_SelectedIndexChanged(object sender, EventArgs e)
        {
            int nSelIndex = lstSectionsChosen.SelectedIndex;

            if (nSelIndex == -1)
            {
                return;
            }

            ClassWholeDocType.sepPart sep = null;

            String strItem = (String)lstSectionsChosen.Items[nSelIndex];

            sep = (ClassWholeDocType.sepPart)inWholeDocType.hshSectionPart[strItem];

            bStopUI2Data = true;
            // load current font and paragraph
            if (sep != null)
            {
                curSepPartSection = sep;

                Data2UI_Section(curSepPartSection);
                Data2UI_Font_Section(curSepPartSection.cFont);
                Data2UI_ParaFmt_Section(curSepPartSection.cParaFmt);
            }

            bStopUI2Data = false;

            return;
        }

        private void chkSectionHeader_CheckedChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Section(ref curSepPartSection);
            }
            return;
        }

        private void chkSectionFooter_CheckedChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Section(ref curSepPartSection);
            }
            return;
        }

        private void cmbSectionChineseFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Section(ref curSepPartSection.cFont);
            }
            return;
        }

        private void cmbSectionAsciiFontName_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Section(ref curSepPartSection.cFont);
            }
            return;
        }

        private void cmbSectionFontSize_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Section(ref curSepPartSection.cFont);
            }
            return;
        }

        private void chkSectionFontBold_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Section(ref curSepPartSection.cFont);
            }
            return;
        }

        private void chkSectionFontItalic_CheckStateChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_Font_Section(ref curSepPartSection.cFont);
            }
            return;
        }

        private void cmbSectionParaAlignStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_Section(ref curSepPartSection.cParaFmt);
            }
            return;
        }

        private void cmbSectionParaLineSpaceRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_Section(ref curSepPartSection.cParaFmt);
            }
            return;
        }

        private void numSectionParaLineSpaceValue_ValueChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_Section(ref curSepPartSection.cParaFmt);
            }
            return;
        }

        private void cmbSectionLineSpaceUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                UI2Data_ParaFmt_Section(ref curSepPartSection.cParaFmt);
            }
            return;
        }

        private void txtChosenHeadingStyleScheme_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                inWholeDocType.headingStyleSchemeName = txtChosenHeadingStyleScheme.Text;
            }
            return;
        }

        private void txtChosenHeadingSnScheme_TextChanged(object sender, EventArgs e)
        {
            if (!bStopUI2Data)
            {
                inWholeDocType.headingSnSchemeName = txtChosenHeadingSnScheme.Text;
            }
            return;
        }


    } // class

} // namespace
