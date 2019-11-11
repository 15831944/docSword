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
    public partial class frmParagraphFormat : Form
    {
        Word.Application app = null;
        Word.Document doc = null;
        ClassParagraphFormat inParaFmt = new ClassParagraphFormat();
        String m_strNoChange = "(保持原值不变)";


        public void setFeatures(Word.Application oApp, ClassParagraphFormat oParaFmt = null, String strNoChange = "(保持原值不变)")
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

            m_strNoChange = strNoChange;

            if (oParaFmt != null)
            {
                oParaFmt.SelCopy2(inParaFmt);
                // inFont.clone(oFont);
            }

            return;
        }


        public ClassParagraphFormat getParaFmt()
        {
            return inParaFmt;
        }


        public frmParagraphFormat()
        {
            InitializeComponent();
        }


        private void frmParagraphFormat_Load(object sender, EventArgs e)
        {
            // 

            foreach (String strAlignStyleItem in Globals.ThisAddIn.m_arrStrAlignStyle)
            {
                cmbAlignStyle.Items.Add(strAlignStyleItem);
            }

            foreach (String strUnitItem in Globals.ThisAddIn.m_arrSpaceUnit)
            {
                cmbIndentLeftUnit.Items.Add(strUnitItem);
                cmbIndentRightUnit.Items.Add(strUnitItem);
                cmbIndentSpecialUnit.Items.Add(strUnitItem);

                cmbBeforeParaSpacingUnit.Items.Add(strUnitItem);
                cmbAfterParaSpacingUnit.Items.Add(strUnitItem);
                cmbLineSpacingUnit.Items.Add(strUnitItem);
            }

            String[] strsSpecialIndent = { "(无)", "首行缩进","悬挂缩进"};

            foreach (String strItem in strsSpecialIndent)
            {
                cmbIndentSpecial.Items.Add(strItem);
            }

            foreach (String strItem in Globals.ThisAddIn.m_arrParaLineSpaceRule)
            {
                cmbLineSpacingRule.Items.Add(strItem);
            }

            String[] strsTextAlignStyle = {"顶端对齐","居中", "基线对齐","底端对齐","自动设置"};
            foreach (String strItem in strsTextAlignStyle)
            {
                cmbTextAlignStyle.Items.Add(strItem);
            }

            cmbTextAlignStyle.Text = "自动设置";

            Data2UI();

            return;
        }

        private void chkAlignStyle_CheckedChanged(object sender, EventArgs e)
        {
            cmbAlignStyle.Enabled = chkAlignStyle.Checked;
            //Ui2Data();

            return;
        }

        private void chkIndentLeft_CheckedChanged(object sender, EventArgs e)
        {
            numIndentLeft.Enabled = chkIndentLeft.Checked;
            cmbIndentLeftUnit.Enabled = chkIndentLeft.Checked;

            //Ui2Data();

            return;
        }

        private void chkIndentRight_CheckedChanged(object sender, EventArgs e)
        {
            numIndentRight.Enabled = chkIndentRight.Checked;
            cmbIndentRightUnit.Enabled = chkIndentRight.Checked;

            //Ui2Data();

            return;
        }

        private void chkIndentSpecial_CheckedChanged(object sender, EventArgs e)
        {
            cmbIndentSpecial.Enabled = chkIndentSpecial.Checked;
            numIndentSpecial.Enabled = chkIndentSpecial.Checked;
            cmbIndentSpecialUnit.Enabled = chkIndentSpecial.Checked;
            //Ui2Data();

            return;
        }

        private void chkParaLineSpaceBefore_CheckedChanged(object sender, EventArgs e)
        {
            numBeforeParaSpacing.Enabled = (chkParaLineSpaceBefore.Checked && !chkSpaceBeforeAuto.Checked);
            cmbBeforeParaSpacingUnit.Enabled = (chkParaLineSpaceBefore.Checked && !chkSpaceBeforeAuto.Checked);
            chkSpaceBeforeAuto.Enabled = chkParaLineSpaceBefore.Checked;

            //Ui2Data();

            return;
        }


        private void chkParaLineSpaceAfter_CheckedChanged(object sender, EventArgs e)
        {
            numAfterParaSpacing.Enabled = (chkParaLineSpaceAfter.Checked && !chkSpaceAfterAuto.Checked);
            cmbAfterParaSpacingUnit.Enabled = (chkParaLineSpaceAfter.Checked && !chkSpaceAfterAuto.Checked);
            chkSpaceAfterAuto.Enabled = chkParaLineSpaceAfter.Checked;

            //Ui2Data();
            return;
        }

        private void chkParaLineSpace_CheckedChanged(object sender, EventArgs e)
        {
            cmbLineSpacingRule.Enabled = chkParaLineSpace.Checked;
            numLineSpacing.Enabled = chkParaLineSpace.Checked;
            cmbLineSpacingUnit.Enabled = chkParaLineSpace.Checked;

            if (chkParaLineSpace.Checked)
            {
                String strText = cmbLineSpacingRule.Text;

                if (strText.Equals("最小值") || strText.Equals("固定值"))
                {
                    numLineSpacing.Enabled = true;
                    cmbLineSpacingUnit.Enabled = true;
                    cmbLineSpacingUnit.Text = "磅";
                }
                else if (strText.Equals("多倍行距"))
                {
                    numLineSpacing.Enabled = true;
                    cmbLineSpacingUnit.Enabled = false;
                    cmbLineSpacingUnit.Text = "行";
                }
                else
                {
                    numLineSpacing.Enabled = false;
                    cmbLineSpacingUnit.Enabled = false;
                    cmbLineSpacingUnit.Text = "行";
                }

            }

            //Ui2Data();

            return;
        }

        private void chkTextAlignStyle_CheckedChanged(object sender, EventArgs e)
        {
            cmbTextAlignStyle.Enabled = chkTextAlignStyle.Checked;
            //Ui2Data();
            return;
        }

        private void chkSpaceBeforeAuto_CheckedChanged(object sender, EventArgs e)
        {
            numBeforeParaSpacing.Enabled = !chkSpaceBeforeAuto.Checked;
            cmbBeforeParaSpacingUnit.Enabled = !chkSpaceBeforeAuto.Checked;
            //Ui2Data();

            return;
        }

        private void chkSpaceAfterAuto_CheckedChanged(object sender, EventArgs e)
        {
            numAfterParaSpacing.Enabled = !chkSpaceAfterAuto.Checked;
            cmbAfterParaSpacingUnit.Enabled = !chkSpaceAfterAuto.Checked;
            //Ui2Data();

            return;
        }


        private void cmbLineSpacingRule_SelectedIndexChanged(object sender, EventArgs e)
        {
            String strText = cmbLineSpacingRule.Text;

            if (strText.Equals("最小值") || strText.Equals("固定值"))
            {
                numLineSpacing.Enabled = true;
                cmbLineSpacingUnit.Enabled = true;
                cmbLineSpacingUnit.Text = "磅";
            }
            else if (strText.Equals("多倍行距"))
            {
                numLineSpacing.Enabled = true;
                cmbLineSpacingUnit.Enabled = false;
                cmbLineSpacingUnit.Text = "行";
            }
            else
            {
                numLineSpacing.Enabled = false;
                cmbLineSpacingUnit.Enabled = false;
                cmbLineSpacingUnit.Text = "行";
            }

            //Ui2Data();

            return;
        }

        private void cmbAlignStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }


        private void numIndentLeft_ValueChanged(object sender, EventArgs e)
        {
            //Ui2Data();
            return;
        }


        private void cmbIndentLeftUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void numIndentRight_ValueChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }


        private void cmbIndentRightUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void cmbIndentSpecial_SelectedIndexChanged(object sender, EventArgs e)
        {
            String strText = cmbIndentSpecial.Text;

            if (strText.Equals("(无)"))
            {
                numIndentSpecial.Enabled = false;
                cmbIndentSpecialUnit.Enabled = false;
            }
            else
            {
                numIndentSpecial.Enabled = true;
                cmbIndentSpecialUnit.Enabled = true;
            }

            //Ui2Data();
            return;
        }

        private void numIndentSpecial_ValueChanged(object sender, EventArgs e)
        {
            //Ui2Data();
            return;
        }

        private void cmbIndentSpecialUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();
            return;
        }

        private void chkSymIndent_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAutoAlignRight_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        // 
        private void cmbBeforeParaSpacingUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();
            return;
        }

        // 
        private void numBeforeParaSpacing_ValueChanged(object sender, EventArgs e)
        {
            //Ui2Data();
            return;
        }


        private void numAfterParaSpacing_ValueChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void cmbAfterParaSpacingUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void numLineSpacing_ValueChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void cmbLineSpacingUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkNoBlank_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAlignMesh_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAlongParaCtrl_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }


        private void chkKeepNext_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkParaNoBreakPage_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkBreakPageBeforePara_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkCancelLineNum_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkCancelBreakWords_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkCtrlFromChinese_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAllowAsciiBreakLineInPara_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAllowCmaOverLimit_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAllowCompressCma_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAutoAdjustLineSpacing_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void chkAutoAdjustNumLineSpacing_CheckedChanged(object sender, EventArgs e)
        {
            //Ui2Data();

            return;
        }

        private void cmbTextAlignStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Ui2Data();
            return;            
        }


        private void loadListData()
        {
            foreach (String strAlignStyleItem in Globals.ThisAddIn.m_arrStrAlignStyle)
            {
                cmbAlignStyle.Items.Add(strAlignStyleItem);
            }

            // String[] strsIndentUnit = { "字符", "磅", "厘米", "毫米" };

            //foreach (String strUnitItem in Globals.ThisAddIn.m_arrSpaceUnit) // strsIndentUnit
            //{
            //    cmbIndentLeftUnit.Items.Add(strUnitItem);
            //    cmbIndentRightUnit.Items.Add(strUnitItem);
            //    cmbIndentSpecialUnit.Items.Add(strUnitItem);
            //}

            // String[] strsLineSpaceUnit = { "行", "磅", "厘米", "毫米", "英寸" };

            foreach (String strUnitItem in Globals.ThisAddIn.m_arrSpaceUnit) // strsLineSpaceUnit
            {
                cmbIndentLeftUnit.Items.Add(strUnitItem);
                cmbIndentRightUnit.Items.Add(strUnitItem);
                cmbIndentSpecialUnit.Items.Add(strUnitItem);

                cmbBeforeParaSpacingUnit.Items.Add(strUnitItem);
                cmbAfterParaSpacingUnit.Items.Add(strUnitItem);
                cmbLineSpacingUnit.Items.Add(strUnitItem);
            }

            String[] strsSpecialIndent = { "(无)", "首行缩进", "悬挂缩进" };

            foreach (String strItem in strsSpecialIndent)
            {
                cmbIndentSpecial.Items.Add(strItem);
            }

            // String[] strsParaLineSpaceRule = { "单倍行距", "1.5 倍行距", "2 倍行距", "最小值", "固定值", "多倍行距" };
            foreach (String strItem in Globals.ThisAddIn.m_arrParaLineSpaceRule)
            {
                cmbLineSpacingRule.Items.Add(strItem);
            }

            String[] strsTextAlignStyle = { "顶端对齐", "居中", "基线对齐", "底端对齐", "自动设置" };
            foreach (String strItem in strsTextAlignStyle)
            {
                cmbTextAlignStyle.Items.Add(strItem);
            }

            return;
        }


        private void Data2UI(ClassParagraphFormat oParaFmt = null)
        {

            if (oParaFmt != null)
            {
                oParaFmt.SelCopy2(inParaFmt);
            }

            //
            // inParaFmt
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.Alignment))
            {
                chkAlignStyle.Checked = true;
                switch (inParaFmt.Alignment)
                {
                    case Word.WdParagraphAlignment.wdAlignParagraphRight:
                        cmbAlignStyle.Text = "居右";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                        cmbAlignStyle.Text = "居中";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                        cmbAlignStyle.Text = "分散对齐";
                        break;

                    case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                        cmbAlignStyle.Text = "两端对齐";
                        break;

                    default:
                    case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                        cmbAlignStyle.Text = "居左";
                        break;
                }
            }
            else
            {
                cmbAlignStyle.Text = "居左";
                chkAlignStyle.Checked = false;
            }

            // Word.ParagraphFormat

            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.CharacterUnitLeftIndent))
            {
                chkIndentLeft.Checked = true;

                numIndentLeft.Value = (decimal)inParaFmt.CharacterUnitLeftIndent; // 2个字符 = 0.35cm，固定值
                cmbIndentLeftUnit.Text = "字符";
            }
            else if (inParaFmt.isSet(ClassParagraphFormat.euMembers.LeftIndent))
            {
                chkIndentLeft.Checked = true;

                numIndentLeft.Value = (decimal)inParaFmt.LeftIndent;
                cmbIndentLeftUnit.Text = "磅";
            }
            else
            {
                chkIndentLeft.Checked = false;
                cmbIndentLeftUnit.Text = "磅";
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.CharacterUnitRightIndent))
            {
                chkIndentRight.Checked = true;

                numIndentRight.Value = (decimal)inParaFmt.CharacterUnitRightIndent;
                cmbIndentRightUnit.Text = "字符";
            }
            else if (inParaFmt.isSet(ClassParagraphFormat.euMembers.RightIndent))
            {
                chkIndentRight.Checked = true;

                numIndentRight.Value = (decimal)inParaFmt.RightIndent;
                cmbIndentRightUnit.Text = "磅";
            }
            else
            {
                chkIndentRight.Checked = false;
                cmbIndentRightUnit.Text = "磅";
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.CharacterUnitFirstLineIndent))
            {
                chkIndentSpecial.Checked = true;

                float fChs = inParaFmt.CharacterUnitFirstLineIndent;

                if (fChs == 0.0f)
                {
                    cmbIndentSpecial.Text = "(无)";
                }
                else if (fChs < 0)
                {
                    cmbIndentSpecial.Text = "悬挂缩进";
                }
                else
                {
                    cmbIndentSpecial.Text = "首行缩进";
                }

                numIndentSpecial.Value = (decimal)Math.Abs(fChs);

                cmbIndentSpecialUnit.Text = "字符";

            }
            else if (inParaFmt.isSet(ClassParagraphFormat.euMembers.FirstLineIndent))
            {
                chkIndentSpecial.Checked = true;

                float fChs = inParaFmt.FirstLineIndent;

                if (fChs == 0.0f)
                {
                    cmbIndentSpecial.Text = "(无)";
                }
                else if (fChs < 0)
                {
                    cmbIndentSpecial.Text = "悬挂缩进";
                }
                else
                {
                    cmbIndentSpecial.Text = "首行缩进";
                }

                float fCents = 0.0f;

                if (fChs == 0.0f)
                {
                    fCents = 0.0f;
                }
                else
                {
                    fCents = app.PointsToCentimeters(fChs) * 28.5f;
                }

                numIndentSpecial.Value = (decimal)Math.Abs(fCents);

                cmbIndentSpecialUnit.Text = "磅";
            }
            else
            {
                chkIndentSpecial.Checked = false;
                cmbIndentSpecial.Text = "(无)";
                cmbIndentSpecialUnit.Text = "磅";
            }


            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.MirrorIndents))
            {
                if (inParaFmt.MirrorIndents != 0)
                {
                    chkSymIndent.CheckState = CheckState.Checked;
                    chkIndentLeft.Text = "内侧";
                    chkIndentRight.Text = "外侧";
                }
                else
                {
                    chkSymIndent.CheckState = CheckState.Unchecked;
                    chkIndentLeft.Text = "左侧";
                    chkIndentRight.Text = "右侧";
                }
            }
            else
            {
                chkSymIndent.CheckState = CheckState.Indeterminate;
                chkIndentLeft.Text = "左侧";
                chkIndentRight.Text = "右侧";
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.AutoAdjustRightIndent))
            {
                if (inParaFmt.AutoAdjustRightIndent != 0)
                {
                    chkAutoAlignRight.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAutoAlignRight.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkAutoAlignRight.CheckState = CheckState.Indeterminate;
            }

            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.SpaceBeforeAuto) ||
                inParaFmt.isSet(ClassParagraphFormat.euMembers.SpaceBefore))
            {
                chkParaLineSpaceBefore.Checked = true;

                if (inParaFmt.SpaceBeforeAuto != 0)
                {
                    chkSpaceBeforeAuto.Checked = true;
                }
                else
                {
                    chkSpaceBeforeAuto.Checked = false;
                    numBeforeParaSpacing.Value = (decimal)inParaFmt.SpaceBefore;
                }

                cmbBeforeParaSpacingUnit.Text = "磅";
            }
            else
            {
                chkParaLineSpaceBefore.Checked = false;
                cmbBeforeParaSpacingUnit.Text = "磅";
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.SpaceAfterAuto) ||
                inParaFmt.isSet(ClassParagraphFormat.euMembers.SpaceAfter))
            {
                chkParaLineSpaceAfter.Checked = true;

                if (inParaFmt.SpaceAfterAuto != 0)
                {
                    chkSpaceAfterAuto.Checked = true;
                }
                else
                {
                    chkSpaceAfterAuto.Checked = false;
                    numAfterParaSpacing.Value = (decimal)inParaFmt.SpaceAfter;
                }

                cmbAfterParaSpacingUnit.Text = "磅";
            }
            else
            {
                chkParaLineSpaceAfter.Checked = false;
                cmbAfterParaSpacingUnit.Text = "磅";
            }

            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacingRule) ||
                inParaFmt.isSet(ClassParagraphFormat.euMembers.LineSpacing))
            {
                chkParaLineSpace.Checked = true;

                switch (inParaFmt.LineSpacingRule)
                {
                    case Word.WdLineSpacing.wdLineSpaceSingle:
                        cmbLineSpacingRule.Text = "单倍行距";
                        break;

                    case Word.WdLineSpacing.wdLineSpace1pt5:
                        cmbLineSpacingRule.Text = "1.5 倍行距";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceDouble:
                        cmbLineSpacingRule.Text = "2 倍行距";
                        break;

                    case Word.WdLineSpacing.wdLineSpaceAtLeast:
                        cmbLineSpacingRule.Text = "最小值";

                        numLineSpacing.Value = (decimal)inParaFmt.LineSpacing;
                        cmbLineSpacingUnit.Text = "磅";

                        break;

                    case Word.WdLineSpacing.wdLineSpaceExactly:
                        cmbLineSpacingRule.Text = "固定值";

                        numLineSpacing.Value = (decimal)inParaFmt.LineSpacing;
                        cmbLineSpacingUnit.Text = "磅";

                        break;

                    case Word.WdLineSpacing.wdLineSpaceMultiple:
                        cmbLineSpacingRule.Text = "多倍行距";

                        numLineSpacing.Value = (decimal)app.PointsToLines(inParaFmt.LineSpacing);
                        cmbLineSpacingUnit.Text = "行";
                        break;

                    default:
                        break;
                }

            }
            else
            {
                chkParaLineSpace.Checked = false;
                cmbLineSpacingRule.Text = "单倍行距";
                cmbLineSpacingUnit.Text = "磅";
            }


            // DisableLineHeightGrid
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.DisableLineHeightGrid))
            {
                if (inParaFmt.DisableLineHeightGrid != 0)
                {
                    chkNoBlank.CheckState = CheckState.Checked;
                }
                else
                {
                    chkNoBlank.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkNoBlank.CheckState = CheckState.Indeterminate;
            }


            // chkAlignMesh
            chkAlignMesh.Enabled = false;
            chkAlignMesh.CheckState = CheckState.Indeterminate;


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.WidowControl))
            {
                if (inParaFmt.WidowControl != 0)
                {
                    chkAloneParaCtrl.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAloneParaCtrl.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkAloneParaCtrl.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.KeepWithNext))
            {
                if (inParaFmt.KeepWithNext != 0)
                {
                    chkKeepNext.CheckState = CheckState.Checked;
                }
                else
                {
                    chkKeepNext.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkKeepNext.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.KeepTogether))
            {
                if (inParaFmt.KeepTogether != 0)
                {
                    chkParaNoBreakPage.CheckState = CheckState.Checked;
                }
                else
                {
                    chkParaNoBreakPage.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkParaNoBreakPage.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.PageBreakBefore))
            {
                if (inParaFmt.PageBreakBefore != 0)
                {
                    chkBreakPageBeforePara.CheckState = CheckState.Checked;
                }
                else
                {
                    chkBreakPageBeforePara.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkBreakPageBeforePara.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.NoLineNumber))
            {
                if (inParaFmt.NoLineNumber != 0)
                {
                    chkCancelLineNum.CheckState = CheckState.Checked;
                }
                else
                {
                    chkCancelLineNum.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkCancelLineNum.CheckState = CheckState.Indeterminate;
            }

            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.Hyphenation))
            {
                if (inParaFmt.Hyphenation != 0)
                {
                    chkCancelBreakWords.CheckState = CheckState.Unchecked;
                }
                else
                {
                    chkCancelBreakWords.CheckState = CheckState.Checked;
                }
            }
            else
            {
                chkCancelBreakWords.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.FarEastLineBreakControl))
            {
                if (inParaFmt.FarEastLineBreakControl != 0)
                {
                    chkCtrlFromChinese.CheckState = CheckState.Checked;
                }
                else
                {
                    chkCtrlFromChinese.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkCtrlFromChinese.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.WordWrap))
            {
                if (inParaFmt.WordWrap != 0)
                {
                    chkAllowAsciiBreakLineInPara.CheckState = CheckState.Unchecked;
                }
                else
                {
                    chkAllowAsciiBreakLineInPara.CheckState = CheckState.Checked;
                }

            }
            else
            {
                chkAllowAsciiBreakLineInPara.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.HangingPunctuation))
            {
                if (inParaFmt.HangingPunctuation != 0)
                {
                    chkAllowCmaOverLimit.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAllowCmaOverLimit.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkAllowCmaOverLimit.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.HalfWidthPunctuationOnTopOfLine))
            {
                if (inParaFmt.HalfWidthPunctuationOnTopOfLine != 0)
                {
                    chkAllowCompressCma.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAllowCompressCma.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkAllowCompressCma.CheckState = CheckState.Indeterminate;
            }

            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndAlpha))
            {
                if (inParaFmt.AddSpaceBetweenFarEastAndAlpha != 0)
                {
                    chkAutoAdjustLineSpacing.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAutoAdjustLineSpacing.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkAutoAdjustLineSpacing.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndDigit))
            {
                if (inParaFmt.AddSpaceBetweenFarEastAndDigit != 0)
                {
                    chkAutoAdjustNumLineSpacing.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAutoAdjustNumLineSpacing.CheckState = CheckState.Unchecked;
                }

            }
            else
            {
                chkAutoAdjustNumLineSpacing.CheckState = CheckState.Indeterminate;
            }


            // 
            if (inParaFmt.isSet(ClassParagraphFormat.euMembers.BaseLineAlignment))
            {
                if (inParaFmt.BaseLineAlignment != 0)
                {
                    chkTextAlignStyle.Checked = true;
                }
                else
                {
                    chkTextAlignStyle.Checked = false;
                }

                switch (inParaFmt.BaseLineAlignment)
                {
                    case Word.WdBaselineAlignment.wdBaselineAlignAuto:
                        cmbTextAlignStyle.Text = "自动设置";
                        break;

                    case Word.WdBaselineAlignment.wdBaselineAlignBaseline:
                        cmbTextAlignStyle.Text = "基线对齐";
                        break;

                    case Word.WdBaselineAlignment.wdBaselineAlignCenter:
                        cmbTextAlignStyle.Text = "居中";
                        break;

                    case Word.WdBaselineAlignment.wdBaselineAlignFarEast50:
                        cmbTextAlignStyle.Text = "底端对齐";
                        break;

                    case Word.WdBaselineAlignment.wdBaselineAlignTop:
                        cmbTextAlignStyle.Text = "顶端对齐";
                        break;

                    default:
                        break;
                }
            }
            else
            {
                cmbTextAlignStyle.Text = "自动设置";
            }

            return;
        }


        private void Ui2Data()
        {
            //
            String strItem = "", strUnit = "";
            float fValue = 0.0f, fRet = 0.0f;

            strItem = cmbAlignStyle.Text;

            if (chkAlignStyle.Checked)
            {
                inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.Alignment);

                if (strItem.Equals("居左"))
                {
                    inParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                else if (strItem.Equals("居中"))
                {
                    inParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                else if (strItem.Equals("居右"))
                {
                    inParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }
                else if (strItem.Equals("分散对齐"))
                {
                    inParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                }
                else if (strItem.Equals("两端对齐"))
                {
                    inParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.Alignment);
            }

            strItem = "";
            strUnit = cmbIndentLeftUnit.Text;
            fValue = (float)numIndentLeft.Value;

            if (chkIndentLeft.Checked)
            {
                if (strUnit.Equals("字符"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitLeftIndent);
                    inParaFmt.CharacterUnitLeftIndent = fValue;
                }
                else
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LeftIndent);
                    fRet = Globals.ThisAddIn.transSpaceUnit(fValue,strUnit);
                    inParaFmt.LeftIndent = fRet;
                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitLeftIndent);
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LeftIndent);
            }

            // 
            strUnit = cmbIndentRightUnit.Text;
            fValue = (float)numIndentRight.Value;

            if (chkIndentRight.Checked)
            {
                if (strUnit.Equals("字符"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitRightIndent);
                    inParaFmt.CharacterUnitRightIndent = fValue;
                }
                else
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.RightIndent);
                    fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit);
                    inParaFmt.RightIndent = fRet;
                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitRightIndent);
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.RightIndent);
            }


            //chkIndentSpecial
            strItem = cmbIndentSpecial.Text;
            strUnit = cmbIndentSpecialUnit.Text;
            fValue = (float)numIndentSpecial.Value;

            if (chkIndentSpecial.Checked)
            {
                if (strItem.Equals("(无)"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.FirstLineIndent);
                    inParaFmt.RightIndent = 0.0f;
                }
                else if(strItem.Equals("悬挂缩进"))
                {
                    if (strUnit.Equals("字符"))
                    {
                        inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitFirstLineIndent);
                        inParaFmt.CharacterUnitFirstLineIndent = -1.0f * fValue;
                    }
                    else
                    {
                        inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.FirstLineIndent);
                        fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit);
                        inParaFmt.FirstLineIndent = -1.0f * fRet;
                    }
                }
                else if(strItem.Equals("首行缩进"))
                {
                    if (strUnit.Equals("字符"))
                    {
                        inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitFirstLineIndent);
                        inParaFmt.CharacterUnitFirstLineIndent = fValue;
                    }
                    else
                    {
                        inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.FirstLineIndent);
                        fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit);
                        inParaFmt.FirstLineIndent = fRet;
                    }
                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.CharacterUnitFirstLineIndent);
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.FirstLineIndent);
            }

            // 
            switch(chkSymIndent.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.MirrorIndents);
                    inParaFmt.MirrorIndents = -1;
                    chkIndentLeft.Text = "内侧";
                    chkIndentRight.Text = "外侧";
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.MirrorIndents);
                    inParaFmt.MirrorIndents = 0;
                    chkIndentLeft.Text = "左侧";
                    chkIndentRight.Text = "右侧";
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.MirrorIndents);
                    chkIndentLeft.Text = "左侧";
                    chkIndentRight.Text = "右侧";
                    break;
            }

            // AutoAdjustRightIndent
            switch (chkAutoAlignRight.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.AutoAdjustRightIndent);
                    inParaFmt.AutoAdjustRightIndent = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.AutoAdjustRightIndent);
                    inParaFmt.AutoAdjustRightIndent = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.AutoAdjustRightIndent);
                    break;
            }

            // 段前
            strUnit = cmbBeforeParaSpacingUnit.Text;
            fValue = (float)numBeforeParaSpacing.Value;
            if (chkParaLineSpaceBefore.Checked)
            {
                if (chkSpaceBeforeAuto.Checked)
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.SpaceBeforeAuto);
                    inParaFmt.SpaceBeforeAuto = -1;
                }
                else
                {
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.SpaceBeforeAuto);
                    inParaFmt.SpaceBeforeAuto = 0;

                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.SpaceBefore);

                    fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                    inParaFmt.SpaceBefore = fRet;

                    if (strUnit.Equals("行"))
                    {
                        inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineUnitBefore);
                        inParaFmt.LineUnitBefore = 1;
                    }
                    else
                    {
                        inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineUnitBefore);
                        inParaFmt.LineUnitBefore = 0;
                    }

                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.SpaceBefore);
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.SpaceBeforeAuto);
            }


            // 段后
            strUnit = cmbAfterParaSpacingUnit.Text;
            fValue = (float)numAfterParaSpacing.Value;
            if (chkParaLineSpaceAfter.Checked)
            {
                if (chkSpaceAfterAuto.Checked)
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.SpaceAfterAuto);
                    inParaFmt.SpaceAfterAuto = -1;
                }
                else
                {
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.SpaceAfterAuto);
                    inParaFmt.SpaceAfterAuto = 0;

                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.SpaceAfter);

                    fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                    inParaFmt.SpaceAfter = fRet;

                    if (strUnit.Equals("行"))
                    {
                        inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineUnitBefore);
                        inParaFmt.LineUnitAfter = 1;
                    }
                    else
                    {
                        inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineUnitBefore);
                        inParaFmt.LineUnitAfter = 0;
                    }
                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.SpaceAfter);
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.SpaceAfterAuto);
            }


            //
            strItem = cmbLineSpacingRule.Text;
            strUnit = cmbLineSpacingUnit.Text;
            fValue = (float)numLineSpacing.Value;

            if (chkParaLineSpace.Checked)
            {
                if(strItem.Equals("单倍行距"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                    inParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                    // inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                }
                else if (strItem.Equals("1.5 倍行距"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                    inParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                }
                else if (strItem.Equals("2 倍行距"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                    inParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                }
                else if (strItem.Equals("最小值"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                    inParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;

                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                    inParaFmt.LineSpacing = fRet;
                }
                else if (strItem.Equals("固定值"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                    inParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;

                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    fRet = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false);
                    inParaFmt.LineSpacing = fRet;
                }
                else if (strItem.Equals("多倍行距"))
                {
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                    inParaFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;

                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);

                    float fLines = Globals.ThisAddIn.transSpaceUnit(fValue, strUnit, false, "行");
                    inParaFmt.LineSpacing = app.LinesToPoints(fLines);
                   
                }
            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacingRule);
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
            }



            // 
            switch (chkNoBlank.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    inParaFmt.DisableLineHeightGrid = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    inParaFmt.DisableLineHeightGrid = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.LineSpacing);
                    break;
            }

            
            // chkAlignMesh
            switch (chkAlignMesh.CheckState)
            {
                case CheckState.Checked:
                    break;

                case CheckState.Unchecked:
                    break;

                case CheckState.Indeterminate:
                    break;
            }


            // 
            switch (chkAloneParaCtrl.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.WidowControl);
                    inParaFmt.WidowControl = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.WidowControl);
                    inParaFmt.WidowControl = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.WidowControl);
                    break;
            }


            // 
            switch (chkKeepNext.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.KeepWithNext);
                    inParaFmt.KeepWithNext = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.KeepWithNext);
                    inParaFmt.KeepWithNext = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.KeepWithNext);
                    break;
            }


            // 
            switch (chkParaNoBreakPage.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.KeepTogether);
                    inParaFmt.KeepTogether = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.KeepTogether);
                    inParaFmt.KeepTogether = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.KeepTogether);
                    break;
            }


            // 
            switch (chkBreakPageBeforePara.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.PageBreakBefore);
                    inParaFmt.PageBreakBefore = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.PageBreakBefore);
                    inParaFmt.PageBreakBefore = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.PageBreakBefore);
                    break;
            }


            // 
            switch (chkCancelLineNum.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.NoLineNumber);
                    inParaFmt.NoLineNumber = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.NoLineNumber);
                    inParaFmt.NoLineNumber = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.NoLineNumber);
                    break;
            }


            // 
            switch (chkCancelBreakWords.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.Hyphenation);
                    inParaFmt.Hyphenation = 0;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.Hyphenation);
                    inParaFmt.Hyphenation = -1;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.Hyphenation);
                    break;
            }


            // 
            switch (chkCtrlFromChinese.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.FarEastLineBreakControl);
                    inParaFmt.FarEastLineBreakControl = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.FarEastLineBreakControl);
                    inParaFmt.FarEastLineBreakControl = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.FarEastLineBreakControl);
                    break;
            }


            // 
            switch (chkAllowAsciiBreakLineInPara.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.WordWrap);
                    inParaFmt.WordWrap = 0;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.WordWrap);
                    inParaFmt.WordWrap = -1;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.WordWrap);
                    break;
            }


            // 
            switch (chkAllowCmaOverLimit.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.HangingPunctuation);
                    inParaFmt.HangingPunctuation = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.HangingPunctuation);
                    inParaFmt.HangingPunctuation = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.HangingPunctuation);
                    break;
            }


            // 
            switch (chkAllowCompressCma.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.HalfWidthPunctuationOnTopOfLine);
                    inParaFmt.HalfWidthPunctuationOnTopOfLine = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.HalfWidthPunctuationOnTopOfLine);
                    inParaFmt.HalfWidthPunctuationOnTopOfLine = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.HalfWidthPunctuationOnTopOfLine);
                    break;
            }


            // 
            switch (chkAutoAdjustLineSpacing.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndAlpha);
                    inParaFmt.AddSpaceBetweenFarEastAndAlpha = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndAlpha);
                    inParaFmt.AddSpaceBetweenFarEastAndAlpha = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndAlpha);
                    break;
            }


            // 
            switch (chkAutoAdjustNumLineSpacing.CheckState)
            {
                case CheckState.Checked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndDigit);
                    inParaFmt.AddSpaceBetweenFarEastAndDigit = -1;
                    break;

                case CheckState.Unchecked:
                    inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndDigit);
                    inParaFmt.AddSpaceBetweenFarEastAndDigit = 0;
                    break;

                case CheckState.Indeterminate:
                    inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.AddSpaceBetweenFarEastAndDigit);
                    break;
            }

            
            // 
            strItem = cmbTextAlignStyle.Text;

            if (chkTextAlignStyle.Checked)
            {
                inParaFmt.AddSelMember((int)ClassParagraphFormat.euMembers.BaseLineAlignment);

                if (strItem.Equals("顶端对齐"))
                {
                    inParaFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignTop;
                }
                else if (strItem.Equals("居中"))
                {
                    inParaFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignCenter;
                }
                else if (strItem.Equals("基线对齐"))
                {
                    inParaFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignBaseline;
                }
                else if (strItem.Equals("底端对齐"))
                {
                    inParaFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignFarEast50;
                }
                else if (strItem.Equals("自动设置"))
                {
                    inParaFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;
                }

            }
            else
            {
                inParaFmt.RemoveSelMember((int)ClassParagraphFormat.euMembers.BaseLineAlignment);
            }

            return;

        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            Ui2Data();
            return;
        }//




    }
}
