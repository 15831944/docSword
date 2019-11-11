using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
//using System.Data.SqlClient;
using System.Collections;

using OfficeAssist.Properties;
using OfficeTools.Common;
using Word = Microsoft.Office.Interop.Word;
using Office=Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;

//using Newtonsoft.Json;

using System.IO;
using System.Globalization;
using System.Diagnostics;
using System.Text.RegularExpressions;
using AutoUpdate;

// @TODO, 2016-01-18
// 1. should support column location
// 2. 

namespace OfficeAssist
{
    /// <summary>
    /// Ribbon菜单及处理函数
    /// </summary>
    public partial class Ribbon1
    {
        private ShareContributorOper m_scOper; // http处理类

        private String m_loginUser = "";
        // private Boolean m_bLogined = false;


        private ThisAddIn m_ownerAddin = null;
        private Hashtable m_hashControls = new Hashtable();

        /// <summary>
        /// Load函数，提供加载初始值等入口
        /// </summary>
        /// <param name="sender">系统缺省</param>
        /// <param name="e"></param>
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

// #if EXT
//             tab1.Label = "doc利器";
// #else
//             tab1.Label = "OfficeAssist";
// #endif

            m_ownerAddin = Globals.ThisAddIn;// 初始化插件对象
            m_scOper = m_ownerAddin.m_HttpOper; // 初始化HTTP处理对象

            if (!m_ownerAddin.m_bLoadedAllData)
            {
                m_ownerAddin.LoadAllData();
            }

            m_ownerAddin.m_bUpdTblCntOnSaving = Settings.Default.bUpdTblCntOnSaving; // 提取bUpdTblCntOnSaving参数
            m_ownerAddin.m_bUpdTblCntOnClosing = Settings.Default.bUpdTblCntOnClosing;// 提取bUpdTblCntOnSaving参数

            // 
            chkBoxUpdTblCntOnSaving.Checked = m_ownerAddin.m_bUpdTblCntOnSaving;// 对应相应UI的check状态
            chkBoxUpdTblCntOnClose.Checked = m_ownerAddin.m_bUpdTblCntOnClosing;// 对应相应UI的check状态

            chkAutoCheckUpdate.Checked = Settings.Default.bAutoUpdate;
            chkGenLocalVer.Checked = Settings.Default.bGenLocalVer;

            // 
            if (m_ownerAddin.m_bAppIsWps)
            {
                ribBtnOutLevel1.Label = "";
                ribBtnOutLevel2.Label = "";
                ribBtnOutLevel3.Label = "";
                ribBtnOutLevel4.Label = "";
                ribBtnOutLevel5.Label = "";
                ribBtnOutLevel6.Label = "";
                ribBtnOutLevel7.Label = "";
                ribBtnOutLevel8.Label = "";
                ribBtnOutLevel9.Label = "";
                ribBtnOutLevelTextBody.Label = "";

                btnNavFirst.Label = "";
                btnNavLast.Label = "";
                btnNavPrev.Label = "";
                btnNavNext.Label = "";

                chkHeadingsStylesPersist.Visible = false;

                m_ownerAddin.loadDocumentChange();
            }

            return;
        }

        /// <summary>
        /// 切换工作区窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toggleTaskWin_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
            	doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }
            
            if (doc != null)
            {
                CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_ownerAddin.HashTaskPane[doc];
                if (myCustomTaskPane != null)
                {
                    myCustomTaskPane.Visible = !myCustomTaskPane.Visible;

                    if (m_ownerAddin.m_bAppIsWps)
                    {
                        m_ownerAddin.HashDocVisible[doc] = myCustomTaskPane.Visible;
                    }
                }//
            }
        }


        /// <summary>
        /// 粘贴格式函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPasteFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }


            Word.Selection sel = doc.ActiveWindow.Selection;

            System.Collections.ArrayList lockStatusArr = new System.Collections.ArrayList();

            // record
            foreach (Word.ContentControl ctrl in sel.Range.ContentControls)
            {
                lockStatusArr.Add(ctrl.LockContentControl);
                lockStatusArr.Add(ctrl.LockContents);

                ctrl.LockContentControl = false;
                ctrl.LockContents = false;
            }

            sel.PasteFormat();

            // restore
            int i = 0;
            foreach (Word.ContentControl ctrl in sel.Range.ContentControls)
            {
                ctrl.LockContentControl = (Boolean)lockStatusArr[i*2];
                ctrl.LockContents = (Boolean)lockStatusArr[i*2 + 1];
                i++;
            }

            lockStatusArr.Clear();
        }


        /// <summary>
        /// 复制格式化信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCopyFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            sel.CopyFormat();
        }

        private System.Collections.Hashtable m_ctrlHash4Cut = new System.Collections.Hashtable();
        private System.Collections.ArrayList m_CutCtrls = new System.Collections.ArrayList();

        /// <summary>
        /// 剪切锁定内容控件
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCut4LockCtrl_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_ctrlHash4Cut.Count > 0 || m_CutCtrls.Count > 0)
            {
                MessageBox.Show("上次剪切内容未粘贴!");
                return;
            }


            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            // cancel lock of content controls
            foreach (Word.ContentControl ctrl in sel.Range.ContentControls)
            {
                System.Int16 uLockStatus = new Int16();

                uLockStatus = 0;

                if (ctrl.LockContents)
                    uLockStatus += 0x01;

                if(ctrl.LockContentControl)
                    uLockStatus += 0x10;

                m_ctrlHash4Cut.Add(ctrl.ID, uLockStatus);
                m_CutCtrls.Add(ctrl);

                ctrl.LockContentControl = false; // unlock
                ctrl.LockContents = false; // unlock
            }

            sel.Cut();
        }

        /// <summary>
        /// 粘贴刚才剪切的被锁定的内容控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPaste4LockCtrl_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;


            sel.Paste();

            // restore lock status
            // 

            foreach (Word.ContentControl ctrl in m_CutCtrls)
            {
                System.Int16 uLockStatus = (System.Int16)m_ctrlHash4Cut[ctrl.ID];

                if ( (uLockStatus & 0x01) == 0x01 )
                {
                    ctrl.LockContents = true;
                }

                if ((uLockStatus & 0x10) == 0x10)
                {
                    ctrl.LockContentControl = true;
                }

            }

            m_ctrlHash4Cut.Clear();
            m_CutCtrls.Clear();
        }

        /// <summary>
        /// 生成自动章节编号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAutoNumber_Click(object sender, RibbonControlEventArgs e)
        {
            //@TODO
            // 
            MessageBox.Show("NOT YET");
        }


        /// <summary>
        /// 插入RICK text box控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInsertRichTextBox_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            // Word.Selection sel = doc.ActiveWindow.Selection;

            Word.ContentControl ctrl = doc.ContentControls.Add(Word.WdContentControlType.wdContentControlRichText);
            // ctrl.MultiLine = true;
            ctrl.SetPlaceholderText(null, null, "[单击此处输入内容]@");
            ctrl.Range.Text = "[单击此处输入内容]";
        }

        /// <summary>
        /// 插入Picture内容控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPictureCntCtrl_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }


            doc.ContentControls.Add(Word.WdContentControlType.wdContentControlPicture);
        }

        /// <summary>
        /// 插入combox 内容控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCombox_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            doc.ContentControls.Add(Word.WdContentControlType.wdContentControlComboBox);
        }


        /// <summary>
        /// 插入ListBox的内容控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnListBox_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            doc.ContentControls.Add(Word.WdContentControlType.wdContentControlDropdownList);
        }


        /// <summary>
        /// 插入DatePicker的内容控件
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDatePicker_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            doc.ContentControls.Add(Word.WdContentControlType.wdContentControlDate);
        }


        /// <summary>
        /// 查看当前内容控件的属性
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProperty_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if((sel.Range.ContentControls.Count > 1  ||
                sel.Range.ContentControls.Count == 0 ) &&
                sel.Range.ParentContentControl == null)
            {
                MessageBox.Show("请选中一个内容控件","注意");
                return;
            }

            Object objIndex = 1;
            Word.ContentControl cntCtrl = null;
            
            if(sel.Range.ContentControls.Count > 0 )
            {
                cntCtrl = sel.Range.ContentControls[objIndex];
            }
            else
            {
                cntCtrl = sel.Range.ParentContentControl;
            }

            // property dialog
            ContentControlPropertyForm propertyFrm = new ContentControlPropertyForm();

            if (propertyFrm.ShowDialog() == DialogResult.OK)
            {
                String strName = propertyFrm.txtBoxCntCtrlName.Text;
                String strTag = propertyFrm.txtBoxCntCtrlTag.Text;

                cntCtrl.Title = strName;
                cntCtrl.Tag = strTag;
            }

            return;

//             //@TODO
//             MessageBox.Show("NOT YET");
//             return;
// 
            //             Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

//             try
//             {
//                 doc = m_ownerAddin.Application.ActiveDocument;
//             }
//             catch (System.Exception ex)
//             {
//                 MessageBox.Show("无活动文档，不能应用");
//                 return;
//             }
//             finally
//             {
//             }
// 
//             Object objIndex = "Control Toolbox"; //"开发工具";
//             Object objId = 1850;//222; // ControlProperties
// 
//             String strOut = "";
// 
//             Office.CommandBar ctBar = m_ownerAddin.Application.CommandBars[objIndex];
//             Office.CommandBarControl ctCmd = ctBar.FindControl(Type.Missing, objId);
//             ctCmd.Execute();

//             foreach(Office.CommandBar dBar in m_ownerAddin.Application.CommandBars)
//             {
//                 strOut = "name:" + dBar.Name + ",namelocal:" + dBar.NameLocal;
//                 System.Diagnostics.Debug.WriteLine(strOut);
//                 foreach (Office.CommandBarControl ctrl in dBar.Controls)
//                 {
//                     strOut = "caption:" + ctrl.Caption + ",id:" + ctrl.Id + ",accName:" + ctrl.accName;
//                     System.Diagnostics.Debug.WriteLine("     " + strOut);
//                 }
//             }
            // devBar.FindControl(Type.Missing, objId);
            // MessageBox.Show("NOT YET");
//            return;
        }


        /// <summary>
        /// 配置对话框提供配置的修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRibConfig_Click(object sender, RibbonControlEventArgs e)
        {
            FormConfig fmConfig = new FormConfig();
            DialogResult ret = fmConfig.ShowDialog();

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc != null)
            {
                CustomTaskPane userPane = (CustomTaskPane)m_ownerAddin.HashTaskPane[doc];
                if (userPane != null)
                {
                    OperationPanel form = (OperationPanel)userPane.Control;
                    form.setShareConfig(fmConfig.m_configDbUrl, fmConfig.m_configTempLoc);
                }

            }

        }


        private const String m_strNavBkmkNamePrefix = "jetBookmark";
        private UInt32 m_uBookmarkSn = 0;

        // private System.Collections.ArrayList m_arrBkmk = new System.Collections.ArrayList();
        // private int m_nBkmkIndex = -1;


        /// <summary>
        /// 在当前位置增加快捷书签
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNavAddBkmk_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;
            
            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;
            Object rng = sel.Range;

            Word.Bookmarks bks = doc.Bookmarks;

            String strBkmkName = m_strNavBkmkNamePrefix + m_uBookmarkSn++;

            while (bks.Exists(strBkmkName))
            {
                strBkmkName = m_strNavBkmkNamePrefix + m_uBookmarkSn++;
            }

            Word.Bookmark newBkmk = bks.Add(strBkmkName,ref rng);

            MessageBox.Show("完成");

            return;
        }


        /// <summary>
        /// 跳转到前一个快捷书签位置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNavPrev_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("没有活动的文档");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Bookmark firstBkmk = null;
            Word.Bookmark lastBkmk = null;
            Word.Bookmark nearstPrevBkmk = null;
            Word.Bookmark nearstNextBkmk = null;

            int nRet = m_ownerAddin.m_commTools.getNavKeyWordBookmk(doc,m_strNavBkmkNamePrefix,ref firstBkmk, ref lastBkmk, ref nearstPrevBkmk, ref nearstNextBkmk);

            Object miss = Type.Missing;

            if (nearstPrevBkmk == null)
            {
                nearstPrevBkmk = lastBkmk;
            }

            if (nearstPrevBkmk != null)
            {
                sel.GoTo(Word.WdGoToItem.wdGoToBookmark, miss, miss, nearstPrevBkmk.Name);
                doc.ActiveWindow.SetFocus();
            }

            return;
        }


        /// <summary>
        /// 跳转到下一个快捷书签位置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNavNext_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("没有活动的文档");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Bookmark firstBkmk = null;
            Word.Bookmark lastBkmk = null;
            Word.Bookmark nearstPrevBkmk = null;
            Word.Bookmark nearstNextBkmk = null;

            int nRet = m_ownerAddin.m_commTools.getNavKeyWordBookmk(doc, m_strNavBkmkNamePrefix, ref firstBkmk, ref lastBkmk, ref nearstPrevBkmk, ref nearstNextBkmk);

            Object miss = Type.Missing;

            if (nearstNextBkmk == null)
            {
                nearstNextBkmk = firstBkmk;
            }

            if (nearstNextBkmk != null)
            {
                sel.GoTo(Word.WdGoToItem.wdGoToBookmark, miss, miss, nearstNextBkmk.Name);
                doc.ActiveWindow.SetFocus();
            }

            return;
        }


        /// <summary>
        /// 清除所有快捷书签
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClearBkmk_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Bookmarks bks = doc.Bookmarks;

            foreach (Word.Bookmark bkmk in doc.Bookmarks)
            {
                if (bkmk.Name.StartsWith(m_strNavBkmkNamePrefix))
                {
                    bkmk.Delete();
                }
            }

            m_uBookmarkSn = 0;

            MessageBox.Show("完成");

            return;
        }


        /// <summary>
        /// 大纲级别的提升
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutlinePromote_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;
            Word.Paragraphs paras = null;

            paras = sel.Paragraphs;

            if (sel.Range.End - sel.Range.Start <= 1) // no selection
            {
                DialogResult ret = MessageBox.Show("没有任何选择，应用到\r\n\r\n    当前段落（选择“是”）\r\n    文档整体（选择“否”）\r\n    取消（选择“取消”）？", "注意", MessageBoxButtons.YesNoCancel);
                if (ret == DialogResult.Yes)
                {
                    // paras = m_ownerAddin.m_commTools.getSpecificHeadingParasInScope(doc) ;//doc.Paragraphs;
                    paras = sel.Paragraphs;
                }
                else if (ret == DialogResult.No)
                {
                    paras = doc.Paragraphs;

                    if (!chkOnlyNonTextBodyPara.Checked) // 
                    {
                        DialogResult ret2 = MessageBox.Show("没有选择“排除正文”，是否确认对可能大量的正文段落进行处理？", "注意", MessageBoxButtons.YesNo);
                        if (ret2 == DialogResult.No)
                        {
                            return;
                        }
                    }
                }
                else
                {
                    return;
                }
            }


            if (paras != null && paras.Count > 0)
            {
                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.StartCustomRecord("批量升级");
                }

                m_ownerAddin.m_commTools.BulkPromote(doc, paras, chkOnlyNonTextBodyPara.Checked);

                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.EndCustomRecord();
                }
            }

            MessageBox.Show("完成");

            return;
        }

        /// <summary>
        /// 大纲级别的降级
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutlineDemote_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;
            Word.Paragraphs paras = null;

            paras = doc.Paragraphs;

            if (sel.Range.End - sel.Range.Start <= 1) // no selection
            {
                DialogResult ret = MessageBox.Show("没有任何选择，应用到\r\n\r\n    当前段落（选择“是”）\r\n    文档整体（选择“否”）\r\n    取消（选择“取消”）？", "注意", MessageBoxButtons.YesNoCancel);
                if (ret == DialogResult.Yes)
                {
                    // paras = m_ownerAddin.m_commTools.getSpecificHeadingParasInScope(doc) ;//doc.Paragraphs;
                    paras = sel.Paragraphs;
                }
                else if (ret == DialogResult.No)
                {
                    paras = doc.Paragraphs;

                    if (!chkOnlyNonTextBodyPara.Checked) // 
                    {
                        DialogResult ret2 = MessageBox.Show("没有选择“排除正文”，是否确认对可能大量的正文段落进行处理？", "注意", MessageBoxButtons.YesNo);
                        if (ret2 == DialogResult.No)
                        {
                            return;
                        }
                    }
                }
                else
                {
                    return;
                }
            }

            if (paras != null && paras.Count > 0)
            {
                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.StartCustomRecord("批量降级");
                }

                m_ownerAddin.m_commTools.BulkDemote(doc, paras, chkOnlyNonTextBodyPara.Checked);

                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.EndCustomRecord();
                }
            }

            MessageBox.Show("完成");

            return;
        }


        /// <summary>
        /// 查找前一个大纲级别
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        int FindPrevParaOutlineLevel(Word.Paragraph para)
        {
            Word.Paragraph prevPara = para.Previous();

            while (prevPara != null && prevPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                prevPara = prevPara.Previous();
            }

            int nOutlineLevel = -1;
            if (prevPara != null)
            {
                nOutlineLevel = (int)prevPara.OutlineLevel;
            }

            return nOutlineLevel;
        }


        /// <summary>
        /// 大纲级别转换到内建样式
        /// 
        /// </summary>
        /// <param name="nOutlineLevel"></param>
        /// <returns></returns>
        Word.WdBuiltinStyle OutlineLevel2BuiltinStyle(int nOutlineLevel)
        {
            Word.WdBuiltinStyle style = Word.WdBuiltinStyle.wdStyleNormal;

            switch ((Word.WdOutlineLevel)nOutlineLevel)
            {
                case Word.WdOutlineLevel.wdOutlineLevel1:
                    style = Word.WdBuiltinStyle.wdStyleHeading1;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel2:
                    style = Word.WdBuiltinStyle.wdStyleHeading2;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel3:
                    style = Word.WdBuiltinStyle.wdStyleHeading3;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel4:
                    style = Word.WdBuiltinStyle.wdStyleHeading4;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel5:
                    style = Word.WdBuiltinStyle.wdStyleHeading5;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel6:
                    style = Word.WdBuiltinStyle.wdStyleHeading6;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel7:
                    style = Word.WdBuiltinStyle.wdStyleHeading7;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel8:
                    style = Word.WdBuiltinStyle.wdStyleHeading8;
                    break;

                case Word.WdOutlineLevel.wdOutlineLevel9:
                    style = Word.WdBuiltinStyle.wdStyleHeading9;
                    break;

                default:
                    break;
            }

            return style;
        }


        /// <summary>
        /// 前一个同级大纲级别的段落样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutlineSamePrev_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Paragraph curPara = sel.Paragraphs[1];

            Word.Paragraph prevPara = curPara.Previous();

            while (prevPara != null && prevPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                prevPara = prevPara.Previous();
            }

            if (prevPara != null)
            {
                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.StartCustomRecord("同前级");
                }

                prevPara.Range.Select();
                sel.CopyFormat();

                curPara.Range.Select();
                sel.PasteFormat();
                sel.Collapse();

                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.EndCustomRecord();
                }

            }

//             int nPrevLvl = FindPrevParaOutlineLevel(curPara);
// 
//             if (nPrevLvl > 0)
//             {
//                 Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nPrevLvl);
//                 Object objStyle = doc.Styles[styleIndex];
//                 sel.Paragraphs.set_Style(objStyle);
//             }

            return;
        }


        /// <summary>
        /// 低前一级的大纲级别
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutlineLow1Prev_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Paragraph curPara = sel.Paragraphs[1];

            int nPrevLvl = FindPrevParaOutlineLevel(curPara);

            if (nPrevLvl > 0 && nPrevLvl < (int)Word.WdOutlineLevel.wdOutlineLevel9)
            {
                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.StartCustomRecord("低前一级");
                }

                Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nPrevLvl+1);
                Object objStyle = doc.Styles[styleIndex];
                sel.Paragraphs.set_Style(objStyle);

                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.EndCustomRecord();
                }
            }
        }


        /// <summary>
        /// 高前一级的大纲级别样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOutlineHigh1Prev_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Paragraph curPara = sel.Paragraphs[1];

            int nPrevLvl = FindPrevParaOutlineLevel(curPara);

            if (nPrevLvl > (int)Word.WdOutlineLevel.wdOutlineLevel1)
            {
                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.StartCustomRecord("高前一级");
                }

                Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nPrevLvl - 1);
                Object objStyle = doc.Styles[styleIndex];
                sel.Paragraphs.set_Style(objStyle);

                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.EndCustomRecord();
                }
            }
        }


        /// <summary>
        /// 赋予当前段落到选择的大纲级别
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listOutline_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            RibbonButton btn = (RibbonButton)sender;

            int nSelIndex = 0;
            
            int.TryParse((String)btn.Tag,out nSelIndex);

            if (nSelIndex <= 0)
                return;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nSelIndex);
            Object objStyle = doc.Styles[styleIndex];
            sel.Paragraphs.set_Style(objStyle);
        }


        /// <summary>
        /// 查找同级的前一段落
        /// </summary>
        /// <param name="para"></param>
        /// <returns></returns>
        Word.Paragraph FindPrevSameLevelPara(Word.Paragraph para)
        {
            Word.Paragraph prevPara = para.Previous();

            while (prevPara != null && prevPara.OutlineLevel != para.OutlineLevel)
            {
                prevPara = prevPara.Previous();
            }

            return prevPara;
        }


        /// <summary>
        /// 同前级章节序号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAutoNumberAsPrev_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Paragraph curPara = sel.Paragraphs[1];

            if (curPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                MessageBox.Show("请先设置段落级别");
                return;
            }

            Word.Paragraph prevPara = FindPrevSameLevelPara(curPara);

            if (prevPara != null)
            {
                Word.ListTemplate prevListTemplate = prevPara.Range.ListFormat.ListTemplate;
                if (prevListTemplate != null)
                {
                    curPara.Range.ListFormat.ApplyListTemplate(prevListTemplate);

                    prevPara.Range.Select();
                    sel.CopyFormat();
                    curPara.Range.Select();
                    sel.PasteFormat();
                }
                else
                {
                    MessageBox.Show("前一个同级段落未定义编号，请建立其编号以便本段落进行复制");
                }

            }
            else
            {
                MessageBox.Show("无前同级段落定义编号可以复制，本段落是本级首段落，请建立编号");
            }

            return;
        }


        /// <summary>
        /// 重新编号章节序号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAllReAutoNumbering_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Paragraph curPara = sel.Paragraphs[1];
            Word.ListTemplate curListTemplate = curPara.Range.ListFormat.ListTemplate;

            if (curPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText ||
                curListTemplate == null)
            {
                MessageBox.Show("请先选择有自动编号的标题段落");
                return;
            }

            // select all headings
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    m_ownerAddin.m_commTools.RecordMultiSel(para.Range);
                }
            }

            m_ownerAddin.m_commTools.ExecMultiSel(doc);

            // apply the list template
            //
            sel = doc.ActiveWindow.Selection;
            
            if (curListTemplate != null)
            {
                Object objContinuePrevList = true;
                Object objApplyTo = Word.WdListApplyTo.wdListApplyToSelection;
                Object objDefaultListBehavior= Word.WdDefaultListBehavior.wdWord9ListBehavior;
                // sel.Range.ListFormat.ApplyListTemplateWithLevel(curListTemplate, objContinuePrevList, objApplyTo, objDefaultListBehavior);

                // NOT effective if more than 1 paragraph selected.
                MessageBox.Show("NOT EFFECTIVE");
            }
            
            return;

        }


        /// <summary>
        /// 登录点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLogin_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_ownerAddin.Application.Documents.Count == 0)
            {
                return ;
            }

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }


            if (m_ownerAddin.m_bLoginedStatus)
            {
                m_ownerAddin.logout(doc);
                btnLogin.Label = "登录文库";
                // update permission
                // 
            }
            else
            {
                String strRetMsg = "";
                LoginForm loginForm = new LoginForm();

                loginForm.txtLoginName.Text = m_ownerAddin.m_strLoginedUser;

                DialogResult res = loginForm.ShowDialog();
                if (res == DialogResult.OK)
                {
                    int nRet = m_ownerAddin.login(loginForm.txtLoginName.Text.Trim(), loginForm.txtPassword.Text.Trim(), ref strRetMsg);

                    if (m_ownerAddin.m_bLoginedStatus)
                    {
                        btnLogin.Label = "注销:" + m_ownerAddin.m_strLoginedUser;

                        // recordAutoLoginInfo();
                    }
                    else
                    {
                        btnLogin.Label = "登录文库";
                        MessageBox.Show("登录失败：" + strRetMsg);
                    }
                }
                else
                {
                    // MessageBox.Show("登录失败：" + strRetMsg);
                }

                loginForm.Dispose();
            }
            
            // btnLogin.
            return;
        }


        /// <summary>
        /// 切换工作区窗口位置
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTogglePanePos_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_ownerAddin.Application.Documents.Count <= 0)
                return;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc != null)
            {
                CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_ownerAddin.HashTaskPane[doc];
                if (myCustomTaskPane != null)
                {
                    if (myCustomTaskPane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionLeft)
                        myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    else
                        myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
                }
            }

            return;
        }


        /// <summary>
        /// 锁定内容控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCntctrlLock_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            foreach (Word.ContentControl ctrl in sel.Range.ContentControls)
            {
                if (!ctrl.LockContentControl)
                {
                    ctrl.LockContentControl = true;
                }

                if (!ctrl.LockContents)
                {
                    ctrl.LockContents = true;
                }

            }

            return;
        }

        /// <summary>
        /// 解锁当前内容控件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCntctrlUnlock_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            foreach (Word.ContentControl ctrl in sel.Range.ContentControls)
            {
                if (ctrl.LockContentControl)
                {
                    ctrl.LockContentControl = false;
                }

                if (ctrl.LockContents)
                {
                    ctrl.LockContents = false;
                }

            }

            return;
        }



        /// <summary>
        /// 设置加页眉线
        /// </summary>
        private void addHeaderLine()
        {
            if (m_ownerAddin.Application.Documents.Count == 0)
                return;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("设置页眉线");
            }


            if (doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
                doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;

            Word.Selection sel = doc.ActiveWindow.Selection;
            sel.EndKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            sel.Select();
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;

            Word.Border btmBorder = sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom];

            btmBorder.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            btmBorder.LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            btmBorder.Color = Word.WdColor.wdColorAutomatic;

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            return;

        }

        /// <summary>
        /// 去除页眉线
        /// </summary>
        private void removeHeaderLine()
        {
            if (m_ownerAddin.Application.Documents.Count == 0)
                return;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("清除页眉线");
            }

            if (doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
                doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;

            Word.Selection sel = doc.ActiveWindow.Selection;
            sel.EndKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            sel.Select();
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;

            Word.Border btmBorder = sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom];

            btmBorder.LineStyle = Word.WdLineStyle.wdLineStyleNone;

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            return;

        }


        /// <summary>
        /// 增加设置页脚线
        /// </summary>
        private void addFooterLine()
        {
            if (m_ownerAddin.Application.Documents.Count == 0)
                return;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("设置页脚线");
            }

            if (doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
                doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;

            Word.Selection sel = doc.ActiveWindow.Selection;
            sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            sel.Select();
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;

            Word.Border btmBorder = sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop];

            btmBorder.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            btmBorder.LineWidth = Word.WdLineWidth.wdLineWidth050pt;
            btmBorder.Color = Word.WdColor.wdColorAutomatic;

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            return;

        }

        /// <summary>
        /// 去除页脚线
        /// </summary>
        private void removeFooterLine()
        {
            if (m_ownerAddin.Application.Documents.Count == 0)
                return;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("清除页脚线");
            }

            if (doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdNormalView ||
                doc.ActiveWindow.ActivePane.View.Type == Word.WdViewType.wdOutlineView)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;

            Word.Selection sel = doc.ActiveWindow.Selection;
            sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            sel.Select();
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;

            Word.Border btmBorder = sel.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop];

            btmBorder.LineStyle = Word.WdLineStyle.wdLineStyleNone;

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            return;

        }

        /// <summary>
        /// 增加页眉线UI按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddHeaderLine_Click(object sender, RibbonControlEventArgs e)
        {
            addHeaderLine();
        }

        /// <summary>
        /// 去除页眉线UI按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRemoveHeaderLine_Click(object sender, RibbonControlEventArgs e)
        {
            removeHeaderLine();
        }

        /// <summary>
        /// 增加页脚线UI按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void btnAddFooterLine_Click(object sender, RibbonControlEventArgs e)
        {
            addFooterLine();
        }


        /// <summary>
        /// 去除页脚线UI按钮点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void btnClearFooterLine_Click(object sender, RibbonControlEventArgs e)
        {
            removeFooterLine();
        }

        /// <summary>
        /// 增加严格居中
        /// </summary>
        private void insertStrictCenter()
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("增加整页标准居中");
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            // 严格居中
            Object objBehav = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object objAutoFitBehav = Word.WdAutoFitBehavior.wdAutoFitFixed;

            int nEnd = sel.End;
            int nStart = sel.Start;

            if (nEnd - nStart > 1)
            {
                sel.Cut();
            }

            Word.Table tbl = doc.Tables.Add(sel.Range, 1, 1, objBehav, objAutoFitBehav);

            if (tbl == null)
            {
                if (m_ownerAddin.AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                    ur.EndCustomRecord();
                }

                return;
            }

            float fHeight = 0.0f;

            //Word.PageSetup pgsetup = doc.PageSetup;
            Word.PageSetup pgsetup = sel.Range.PageSetup;

            fHeight = (pgsetup.PageHeight - pgsetup.TopMargin - pgsetup.BottomMargin) / 28.34f;

            // tbl.Style = "网格型";
            tbl.ApplyStyleLastRow = false;
            tbl.ApplyStyleFirstColumn = false;
            tbl.ApplyStyleLastColumn = false;
            tbl.ApplyStyleRowBands = false;
            tbl.ApplyStyleColumnBands = false;

            tbl.Rows.HeightRule = Word.WdRowHeightRule.wdRowHeightAtLeast;
            tbl.Rows.Height = m_ownerAddin.Application.CentimetersToPoints(fHeight);
            // tbl.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            // tbl.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            tbl.Cell(1, 1).Select();

            if (nEnd - nStart > 1)
            {
                sel.Paste();
            }
            else
            {
                tbl.Cell(1, 1).Range.Text = "[请点击输入内容]";
            }

            tbl.Cell(1, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            tbl.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            tbl.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tbl.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tbl.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tbl.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tbl.Borders[Word.WdBorderType.wdBorderDiagonalDown].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tbl.Borders[Word.WdBorderType.wdBorderDiagonalUp].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            tbl.Borders.Shadow = false;

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            MessageBox.Show("完成");

            return;
        }

        /// <summary>
        /// 增加严格居中UI BUTTON
        /// </summary>
        private void btnStrictCenter_Click(object sender, RibbonControlEventArgs e)
        {
            insertStrictCenter();
            return;
        }


        private void btnCenterAllPics_Click_v1(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                return;
            }

            DialogResult dr = MessageBox.Show("确认将对单独成行的图片进行居中处理？", "确认", MessageBoxButtons.YesNo);

            if (dr == DialogResult.No)
            {
                return;
            }


            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("居中图片");
            }

            Word.Selection sel = doc.ActiveWindow.Selection;
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            //ArrayList arrPics = null;

            ArrayList arrIsolatePicsNotInTbl = new ArrayList();
            ArrayList arrNotIsolatePicsNotInTbl = new ArrayList();
            ArrayList arrIsolatePicsInTbl = new ArrayList();
            ArrayList arrNotIsolatePicsInTbl = new ArrayList();
            ArrayList arrPics = null;


            if (sel.Range.End - sel.Range.Start <= 1) // no selection
            {
                DialogResult ret = MessageBox.Show("没有任何选择，确认应用文档整体？", "注意", MessageBoxButtons.YesNo);
                if (ret == DialogResult.Yes)
                {
                    // Word.Paragraphs paras = doc.Paragraphs;
                    // m_ownerAddin.m_commTools.alignAllPicsInSel(paras, Word.WdParagraphAlignment.wdAlignParagraphCenter);

                    // arrPics = m_ownerAddin.m_commTools.getSpecificPicsParasInScope(doc);
                    arrPics = m_ownerAddin.m_commTools.getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                                                                            arrIsolatePicsInTbl, arrNotIsolatePicsInTbl,true);

                }
            }
            else
            {
                // Word.Paragraphs paras = sel.Range.Paragraphs;
                // m_ownerAddin.m_commTools.alignAllPicsInSel(paras, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                // arrPics = m_ownerAddin.m_commTools.getSpecificPicsParasInScope(doc, sel.Range);
                arrPics = m_ownerAddin.m_commTools.getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                                                                        arrIsolatePicsInTbl, arrNotIsolatePicsInTbl,true,sel.Range);
            }

            int nTotalPicParasCnt = arrPics.Count;
            int nIsolatePicParasNotInTblCnt = arrIsolatePicsNotInTbl.Count;
            int nNotIsolatePicParasNotInTblCnt = arrNotIsolatePicsNotInTbl.Count;
            int nIsolatePicParasInTblCnt = arrIsolatePicsInTbl.Count;
            int nNotIsolatePicParasInTblCnt = arrNotIsolatePicsInTbl.Count;


            if (arrIsolatePicsInTbl.Count > 0)
            {
                foreach (Word.Paragraph picPara in arrIsolatePicsInTbl)
                {
                    arrIsolatePicsNotInTbl.Add(picPara);
                }
            }

            float fZeroPoints = 0.0f;// app.CentimetersToPoints(0.0f);

            // foreach (Word.Paragraph picPara in arrPics)
            foreach (Word.Paragraph picPara in arrIsolatePicsNotInTbl)
            {
                //if (m_ownerAddin.m_commTools.isIsolatePic(picPara))
                //{
                    doc.ActiveWindow.ScrollIntoView(picPara.Range);
                    picPara.Range.GoTo();

                    picPara.LeftIndent = fZeroPoints;
                    picPara.RightIndent = fZeroPoints;
                    picPara.SpaceBefore = 0.0f;
                    picPara.SpaceBeforeAuto = 0;
                    picPara.SpaceAfter = 0.0f;
                    picPara.SpaceAfterAuto = 0;
                    picPara.LineSpacingRule = picPara.LineSpacingRule;//Word.WdLineSpacing.wdLineSpaceSingle;
                    picPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    picPara.WidowControl = picPara.WidowControl;//0;
                    picPara.KeepWithNext = picPara.KeepWithNext;//0;//False
                    picPara.KeepTogether = picPara.KeepTogether;//0;//False
                    picPara.PageBreakBefore = picPara.PageBreakBefore;//0;//False
                    picPara.NoLineNumber = picPara.NoLineNumber;//0;//False
                    picPara.Hyphenation = picPara.Hyphenation;// -1;//True
                    picPara.FirstLineIndent = fZeroPoints;
                    picPara.OutlineLevel = picPara.OutlineLevel;//Word.WdOutlineLevel.wdOutlineLevelBodyText;
                    picPara.CharacterUnitLeftIndent = 0;
                    picPara.CharacterUnitRightIndent = 0;
                    picPara.CharacterUnitFirstLineIndent = 0;
                    picPara.LineUnitBefore = picPara.LineUnitBefore;//0;
                    picPara.LineUnitAfter = picPara.LineUnitAfter;// 0;
                    picPara.MirrorIndents = picPara.MirrorIndents;// 0;//False;
                    picPara.TextboxTightWrap = picPara.TextboxTightWrap;// Word.WdTextboxTightWrap.wdTightNone;
                    picPara.AutoAdjustRightIndent = picPara.AutoAdjustRightIndent;//0;//False
                    picPara.DisableLineHeightGrid = picPara.DisableLineHeightGrid;//0;//False
                    picPara.FarEastLineBreakControl = picPara.FarEastLineBreakControl;// -1;//True
                    picPara.WordWrap = picPara.WordWrap;// -1;//True
                    picPara.HangingPunctuation = picPara.HangingPunctuation;// -1;//True
                    picPara.HalfWidthPunctuationOnTopOfLine = picPara.HalfWidthPunctuationOnTopOfLine;//0;//False
                    picPara.AddSpaceBetweenFarEastAndAlpha = picPara.AddSpaceBetweenFarEastAndAlpha;//0;//False
                    picPara.AddSpaceBetweenFarEastAndDigit = picPara.AddSpaceBetweenFarEastAndDigit;//0;//False
                    picPara.BaseLineAlignment = picPara.BaseLineAlignment;//Word.WdBaselineAlignment.wdBaselineAlignAuto;
                //}
            }

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }



            MessageBox.Show("完成\r\n\r\n" + "总计:" + nTotalPicParasCnt + "个图片\r\n成功：" +
                            (nIsolatePicParasNotInTblCnt + nIsolatePicParasInTblCnt) + "(正文内:" + nIsolatePicParasNotInTblCnt + "表内：" + nIsolatePicParasInTblCnt + ")" +
                            "个\r\n忽略(未单独成行):" + (nNotIsolatePicParasNotInTblCnt + nNotIsolatePicParasInTblCnt) + "个");

            return;
        }

        /// <summary>
        /// 居中图片UI BUTTON
        /// </summary>
        private void btnCenterAllPics_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                return;
            }

            DialogResult dr = MessageBox.Show("确认将对单独成行的图片进行居中处理？", "确认", MessageBoxButtons.YesNo);

            if (dr == DialogResult.No)
            {
                return;
            }


            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("居中图片");
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;

            ArrayList arrPics = null;
            ArrayList arrIsolatePicsNotInTbl = new ArrayList();
            ArrayList arrNotIsolatePicsNotInTbl = new ArrayList();
            ArrayList arrIsolatePicsInTbl = new ArrayList();
            ArrayList arrNotIsolatePicsInTbl = new ArrayList();
            // ArrayList arrPics = null;

            if (sel.Range.End - sel.Range.Start <= 1) // no selection
            {
                DialogResult ret = MessageBox.Show("没有任何选择，确认应用文档整体？", "注意", MessageBoxButtons.YesNo);
                if (ret == DialogResult.Yes)
                {
                    // Word.Paragraphs paras = doc.Paragraphs;
                    // m_ownerAddin.m_commTools.alignAllPicsInSel(paras, Word.WdParagraphAlignment.wdAlignParagraphCenter);

                    //arrPics = m_ownerAddin.m_commTools.getSpecificPicsParasInScope(doc);
                    arrPics = m_ownerAddin.m_commTools.getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                                                                            arrIsolatePicsInTbl, arrNotIsolatePicsInTbl, true);
                }
            }
            else
            {
                // Word.Paragraphs paras = sel.Range.Paragraphs;
                // m_ownerAddin.m_commTools.alignAllPicsInSel(paras, Word.WdParagraphAlignment.wdAlignParagraphCenter);
                // arrPics = m_ownerAddin.m_commTools.getSpecificPicsParasInScope(doc, sel.Range);
                arrPics = m_ownerAddin.m_commTools.getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                                                                        arrIsolatePicsInTbl, arrNotIsolatePicsInTbl, true, sel.Range);
            }


            int nTotalPicParasCnt = arrPics.Count;
            int nIsolatePicParasNotInTblCnt = arrIsolatePicsNotInTbl.Count;
            int nNotIsolatePicParasNotInTblCnt = arrNotIsolatePicsNotInTbl.Count;
            int nIsolatePicParasInTblCnt = arrIsolatePicsInTbl.Count;
            int nNotIsolatePicParasInTblCnt = arrNotIsolatePicsInTbl.Count;

            if (arrIsolatePicsInTbl.Count > 0)
            {
                foreach (Word.Paragraph picPara in arrIsolatePicsInTbl)
                {
                    arrIsolatePicsNotInTbl.Add(picPara);
                }
            }

            dynamic dgParaFmt = app.Dialogs[Word.WdWordDialog.wdDialogFormatParagraph];

            Boolean bPagination = app.Options.Pagination;
            doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            //doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;

            int nTotal = 0, nSuc = 0, nMultiSelCnt = 0, nMCnt = 0,nIgnoreIcons = 0,nSelCntInCycle = 0;
            float fZeroPoints = 0.0f;// app.CentimetersToPoints(0.0f);
            Boolean bIgnore = false;

            // foreach (Word.Paragraph picPara in arrPics)
            foreach (Word.Paragraph picPara in arrIsolatePicsNotInTbl)
            {
                bIgnore = false;

                nTotal++;

                foreach(Word.InlineShape inShp in picPara.Range.InlineShapes)
                {
                    if(inShp.Type == Word.WdInlineShapeType.wdInlineShapePictureBullet || 
                       (inShp.OLEFormat != null && inShp.OLEFormat.DisplayAsIcon))
                    {
                        bIgnore = true;
                        break;
                    }
                }

                //if (m_ownerAddin.m_commTools.isIsolatePic(picPara))
                {
                    nMultiSelCnt++;

                    doc.ActiveWindow.ScrollIntoView(picPara.Range);
                    //picPara.Range.GoTo();

                    if (!bIgnore)
                    {
                        m_ownerAddin.m_commTools.RecordMultiSel(picPara.Range);
                        nSuc++;
                        nSelCntInCycle++;
                    }
                    else
                    {
                        nIgnoreIcons++;
                    }

                    if (nMultiSelCnt == 50 || nTotal == arrIsolatePicsNotInTbl.Count)
                    {
                        nMCnt++;

                        nMultiSelCnt = 0;

                        // 
                        if (nSelCntInCycle > 0)
                        {
                            m_ownerAddin.m_commTools.ExecMultiSel(doc);
                        }
                        else
                        {
                            nSelCntInCycle = 0;
                            continue;
                        }

                        nSelCntInCycle = 0;

                        //     LeftIndent, RightIndent, Before, After, LineSpacingRule, LineSpacing, Alignment,
                        //     WidowControl, KeepWithNext, KeepTogether, PageBreak, NoLineNum, DontHyphen,
                        //     Tab, FirstIndent, OutlineLevel, Kinsoku, WordWrap, OverflowPunct, TopLinePunct,
                        //     AutoSpaceDE, LineHeightGrid, AutoSpaceDN, CharAlign, CharacterUnitLeftIndent,
                        //     AdjustRight, CharacterUnitFirstIndent, CharacterUnitRightIndent, LineUnitBefore,
                        //     LineUnitAfter, NoSpaceBetweenParagraphsOfSameStyle, OrientationBi
                        /*
                        sel.ParagraphFormat.LeftIndent = fZeroPoints;
                        sel.ParagraphFormat.RightIndent = fZeroPoints;
                        sel.ParagraphFormat.SpaceBefore = 0.0f;
                        sel.ParagraphFormat.SpaceBeforeAuto = 0;
                        sel.ParagraphFormat.SpaceAfter = 0.0f;
                        sel.ParagraphFormat.SpaceAfterAuto = 0;
                        sel.ParagraphFormat.FirstLineIndent = fZeroPoints;
                        sel.ParagraphFormat.CharacterUnitLeftIndent = 0;
                        sel.ParagraphFormat.CharacterUnitRightIndent = 0;
                        sel.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                        sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        */

                        if (m_ownerAddin.m_bAppIsWps)
                        {
                            sel.ParagraphFormat.LeftIndent = fZeroPoints;
                            sel.ParagraphFormat.RightIndent = fZeroPoints;
                            //sel.ParagraphFormat.SpaceBefore = 0.0f;
                            //sel.ParagraphFormat.SpaceBeforeAuto = 0;
                            //sel.ParagraphFormat.SpaceAfter = 0.0f;
                            //sel.ParagraphFormat.SpaceAfterAuto = 0;
                            sel.ParagraphFormat.FirstLineIndent = fZeroPoints;
                            sel.ParagraphFormat.CharacterUnitLeftIndent = 0;
                            sel.ParagraphFormat.CharacterUnitRightIndent = 0;
                            sel.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                            sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        else
                        {

                            try
                            {
                                dgParaFmt.LeftIndent = fZeroPoints;
                                dgParaFmt.RightIndent = fZeroPoints;
                                dgParaFmt.Before = 0.0f;
                                dgParaFmt.After = 0.0f;
                                dgParaFmt.FirstIndent = fZeroPoints;
                                dgParaFmt.CharacterUnitLeftIndent = 0;
                                dgParaFmt.CharacterUnitRightIndent = 0;
                                dgParaFmt.CharacterUnitFirstIndent = 0;
                                dgParaFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                                dgParaFmt.AutoSpaceDE = 0; //dgParaFmt.SpaceBeforeAuto = 0;
                                dgParaFmt.AutoSpaceDN = 0; //dgParaFmt.SpaceAfterAuto = 0;

                                dgParaFmt.Execute();
                            }
                            catch (System.Exception ex)
                            {
                                // center all
                                sel.ParagraphFormat.LeftIndent = fZeroPoints;
                                sel.ParagraphFormat.RightIndent = fZeroPoints;
                                sel.ParagraphFormat.SpaceBefore = 0.0f;
                                sel.ParagraphFormat.SpaceBeforeAuto = 0;
                                sel.ParagraphFormat.SpaceAfter = 0.0f;
                                sel.ParagraphFormat.SpaceAfterAuto = 0;
                                sel.ParagraphFormat.FirstLineIndent = fZeroPoints;
                                sel.ParagraphFormat.CharacterUnitLeftIndent = 0;
                                sel.ParagraphFormat.CharacterUnitRightIndent = 0;
                                sel.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
                                sel.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                                //sel.ParagraphFormat.LineSpacingRule = sel.ParagraphFormat.LineSpacingRule;//Word.WdLineSpacing.wdLineSpaceSingle;
                                //sel.ParagraphFormat.WidowControl = sel.ParagraphFormat.WidowControl;//0;
                                //sel.ParagraphFormat.KeepWithNext = sel.ParagraphFormat.KeepWithNext;//0;//False
                                //sel.ParagraphFormat.KeepTogether = sel.ParagraphFormat.KeepTogether;//0;//False
                                //sel.ParagraphFormat.PageBreakBefore = sel.ParagraphFormat.PageBreakBefore;//0;//False
                                //sel.ParagraphFormat.NoLineNumber = sel.ParagraphFormat.NoLineNumber;//0;//False
                                //sel.ParagraphFormat.Hyphenation = sel.ParagraphFormat.Hyphenation;// -1;//True
                                //sel.ParagraphFormat.OutlineLevel = sel.ParagraphFormat.OutlineLevel;//Word.WdOutlineLevel.wdOutlineLevelBodyText;
                                //sel.ParagraphFormat.LineUnitBefore = sel.ParagraphFormat.LineUnitBefore;//0;
                                //sel.ParagraphFormat.LineUnitAfter = sel.ParagraphFormat.LineUnitAfter;// 0;
                                //sel.ParagraphFormat.MirrorIndents = sel.ParagraphFormat.MirrorIndents;// 0;//False;
                                //sel.ParagraphFormat.TextboxTightWrap = sel.ParagraphFormat.TextboxTightWrap;// Word.WdTextboxTightWrap.wdTightNone;
                                //sel.ParagraphFormat.AutoAdjustRightIndent = sel.ParagraphFormat.AutoAdjustRightIndent;//0;//False
                                //sel.ParagraphFormat.DisableLineHeightGrid = sel.ParagraphFormat.DisableLineHeightGrid;//0;//False
                                //sel.ParagraphFormat.FarEastLineBreakControl = sel.ParagraphFormat.FarEastLineBreakControl;// -1;//True
                                //sel.ParagraphFormat.WordWrap = sel.ParagraphFormat.WordWrap;// -1;//True
                                //sel.ParagraphFormat.HangingPunctuation = sel.ParagraphFormat.HangingPunctuation;// -1;//True
                                //sel.ParagraphFormat.HalfWidthPunctuationOnTopOfLine = sel.ParagraphFormat.HalfWidthPunctuationOnTopOfLine;//0;//False
                                //sel.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha = sel.ParagraphFormat.AddSpaceBetweenFarEastAndAlpha;//0;//False
                                //sel.ParagraphFormat.AddSpaceBetweenFarEastAndDigit = sel.ParagraphFormat.AddSpaceBetweenFarEastAndDigit;//0;//False
                                //sel.ParagraphFormat.BaseLineAlignment = sel.ParagraphFormat.BaseLineAlignment;//Word.WdBaselineAlignment.wdBaselineAlignAuto;
                            }

                            //picPara.IndentCharWidth(0);
                            //picPara.IndentFirstLineCharWidth(0);

                            //picPara.LeftIndent = fZeroPoints;
                            //picPara.RightIndent = fZeroPoints;
                            //picPara.SpaceBefore = 0.0f;
                            //picPara.SpaceBeforeAuto = 0;
                            //picPara.SpaceAfter = 0.0f;
                            //picPara.SpaceAfterAuto = 0;
                            //picPara.LineSpacingRule = picPara.LineSpacingRule;//Word.WdLineSpacing.wdLineSpaceSingle;
                            //picPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            //picPara.WidowControl = picPara.WidowControl;//0;
                            //picPara.KeepWithNext = picPara.KeepWithNext;//0;//False
                            //picPara.KeepTogether = picPara.KeepTogether;//0;//False
                            //picPara.PageBreakBefore = picPara.PageBreakBefore;//0;//False
                            //picPara.NoLineNumber = picPara.NoLineNumber;//0;//False
                            //picPara.Hyphenation = picPara.Hyphenation;// -1;//True
                            //picPara.FirstLineIndent = fZeroPoints;
                            //picPara.OutlineLevel = picPara.OutlineLevel;//Word.WdOutlineLevel.wdOutlineLevelBodyText;
                            //picPara.CharacterUnitLeftIndent = 0;
                            //picPara.CharacterUnitRightIndent = 0;
                            //picPara.CharacterUnitFirstLineIndent = 0;
                            //picPara.LineUnitBefore = picPara.LineUnitBefore;//0;
                            //picPara.LineUnitAfter = picPara.LineUnitAfter;// 0;
                            //picPara.MirrorIndents = picPara.MirrorIndents;// 0;//False;
                            //picPara.TextboxTightWrap = picPara.TextboxTightWrap;// Word.WdTextboxTightWrap.wdTightNone;
                            //picPara.AutoAdjustRightIndent = picPara.AutoAdjustRightIndent;//0;//False
                            //picPara.DisableLineHeightGrid = picPara.DisableLineHeightGrid;//0;//False
                            //picPara.FarEastLineBreakControl = picPara.FarEastLineBreakControl;// -1;//True
                            //picPara.WordWrap = picPara.WordWrap;// -1;//True
                            //picPara.HangingPunctuation = picPara.HangingPunctuation;// -1;//True
                            //picPara.HalfWidthPunctuationOnTopOfLine = picPara.HalfWidthPunctuationOnTopOfLine;//0;//False
                            //picPara.AddSpaceBetweenFarEastAndAlpha = picPara.AddSpaceBetweenFarEastAndAlpha;//0;//False
                            //picPara.AddSpaceBetweenFarEastAndDigit = picPara.AddSpaceBetweenFarEastAndDigit;//0;//False
                            //picPara.BaseLineAlignment = picPara.BaseLineAlignment;//Word.WdBaselineAlignment.wdBaselineAlignAuto;

                        }// else

                    }
                }
            }

            doc.ActiveWindow.View.Type = oViewType;

            app.Options.Pagination = bPagination;

            // restore original position
            sel.Start = nOStart;
            sel.End = nOEnd;
            //// sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range); // 视角恢复

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            // MessageBox.Show("完成\r\n\r\n" + "总计:" + nTotal + "个图片\r\n成功：" +
            //                nSuc + "个\r\n失败(未单独成行):" + (nTotal - nSuc) + "个\r\n多选：" + nMCnt);

            //MessageBox.Show("完成\r\n\r\n" + "总计:" + nTotalPicParasCnt + "个图片段落\r\n成功：" +
            //                (nIsolatePicParasNotInTblCnt + nIsolatePicParasInTblCnt) + "个(正文内:" + nIsolatePicParasNotInTblCnt + ",表内：" + nIsolatePicParasInTblCnt + ")" +
            //                "\r\n忽略(未单独成行图片段落):" + (nNotIsolatePicParasNotInTblCnt + nNotIsolatePicParasInTblCnt) + "个");

            MessageBox.Show("完成\r\n\r\n" + "总计:" + nTotalPicParasCnt + "个图片段落\r\n成功：" +
                            nSuc + "个(正文内:" + (nIsolatePicParasNotInTblCnt - nIgnoreIcons) + "个；表内：" + nIsolatePicParasInTblCnt + ")" +
                            "\r\n忽略图标图片段落：" + nIgnoreIcons + "个" + 
                            "\r\n忽略(未单独成行图片段落):" + (nNotIsolatePicParasNotInTblCnt + nNotIsolatePicParasInTblCnt) + "个");

            return;
        }


        /// <summary>
        /// 数字转中文UI BUTTON
        /// </summary>
        private void btnNum2SimpCh_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                return;
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if (sel.Range.End - sel.Range.Start == 1) // no selection
            {
                MessageBox.Show("请先选择要转换的内容");
                return;
            }

            String strArabicNum;
            String strSimpChNum;
            String strBigSimpChNum;

            m_ownerAddin.m_commTools.digitTranslate(sel.Range.Text, out strArabicNum, out strSimpChNum, out strBigSimpChNum);

            sel.Range.Text = strSimpChNum;
            return;
        }

        /// <summary>
        /// 数值转中文繁体UI BUTTON
        /// </summary>
        private void btnNum2BigSimpCh_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                return;
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if (sel.Range.End - sel.Range.Start == 1) // no selection
            {
                MessageBox.Show("请先选择要转换的内容");
                return;
            }

            String strArabicNum;
            String strSimpChNum;
            String strBigSimpChNum;

            m_ownerAddin.m_commTools.digitTranslate(sel.Range.Text, out strArabicNum, out strSimpChNum, out strBigSimpChNum);

            sel.Range.Text = strBigSimpChNum;

            return;
        }

        /// <summary>
        /// 数字转阿拉伯数字UI BUTTON
        /// </summary>
        /// 
        private void btnNum2Arabic_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                return;
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if (sel.Range.End - sel.Range.Start == 1) // no selection
            {
                MessageBox.Show("请先选择要转换的内容");
                return;
            }

            String strArabicNum;
            String strSimpChNum;
            String strBigSimpChNum;

            m_ownerAddin.m_commTools.digitTranslate(sel.Range.Text, out strArabicNum, out strSimpChNum, out strBigSimpChNum);

            sel.Range.Text = strArabicNum;

            return;
        }

        /// <summary>
        /// 切换关联page ui button
        /// </summary>
        private void btnToggleRelPage_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_ownerAddin.Application.Documents.Count == 0)
            {
                return;
            }

            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc != null)
            {
                CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_ownerAddin.HashTaskPane[doc];
                if (myCustomTaskPane != null)
                {
                    myCustomTaskPane.Visible = true;

                    OperationPanel opPan = (OperationPanel)myCustomTaskPane.Control;

                    // opPan.tabCtrl.SelectedIndex = 1;
                    // opPan.tabCtrl.TabPages[]
                    opPan.tabCtrl.SelectedTab = opPan.tabPageRel;
                }
            }

            return;
        }

        /// <summary>
        /// 菜单关闭
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Ribbon1_Close(object sender, EventArgs e)
        {
            Settings.Default.Save();
            return;
        }

        /// <summary>
        /// 记录m_bUpdTblCntOnSaving点击值
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkBoxUpdTblCntOnSaving_Click(object sender, RibbonControlEventArgs e)
        {
            m_ownerAddin.m_bUpdTblCntOnSaving   = chkBoxUpdTblCntOnSaving.Checked;
            Settings.Default.bUpdTblCntOnSaving = chkBoxUpdTblCntOnSaving.Checked;
            //Settings.Default.Save();
        }


        /// <summary>
        /// 记录m_bUpdTblCntOnClosing点击值
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkBoxUpdTblCntOnClose_Click(object sender, RibbonControlEventArgs e)
        {
            m_ownerAddin.m_bUpdTblCntOnClosing   = chkBoxUpdTblCntOnClose.Checked;
            Settings.Default.bUpdTblCntOnClosing = chkBoxUpdTblCntOnClose.Checked;
            //Settings.Default.Save();
        }

        
        // read text file and write it into DOC
        /// <summary>
        ///  特定排版变换转换函数
        /// </summary>
        /// <param name="strTxtFile"></param>
        /// <param name="dstDoc"></param>
        private void painpanBody2(String strTxtFile, Word.Document dstDoc)
        {

            StreamReader sr = new StreamReader(strTxtFile, Encoding.Default);

            object oMissing = System.Reflection.Missing.Value;

            String strText = "", strCnt = "";
            int nLvl = 0;

            Hashtable hashHeading = new Hashtable();

            int nHeadingPara = -1, n1stPage = -1;
            Boolean bCloseDir = false;


            Word.Paragraph curPara = null;
            int i = 0;

            while ((strText = sr.ReadLine()) != null)
            {
                curPara = dstDoc.Content.Paragraphs.Add(oMissing);
                curPara.Range.InsertParagraphAfter();
                i++;

                if (strText.StartsWith("<目录>"))
                {
                    bCloseDir = true;
                    strCnt = strText.Replace("<目录>", "");

                    String[] strHeading = strCnt.Split('\\');

                    nLvl = strHeading.GetLength(0);

                    if (hashHeading.Contains(strCnt))
                    {
                        curPara.Range.Text = "";
                    }
                    else
                    {
                        Boolean bExist = true;

                        for (int j = 0; j < strHeading.GetLength(0); j++)
                        {
                            if (!hashHeading.Contains(strHeading[j]))
                            {
                                bExist = false;
                            }
                        }

                        if (bExist)
                        {
                            curPara.Range.Text = "";
                        }
                        else
                        {
                            curPara.Range.Text = strCnt;//Environment.NewLine;
                        }

                        if (!curPara.Range.Text.Equals("\r") &&
                            !curPara.Range.Text.Equals("\r\n"))
                        {
                            Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nLvl);
                            Object objStyle = dstDoc.Styles[styleIndex];
                            curPara.set_Style(objStyle);

                            hashHeading[strCnt] = nLvl;

                            if (nHeadingPara == -1)
                                nHeadingPara = i;
                        }

                    }
                }
                else if (strText.StartsWith("<篇名>"))
                {

                    strCnt = strText.Replace("<篇名>", "");

                    curPara.Range.Text = strCnt; // Environment.NewLine;


                    if (bCloseDir)
                    {
                        if (!curPara.Range.Text.Equals("\r") &&
                            !curPara.Range.Text.Equals("\r\n"))
                        {
                            Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nLvl + 1);
                            Object objStyle = dstDoc.Styles[styleIndex];
                            curPara.set_Style(objStyle);

                            hashHeading[strCnt] = nLvl;

                            if (nHeadingPara == -1)
                                nHeadingPara = i;
                        }
                    }
                    else
                    {
                        if (n1stPage == -1)
                            n1stPage = i;
                        // separate chapter
                        curPara.Range.Font.Size = 42;
                    }

                    bCloseDir = false;
                }
                else if (strText.StartsWith("内容："))
                {
                    String strContent = strText.Replace("内容：", "");
                    curPara.Range.Text = strContent;

                    bCloseDir = false;
                }
                else
                {
                    curPara.Range.Text = strText;
                    bCloseDir = false;
                }


                if (curPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText && i != n1stPage)
                {
                    curPara.Format.FirstLineIndent = m_ownerAddin.Application.CentimetersToPoints(0.35f);
                    curPara.Format.CharacterUnitFirstLineIndent = 2;
                }


            }// while 
            
            sr.Close();


            // replace all "\x"
            //Selection.Find.Replacement.ClearFormatting
            //With Selection.Find
            //    .Text = "\x"
            //    .Replacement.Text = ""
            //    .Forward = True
            //    .Wrap = wdFindContinue
            //    .Format = False
            //    .MatchCase = False
            //    .MatchWholeWord = False
            //    .MatchByte = True
            //    .MatchWildcards = False
            //    .MatchSoundsLike = False
            //    .MatchAllWordForms = False
            //End With
            //Selection.Find.Execute Replace:=wdReplaceAll
            
           
            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            sel.Find.Replacement.ClearFormatting();
            sel.Find.Text = "\\x";
            sel.Find.Replacement.Text = "";

            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindContinue;
            sel.Find.Format = false;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            Object objReplace = Word.WdReplace.wdReplaceAll;
            Object objDefault = Type.Missing;

            sel.Find.Execute(objDefault, objDefault, objDefault, objDefault, objDefault, objDefault,
                             objDefault, objDefault, objDefault, objDefault, objReplace);



            Word.Application app = m_ownerAddin.Application;
            // 自动编号 
            Word.ListGallery listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];

            Object objIndex = 1;
            Word.ListLevels lstLvels = listGallery.ListTemplates[objIndex].ListLevels;

//             Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;
// 
//             try
//             {
//                 doc = m_ownerAddin.Application.ActiveDocument;
//             }
//             catch (System.Exception ex)
//             {
//                 MessageBox.Show("无活动文档，不能应用");
//                 return;
//             }
//             finally
//             {
//             }

            //Word.Selection sel = doc.ActiveWindow.Selection;


            lstLvels[1].NumberFormat = "%1";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 0;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 1";

            lstLvels[2].NumberFormat = "%1.%2";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 1;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 2";

            lstLvels[3].NumberFormat = "%1.%2.%3";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 2;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 3";

            lstLvels[4].NumberFormat = "%1.%2.%3.%4";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 3;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 4";


            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 4;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 5";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 5;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 6";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 6;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 7";

            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 7;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 8";


            lstLvels[9].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[9].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[9].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[9].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[9].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[9].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[9].TabPosition = 0f;
            lstLvels[9].ResetOnHigher = 8;
            lstLvels[9].StartAt = 1;
            lstLvels[9].LinkedStyle = "标题 9";

            Object objContinue = false;
            Object objApplyTo = Word.WdListApplyTo.wdListApplyToWholeList;// .wdListApplyToSelection; // wdListApplyToWholeList;
            Object objDefaultBehav = Word.WdDefaultListBehavior.wdWord10ListBehavior;

            if (nHeadingPara != -1)
            {
                dstDoc.Paragraphs[nHeadingPara].Range.ListFormat.ApplyListTemplateWithLevel(listGallery.ListTemplates[objIndex],
                            objContinue, objApplyTo, objDefaultBehav);
            }

            return;

        }


        /// <summary>
        /// 排版函数
        /// </summary>
        /// <param name="doc"></param>
        private void paipanBody(Word.Document doc)
        {
            String strText = "", strCnt = "";
            int nLvl = 0;

            Hashtable hashHeading = new Hashtable();
            //Word.Paragraph para = null;

            int nHeadingPara = -1, n1stPage = -1;
            Boolean bCloseDir = false;

            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                // para = doc.Paragraphs[i];

                strText = doc.Paragraphs[i].Range.Text.Trim();

                if (strText.StartsWith("<目录>"))
                {
                    bCloseDir = true;
                    strCnt = strText.Replace("<目录>", "");

                    String[] strHeading = strCnt.Split('\\');

                    nLvl = strHeading.GetLength(0);

                    if (hashHeading.Contains(strCnt))
                    {
                        doc.Paragraphs[i].Range.Text = "\r";
                    }
                    else
                    {
                        Boolean bExist = true;

                        for (int j = 0; j < strHeading.GetLength(0); j++)
                        {
                            if (!hashHeading.Contains(strHeading[j]))
                            {
                                bExist = false;
                            }
                        }

                        if (bExist)
                        {
                            doc.Paragraphs[i].Range.Text = "\r";
                        }
                        else
                        {
                            doc.Paragraphs[i].Range.Text = strCnt + "\r";//Environment.NewLine;
                        }

                        if (!doc.Paragraphs[i].Range.Text.Equals("\r") &&
                            !doc.Paragraphs[i].Range.Text.Equals("\r\n"))
                        {
                            Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nLvl);
                            Object objStyle = doc.Styles[styleIndex];
                            doc.Paragraphs[i].set_Style(objStyle);

                            hashHeading[strCnt] = nLvl;

                            if (nHeadingPara == -1)
                                nHeadingPara = i;
                        }

                    }
                }
                else if (strText.StartsWith("<篇名>"))
                {
                    
                    strCnt = strText.Replace("<篇名>", "");

                    doc.Paragraphs[i].Range.Text = strCnt + "\r"; // Environment.NewLine;


                    if (bCloseDir)
                    {
                        Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nLvl + 1);
                        Object objStyle = doc.Styles[styleIndex];
                        doc.Paragraphs[i].set_Style(objStyle);

                        hashHeading[strCnt] = nLvl;

                        if (nHeadingPara == -1)
                            nHeadingPara = i;
                    }
                    else
                    {
                        if (n1stPage == -1)
                            n1stPage = i;
                        // separate chapter
                        doc.Paragraphs[i].Range.Font.Size = 42;
                    }

                    bCloseDir = false;
                }
                else if (strText.StartsWith("内容："))
                {
                    String strContent = strText.Replace("内容：", "");
                    doc.Paragraphs[i].Range.Text = strContent + "\r";

                    bCloseDir = false;
                }
                else
                {
                    bCloseDir = false;
                }

                /*
                 * else if (strText.StartsWith("\\x"))
                {
                    // if (strText.EndsWith("\\x"))
                    {
                        // 
                        String strContent = strText.Replace("\\x", "");
                        doc.Paragraphs[i].Range.Text = strContent + "\r";

                        // doc.Paragraphs[i].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }
                 * */

                if (doc.Paragraphs[i].OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText && i != n1stPage)
                {
                    doc.Paragraphs[i].Format.FirstLineIndent = m_ownerAddin.Application.CentimetersToPoints(0.35f);
                    doc.Paragraphs[i].Format.CharacterUnitFirstLineIndent = 2;
                }

            }

            // replace all "\x"
            //Selection.Find.Replacement.ClearFormatting
            //With Selection.Find
            //    .Text = "\x"
            //    .Replacement.Text = ""
            //    .Forward = True
            //    .Wrap = wdFindContinue
            //    .Format = False
            //    .MatchCase = False
            //    .MatchWholeWord = False
            //    .MatchByte = True
            //    .MatchWildcards = False
            //    .MatchSoundsLike = False
            //    .MatchAllWordForms = False
            //End With
            //Selection.Find.Execute Replace:=wdReplaceAll
            Word.Selection sel = doc.ActiveWindow.Selection;

            sel.Find.Replacement.ClearFormatting();
            sel.Find.Text = "\\x";
            sel.Find.Replacement.Text = "";
            
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindContinue;
            sel.Find.Format = false;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;
            
            Object objReplace = Word.WdReplace.wdReplaceAll;
            Object objDefault = Type.Missing;
            
            sel.Find.Execute(objDefault,objDefault,objDefault,objDefault,objDefault,objDefault,
                             objDefault,objDefault,objDefault,objDefault,objReplace);



            Word.Application app = m_ownerAddin.Application;
            // 自动编号 
            Word.ListGallery listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];

            Object objIndex = 1;
            Word.ListLevels lstLvels = listGallery.ListTemplates[objIndex].ListLevels;

//             Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;
// 
//             try
//             {
//                 doc = m_ownerAddin.Application.ActiveDocument;
//             }
//             catch (System.Exception ex)
//             {
//                 MessageBox.Show("无活动文档，不能应用");
//                 return;
//             }
//             finally
//             {
//             }
            //Word.Selection sel = doc.ActiveWindow.Selection;


            lstLvels[1].NumberFormat = "%1";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 0;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 1";

            lstLvels[2].NumberFormat = "%1.%2";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 1;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 2";

            lstLvels[3].NumberFormat = "%1.%2.%3";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 2;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 3";

            lstLvels[4].NumberFormat = "%1.%2.%3.%4";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 3;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 4";


            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 4;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 5";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 5;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 6";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 6;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 7";

            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 7;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 8";


            lstLvels[9].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[9].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[9].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[9].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[9].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[9].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[9].TabPosition = 0f;
            lstLvels[9].ResetOnHigher = 8;
            lstLvels[9].StartAt = 1;
            lstLvels[9].LinkedStyle = "标题 9";

            Object objContinue = false;
            Object objApplyTo = Word.WdListApplyTo.wdListApplyToWholeList;// .wdListApplyToSelection; // wdListApplyToWholeList;
            Object objDefaultBehav = Word.WdDefaultListBehavior.wdWord10ListBehavior;

            if (nHeadingPara != -1)
            {
                doc.Paragraphs[nHeadingPara].Range.ListFormat.ApplyListTemplateWithLevel(listGallery.ListTemplates[objIndex],
                            objContinue, objApplyTo, objDefaultBehav);
            }

            return;
        }


        /// <summary>
        /// 测试处理单个文件
        /// </summary>
        /// <param name="strDoc"></param>
        /// <param name="strNewLoc"></param>
        private void testHandleOneFile(String strDoc, String strNewLoc)
        {
            Word.Application app = m_ownerAddin.Application;

            Object sDocLoc = strDoc;
            Object nothing = System.Reflection.Missing.Value;
            Object filePath = strDoc;
            Object visible = false;
            Object objDefault = Type.Missing;

            //@TODO, open 不成，会close掉; 要add再open，始终保留一个DOC
            Word.Document doc = app.Documents.Add(objDefault, objDefault, objDefault,visible);

//             Word.Document doc = app.Documents.Open(ref filePath, ref nothing,
//                                       ref nothing, ref nothing,
//                                       ref nothing, ref nothing,
//                                       ref nothing, ref nothing,
//                                       ref nothing, ref nothing,
//                                       ref nothing, ref visible,
//                                       ref nothing, ref nothing,
//                                       ref nothing, ref nothing);

            if (doc == null)
            {
                MessageBox.Show("创建文档失败");
                return;
            }


            // <目录>卷一\上经
            // <篇名>牛膝

            // paipanBody(doc);
            painpanBody2(strDoc, doc);


            String fileName = Path.GetFileNameWithoutExtension(strDoc);

            String newFileName = strNewLoc + "\\" + fileName + ".doc";

            Object objFormat = Word.WdSaveFormat.wdFormatDocument97;
            doc.SaveAs(newFileName, objFormat);
            doc.Close();

            return;
        }

        /// <summary>
        /// 打开指定的TEXT文档，准备转换
        /// </summary>
        private void test_txt2Doc()
        {
            OpenFileDialog dig = new OpenFileDialog();

            dig.Title = "";
            dig.Filter = "(*.txt)|*.txt";

            String strFileLoc = "";
            if (dig.ShowDialog() == DialogResult.OK)
            {
                strFileLoc = dig.FileName;
            }
            else
            {
                return;
            }

            String strPath = Path.GetDirectoryName(strFileLoc);

            DirectoryInfo TheFolder = new DirectoryInfo(strPath);
            //遍历文件夹
            //foreach (DirectoryInfo NextFolder in TheFolder.GetDirectories())
            //    this.listBox1.Items.Add(NextFolder.Name);

            //遍历文件
            DateTime dt = System.DateTime.Now;

            String strNewFolder = TheFolder + "\\" + "doc_" + dt.ToString("yyyy-MM-dd_HH-mm-ss", DateTimeFormatInfo.InvariantInfo);

            Directory.CreateDirectory(strNewFolder);
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                testHandleOneFile(TheFolder + "\\" + NextFile.Name, strNewFolder);
            }

            MessageBox.Show("Done!");

            return;
        }

        /// <summary>
        /// 点击测试UI BUTTON
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRibTest_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;


            return;
        }

        /// <summary>
        /// 居中表格的UI BUTTON
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRibCenterTables_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                return;
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("居中表格");
            }

            Word.Selection sel = doc.ActiveWindow.Selection;
            
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Tables tbls = null;

            if (sel.Range.End - sel.Range.Start <= 1) // no selection
            {
                DialogResult ret = MessageBox.Show("没有任何选择，确认应用文档整体？", "注意", MessageBoxButtons.YesNo);
                if (ret == DialogResult.Yes)
                {
                    tbls = doc.Tables;
                    //m_ownerAddin.m_commTools.alignAllTablesInSel(tbls, Word.WdRowAlignment.wdAlignRowCenter);
                }
            }
            else
            {
                tbls = sel.Range.Tables;
                // m_ownerAddin.m_commTools.alignAllTablesInSel(tbls, Word.WdRowAlignment.wdAlignRowCenter);
            }

            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int nHideTbl = 0;
            foreach (Word.Table tbl in tbls)
            {
                if(m_ownerAddin.m_commTools.isHideTbl(tbl))
                {
                    nHideTbl++;
                    continue;
                }

                doc.ActiveWindow.ScrollIntoView(tbl.Range);
                tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            }

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
            else
            {
                doc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            String strMsg = "";

            strMsg = "完成：" + (tbls.Count - nHideTbl) + "个表格";

            MessageBox.Show(strMsg);

            return;
        }


        /// <summary>
        /// 复制章节样式集
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCopyHeadingStyles_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;

            if (app.Documents.Count == 0)
            {
                return;
            }

            Word.Document dstDoc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                dstDoc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            Word.Range rng = sel.Range;

            if (sel.Range.End - sel.Range.Start <= 1)
            {
                DialogResult res = MessageBox.Show("当前无选择，确定范围从整个文档？", "请确定范围", MessageBoxButtons.YesNo);

                if (res == DialogResult.Yes)
                {
                    rng = dstDoc.Range();
                }
                else
                {
                    return;
                }
            }

            m_ownerAddin.copyMultiStyles(rng);
            MessageBox.Show("完成");

            return;
        }


        /// <summary>
        /// 粘贴章节样式
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPasteHeadingStyles_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;

            if (app.Documents.Count == 0)
            {
                return;
            }

            Word.Document dstDoc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                dstDoc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            Word.Range rng = sel.Range;

            if (sel.Range.End - sel.Range.Start <= 1)
            {
                DialogResult res = MessageBox.Show("当前无选择，确定引用到整个文档？", "请确定范围", MessageBoxButtons.YesNo);

                if (res == DialogResult.Yes)
                {
                    rng = dstDoc.Range();
                }
                else
                {
                    return;
                }
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("粘贴章节样式");
            }

            String strRet = m_ownerAddin.applyMultiStyles(rng);

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            MessageBox.Show("完成\r\n" + strRet);

            return;
        }


        private void wps_insertSepTblCntSec()
        {
            DialogResult res = MessageBox.Show("确认将文档从当前位置分成3个独立节（目录前如封面、目录、目录后如正文）？", "确认", MessageBoxButtons.YesNo);

            if (res == DialogResult.No)
                return;

            Word.Application app = m_ownerAddin.Application; 

            if (app.Documents.Count == 0)
            {
                return;
            }

            Word.Document dstDoc = null;

            try
            {
                dstDoc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("因为无活动文档，不能应用");
                return;
            }
            finally
            {

            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            if (dstDoc.TablesOfContents.Count > 0)
            {
                MessageBox.Show("已经有目录，不能再创建", "失败");
                return;
            }

            if (dstDoc.ActiveWindow.ActivePane.View.SeekView != Word.WdSeekView.wdSeekMainDocument)
            {
                MessageBox.Show("请在正文区内", "失败");
                return;
            }

            Boolean bInTbl = sel.Range.get_Information(Word.WdInformation.wdWithInTable);
            if (bInTbl)
            {
                MessageBox.Show("不支持在表格中创建", "失败");
                return;
            }

            if (m_ownerAddin.AppVersion >= 11) // wps2015
            {
                Word.UndoRecord ur = app.UndoRecord;
                ur.StartCustomRecord("插入独立目录节");
            }


            Word.WdParagraphAlignment oAlign = sel.ParagraphFormat.Alignment;
            Object objUnit = Word.WdUnits.wdLine;

            sel.HomeKey(objUnit, Word.WdMovementType.wdMove);
            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            sel.InsertParagraph();
            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            if (sel.Range.Paragraphs[1].OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                //sel.MoveUp();
                sel.Range.Paragraphs[1].Range.Select();
                Word.WdBuiltinStyle styleIndex = Word.WdBuiltinStyle.wdStyleNormal;

                Object objStyle = dstDoc.Styles[styleIndex];
                sel.Paragraphs.set_Style(objStyle);

                // sel.MoveDown();
                sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            }

            sel.EndKey(objUnit, Word.WdMovementType.wdMove);
            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);


            object oMissing = Type.Missing;

            Word.Paragraph newPara = sel.Paragraphs.Add(oMissing);

            // newPara.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            newPara.Range.InsertParagraphAfter();

            newPara.Range.Text = "目  录";
            newPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            newPara.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            newPara.Range.Font.Name = "黑体";
            newPara.Range.Font.Size = 16;

            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // insert break


            Object objBreakType = Word.WdBreakType.wdSectionBreakNextPage;
            sel.InsertBreak(objBreakType);

            // insert or table of content
            Object objTrue = true;
            Object objMissing = Type.Missing;
            Object objAddedStyle = "";
            Object obj1Num = 1, obj3Num = 3;

            sel.MoveDown();

            dstDoc.TablesOfContents.Add(sel.Range, ref objTrue, ref obj1Num, ref obj3Num, ref objMissing, ref objMissing,
                                        ref objTrue, ref objTrue, ref objMissing, ref objTrue, ref objTrue, ref objTrue);
            dstDoc.TablesOfContents[1].TabLeader = Word.WdTabLeader.wdTabLeaderDots;
            // dstDoc.TablesOfContents.Format = Word.WdTocFormat.wdTOCClassic;

            sel.Start = dstDoc.TablesOfContents[1].Range.End;
            sel.End = sel.Start;
            sel.Range.GoTo();

            sel.InsertParagraph();
            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // insert break
            objBreakType = Word.WdBreakType.wdSectionBreakNextPage;
            sel.InsertBreak(objBreakType);


            sel.Start = dstDoc.TablesOfContents[1].Range.Start;
            sel.End = sel.Start;
            sel.Range.GoTo();

            int nSecIdx = sel.get_Information(Word.WdInformation.wdActiveEndSectionNumber);

            if (nSecIdx >= 2)
            {
                int nPreSecIdx = nSecIdx - 1;
                int nNextSecIdx = nSecIdx + 1;

                Word.HeaderFooter header = dstDoc.Sections[nPreSecIdx].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                Word.HeaderFooter footer = dstDoc.Sections[nPreSecIdx].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                header.LinkToPrevious = false;
                footer.LinkToPrevious = false;

                foreach (Word.Shape shp in header.Shapes)
                {
                    shp.Delete();
                }

                foreach (Word.Shape shp in footer.Shapes)
                {
                    shp.Delete();
                }

                ///
                header = dstDoc.Sections[nSecIdx].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer = dstDoc.Sections[nSecIdx].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                header.LinkToPrevious = false;
                footer.LinkToPrevious = false;

                foreach (Word.Shape shp in header.Shapes)
                {
                    shp.Delete();
                }

                foreach (Word.Shape shp in footer.Shapes)
                {
                    shp.Delete();
                }


                /// 
                header = dstDoc.Sections[nNextSecIdx].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer = dstDoc.Sections[nNextSecIdx].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                header.LinkToPrevious = false;
                footer.LinkToPrevious = false;

                foreach (Word.Shape shp in header.Shapes)
                {
                    shp.Delete();
                }

                foreach (Word.Shape shp in footer.Shapes)
                {
                    shp.Delete();
                }

                // 

                dstDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;

                footer = sel.HeaderFooter;
                footer.LinkToPrevious = false;

                //sel.TypeText("AAA");

                dstDoc.ActiveWindow.ActivePane.View.NextHeaderFooter();

                footer = sel.HeaderFooter;
                footer.LinkToPrevious = false;
                footer.Range.Select();
                sel.Delete();

                //sel.TypeText("BBB");

                dstDoc.ActiveWindow.ActivePane.View.PreviousHeaderFooter();

                footer = sel.HeaderFooter;

                // footer = dstDoc.Sections[nSecIdx].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                Word.Shape pgNumShp = footer.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 144, 144, footer.Range);
                if (pgNumShp != null)
                {
                    pgNumShp.Fill.Visible = Office.MsoTriState.msoFalse;
                    pgNumShp.Line.Visible = Office.MsoTriState.msoFalse;
                    pgNumShp.TextFrame.AutoSize = 1;
                    pgNumShp.TextFrame.WordWrap = 0;
                    pgNumShp.TextFrame.MarginLeft = 0;
                    pgNumShp.TextFrame.MarginRight = 0;
                    pgNumShp.TextFrame.MarginTop = 0;
                    pgNumShp.TextFrame.MarginBottom = 0;

                    pgNumShp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                    pgNumShp.Left = -999995;
                    pgNumShp.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
                    pgNumShp.Top = 0;
                    pgNumShp.WrapFormat.Type = Word.WdWrapType.wdWrapNone;

                    pgNumShp.TextFrame.TextRange.Text = "X";
                    pgNumShp.TextFrame.TextRange.Fields.Add(pgNumShp.TextFrame.TextRange, Word.WdFieldType.wdFieldPage, "", true);

                    footer.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
                    footer.PageNumbers.RestartNumberingAtSection = true;
                    footer.PageNumbers.StartingNumber = 1;
                }

                dstDoc.ActiveWindow.ActivePane.View.NextHeaderFooter();

                footer = sel.HeaderFooter;

                pgNumShp = footer.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 144, 144, footer.Range);
                if (pgNumShp != null)
                {
                    pgNumShp.Fill.Visible = Office.MsoTriState.msoFalse;
                    pgNumShp.Line.Visible = Office.MsoTriState.msoFalse;
                    pgNumShp.TextFrame.AutoSize = 1;
                    pgNumShp.TextFrame.WordWrap = 0;
                    pgNumShp.TextFrame.MarginLeft = 0;
                    pgNumShp.TextFrame.MarginRight = 0;
                    pgNumShp.TextFrame.MarginTop = 0;
                    pgNumShp.TextFrame.MarginBottom = 0;

                    pgNumShp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                    pgNumShp.Left = -999995;
                    pgNumShp.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
                    pgNumShp.Top = 0;
                    pgNumShp.WrapFormat.Type = Word.WdWrapType.wdWrapNone;

                    pgNumShp.TextFrame.TextRange.Text = "X";
                    pgNumShp.TextFrame.TextRange.Fields.Add(pgNumShp.TextFrame.TextRange, Word.WdFieldType.wdFieldPage, "", true);

                    footer.PageNumbers.NumberStyle = Word.WdPageNumberStyle.wdPageNumberStyleArabic;
                    footer.PageNumbers.RestartNumberingAtSection = true;
                    footer.PageNumbers.StartingNumber = 1;
                }

            }
            else
            {
                MessageBox.Show("出错！请联系系统管理员");
            }

            dstDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            // update table of content
            dstDoc.TablesOfContents[1].Update();

            if (m_ownerAddin.AppVersion >= 11) // wps2015
            {
                Word.UndoRecord ur = app.UndoRecord;
                ur.EndCustomRecord();
            }

            MessageBox.Show("完成！已将文档分成3个独立节（目录前如封面、目录、目录后如正文）及设置各节页码，用户可自行重新设置页眉页脚页码");

            return;
        }


        // 在当前光标处插入独立的目录节，将目录之上和之下分成不关联的3个节
        // 并且设置目录所在节、以下目录后的节的页码从1开始
        // 但因Office2007没有可访问的building block的页码，无法完成代码的页码插入
        // 只能人工
        /// <summary>
        /// 插入独立的目录节
        /// </summary>
        private void insertSeparateTblContentSection()
        {
            DialogResult res = MessageBox.Show("确认将文档从当前位置分成3个独立节（目录前如封面、目录、目录后如正文）？", "确认", MessageBoxButtons.YesNo);

            if (res == DialogResult.No)
                return;

            Word.Application app = m_ownerAddin.Application;

            if (app.Documents.Count == 0)
            {
                return;
            }

            Word.Document dstDoc = null;

            try
            {
                dstDoc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("因为无活动文档，不能应用");
                return;
            }
            finally
            {

            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            if (dstDoc.TablesOfContents.Count > 0)
            {
                MessageBox.Show("已经有目录，不能再创建", "失败");
                return;
            }

            Boolean bInTbl = sel.Range.get_Information(Word.WdInformation.wdWithInTable);
            if (bInTbl)
            {
                MessageBox.Show("不支持在表格中创建", "失败");
                return;
            }

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("插入独立目录节");
            }

            Word.WdParagraphAlignment oAlign = sel.ParagraphFormat.Alignment;
            Object objUnit = Word.WdUnits.wdLine;

            sel.HomeKey(objUnit, Word.WdMovementType.wdMove);
            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            sel.InsertParagraph();
            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            if (sel.Range.Paragraphs[1].OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                //sel.MoveUp();
                sel.Range.Paragraphs[1].Range.Select();
                Word.WdBuiltinStyle styleIndex = Word.WdBuiltinStyle.wdStyleNormal;

                Object objStyle = dstDoc.Styles[styleIndex];
                sel.Paragraphs.set_Style(objStyle);

                // sel.MoveDown();
                sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            }

            sel.EndKey(objUnit, Word.WdMovementType.wdMove);
            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);


            object oMissing = Type.Missing;

            Word.Paragraph newPara = sel.Paragraphs.Add(oMissing);

            // newPara.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            newPara.Range.InsertParagraphAfter();

            newPara.Range.Text = "目  录";
            newPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            newPara.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            newPara.Range.Font.Name = "黑体";
            newPara.Range.Font.Size = 16;

            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // insert break


            Object objBreakType = Word.WdBreakType.wdSectionBreakNextPage;
            sel.InsertBreak(objBreakType);

            // insert or table of content
            Object objTrue = true;
            Object objMissing = Type.Missing;
            Object objAddedStyle = "";
            Object obj1Num = 1, obj3Num = 3;

            sel.MoveDown();

            dstDoc.TablesOfContents.Add(sel.Range, ref objTrue, ref obj1Num, ref obj3Num, ref objMissing, ref objMissing,
                                        ref objTrue, ref objTrue, ref objMissing, ref objTrue, ref objTrue, ref objTrue);
            dstDoc.TablesOfContents[1].TabLeader = Word.WdTabLeader.wdTabLeaderDots;
            // dstDoc.TablesOfContents.Format = Word.WdTocFormat.wdTOCClassic;


            // update footer(link to previous and page number)
            dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
            sel.HeaderFooter.LinkToPrevious = false;


            //dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;


            dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            sel.HeaderFooter.LinkToPrevious = false;


            //Word.Field pageFld = sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldPage, @" \* ROMAN ");
            //sel.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            
            //sel.HeaderFooter.PageNumbers.RestartNumberingAtSection = true;
            //sel.HeaderFooter.PageNumbers.StartingNumber = 1;
            sel.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //sel.HeaderFooter.PageNumbers.RestartNumberingAtSection = true;
            dynamic dgPageNumberFormat = app.Dialogs[Word.WdWordDialog.wdDialogFormatPageNumber];
            dynamic dgInsertPageNumber = app.Dialogs[Word.WdWordDialog.wdDialogInsertPageNumbers];

            //if (sel.HeaderFooter.PageNumbers.Count == 0)
            {
                dgPageNumberFormat.ChapterNumber = 0;
                dgPageNumberFormat.NumRestart = 1;
                dgPageNumberFormat.NumFormat = 1;
                dgPageNumberFormat.StartingNum = 1;
                dgPageNumberFormat.Level = 0;
                dgPageNumberFormat.Separator = 0;
                dgPageNumberFormat.DoubleQuote = 0;
                dgPageNumberFormat.PgNumberingStyle = 1;

                dgPageNumberFormat.Execute();

                foreach (Word.Field tmpFld in sel.HeaderFooter.Range.Fields)
                {
                    if (tmpFld.Type == Word.WdFieldType.wdFieldPage)
                    {
                        tmpFld.Delete();
                    }
                }

                sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldPage, @" \* ROMAN ");
                //dgPageNumberFormat.Show();

                //dgInsertPageNumber.Type = 294;
                //dgInsertPageNumber.Position = 1;
                //dgInsertPageNumber.FirstPage = 1;
                //dgInsertPageNumber.Execute();
                //dgInsertPageNumber.Show();

                sel.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            

            dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
            sel.InsertParagraph();
            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // insert break
            objBreakType = Word.WdBreakType.wdSectionBreakNextPage;
            sel.InsertBreak(objBreakType);

            dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
            sel.HeaderFooter.LinkToPrevious = false;

            //dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;

            //sel.HeaderFooter.LinkToPrevious = false;
            //Word.Field pageFld2 = sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldPage);
            //sel.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


            // dstDoc.ActiveWindow.View.NextHeaderFooter();
            sel.HeaderFooter.LinkToPrevious = false;

            //sel.HeaderFooter.PageNumbers.RestartNumberingAtSection = true;
            //sel.HeaderFooter.PageNumbers.StartingNumber = 1;
            sel.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //if (sel.HeaderFooter.PageNumbers.Count == 0)
            {
                dgPageNumberFormat.ChapterNumber = 0;
                dgPageNumberFormat.NumRestart = 1;
                dgPageNumberFormat.NumFormat = 0;
                dgPageNumberFormat.StartingNum = 1;
                dgPageNumberFormat.Level = 0;
                dgPageNumberFormat.Separator = 0;
                dgPageNumberFormat.DoubleQuote = 0;
                dgPageNumberFormat.PgNumberingStyle = 0;

                dgPageNumberFormat.Execute();

                foreach (Word.Field tmpFld in sel.HeaderFooter.Range.Fields)
                {
                    if (tmpFld.Type == Word.WdFieldType.wdFieldPage)
                    {
                        tmpFld.Delete();
                    }
                }

                sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldPage);
                //dgPageNumberFormat.Show();

                //dgInsertPageNumber.Type = 294;
                //dgInsertPageNumber.Position = 1;
                //dgInsertPageNumber.FirstPage = 1;
                //dgInsertPageNumber.Execute();
                //dgInsertPageNumber.Show();

                sel.HeaderFooter.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            
            dstDoc.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            // update table of content
            dstDoc.TablesOfContents[1].Update();

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            MessageBox.Show("完成！已将文档分成3个独立节（目录前如封面、目录、目录后如正文）及设置各节页码，用户可自行重新设置页眉页脚页码");

            return;
        }

        /// <summary>
        /// 插入分离的表格内容节
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRibInsertSeparateTblContent_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_ownerAddin.m_bAppIsWps)
            {
                wps_insertSepTblCntSec();
            }
            else
            {
                insertSeparateTblContentSection();
            }
            return;
        }


        /// <summary>
        /// 记录自动登录的信息
        /// </summary>
        private void recordAutoLoginInfo()
        {
            String strEncodedName = "", strEncodedPass = "";

            strEncodedName = ClassEncryptUtils.DESEncrypt(m_ownerAddin.m_strLoginedUser, m_ownerAddin.m_stryp, m_ownerAddin.m_stryv);
            strEncodedPass = ClassEncryptUtils.DESEncrypt(m_ownerAddin.m_strLoginedPass, m_ownerAddin.m_stryp, m_ownerAddin.m_stryv);

            if (strEncodedName == null || strEncodedPass == null)
                return;

            Settings.Default.bALn = true;
            Settings.Default.strLnNm = strEncodedName;
            Settings.Default.strLnws = strEncodedPass;
            Settings.Default.Save();
            return;
        }


        /// <summary>
        /// 自动登录check box处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkAutoLogin_Click(object sender, RibbonControlEventArgs e)
        {

            if(chkAutoLogin.Checked)
            {
                if (m_ownerAddin.m_bLoginedStatus)
                {
                    recordAutoLoginInfo();

//                     String strEncodedName = "",strEncodedPass = "";
// 
//                     strEncodedName = ClassEncryptUtils.DESEncrypt(m_ownerAddin.m_strLoginedUser, m_ownerAddin.m_strPassKey, m_ownerAddin.m_strPassIv);
//                     strEncodedPass = ClassEncryptUtils.DESEncrypt(m_ownerAddin.m_strLoginedPass, m_ownerAddin.m_strPassKey, m_ownerAddin.m_strPassIv);
// 
//                     if (strEncodedName == null || strEncodedPass == null)
//                         return;
// 
//                     Settings.Default.bALn = true;
//                     Settings.Default.strLnNm = strEncodedName;
//                     Settings.Default.strLnws = strEncodedPass;
                    //Settings.Default.Save();
                }
                else
                {
                    // chkAutoLogin.Checked = false;
                    // MessageBox.Show("请在成功登录后勾选");
                }
            }
            else
            {
                Settings.Default.bALn = false;
                Settings.Default.strLnNm = "";
                Settings.Default.strLnws = "";
                Settings.Default.Save();
            }

        }

        /// <summary>
        /// 保存章节样式到内置章节样式中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ribbtnSaveCurHeadingStyle2Style_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;

            if (app.Documents.Count == 0)
            {
                return;
            }

            Word.Document dstDoc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                dstDoc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            String strHeading = "标题 ";
            Word.Style st = null;

            String strDestFile = "";

            ArrayList arrHeadingsParas = new ArrayList();

            if (sel.Range.End - sel.Range.Start <= 1)
            {
                DialogResult res = MessageBox.Show("未做任何选择，确认提取全文的章节样式？","确认",MessageBoxButtons.YesNo,MessageBoxIcon.Question);

                if (res == DialogResult.No)
                {
                    return;
                }

                arrHeadingsParas = m_ownerAddin.m_commTools.get9HeadingParas(dstDoc);
            }
            else
            {
                foreach (Word.Paragraph pItem in sel.Paragraphs)
                {
                    arrHeadingsParas.Add(pItem);
                }
                
            }

            Boolean[] bed = new Boolean[10] {false,false,false,false,false,false,false,false,false,false};
            int nLvl = -1;

            if (m_ownerAddin.AppVersion >= 15 && arrHeadingsParas.Count > 0) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord("保存当前章节样式");
            }

            foreach(Word.Paragraph para in arrHeadingsParas)
            {
                if (para.Range.Text.Trim().Equals("") || para.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    continue;

                nLvl = (int)para.OutlineLevel;
                // copy form into styles
                if(para.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    strHeading = "正文";
                }
                else
                {
                    strHeading = "标题 " + nLvl;
                }

                try
                {
                	st = dstDoc.Styles[strHeading];
                }
                catch (System.Exception ex)
                {
                    continue;
                }
                finally
                {
                    
                }

                if (st != null && !bed[nLvl])
                {
                    m_ownerAddin.m_commTools.copyFontStyle(para.Range.Font, st.Font);
                    m_ownerAddin.m_commTools.copyParagraphFormat(para.Range.ParagraphFormat, st.ParagraphFormat);
                    
                    
                    // st.ListTemplate.ListLevels


                    bed[nLvl] = true;
                }


                if (!m_ownerAddin.m_bAppIsWps && chkHeadingsStylesPersist.Checked)
                {
                    String strLoc = Environment.GetEnvironmentVariable("APPDATA");
                    if (strLoc == null)
                    {
                        MessageBox.Show("目标模板的路径变量APPDATA获取失败，请确认有此变量");
                        return;
                    }

                    strDestFile = strLoc + @"\Microsoft\Templates\Normal.dotm";
                    
                    try
                    {
                    	app.OrganizerCopy(dstDoc.FullName, strDestFile, strHeading,Word.WdOrganizerObject.wdOrganizerObjectStyles);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {

                    }
                }
            }

            if (m_ownerAddin.AppVersion >= 15 && arrHeadingsParas.Count > 0) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            MessageBox.Show("完成提取当前章节样式赋到‘标题 1~9’样式的设置，请在样式表中查看");

            return;
        }


        private void aboutInfoRefresh(ref AboutBox aboutBox)
        {
            String strInfo = "";
            String strExpInfo = m_ownerAddin.m_edtCenter.GetInfo();

            strInfo = "版本序号：" + m_ownerAddin.m_commTools.MachineIdMD5 + "\r\n" + strExpInfo;

            aboutBox.labelProductName.Text = "产品名称：doc利器";
            aboutBox.labelVersion.Text = @"版本号：" + Resources.VERSION;

            aboutBox.textBoxDescription.Text = strInfo + "\r\n\r\n适用于：WORD2007及以上版本、WPS专业版。\r\n作者：李栋、胡炜佳";

            return;
        }

        /// <summary>
        /// ABOUT关于对话框
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ribbtnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            AboutBox aboutBox = new AboutBox();

            aboutInfoRefresh(ref aboutBox);

            aboutBox.ShowDialog();

            return;
        }


        /// <summary>
        /// 复制章节结构到剪贴板
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ribbtnCopyHeadingsStructure_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document curDoc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                curDoc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = curDoc.ActiveWindow.Selection;

            if (curDoc == null)
            {
                return;
            }

            ArrayList arrParas = new ArrayList();

            if (sel.Range.End - sel.Range.Start <= 1)
            {
                DialogResult res = MessageBox.Show("当前无选择，确定范围从整个文档？", "请确定范围", MessageBoxButtons.YesNo);

                if (res == DialogResult.Yes)
                {
                    arrParas = m_ownerAddin.m_commTools.getSpecificHeadingParasInScope(curDoc);// getHeadingParas(curDoc);
                }
                else
                {
                    return;
                }
            }
            else
            {
                arrParas = m_ownerAddin.m_commTools.getSpecificHeadingParasInScope(curDoc,sel.Range);

//                 foreach (Word.Paragraph para in sel.Paragraphs)
//                 {
//                     if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
//                     {
//                         arrParas.Add(para);
//                     }
//                 }
            }

            if (arrParas.Count == 0)
            {
                MessageBox.Show("没有章节");
                return;
            }

            foreach (Word.Paragraph para in arrParas)
            {
                m_ownerAddin.m_commTools.RecordMultiSel(para.Range);
            }

            m_ownerAddin.m_commTools.ExecMultiSel(curDoc);

            sel.Copy();

            MessageBox.Show(@"已复制到剪贴板");

            return;
        }


        /// <summary>
        /// 帮助UI BUTTON处理函数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ribBtnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            String strLocFileUrl = "使用手册*.pdf";

            string[] filedir = Directory.GetFiles(strBaseDir, strLocFileUrl, SearchOption.TopDirectoryOnly);

            if (filedir.GetLength(0) == 0)
            {
                MessageBox.Show("帮助文件丢失，请检查安装目录下的使用手册文档");
                return;
            }

            strLocFileUrl = filedir[0];

            Process proc = null;

            try
            {
                proc = System.Diagnostics.Process.Start(strLocFileUrl);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }

            if (proc == null)
            {
                MessageBox.Show("请安装PDF阅读器");
            }


            return;

 /*           // open help file
            // open it
            Word.Application wordApplication = new Word.Application();
            Word.Document wordDocument = new Word.Document();
            Object nothing = System.Reflection.Missing.Value;
            Object filePath = strLocFileUrl;
            wordApplication.Documents.Open(ref filePath, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref nothing, ref   nothing, ref   nothing, ref   nothing, ref   nothing, ref  nothing, ref   nothing, ref   nothing);
            wordDocument = wordApplication.ActiveDocument;
            wordApplication.Visible = true;


            return;*/
        }

        private void ribBtnTutorial_Click(object sender, RibbonControlEventArgs e)
        {
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            String strLocFileUrl = "*入门*.pdf";

            string[] filedir = Directory.GetFiles(strBaseDir, strLocFileUrl, SearchOption.TopDirectoryOnly);

            if (filedir.GetLength(0) == 0)
            {
                MessageBox.Show("帮助文件丢失，请检查安装目录下的帮助文档");
                return;
            }

            strLocFileUrl = filedir[0];

            Process proc = null;

            try
            {
                proc = System.Diagnostics.Process.Start(strLocFileUrl);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }

            if (proc == null)
            {
                MessageBox.Show("请安装PDF阅读器");
            }


            /*
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            String strLocFileUrl = "UpdateTool.exe";

            string[] filedir = Directory.GetFiles(strBaseDir, strLocFileUrl, SearchOption.TopDirectoryOnly);

            if (filedir.GetLength(0) == 0)
            {
                MessageBox.Show("帮助文件丢失，请检查安装目录下的帮助文档");
                return;
            }

            strLocFileUrl = filedir[0];

            System.Diagnostics.Process.Start(strLocFileUrl, "http://10.115.246.179:30000/WebService.asmx");

            return;
            */

            /*
            System.Diagnostics.ProcessStartInfo cp = new System.Diagnostics.ProcessStartInfo();
            cp.FileName =  AppDomain.CurrentDomain.BaseDirectory + @"UpdateTool.exe";
            //cp.FileName =  @"C:\Users\Administrator\Desktop\Debug\" + @"UpdateTool.exe";

            cp.Arguments = "";// "http://10.115.246.179:30000/WebService.asmx";
            cp.UseShellExecute = false;
            cp.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory;
//             cp.UserName = "administrator";

            cp.RedirectStandardInput = true;
            cp.RedirectStandardOutput = true;
            cp.RedirectStandardError = true;
            //cp.ErrorDialog = false;
            //cp.CreateNoWindow = true;

            System.Diagnostics.Process.Start(cp);

            / *
            try
            {
                System.Diagnostics.Process.Start(cp);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }
            */

            return;
        }


        private void ribBtnCheckUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            String strRet = "0";

            // strRet = m_ownerAddin.m_uiCtrler.checkUpdate();
            strRet = m_ownerAddin.m_edtCenter.CheckUpdate();

            int nRet = 0;

            if(int.TryParse(strRet,out nRet))
            {
                switch (nRet)
                {
                    case 0:
                        break;
                    case -1:
                        MessageBox.Show("没有最新版本，不需要更新");
                        break;
                    case -2:
                        MessageBox.Show("无法检查更新，没有网络连接");
                        break;
                    case -3:
                        MessageBox.Show("下载更新配置文件失败，请检查网络和文件夹权限");
                        break;

                    default:
                        break;
                }
            }
            else
            {
                // MessageBox.Show("NEVER TO SEE");
            }

            return;
        }

        private void chkAutoCheckUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.bAutoUpdate = chkAutoCheckUpdate.Checked;
            Settings.Default.Save();

            return;
        }

        private void RibbtnOpenCurDocDir_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;

            if (app.Documents.Count == 0)
            {
                return;
            }

            Word.Document dstDoc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                dstDoc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;

            String strPath = Path.GetDirectoryName(dstDoc.FullName);

            if (!strPath.Equals(""))
            {
                System.Diagnostics.Process.Start(strPath);
            }
            else
            {
                MessageBox.Show("未保存的新建文档尚无目录");
            }

            return;
        }


        private void RibbtnRegister_Click(object sender, RibbonControlEventArgs e)
        {
            frmRegister frmReg = new frmRegister();

            frmReg.txtRegisterAccount.Text = "";
            frmReg.txtActivateSn.Text = "";
            frmReg.txtRegisterInfo.Text = "";


            String strAccount = "", strActSn = "", strInfo = "";

            strInfo = m_ownerAddin.m_edtCenter.GetPrivEditionInfo(ref strAccount, ref strActSn);


            frmReg.txtRegisterAccount.Text = strAccount;
            String strSn = strActSn;

            if (strSn.Length > 4)
            {
                frmReg.txtActivateSn.Text = strSn.Substring(0, strSn.Length - 4) + "****"; // 
            }
            else if (!String.IsNullOrWhiteSpace(strAccount))
            {
                frmReg.txtActivateSn.Text = "****";
            }


            frmReg.txtRegisterInfo.Text = "版本：" + Settings.Default.strVerName + "\r\n";

            frmReg.txtRegisterInfo.Text += strInfo;


            DialogResult res = frmReg.ShowDialog();

            if (res == DialogResult.Cancel)
            {
                return;
            }

            String strRetMsg = "";
            
            // int nRet = m_ownerAddin.m_uiCtrler.activateSoft(frmReg.txtRegisterAccount.Text.Trim(), frmReg.txtActivateSn.Text.Trim(),ref strRetMsg);

            int nRet = m_ownerAddin.m_edtCenter.ActivateSoft(frmReg.txtRegisterAccount.Text.Trim(), frmReg.txtActivateSn.Text.Trim(), ref strRetMsg);

            Word.Document doc = null;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
                m_ownerAddin.updateUI(doc);
            }
            catch (System.Exception ex)
            {

            }
            finally
            {
            }

            if (nRet == 0)
            {
                MessageBox.Show("激活成功！");
            }
            else
            {
                MessageBox.Show("激活失败，原因："+strRetMsg);
            }

/*
            QueryResult qRet = null;
            if (!String.IsNullOrWhiteSpace(m_ownerAddin.m_cfgAutoUpdateSvrUrl))
            {
                AutoUpdate.AutoUpdateClass checkUpdate = new AutoUpdate.AutoUpdateClass(m_ownerAddin.m_cfgAutoUpdateSvrUrl);
                qRet = checkUpdate.ActiveProject(frmReg.txtRegisterAccount.Text, frmReg.txtActivateSn.Text, strMachineId);
            }

            // record
            if (qRet == null || !qRet.IsSuccess)
            {
                Settings.Default.bRegSuc = false;
                Settings.Default.Save();

                if (qRet == null)
                {
                    MessageBox.Show("激活失败，请保持与注册服务器的网络畅通");
                }
                else
                {
                    MessageBox.Show("激活失败，错误原因：" + qRet.ErrorInfo);
                }
            }
            else
            {
                strRegAccount = ClassEncryptUtils.DESEncrypt(frmReg.txtRegisterAccount.Text, m_ownerAddin.m_stryp, m_ownerAddin.m_stryv);
                strRegActSn = ClassEncryptUtils.DESEncrypt(frmReg.txtActivateSn.Text, m_ownerAddin.m_stryp, m_ownerAddin.m_stryv);

                if (!String.IsNullOrWhiteSpace(strRegAccount) && !String.IsNullOrWhiteSpace(strRegActSn))
                {
                    Settings.Default.bRegSuc = true;
                    Settings.Default.strRegAcnt = strRegAccount;
                    Settings.Default.strRegActSn = strRegActSn;
                    Settings.Default.dtRegExp = qRet.Date;
                    Settings.Default.Save();
                }

                // permission file write
//                 if ((strRet.Length & 0x1F) != 0)
//                 {
//                     MessageBox.Show("权限码异常");
// 
//                 }

                // write into lic.dat
                // m_ownerAddin.m_licDat
                String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
                String strLicDat = strBaseDir + @"\" + m_ownerAddin.m_licDat;

                String strEncLicDat = m_ownerAddin.EncodeLicData(qRet.Data1, strMachineId);

                StreamWriter sw = new StreamWriter(strLicDat);
                sw.Write(strEncLicDat);
                sw.Close();


                m_ownerAddin.m_hashRegPerm = m_ownerAddin.getRegisterPermItems(qRet.Data1);

                // update ui
                Word.Application app = m_ownerAddin.Application;
                Word.Document curDoc = app.ActiveDocument;

                CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_ownerAddin.HashTaskPane[curDoc];
                OperationPanel userPane = null;

                if (myCustomTaskPane != null)
                {
                    userPane = (OperationPanel)myCustomTaskPane.Control;

                    if (userPane != null)
                    {
                        // userPane.m_hashControls
                        // m_hashKeyControls
                        String strNameMD5 = "";
                        Object ctrlObj = null;
                        Control ctrl = null;
                        ToolStripItem item = null;
                        RibbonControl ribCtrl = null;

                        foreach (DictionaryEntry ent in userPane.m_hashKeyControls)
                        {
                            strNameMD5 = (String)ent.Key;
                            ctrlObj = (Object)userPane.m_hashKeyControls[strNameMD5];

                            if (m_ownerAddin.m_hashRegPerm.Contains(strNameMD5))
                            {
                                if (ctrlObj != null)
                                {
                                    if (ctrlObj is Control)
                                    {
                                        ctrl = (Control)ctrlObj;
                                        ctrl.Enabled = true;
                                        / *
                                        if (ctrl is TabPage)
                                        {
                                            // 递归enable child controls
                                            m_ownerAddin.transChildCtrl(ctrl, true);
                                        }
                                        else if (!(ctrl is TabControl))
                                        {
                                            // 递归enable child controls
                                            m_ownerAddin.transChildCtrl(ctrl, true);
                                        }
                                         * * /
                                    }
                                    else if (ctrlObj is ToolStripItem)
                                    {
                                        item = (ToolStripItem)ctrlObj;
                                        item.Enabled = true;
                                    }
                                    else if (ctrlObj is RibbonGroup)
                                    {
                                        // 
                                    }
                                    else if (ctrlObj is RibbonControl)
                                    {
                                        ribCtrl = (RibbonControl)ctrlObj;
                                        ribCtrl.Enabled = true;
                                    }
                                }// if

                            }
                            else
                            {
                                if (ctrlObj != null)
                                {
                                    if (ctrlObj is Control)
                                    {
                                        ctrl = (Control)ctrlObj;
                                        ctrl.Enabled = false;

                                        / *
                                        if (ctrl is TabPage)
                                        {
                                            // 递归enable child controls
                                            m_ownerAddin.transChildCtrl(ctrl, false);
                                        }
                                        else if (!(ctrl is TabControl))
                                        {
                                            // 递归enable child controls
                                            m_ownerAddin.transChildCtrl(ctrl, false);
                                        }
                                         * * /
                                    }
                                    else if (ctrlObj is ToolStripItem)
                                    {
                                        item = (ToolStripItem)ctrlObj;
                                        item.Enabled = false;
                                    }
                                    else if (ctrlObj is RibbonGroup)
                                    {
                                        // 
                                    }
                                    else if (ctrlObj is RibbonControl)
                                    {
                                        ribCtrl = (RibbonControl)ctrlObj;
                                        ribCtrl.Enabled = false;
                                    }
                                }// if
                            }

                        }

                        userPane.RefreshRelsByPermission();
                        userPane.RefreshShareLibByPermission();
                        userPane.refreshMyComputerFolders();
                        userPane.recordCommonShareLibTree();

                        userPane.Invalidate();
                    }

                }

                Word.Document oDoc = null, openDoc = null;

                // permission broadcast
                foreach (DictionaryEntry ent in m_ownerAddin.HashTaskPane)
                {
                    oDoc = (Word.Document)ent.Key;
                    myCustomTaskPane = (CustomTaskPane)ent.Value;

                    try
                    {
                        openDoc = app.Documents[oDoc];
                    }
                    catch (System.Exception ex)
                    {
                        continue;
                    }
                    finally
                    {
                    }

                    if (oDoc == curDoc || openDoc == null)
                    {
                        continue;
                    }

                    if (myCustomTaskPane != null)
                    {
                        userPane = (OperationPanel)myCustomTaskPane.Control;

                        if (userPane != null)
                        {
                            // userPane.m_hashControls
                            // m_hashKeyControls
                            String strNameMD5 = "";
                            Object ctrlObj = null;
                            Control ctrl = null;
                            ToolStripItem item = null;
                            RibbonControl ribCtrl = null;

                            foreach (DictionaryEntry ent2 in userPane.m_hashKeyControls)
                            {
                                strNameMD5 = (String)ent2.Key;
                                ctrlObj = (Object)userPane.m_hashKeyControls[strNameMD5];

                                if (m_ownerAddin.m_hashRegPerm.Contains(strNameMD5))
                                {
                                    if (ctrlObj != null)
                                    {
                                        if (ctrlObj is Control)
                                        {
                                            ctrl = (Control)ctrlObj;
                                            ctrl.Enabled = true;
                                            / *
                                            if (ctrl is TabPage)
                                            {
                                                // 递归enable child controls
                                                m_ownerAddin.transChildCtrl(ctrl, true);
                                            }
                                            else if (!(ctrl is TabControl))
                                            {
                                                // 递归enable child controls
                                                m_ownerAddin.transChildCtrl(ctrl, true);
                                            }
                                             * * /
                                        }
                                        else if (ctrlObj is ToolStripItem)
                                        {
                                            item = (ToolStripItem)ctrlObj;
                                            item.Enabled = true;
                                        }
                                        else if (ctrlObj is RibbonGroup)
                                        {
                                            // 
                                        }
                                        else if (ctrlObj is RibbonControl)
                                        {
                                            ribCtrl = (RibbonControl)ctrlObj;
                                            ribCtrl.Enabled = true;
                                        }
                                    }// if

                                }
                                else
                                {
                                    if (ctrlObj != null)
                                    {
                                        if (ctrlObj is Control)
                                        {
                                            ctrl = (Control)ctrlObj;
                                            ctrl.Enabled = false;
                                            / *
                                            if (ctrl is TabPage)
                                            {
                                                // 递归enable child controls
                                                m_ownerAddin.transChildCtrl(ctrl, false);
                                            }
                                            else if (!(ctrl is TabControl))
                                            {
                                                // 递归enable child controls
                                                m_ownerAddin.transChildCtrl(ctrl, false);
                                            }
                                             * * /
                                        }
                                        else if (ctrlObj is ToolStripItem)
                                        {
                                            item = (ToolStripItem)ctrlObj;
                                            item.Enabled = false;
                                        }
                                        else if (ctrlObj is RibbonGroup)
                                        {
                                            // 
                                        }
                                        else if (ctrlObj is RibbonControl)
                                        {
                                            ribCtrl = (RibbonControl)ctrlObj;
                                            ribCtrl.Enabled = false;
                                        }
                                    }// if
                                }

                            }

                            userPane.RefreshRelsByPermission();
                            userPane.cloneShareLibTree();
                            userPane.Invalidate();
                        }

                    }
                }

                MessageBox.Show("激活成功");
            }*/


            return;
        }

        private void nav2(Word.WdGoToItem gtItem, Word.WdGoToDirection gtDir)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("没有活动的文档");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            sel.GoTo(gtItem, gtDir, 1, "");
            doc.ActiveWindow.SetFocus();
            return;
        }


        private void btnNavFirst_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("没有活动的文档");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Bookmark firstBkmk = null;
            Word.Bookmark lastBkmk = null;
            Word.Bookmark nearstPrevBkmk = null;
            Word.Bookmark nearstNextBkmk = null;

            int nRet = m_ownerAddin.m_commTools.getNavKeyWordBookmk(doc, m_strNavBkmkNamePrefix, ref firstBkmk, ref lastBkmk, ref nearstPrevBkmk, ref nearstNextBkmk);

            Object miss = Type.Missing;

            if (firstBkmk != null)
            {
                sel.GoTo(Word.WdGoToItem.wdGoToBookmark, miss, miss, firstBkmk.Name);
                doc.ActiveWindow.SetFocus();
            }

            return;
        }

        private void btnNavLast_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("没有活动的文档");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.Bookmark firstBkmk = null;
            Word.Bookmark lastBkmk = null;
            Word.Bookmark nearstPrevBkmk = null;
            Word.Bookmark nearstNextBkmk = null;

            int nRet = m_ownerAddin.m_commTools.getNavKeyWordBookmk(doc, m_strNavBkmkNamePrefix, ref firstBkmk, ref lastBkmk, ref nearstPrevBkmk, ref nearstNextBkmk);

            Object miss = Type.Missing;

            if (lastBkmk != null)
            {
                sel.GoTo(Word.WdGoToItem.wdGoToBookmark, miss, miss, lastBkmk.Name);
                doc.ActiveWindow.SetFocus();
            }

            return;
        }

        private void ribbtnOpenVerDir_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("没有活动的文档");
                return;
            }
            finally
            {
            }

            String strExt = Path.GetExtension(doc.FullName);

            if (String.IsNullOrWhiteSpace(strExt))
            {
                MessageBox.Show("请先保存当前文档");
                return;
            }

            String strDir = Path.GetDirectoryName(doc.FullName);

            String strVerDir = strDir + "\\localVer";

            if (!Directory.Exists(strVerDir))
            {
                // Directory.CreateDirectory(strVerDir);
                MessageBox.Show("没有本地版本");
                return;
            }

            try
            {
            	System.Diagnostics.Process.Start(strVerDir);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }

            return;
        }

        private void chkGenLocalVer_Click(object sender, RibbonControlEventArgs e)
        {
            Settings.Default.bGenLocalVer = chkGenLocalVer.Checked;
            Settings.Default.Save();
            return;
        }

        private void ribBtnRemoveJetNav_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = app.ActiveWindow.Selection;

            Word.Bookmarks bks = sel.Range.Bookmarks;// doc.Bookmarks;

            Boolean bFnd = false;

            foreach (Word.Bookmark bkmk in bks)
            {
                if (bkmk.Name.StartsWith(m_strNavBkmkNamePrefix))
                {
                    bFnd = true;
                    bkmk.Delete();
                }
            }

            // m_uBookmarkSn = 0;
            if (bFnd)
            {
                MessageBox.Show("完成");
            }
            else
            {
                MessageBox.Show("选择区无快捷导航记录点");
            }

            return;
        }



        private char[] m_trimChars = new char[] { ' ', '\t', '\r', '\n', '\a' };

        private String splitPostfixNum(String strSrc, ref String strNum, ref int nNumWidth)
        {
            if (String.IsNullOrWhiteSpace(strSrc))
            {
                strNum = "";
                nNumWidth = 0;

                return strSrc;
            }

            String strPrefix = strSrc;
            char[] chArr = strSrc.ToCharArray();
            int nIndex = -1;
            Boolean bAllNum = true;
            nNumWidth = 0;

            for (int i = chArr.GetLength(0) - 1; i >= 0; i--)
            {
                if (chArr[i] >= '0' && chArr[i] <= '9')
                {
                    bAllNum = true;
                }
                else
                {
                    nIndex = i;
                    bAllNum = false;
                    break;
                }
            }

            if (bAllNum)
            {
                strNum = strSrc.Substring(nIndex + 1);
                strPrefix = "";

                int nNum = 0;

                if (int.TryParse(strNum, out nNum))
                {
                    String strActualNum = nNum.ToString();

                    if (!strActualNum.Equals(strNum))
                    {
                        nNumWidth = strNum.Length;
                    }
                    else
                    {
                        nNumWidth = 0;
                    }

                }
            }
            else
            {
                if (nIndex != -1)
                {
                    strNum = strSrc.Substring(nIndex + 1);
                    strPrefix = strSrc.Substring(0, nIndex + 1);

                    int nNum = 0;

                    if (int.TryParse(strNum, out nNum))
                    {
                        String strActualNum = nNum.ToString();

                        if (!strActualNum.Equals(strNum))
                        {
                            nNumWidth = strNum.Length;
                        }
                        else
                        {
                            nNumWidth = 0;
                        }

                    }

                }
                else
                {
                    strNum = "";
                    strPrefix = strSrc;

                    nNumWidth = 0;
                }
            }

            return strPrefix;
        }


        private void ribBtnFillSn_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;// 测试代码
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            // Word.Range scopeRng = null;
            Word.Range oRng = sel.Range;

            String strInit = sel.Paragraphs[1].Range.Text;
            ArrayList arrDstParas = new ArrayList();

            if (sel.Paragraphs.Count <= 1)
            {
                Boolean bInTbl = sel.Range.get_Information(Word.WdInformation.wdWithInTable);

                if (!bInTbl)
                {
                    MessageBox.Show("未选择段落时，请在表格中使用");
                    return;
                }

                Word.Table tbl = sel.Tables[1];

                int nRowNum = sel.Range.get_Information(Word.WdInformation.wdStartOfRangeRowNumber);
                int nColNum = sel.Range.get_Information(Word.WdInformation.wdStartOfRangeColumnNumber);

                Word.Cell cel = null;
                for (int i = nRowNum; i <= tbl.Rows.Count; i++)
                {
                    try
                    {
                        cel = tbl.Cell(i, nColNum);

                        if (cel != null)
                        {
                            arrDstParas.Add(cel.Range.Paragraphs[1]);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        continue;
                    }
                    finally
                    {

                    }
                }
            }
            else
            {
                foreach (Word.Paragraph para in sel.Paragraphs)
                {
                    arrDstParas.Add(para);
                }
            }

            FillSn(arrDstParas,"智能填充");

            return;
        }

        private Boolean IsTimeFormat(string str)
        {
            Boolean isDateTime = false;
            // XX时XX分XX秒  
            if (Regex.IsMatch(str, "^(?<hour>\\d{1,2})(时|点)(?<minute>\\d{1,2})分((?<second>\\d{1,2})秒)?$"))
                isDateTime = true;
            // XX时XX分XX秒  
            else if (Regex.IsMatch(str, "^((零|一|二|三|四|五|六|七|八|九|十){1,3})(时|点)((零|一|二|三|四|五|六|七|八|九|十){1,3})分(((零|一|二|三|四|五|六|七|八|九|十){1,3})秒)?$"))
                isDateTime = true;
            // XX分XX秒  
            else if (Regex.IsMatch(str, "^(?<minute>\\d{1,2})分(?<second>\\d{1,2})秒$"))
                isDateTime = true;
            // XX分XX秒  
            else if (Regex.IsMatch(str, "^((零|一|二|三|四|五|六|七|八|九|十){1,3})分((零|一|二|三|四|五|六|七|八|九|十){1,3})秒$"))
                isDateTime = true;
            // XX时  
            else if (Regex.IsMatch(str, "\\b(?<hour>\\d{1,2})(时|点钟)\\b"))
                isDateTime = true;
            else
                isDateTime = false;

            return isDateTime;
        }



        private Boolean IsDateFormat(String str, ref String strDateFmt, ref Boolean bDay)
        {
            Boolean isDateTime = false;
            bDay = false;

            // yyyy/MM/dd  
            if (Regex.IsMatch(str, "^(?<year>\\d{2,4})/(?<month>\\d{1,2})(/(?<day>\\d{1,2}))?$"))
            {
                Regex rgx = new Regex("^(?<year>\\d{2,4})/(?<month>\\d{1,2})(/(?<day>\\d{1,2}))?$");

                MatchCollection mcths = rgx.Matches(str);

                String strYear = mcths[0].Groups[2].Value;
                String strMon = mcths[0].Groups[3].Value;
                String strDay = mcths[0].Groups[4].Value;

                int i = 1;
                String strYFmt = "", strMFmt = "", strDFmt = "";
                for (i = 1; i <= strYear.Length; i++)
                {
                    strYFmt += "y";
                }

                int nMon = int.Parse(strMon);

                if (nMon < 10)
                {
                    if (nMon.ToString().Length != strMon.Length)
                    {
                        strMFmt = "MM";
                    }
                    else
                    {
                        strMFmt = "M";
                    }
                }
                else
                {
                    strMFmt = "M";
                }

                if (strDay.Length > 0)
                {
                    int nDay = int.Parse(strDay);

                    if (nDay < 10)
                    {
                        if (nDay.ToString().Length != strDay.Length)
                        {
                            strDFmt = "dd";
                        }
                        else
                        {
                            strDFmt = "d";
                        }
                    }
                    else
                    {
                        strDFmt = "d";
                    }
                }

                strDateFmt = "";

                if (strYFmt.Length > 0)
                    strDateFmt += strYFmt;

                if (strMFmt.Length > 0)
                    strDateFmt += "/" + strMFmt;

                if (strDFmt.Length > 0)
                {
                    bDay = true;
                    strDateFmt += "/" + strDFmt;
                }

                isDateTime = true;
            }
            // yyyy-MM-dd   
            else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})-(?<month>\\d{1,2})(-(?<day>\\d{1,2}))?$"))
            {
                Regex rgx = new Regex("^(?<year>\\d{2,4})-(?<month>\\d{1,2})(-(?<day>\\d{1,2}))?$");

                MatchCollection mcths = rgx.Matches(str);

                String strYear = mcths[0].Groups[2].Value;
                String strMon = mcths[0].Groups[3].Value;
                String strDay = mcths[0].Groups[4].Value;

                int i = 1;
                String strYFmt = "", strMFmt = "", strDFmt = "";
                for (i = 1; i <= strYear.Length; i++)
                {
                    strYFmt += "y";
                }

                int nMon = int.Parse(strMon);

                if (nMon < 10)
                {
                    if (nMon.ToString().Length != strMon.Length)
                    {
                        strMFmt = "MM";
                    }
                    else
                    {
                        strMFmt = "M";
                    }
                }
                else
                {
                    strMFmt = "M";
                }

                if (strDay.Length > 0)
                {
                    int nDay = int.Parse(strDay);

                    if (nDay < 10)
                    {
                        if (nDay.ToString().Length != strDay.Length)
                        {
                            strDFmt = "dd";
                        }
                        else
                        {
                            strDFmt = "d";
                        }
                    }
                    else
                    {
                        strDFmt = "d";
                    }
                }

                strDateFmt = "";

                if (strYFmt.Length > 0)
                    strDateFmt += strYFmt;

                if (strMFmt.Length > 0)
                    strDateFmt += "-" + strMFmt;

                if (strDFmt.Length > 0)
                {
                    bDay = true;
                    strDateFmt += "-" + strDFmt;
                }

                isDateTime = true;
            }
            // yyyy.MM.dd   
            else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})[.](?<month>\\d{1,2})([.](?<day>\\d{1,2}))?$"))
            {
                Regex rgx = new Regex("^(?<year>\\d{2,4})[.](?<month>\\d{1,2})([.](?<day>\\d{1,2}))?$");

                MatchCollection mcths = rgx.Matches(str);

                String strYear = mcths[0].Groups[2].Value;
                String strMon = mcths[0].Groups[3].Value;
                String strDay = mcths[0].Groups[4].Value;

                int i = 1;
                String strYFmt = "", strMFmt = "", strDFmt = "";
                for (i = 1; i <= strYear.Length; i++)
                {
                    strYFmt += "y";
                }

                int nMon = int.Parse(strMon);

                if (nMon < 10)
                {
                    if (nMon.ToString().Length != strMon.Length)
                    {
                        strMFmt = "MM";
                    }
                    else
                    {
                        strMFmt = "M";
                    }
                }
                else
                {
                    strMFmt = "M";
                }

                if (strDay.Length > 0)
                {
                    int nDay = int.Parse(strDay);

                    if (nDay < 10)
                    {
                        if (nDay.ToString().Length != strDay.Length)
                        {
                            strDFmt = "dd";
                        }
                        else
                        {
                            strDFmt = "d";
                        }
                    }
                    else
                    {
                        strDFmt = "d";
                    }
                }

                strDateFmt = "";

                if (strYFmt.Length > 0)
                    strDateFmt += strYFmt;

                if (strMFmt.Length > 0)
                    strDateFmt += "." + strMFmt;

                if (strDFmt.Length > 0)
                {
                    bDay = true;
                    strDateFmt += "." + strDFmt;
                }

                isDateTime = true;
            }
            // yyyy年MM月dd日  
            else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})年(?<month>\\d{1,2})月((?<day>\\d{1,2})日)?$"))
            {
                Regex rgx = new Regex("^(?<year>\\d{2,4})年(?<month>\\d{1,2})月((?<day>\\d{1,2})日)?$");

                MatchCollection mcths = rgx.Matches(str);

                String strYear = mcths[0].Groups[2].Value;
                String strMon = mcths[0].Groups[3].Value;
                String strDay = mcths[0].Groups[4].Value;

                int i = 1;
                String strYFmt = "", strMFmt = "", strDFmt = "";
                for (i = 1; i <= strYear.Length; i++)
                {
                    strYFmt += "y";
                }

                int nMon = int.Parse(strMon);

                if (nMon < 10)
                {
                    if (nMon.ToString().Length != strMon.Length)
                    {
                        strMFmt = "MM";
                    }
                    else
                    {
                        strMFmt = "M";
                    }
                }
                else
                {
                    strMFmt = "M";
                }

                if (strDay.Length > 0)
                {
                    int nDay = int.Parse(strDay);

                    if (nDay < 10)
                    {
                        if (nDay.ToString().Length != strDay.Length)
                        {
                            strDFmt = "dd";
                        }
                        else
                        {
                            strDFmt = "d";
                        }
                    }
                    else
                    {
                        strDFmt = "d";
                    }
                }

                strDateFmt = "";

                if (strYFmt.Length > 0)
                    strDateFmt += strYFmt + "年";

                if (strMFmt.Length > 0)
                    strDateFmt += strMFmt + "月";

                if (strDFmt.Length > 0)
                {
                    bDay = true;
                    strDateFmt += strDFmt + "日";
                }

                isDateTime = true;
            }
            /*
            // yyyy年MM月dd日  
            else if (Regex.IsMatch(str, "^((?<year>\\d{2,4})年)?(正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}日)?$"))
                isDateTime = true;

            // yyyy年MM月dd日  
            else if (Regex.IsMatch(str, "^(零|〇|一|二|三|四|五|六|七|八|九|十){2,4}年((正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}(日)?)?)?$"))
                isDateTime = true;
            // yyyy年  
            //else if (Regex.IsMatch(str, "^(?<year>\\d{2,4})年$"))  
            //    isDateTime = true;  

            // 农历1  
            else if (Regex.IsMatch(str, "^(甲|乙|丙|丁|戊|己|庚|辛|壬|癸)(子|丑|寅|卯|辰|巳|午|未|申|酉|戌|亥)年((正|一|二|三|四|五|六|七|八|九|十|十一|十二)月((一|二|三|四|五|六|七|八|九|十){1,3}(日)?)?)?$"))
                isDateTime = true;
            // 农历2  
            else if (Regex.IsMatch(str, "^((甲|乙|丙|丁|戊|己|庚|辛|壬|癸)(子|丑|寅|卯|辰|巳|午|未|申|酉|戌|亥)年)?(正|一|二|三|四|五|六|七|八|九|十|十一|十二)月初(一|二|三|四|五|六|七|八|九|十)$"))
                isDateTime = true;
             * */
            else
            {
                isDateTime = false;
            }

            return isDateTime;
        }

        private int splitNum(String strValue, ref String strPrefix, ref String strPostfix, ref String strNum)
        {
            int nNum = -1;

            char[] chArr = strValue.ToCharArray();
            int nLen = chArr.GetLength(0);
            int nEndPos = -1, nStartPos = -1;

            for (int i = nLen - 1; i >= 0; i--)
            {
                if (chArr[i] >= '0' && chArr[i] <= '9')
                {
                    if (nEndPos == -1)
                    {
                        nEndPos = i;
                    }
                    else
                    {
                        nStartPos = i;
                    }
                }
                else
                {
                    if (nEndPos != -1)
                    {
                        break;
                    }
                }
            }

            if (nEndPos == -1)
            {
                if (nStartPos == -1)
                {
                    nNum = -1;
                    strNum = "";
                    strPrefix = "";
                    strPostfix = "";

                    return nNum;
                }
                else
                {
                    // impossible
                }
            }
            else
            {
                if (nStartPos == -1)
                {
                    // nStartPos = 0;
                    nStartPos = nEndPos;
                }
                else
                {

                }
            }


            strPrefix = strValue.Substring(0, nStartPos);
            strPostfix = strValue.Substring(nEndPos + 1, nLen - nEndPos - 1);
            strNum = strValue.Substring(nStartPos, nEndPos - nStartPos + 1);

            if (int.TryParse(strNum, out nNum))
            {
                return nNum;
            }
            
            return nNum;
        }

        private void FillSn(ArrayList arrDstParas,String strUndoName = "")
        {
            Word.Application app = Globals.ThisAddIn.Application;// 测试代码
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            Boolean bNumOrDate = true; // true -- num, false -- date
            int nStartNum = 0;
            String strNumPrefix = "", strNumPostfix = "";
            String strDateFmt = "yyyy-M-d";
            Boolean bDayOrMonth = false;
            String strNum = "", strNumFmt = "0";
            Boolean bJump1stPara = false;
            DateTime dt = DateTime.Now;


            if (arrDstParas.Count == 0)
                return;

            String strInit = ((Word.Paragraph)arrDstParas[0]).Range.Text;

            strInit = strInit.Trim(m_trimChars);

            if (String.IsNullOrWhiteSpace(strInit))
            {
                bNumOrDate = true;
                bJump1stPara = false;
            }
            else
            {
                bJump1stPara = true;
                Boolean bValidDate = false;

                if (IsDateFormat(strInit, ref strDateFmt, ref bDayOrMonth))
                {
                    if (DateTime.TryParse(strInit, out dt))
                    {
                        // parse string's date format (yyyy-mm-dd, or ...)

                        bNumOrDate = false;
                        bValidDate = true;
                    }
                }

                if (!bValidDate)
                {
                    // parse string num
                    int nRetNum = splitNum(strInit, ref strNumPrefix, ref strNumPostfix, ref strNum);

                    if (nRetNum == -1)
                    {
                    }
                    else
                    {
                        nStartNum = nRetNum;

                        if (nRetNum.ToString("0").Length != strNum.Length)
                        {
                            strNumFmt = "";
                            for (int i = 0; i < strNum.Length; i++)
                            {
                                strNumFmt += "0";
                            }
                        }
                    }

                    bNumOrDate = true;
                }
            }

            if (bJump1stPara)
            {
                if (arrDstParas.Count > 0)
                {
                    arrDstParas.RemoveAt(0);
                }
            }


            if (m_ownerAddin.AppVersion >= 15 && arrDstParas.Count > 0) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.StartCustomRecord(strUndoName);
            }


            foreach (Word.Paragraph paraItem in arrDstParas)
            {
                if (bNumOrDate)
                {
                    nStartNum++;
                    paraItem.Range.Select();
                    sel.TypeText(strNumPrefix + nStartNum.ToString(strNumFmt) + strNumPostfix);
                }
                else
                {
                    if (bDayOrMonth)
                    {
                        dt = dt.AddDays(1.0);
                    }
                    else
                    {
                        dt = dt.AddMonths(1);
                    }

                    paraItem.Range.Select();
                    sel.TypeText(dt.ToString(strDateFmt));
                }
            }

            if (m_ownerAddin.AppVersion >= 15 && arrDstParas.Count > 0) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }


            return;
        }



        private void ribBtnFillSn2EndRow_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;// 测试代码
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            // Word.Range scopeRng = null;
            Word.Range oRng = sel.Range;

            String strInit = sel.Paragraphs[1].Range.Text;
            ArrayList arrDstParas = new ArrayList();

            Boolean bInTbl = sel.Range.get_Information(Word.WdInformation.wdWithInTable);

            if (!bInTbl)
            {
                MessageBox.Show("请在表格中使用");
                return;
            }

            Word.Table tbl = sel.Tables[1];

            int nRowNum = sel.Range.get_Information(Word.WdInformation.wdStartOfRangeRowNumber);
            int nColNum = sel.Range.get_Information(Word.WdInformation.wdStartOfRangeColumnNumber);

            Word.Cell cel = null;
            for (int i = nRowNum; i <= tbl.Rows.Count; i++)
            {
                try
                {
                    cel = tbl.Cell(i, nColNum);

                    if (cel != null)
                    {
                        arrDstParas.Add(cel.Range.Paragraphs[1]);
                    }
                }
                catch (System.Exception ex)
                {
                    continue;
                }
                finally
                {

                }
            }

            FillSn(arrDstParas, "填充至表末行");
            return;
        }


        private void ribBtnFillSelection_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;// 测试代码
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            // Word.Range scopeRng = null;
            ArrayList arrDstParas = new ArrayList();

            foreach (Word.Paragraph para in sel.Paragraphs)
            {
                arrDstParas.Add(para);
            }

            FillSn(arrDstParas, "填充选择区");
            return;
        }


        private void FillSnNum(Word.Selection sel, String strPrefix, String strPostfix, int nStart, String strFmt, ArrayList arrParas)
        {
            Word.Paragraph tmpPara = null;
            for (int i = 0; i < arrParas.Count; i++)
            {
                tmpPara = (Word.Paragraph)arrParas[i];

                tmpPara.Range.Select();
                nStart++;

                // if (nNumWidth == 0)
                if(String.IsNullOrWhiteSpace(strFmt))
                {
                    sel.TypeText(strPrefix + nStart + strPostfix);
                }
                else
                {
                    sel.TypeText(strPrefix + nStart.ToString(strFmt) + strPostfix);
                }
            }

            return;
        }


        private void FillSnDate()
        {
            DateTime dt = DateTime.Now;

            //dt.Add
            return;
        }

        private void setOutlineLevel(int nSelIndex)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }


            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;

                if (nSelIndex < (int)Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    ur.StartCustomRecord("设置大纲级别：" + nSelIndex);
                }
                else
                {
                    ur.StartCustomRecord("设置大纲级别：正文");
                }
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(nSelIndex);
            Object objStyle = doc.Styles[styleIndex];
            sel.Paragraphs.set_Style(objStyle);

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            return;
        }

        private void ribBtnOutLevel1_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(1);
            return;
        }

        private void ribBtnOutLevel2_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(2);
            return;
        }

        private void ribBtnOutLevel3_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(3);
            return;
        }

        private void ribBtnOutLevel4_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(4);
            return;
        }

        private void ribBtnOutLevel5_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(5);
            return;
        }

        private void ribBtnOutLevel6_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(6);
            return;
        }

        private void ribBtnOutLevel7_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(7);
            return;
        }

        private void ribBtnOutLevel8_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(8);
            return;
        }

        private void ribBtnOutLevel9_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(9);
            return;
        }

        private void ribBtnOutLevelTextBody_Click(object sender, RibbonControlEventArgs e)
        {
            setOutlineLevel(10);
            return;
        }


        private void unitedHeadersFooters(Boolean bFooter = false)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
            // 切换到normal view
            doc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;


            Word.Selection sel = doc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Range curRng = sel.Range;

            sel.Collapse();

            // 将当前节的页脚统一到其它节
            int nIndex = -1, nCurSecIndex = -1;
            foreach (Word.Section sec in doc.Sections)
            {
                nIndex++;

                if (curRng.InRange(sec.Range) && nCurSecIndex == -1)
                {
                    nCurSecIndex = nIndex;
                    break;
                }
            }

            frmChooseSections frmChSecs = new frmChooseSections();

            frmChSecs.lblInfo.Text = "当前在：节" + (nCurSecIndex + 1);

            frmChSecs.chkListBox.Items.Clear();

            int nCnt = doc.Sections.Count;

            for (int i = 0; i < nCnt; i++)
            {
                if (nCurSecIndex == i)
                    continue;

                frmChSecs.chkListBox.Items.Add("节" + (i + 1));
            }

            DialogResult res = frmChSecs.ShowDialog();

            if (res == DialogResult.No)
            {
                return;
            }

            if (frmChSecs.chkListBox.CheckedItems.Count == 0)
            {
                MessageBox.Show("请选择至少一项");
                return;
            }


            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                if (bFooter)
                {
                    ur.StartCustomRecord("统一用当前页脚");
                }
                else
                {
                    ur.StartCustomRecord("统一用当前页眉");
                }
            }

            Boolean[] bSecs = new Boolean[nCnt];

            for (int i = 0; i < nCurSecIndex; i++)
            {
                bSecs[i] = frmChSecs.chkListBox.GetItemChecked(i);
            }

            bSecs[nCurSecIndex] = false;

            for (int i = nCurSecIndex + 1; i < nCnt; i++)
            {
                bSecs[i] = frmChSecs.chkListBox.GetItemChecked(i - 1);
            }

            if (bFooter)
            {
                doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
            }
            else
            {
                doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
            }

            sel.WholeStory();
            sel.Copy();

            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            // 将当前节的页眉统一到其它节
            nIndex = -1;
            foreach (Word.Section sec in doc.Sections)
            {
                nIndex++;

//                 if (curRng.InRange(sec.Range) || !bSecs[nIndex])
//                 {
//                     continue;
//                 }

                if (!bSecs[nIndex])
                {
                    continue;
                }

                sel.Start = sec.Range.Start;
                sel.End = sec.Range.Start;
                sec.Range.GoTo();
                doc.ActiveWindow.ScrollIntoView(sec.Range);

                if (bFooter)
                {
                    doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageFooter;
                }
                else
                {
                    doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader;
                }

                sel.WholeStory();

                sel.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);
                // sel.Paste();

                sel.EndKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
                sel.MoveLeft();
                sel.Delete();

                doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            }

            // 恢复特定view
            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
            else
            {
                // doc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
                doc.ActiveWindow.View.Type = oViewType;
            }

            // restore original position
            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range); // 视角恢复

            if (m_ownerAddin.AppVersion >= 15) // 2013
            {
                Word.UndoRecord ur = m_ownerAddin.Application.UndoRecord;
                ur.EndCustomRecord();
            }

            MessageBox.Show("完成");

            return;
        }


        private void ribbtnUnitedHeaders_Click(object sender, RibbonControlEventArgs e)
        {
            unitedHeadersFooters();
           
            return;
        }


        private void ribbtnUnitedFooters_Click(object sender, RibbonControlEventArgs e)
        {
            unitedHeadersFooters(true);

            return;
        }

        private void btnLocalVerMileStone_Click(object sender, RibbonControlEventArgs e)
        {

            Word.Document Doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                Doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            String strExt = Path.GetExtension(Doc.FullName);
            String strPath = Path.GetFullPath(Doc.FullName);

            // if (/*Settings.Default.bGenLocalVer && */m_ownerAddin.searchPermission("chkGenLocalVer") > 0)
            if (/*Settings.Default.bGenLocalVer && */m_ownerAddin.searchPermission("groupLocalVer") > 0)
            {
                Boolean bRet = Settings.Default.bGenLocalVer;

                if (String.IsNullOrWhiteSpace(strExt))
                {
                    MessageBox.Show("请先保存文档后再使用本功能");
                    return;
                }

                String strDir = Path.GetDirectoryName(Doc.FullName);

                String strVerDir = strDir + "\\localVer\\";

                if (!Directory.Exists(strVerDir))
                {
                    try
                    {
                        Directory.CreateDirectory(strVerDir);
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                    finally
                    {
                    }
                }

                if (Directory.Exists(strVerDir))
                {
                    String strOnlyNameNoExt = Path.GetFileNameWithoutExtension(Doc.FullName);
                    String strTrimTimestampNameNoExt = strOnlyNameNoExt;

                    // 截断当前文档右侧符合此模式的时间戳，避免连续
                    // 
                    String[] strsParts = strOnlyNameNoExt.Split('_');
                    int nNum = strsParts.GetLength(0);

                    if (nNum >= 3)
                    {
                        String strTimeStamp = strsParts[nNum - 3] + strsParts[nNum - 2] + strsParts[nNum - 1];

                        String[] format = { "yyyyMMddhhmmssff" };
                        DateTime date;
                        if (DateTime.TryParseExact(strTimeStamp,
                                                   format,
                                                   System.Globalization.CultureInfo.InvariantCulture,
                                                   System.Globalization.DateTimeStyles.None,
                                                   out date))
                        {
                            int nIndex = strOnlyNameNoExt.LastIndexOf("_" + strsParts[nNum - 3]);
                            if (nIndex > 0)
                            {
                                strTrimTimestampNameNoExt = strOnlyNameNoExt.Substring(0, nIndex);
                            }
                        }
                    }
                    // 

                    String strPostx = DateTime.Now.ToString("yyyyMMdd_hhmmss_ff");

                    if (nNum > 3 && "K".Equals(strsParts[nNum - 4]))
                    {
                        int nIndex = strTrimTimestampNameNoExt.LastIndexOf("_K");
                        if (nIndex > 0)
                        {
                            strTrimTimestampNameNoExt = strTrimTimestampNameNoExt.Substring(0, nIndex);
                        }
                    }

                    String strTmpFile = strVerDir + strTrimTimestampNameNoExt + "_K_" + strPostx + strExt;

                    if (File.Exists(strTmpFile)) // 判断存在
                    {
                        try
                        {
                            File.Delete(strTmpFile);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }
                        finally
                        {
                        }
                    }

                    if (!File.Exists(strTmpFile))
                    {
                        try
                        {
                            Settings.Default.bGenLocalVer = false;
                            Doc.Save();
                            Settings.Default.bGenLocalVer = bRet;

                            File.Copy(Doc.FullName, strTmpFile);
                            MessageBox.Show("创建关键版本成功：\r\n" + strTmpFile);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }
                        finally
                        {
                        }

                    }// if
                }
            }
            else
            {
                MessageBox.Show("无此功能权限","错误");
            }

            return;
        }

        private void ribBtnPrevEditPos_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            app.Browser.Target = Word.WdBrowseTarget.wdBrowseEdit;

            app.Browser.Previous();

            return;
        }


        private void ribBtnNextEditPos_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = m_ownerAddin.Application;
            app.Browser.Target = Word.WdBrowseTarget.wdBrowseEdit;

            app.Browser.Next();

            return;
        }

        private void ribBtnJump2Toc_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if (doc.TablesOfContents.Count > 0)
            {
                Word.Range rng = doc.TablesOfContents[1].Range;

                sel.Start = rng.Start;
                sel.End = sel.Start;

                sel.Range.GoTo();
                doc.ActiveWindow.ScrollIntoView(rng);
            }
            else
            {
                MessageBox.Show("本文档没有目录");
            }

            return;
        }


        private Boolean drawNumber(String strTxt, ref String strStartNum)
        {
            char[] chs = strTxt.ToCharArray();
            int nStartIndex = -1, nEndIndex = -1;
            int nIndex = -1;

            foreach (char ch in chs)
            {
                nIndex ++;

                if (ch == '-' || ch == '+')
                {
                    if (nStartIndex == -1)
                    {
                        nStartIndex = nIndex;
                        nEndIndex = nIndex;
                    }
                    else
                    {
                        break;
                    }
                }
                else if (ch == '.')
                {
                    if (nStartIndex == -1)
                    {
                        nStartIndex = nIndex;
                    }

                    nEndIndex = nIndex;
                }
                else if (ch >= '0' && ch <= '9')
                {
                    // 
                    if (nStartIndex == -1)
                    {
                        nStartIndex = nIndex;
                    }

                    nEndIndex = nIndex;
                }
                else
                {
                    if (nStartIndex != -1)
                    {
                        break;
                    }
                }

                //if ((ch >= '0' && ch <= '9') || ch == '.' || ch == 'e' || ch == '-' || ch == '+')
                //{
                //    if (nStartIndex == -1)
                //    {
                //        nStartIndex = nIndex;
                //        nEndIndex = nIndex;
                //    }
                    
                //    nEndIndex = nIndex;
                //}
                //else
                //{
                //    if (nStartIndex != -1)
                //    {
                //        break;
                //    }
                //}

            }

            if (nStartIndex != -1 )
            {
                strStartNum = strTxt.Substring(nStartIndex, nEndIndex - nStartIndex + 1);
                return true;
            }

            return false;
        }


        private void rbBtnCalculate_Click(object sender, RibbonControlEventArgs e)
        {

            Word.Application app = m_ownerAddin.Application;
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if (sel.End - sel.Start <= 1)
            {
                MessageBox.Show("请选择至少2个带数值的段落");
                return;
            }


            FormBasicCalculate frmCalc = new FormBasicCalculate();

            String strCnt = "";

            if (m_ownerAddin.m_bAppIsWps)
            {
                Boolean bInTbl = sel.Range.get_Information(Word.WdInformation.wdWithInTable);

                if (bInTbl)
                {
                    foreach (Word.Cell para in sel.Cells)
                    {
                        strCnt = para.Range.Text.Trim(m_ownerAddin.m_commTools.m_trimChars);
                        frmCalc.lstBoxSelDataItems.Items.Add(strCnt);
                    }
                }
                else
                {
                    foreach (Word.Paragraph para in sel.Paragraphs)
                    {
                        strCnt = para.Range.Text.Trim(m_ownerAddin.m_commTools.m_trimChars);
                        frmCalc.lstBoxSelDataItems.Items.Add(strCnt);
                    }
                }
            }
            else
            {

                foreach (Word.Paragraph para in sel.Paragraphs)
                {
                    strCnt = para.Range.Text.Trim(m_ownerAddin.m_commTools.m_trimChars);

                    frmCalc.lstBoxSelDataItems.Items.Add(strCnt);
                }
            }



            String strSum = "", strAvg = "", strMax = "", strMin = "", strCount = "";
            double dbVal = 0.0, dbMax = double.MinValue, dbMin = double.MaxValue;
            double dbSum = 0.0, dbAvg = 0.0;
            int nCnt = 0;

            String strNum = "",strSelNum = "";

            foreach (String strItem in frmCalc.lstBoxSelDataItems.Items)
            {
                // find the first number
                if (drawNumber(strItem,ref strNum))
                {
                    if (double.TryParse(strNum, out dbVal))
                    {
                        strSelNum += strNum + " ";

                        dbSum += dbVal;
                        nCnt++;

                        if (dbVal > dbMax)
                        {
                            dbMax = dbVal;
                        }

                        if (dbVal < dbMin)
                        {
                            dbMin = dbVal;
                        }

                    }
                }
            }

            if (nCnt > 0)
            {
                dbAvg = dbSum / nCnt;

                strSum = dbSum.ToString();
                strAvg = dbAvg.ToString();
                strMax = dbMax.ToString();
                strMin = dbMin.ToString();
                strCount = nCnt.ToString();
            }
            else
            {
                strSelNum = "(无)";
                strSum = "--";
                strAvg = "--";
                strMax = "--";
                strMin = "--";
                strCount = "--";
            }

            String strLastResult = "";

            strLastResult = "合  计：" + strSum + "\r\n" +
                            "计  数：" + strCount + "\r\n" +
                            "均  值：" + strAvg + "\r\n" +
                            "最大值：" + strMax + "\r\n" +
                            "最小值：" + strMin + "\r\n\r\n" +
                            "合法数值：" + strSelNum; 

            frmCalc.txtBoxCalcResult.Text = strLastResult;

            frmCalc.ShowDialog();
            // frmCalc.txtBoxCalcResult.SelectAll();

            return;
        }

        private void ribLoadSoloLic_Click(object sender, RibbonControlEventArgs e)
        {
            FormLoadLicFile frmLoadLicFile = new FormLoadLicFile();

            frmLoadLicFile.setItems(m_ownerAddin.m_commTools, m_ownerAddin.m_edtCenter);

            DialogResult res = frmLoadLicFile.ShowDialog();

            if (res == DialogResult.OK && frmLoadLicFile.m_bLicLegal)
            {
                String strLicFile = frmLoadLicFile.txtBoxSelectedLicFileLoc.Text;

                int nRet = m_ownerAddin.m_edtCenter.LoadSoloLic(strLicFile);

                if (nRet != 0)
                {
                    MessageBox.Show("加载失败，请确保当前使用的许可文件未被占用");
                    return;
                }

                // 刷新UI
                Word.Application app = m_ownerAddin.Application;
                Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

                try
                {
                    doc = app.ActiveDocument;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("无活动文档，不能应用");
                    return;
                }
                finally
                {
                }

                m_ownerAddin.updateUI(doc);

                MessageBox.Show("加载完成");
            }

            return;
        }

        private void ribBtnViewOutlineLevel_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

            try
            {
                doc = m_ownerAddin.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return;
            }
            finally
            {
            }

            Word.Selection sel = doc.ActiveWindow.Selection;
            Word.Paragraph para = sel.Paragraphs[1];

            dynamic style = para.get_Style();

            if (para.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                MessageBox.Show("当前段落大纲级别：正文\r\n\r\n" + "样式描述：" + style.Description);
            }
            else
            {
                String strHeadingStyleName = "";

                try
                {
                    strHeadingStyleName = ",样式名称：" + style.NameLocal;
                }
                catch (System.Exception ex)
                {
                    strHeadingStyleName = "";
                }


                MessageBox.Show("当前段落大纲级别：章节，" + (int)para.OutlineLevel + "级" + strHeadingStyleName +
                                "\r\n\r\n样式描述：" + style.Description);
            }

            return;
        }




     }
}
