using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using OfficeTools.Common;
using Office=Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using System.Collections;
using System.Collections.Specialized;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using OfficeAssist.Properties;
//using OfficeAssist.localdbDataSetTableAdapters;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using AutoUpdate;
using NHibernate;
using OfficeAssist.localDB.Model;
using OfficeAssist.localDB.Util;
using OfficeAssist.docPub;

namespace OfficeAssist
{
    public partial class ThisAddIn
    {
        public ArrayList m_arrFont = new ArrayList();
        public ArrayList m_arrParaFmt = new ArrayList();
        public ClassListLevel[] m_listLevels = new ClassListLevel[9];
        public Boolean m_bExistListLevels = false;

        public static Boolean m_bLog = false;

        public Boolean m_bAppIsWps = false;

        // 支持从一个打开的文档中直接copy多个样式
        //（listlevel前导编号、章节FONT和段落样式、正文包括不带/带前导符或前导编号)
        // 
        private Hashtable m_hashHeadingFont = new Hashtable();
        private Hashtable m_hashHeadingParaFormat = new Hashtable();

        private ArrayList m_srcArrListLevels = new ArrayList();

        public String m_cfgTempFileLoc = "";

        public ClassOfficeCommon m_commTools = null;// new ClassOfficeCommon();

        public readonly String m_stryp = "vubinomp";
        public readonly String m_stryv = "!%@^#&$*";
        // public readonly String m_licDat = "cil.dat";
        
        // public Hashtable m_hashRegPerm = null;

        //public Boolean m_bLicIllegal = false;

        //public Boolean m_bTryExpired = false;

        //        public Hashtable m_hashUniformStyleHistoryStyleDocs = new Hashtable();
//        public ArrayList m_arrUniformStyleHistoryStyleDocs = new ArrayList();
        public readonly uint m_nMaxUniformStyleHistoryStyleDocs = 10;

        public Boolean m_bDontShowPane = false;

        // public TableAdapterManager 
        //public TableAdapterManager m_tblAdapterMgr = new TableAdapterManager();
        //public localdbDataSet m_localDb = new localdbDataSet();
        private Boolean m_bInitedDataBase = false;

        public Hashtable m_hashHeadingSnPreBuiltInScheme = new Hashtable();
        public Hashtable m_hashHeadingSnUserDefineScheme = new Hashtable();
        private Boolean m_bLoadedHeadingSn = false;


        public Hashtable m_hashHeadingStylePreBuiltInScheme = new Hashtable();
        public Hashtable m_hashHeadingStyleUserDefineScheme = new Hashtable();
        private Boolean m_bLoadedHeadingStyle = false;

        private Hashtable m_hashWordFontNum = new Hashtable();
        private Hashtable m_hashPicFileType = new Hashtable();

        private Hashtable m_hashDocFileType = new Hashtable();

        private TreeNodeCollection m_trvShareLibNodes = null; // share lib,wenku

        public ArrayList m_arrWordFontSize = new ArrayList();
        public Hashtable m_hashFontSize2Name = new Hashtable(); // 字体尺寸到字号
        public Hashtable m_hashFontSizeName2Size = new Hashtable(); // 字体字号到尺寸


        public readonly String[] m_arrStrAlignStyle = { "左对齐", "居中", "右对齐", "两端对齐", "分散对齐" };

        public readonly String[] m_arrParaLineSpaceRule = { "单倍行距", "1.5 倍行距", "2 倍行距", "最小值", "固定值", "多倍行距" };

        public readonly String[] m_arrSpaceUnit = { "字符", "磅", "厘米", "毫米", "英寸", "行" };

        public readonly String[] m_arrFirstIndent = { "首行缩进", "悬挂缩进" };

        public readonly String[] m_arrTiZhuPos = { "上", "下" };

        public readonly String[] m_arrTiZhuAlign = { "中", "左", "右" };

        public readonly String[] m_arrHdSnStyles = {"（无）","1,2,3,…","I,II,III,…","i,ii,iii,…","A,B,C,…","a,b,c,…",
                                           "一,二,三（简）…","壹,贰,叁 …","甲,乙,丙 …","子,丑,寅 …",
                                           "1st,2nd,3rd …","One,Two,Three …","First,Second,Third …",
                                           "01,02,03,…","①,②…,⑳,21,22…"};

        public readonly String[] m_arrPageNumStyles = {"1，2，3，…","- 1 -，- 2 -，- 3 -，…","全角 …","a，b，c，…",
                                              "A，B，C，…","i，ii，iii，…","I，II，III，…","一，二，三（简） …",
                                              "壹，贰，叁 …","甲，乙，丙 …","子，丑，寅 …"};

        public readonly String[] m_arrPageNumSplittors = { "- （连字符）", ". （句点）", ": （冒号）", "—（长划线）", "–（短划线）" };

        public Hashtable m_hashIndex2ListStyle = new Hashtable();
        public Hashtable m_hashListStyle2Index = new Hashtable();


        public docPubDataLevelMgr m_docPubMgr = new docPubDataLevelMgr();
        public Boolean m_bDocPubSchemeNamesLoaded = false;
        public Hashtable m_hshDocPubNodeSn = new Hashtable();

        private int m_nAppVersion = 0;

        public int AppVersion
        {
            get 
            {
                return m_nAppVersion;
            }
        }


        /*
        public String m_cfgDocRepositoryUrl = ""; // 文库地址
        public String m_cfgCerServerUrl = "";     // Lic server
        public String m_cfgAutoUpdateSvrUrl = ""; // 自动更新地址

        public Boolean m_bVerIndividual = true; // false -- enterprise;
        public Boolean m_bWithDocRepository = false;

        public Boolean  m_bEnable = false;
        public DateTime m_dtExpireDate = DateTime.Now;
        */

        //public ClassUIController m_uiCtrler = new ClassUIController();

        public classMultiEditionCenter m_edtCenter = null;


        // 2017-02-20 deprecated
        public enum ShareLibFolderPermission
        {
            fpNewFolder = 1,
            fpUpload = 2,
            fpEdit = 3,
            fpDelete = 4,
            fpCopy = 5,
            fpCut = 6,
            fpDownload = 7,
            fpRefShare = 8,
            fpRefPrivateLib = 9,
            fpShare = 10,
            fpCancelShare = 11,
            fpSend2PrivateLib = 12,
            fpOpenOrCloseShare = 13,
            fpMoveUp = 14,
            fpMoveDown = 15,

        /*
            VALUE	TEXT
            1	新建文件夹
            2	上传
            3	编辑
            4	删除
            5	复制
            6	剪切
            7	下载
            8	引用共享
            9	引用个人库
            10	共享
            11	取消共享
            12	发送至个人库
            13	开通或关闭共享
            14	上移
            15	下移
        */
            
        }


        public enum DocRepositoryFolderPermission
        {
            fpFullControl = 0,
            fpVisible = 1,
            fpCreateFile = 2,
            fpCreateFolder = 3,
            fpPreviewFile = 4,
            fpPrintFile = 5, 
            fpDownloadFile = 6,
            fpUpdateFile = 7,
            fpRemoveFile = 8,
            fpShare = 9

            /*
            1	可见
            2	创建文件
            3	创建文件夹
            4	预览文件
            5	打印文件
            6	下载文件
            7	修改文件
            8	删除文件
            9	共享
            99	完全控制
            */
        }


        public enum DocRepositoryFilePermission
        {
            fpFullControl = 0,
            fpVisible = 1,
            
            fpPreview = 4,
            fpPrint = 5,
            fpDownload = 6,
            fpUpdate = 7,
            fpRemove = 8

            /*
            1	可见
            4	预览
            5	打印
            6	下载
            7	修改
            8	删除
            99	完全控制
            */
        }


        // public ShareLibFolderPermission m_

        public void loadAllHeadingSnSchemesNH()
        {
            if (m_bLoadedHeadingSn)
            {
                return;
            }

            String strSchemeName = "";
            int nRet = 0;


            m_hashHeadingSnPreBuiltInScheme.Clear();
            m_hashHeadingSnUserDefineScheme.Clear();


            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblListLevelSchemes> lstItems = null;

            try
            {
                IQuery hsSchemes = session.CreateQuery("from tblListLevelSchemes");
                lstItems = hsSchemes.List<tblListLevelSchemes>();
                
                foreach(tblListLevelSchemes schemeItem in lstItems)
                {
                    strSchemeName = schemeItem.schemeName;

                    if (strSchemeName.Equals(""))
                    {
                        System.Console.WriteLine("ListLevels name EMPTY, scheme id={0}", schemeItem.ID);
                        continue;
                    }

                    ClassListLevel[] listLevelScheme = new ClassListLevel[9];

                    for (int i = 0; i < 9; i++)
                    {
                        listLevelScheme[i] = new ClassListLevel();
                        //listLevelScheme[i].Font = new ClassFont();
                        listLevelScheme[i].Font.Name = "";
                    }

                    nRet = reloadHeadingSnSchemeNH(strSchemeName, ref listLevelScheme);

                    if (nRet < 0)
                    {
                        continue;
                    }

                    if (schemeItem.isPreBuiltIn)
                    {
                        m_hashHeadingSnPreBuiltInScheme[strSchemeName] = listLevelScheme;
                    }
                    else
                    {
                        m_hashHeadingSnUserDefineScheme[strSchemeName] = listLevelScheme;
                    }
                }

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            m_bLoadedHeadingSn = true;

            return;
        }

        /*
        public void loadAllHeadingSnSchemes_v1()
        {

            if (m_bLoadedHeadingSn)
            {
                return;
            }

            String strSchemeName = "";
            int nRet = 0;


            m_hashHeadingSnPreBuiltInScheme.Clear();
            m_hashHeadingSnUserDefineScheme.Clear();

            try
            {
	            foreach (localdbDataSet.tblListLevelSchemesRow dtRow in m_localDb.tblListLevelSchemes.Rows)
	            {
	                strSchemeName = dtRow.schemeName.Trim();
	
	                if (strSchemeName.Equals(""))
	                {
	                    System.Console.WriteLine("ListLevels name EMPTY, scheme id={0}", dtRow.ID);
	                    continue;
	                }
	
	                ClassListLevel[] listLevelScheme = new ClassListLevel[9];
	
	                for (int i = 0; i < 9; i++)
	                {
	                    listLevelScheme[i] = new ClassListLevel();
	                    //listLevelScheme[i].Font = new ClassFont();
	                    listLevelScheme[i].Font.Name = "";
	                }
	
	                nRet = reloadHeadingSnScheme_v1(strSchemeName, ref listLevelScheme);
	
	                if (nRet < 0)
	                {
	                    continue;
	                }
	
	                if (dtRow.isPreBuiltIn)
	                {
	                    m_hashHeadingSnPreBuiltInScheme[strSchemeName] = listLevelScheme;
	                }
	                else
	                {
	                    m_hashHeadingSnUserDefineScheme[strSchemeName] = listLevelScheme;
	                }
	
	            } // foreach
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
                
            }

            m_bLoadedHeadingSn = true;

            return;
        }
        */

        public void loadAllHeadingStyleSchemesNH()
        {
            if (m_bLoadedHeadingStyle)
            {
                return;
            }

            String strSchemeName = "";
            int nRet = 0;


            m_hashHeadingStylePreBuiltInScheme.Clear();
            m_hashHeadingStyleUserDefineScheme.Clear();

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblHeadingStyleScheme> lstItems = null;

            try
            {
                IQuery hsSchemes = session.CreateQuery("from tblHeadingStyleScheme");
                lstItems = hsSchemes.List<tblHeadingStyleScheme>();

                foreach (tblHeadingStyleScheme schemeItem in lstItems)
                {
                    strSchemeName = schemeItem.schemeName;

                    if (strSchemeName.Equals(""))
                    {
                        System.Console.WriteLine("ListLevels name EMPTY, scheme id={0}", schemeItem.ID);
                        continue;
                    }

                    ClassHeadingStyle[] hsScheme = new ClassHeadingStyle[10];

                    for (int i = 0; i < 10; i++)
                    {
                        hsScheme[i] = new ClassHeadingStyle();
                    }

                    nRet = reloadHeadingStyleSchemeNH(strSchemeName, ref hsScheme);

                    if (nRet < 0)
                    {
                        continue;
                    }

                    if (schemeItem.bPreBuiltIn)
                    {
                        m_hashHeadingStylePreBuiltInScheme[strSchemeName] = hsScheme;
                    }
                    else
                    {
                        m_hashHeadingStyleUserDefineScheme[strSchemeName] = hsScheme;
                    }

                } // foreach
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            m_bLoadedHeadingStyle = true;

            return;
        }

        /*
        public void loadAllHeadingStyleSchemes_v1()
        {
            if (m_bLoadedHeadingStyle)
            {
                return;
            }

            String strSchemeName = "";
            int nRet = 0;


            m_hashHeadingStylePreBuiltInScheme.Clear();
            m_hashHeadingStyleUserDefineScheme.Clear();

            try
            {
                foreach (localdbDataSet.tblHeadingStyleSchemeRow dtRow in m_localDb.tblHeadingStyleScheme.Rows)
                {
                    strSchemeName = dtRow.schemeName.Trim();

                    if (strSchemeName.Equals(""))
                    {
                        System.Console.WriteLine("HeadingStyle name EMPTY, scheme id={0}", dtRow.ID);
                        continue;
                    }

                    ClassHeadingStyle[] hsScheme = new ClassHeadingStyle[10];

                    for (int i = 0; i < 10; i++)
                    {
                        hsScheme[i] = new ClassHeadingStyle();
                    }

                    nRet = reloadHeadingStyleScheme_v1(strSchemeName, ref hsScheme);

                    if (nRet < 0)
                    {
                        continue;
                    }

                    if (dtRow.bPreBuiltIn)
                    {
                        m_hashHeadingStylePreBuiltInScheme[strSchemeName] = hsScheme;
                    }
                    else
                    {
                        m_hashHeadingStyleUserDefineScheme[strSchemeName] = hsScheme;
                    }

                } // foreach
            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }

            m_bLoadedHeadingStyle = true;

            return;
        }
        */

        public void saveAllHeadingStyleSchemesNH()
        {
            int nSchemeCnt = 0;

            ISession session = dbNHmgr.getSession();
            IList<tblListLevel> lstItems = null;

            foreach (DictionaryEntry entry in m_hashHeadingStyleUserDefineScheme)
            {
                String strName = (String)entry.Key;
                ClassHeadingStyle[] headingStyles = (ClassHeadingStyle[])entry.Value;

                try
                {
                    IQuery hsSchemes = session.CreateQuery("from tblHeadingStyleScheme where schemeName =:sName").SetString("sName", strName);
                    lstItems = hsSchemes.List<tblListLevel>();

                    nSchemeCnt = lstItems.Count;

                    if (nSchemeCnt == 0)
                    {
                        addHeadingStyleSchemeNH(strName, headingStyles);
                    }
                    else if (nSchemeCnt == 1)
                    {
                        updateHeadingStyleSchemeNH(strName, headingStyles);
                    }
                    else
                    {

                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    
                }

            }// foreach

            session.Close();

            return;
        }

        /*
        public void saveAllHeadingStyleSchemes_v1()
        {
            int nSchemeCnt = 0;

            foreach (DictionaryEntry entry in m_hashHeadingStyleUserDefineScheme)
            {
                String strName = (String)entry.Key;
                ClassHeadingStyle[] headingStyles = (ClassHeadingStyle[])entry.Value;

                // search db
                // if exist then update
                // else insert
                localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select("schemeName='" + strName + "'");

                nSchemeCnt = schemeRows.GetLength(0);

                if (nSchemeCnt == 0)
                {
                    addHeadingStyleScheme_v1(strName, headingStyles);
                }
                else if (nSchemeCnt == 1)
                {
                    updateHeadingStyleScheme_v1(strName, headingStyles);
                }
                else
                {

                }

            }// foreach

            return;
        }
        */

        public void saveAllHeadingSnSchemesNH()
        {
            int nSchemeCnt = 0;

            ISession session = dbNHmgr.getSession();
            IList<tblListLevel> lstItems = null;

            foreach (DictionaryEntry entry in m_hashHeadingSnUserDefineScheme)
            {
                String strName = (String)entry.Key;
                ClassListLevel[] listLevels = (ClassListLevel[])entry.Value;

                try
                {
                    IQuery hsSchemes = session.CreateQuery("from tblListLevelSchemes where schemeName =:sName").SetString("sName", strName);
                    lstItems = hsSchemes.List<tblListLevel>();

                    nSchemeCnt = lstItems.Count;

                    if (nSchemeCnt == 0)
                    {
                        addHeadingSnSchemeNH(strName, listLevels);
                    }
                    else if (nSchemeCnt == 1)
                    {
                        updateHeadingSnSchemeNH(strName, listLevels);
                    }
                    else
                    {

                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    
                }

            }// foreach

            session.Close();

            return;
        }

        /*
        public void saveAllHeadingSnSchemes_v1()
        {
            int nSchemeCnt = 0;

            foreach (DictionaryEntry entry in m_hashHeadingSnUserDefineScheme)
            {
                String strName = (String)entry.Key;
                ClassListLevel[] listLevels = (ClassListLevel[])entry.Value;

                // search db
                // if exist then update
                // else insert
                localdbDataSet.tblListLevelSchemesRow[] schemeRows = (localdbDataSet.tblListLevelSchemesRow[])m_localDb.tblListLevelSchemes.Select("schemeName='" + strName + "'");

                nSchemeCnt = schemeRows.GetLength(0);

                if (nSchemeCnt == 0)
                {
                    addHeadingSnScheme_v1(strName, listLevels);
                }
                else if (nSchemeCnt == 1)
                {
                    updateHeadingSnScheme_v1(strName, listLevels);
                }
                else
                {
                    
                }
                
            }// foreach

            return;
        }
        */


        public int removeHeadingStyleSchemeNH(Boolean bPreBuiltIn = true)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;

            IList<tblHeadingStyleScheme> htItems = null;

            IList<tblHeadingStyleFont> lstFntItems = null;
            IList<tblHeadingStyleParagraphFormat> lstParaFmt = null;

            try
            {
                //tx = session.BeginTransaction();

                IQuery qSchemes = session.CreateQuery("from tblHeadingStyleScheme where bPreBuiltIn =:sValue").SetBoolean("sValue", bPreBuiltIn);
                htItems = qSchemes.List<tblHeadingStyleScheme>();

                foreach (tblHeadingStyleScheme schemeItem in htItems)
                {
                    tx = session.BeginTransaction();

                    IQuery qFnt = session.CreateQuery("from tblHeadingStyleFont where schemeName =:sName").SetString("sName", schemeItem.schemeName);
                    lstFntItems = qFnt.List<tblHeadingStyleFont>();

                    foreach (tblHeadingStyleFont fntItem in lstFntItems) // remove listlevels
                    {
                        session.Delete(fntItem);
                    }

                    IQuery qParaFmt = session.CreateQuery("from tblHeadingStyleParagraphFormat where schemeName =:sName").SetString("sName", schemeItem.schemeName);
                    lstParaFmt = qParaFmt.List<tblHeadingStyleParagraphFormat>();

                    foreach (tblHeadingStyleParagraphFormat paraFmtItem in lstParaFmt) // remove listlevels
                    {
                        session.Delete(paraFmtItem);
                    }

                    // remove scheme
                    if (bPreBuiltIn)
                    {
                        m_hashHeadingStylePreBuiltInScheme.Remove(schemeItem.schemeName);
                    }
                    else
                    {
                        m_hashHeadingStyleUserDefineScheme.Remove(schemeItem.schemeName);
                    }

                    session.Delete(schemeItem);

                    tx.Commit();
                }
            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            return nRet;
        }

        /*
        public int removeHeadingStyleScheme_v1(Boolean bPreBuiltIn = true)
        {
            int nRet = 0, nTmp = 0;

            String strName = "";
            Boolean bRowPreBuiltIn = false;

            try
            {
                localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select("bPreBuiltIn=" + bPreBuiltIn);

                if (schemeRows.GetLength(0) == 0)
                {
                    return 1;
                }

                foreach (localdbDataSet.tblHeadingStyleSchemeRow row in schemeRows)
                {
                    localdbDataSet.tblHeadingStyleFontRow[] hsFntRows = (localdbDataSet.tblHeadingStyleFontRow[])m_localDb.tblHeadingStyleFont.Select("schemeName='" + row.schemeName + "'");

                    foreach (localdbDataSet.tblHeadingStyleFontRow fntRow in hsFntRows)
                    {
                        fntRow.BeginEdit();
                        fntRow.Delete();
                        fntRow.EndEdit();

                        // m_localDb.tblHeadingStyleFont.RemovetblHeadingStyleFontRow(fntRow);
                        // nTmp = m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Update(fntRow);
                    }

                    nTmp = m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Update(m_localDb.tblHeadingStyleFont);

                    localdbDataSet.tblHeadingStyleParagraphFormatRow[] hsParaFmtRows = (localdbDataSet.tblHeadingStyleParagraphFormatRow[])m_localDb.tblHeadingStyleParagraphFormat.Select("schemeName='" + row.schemeName + "'");

                    foreach (localdbDataSet.tblHeadingStyleParagraphFormatRow paraFmtRow in hsParaFmtRows)
                    {
                        paraFmtRow.BeginEdit();
                        paraFmtRow.Delete();
                        paraFmtRow.EndEdit();

                        // m_localDb.tblHeadingStyleParagraphFormat.RemovetblHeadingStyleParagraphFormatRow(paraFmtRow);
                        // nTmp = m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Update(paraFmtRow);
                    }

                    nTmp = m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Update(m_localDb.tblHeadingStyleParagraphFormat);


                    bRowPreBuiltIn = row.bPreBuiltIn;
                    strName = row.schemeName;

                    row.BeginEdit();
                    row.Delete();
                    row.EndEdit();

                    // m_localDb.tblHeadingStyleScheme.RemovetblHeadingStyleSchemeRow(row);
                    // nTmp = m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Update(row);

                    if (bRowPreBuiltIn)
                    {
                        m_hashHeadingStylePreBuiltInScheme.Remove(strName);
                    }
                    else
                    {
                        m_hashHeadingStyleUserDefineScheme.Remove(strName);
                    }
                }

                nTmp = m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Update(m_localDb.tblHeadingStyleScheme);


                m_localDb.tblHeadingStyleFont.AcceptChanges();
                m_localDb.tblHeadingStyleParagraphFormat.AcceptChanges();
                m_localDb.tblHeadingStyleScheme.AcceptChanges();
            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }

            return nRet;
        }
        */

        public int removeHeadingSnSchemeNH(Boolean bPreBuiltIn = true)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblListLevelSchemes> lstItems = null;
            IList<tblListLevel> lstLvlItems = null;

            try
            {
                //tx = session.BeginTransaction();

                IQuery hsSchemes = session.CreateQuery("from tblListLevelSchemes where isPreBuiltIn =:sValue").SetBoolean("sValue",bPreBuiltIn);
                lstItems = hsSchemes.List<tblListLevelSchemes>();

                if (lstItems.Count == 0) // 
                {
                    //session.Close();

                    return -1;
                }

                tx = session.BeginTransaction();

                // search by schemeName in listlevels
                foreach(tblListLevelSchemes schemeItem in lstItems)
                {
                    IQuery hsLvlItems = session.CreateQuery("from tblListLevel where schemeName =:sName").SetString("sName", schemeItem.schemeName);
                    lstLvlItems = hsLvlItems.List<tblListLevel>();

                    foreach (tblListLevel lvlItem in lstLvlItems) // remove listlevels
                    {
                        session.Delete(lvlItem);
                    }

                    if (bPreBuiltIn)
                    {
                        m_hashHeadingSnPreBuiltInScheme.Remove(schemeItem.schemeName);
                    }
                    else
                    {
                        m_hashHeadingSnUserDefineScheme.Remove(schemeItem.schemeName);
                    }

                    // remove scheme
                    session.Delete(schemeItem);
                }

                tx.Commit();

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            return nRet;
        }

        /*
        public int removeHeadingSnScheme_v1(Boolean bPreBuiltIn = true)
        {
            int nRet = 0, nTmp = 0;

            String strName = "";
            Boolean bRowPreBuiltIn = false;

            try
            {
                localdbDataSet.tblListLevelSchemesRow[] schemeRows = (localdbDataSet.tblListLevelSchemesRow[])m_localDb.tblListLevelSchemes.Select("isPreBuiltIn=" + bPreBuiltIn);
	
	            if (schemeRows.GetLength(0) == 0)
	            {
	                return 1;
	            }
	
	            foreach (localdbDataSet.tblListLevelSchemesRow row in schemeRows)
	            {
	                localdbDataSet.tblListLevelRow[] listlevelRows = (localdbDataSet.tblListLevelRow[])m_localDb.tblListLevel.Select("schemeName='" + row.schemeName + "'");
	
	                foreach (localdbDataSet.tblListLevelRow listLevelRow in listlevelRows)
	                {
	                    // remove listlevel
	                    listLevelRow.BeginEdit();
	                    listLevelRow.Delete();
	                    listLevelRow.EndEdit();
	
	                    // m_localDb.tblListLevel.RemovetblListLevelRow(listLevelRow);
	                    nTmp = m_tblAdapterMgr.tblListLevelTableAdapter.Update(listLevelRow);
	                }
	
	                bRowPreBuiltIn = row.isPreBuiltIn;
	                strName = row.schemeName;
	
	                row.BeginEdit();
	                row.Delete();
	                row.EndEdit();
	
	                //m_localDb.tblListLevelSchemes.RemovetblListLevelSchemesRow(row);
	                nTmp = m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Update(row);
	
	                if (bRowPreBuiltIn)
	                {
	                    m_hashHeadingSnPreBuiltInScheme.Remove(strName);
	                }
	                else
	                {
	                    m_hashHeadingSnUserDefineScheme.Remove(strName);
	                }
	
	            }
	
	            // nTmp = m_tblAdapterMgr.tblListLevelTableAdapter.Update(m_localDb.tblListLevel);
	            // nTmp = m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Update(m_localDb.tblListLevelSchemes);
	
	            m_localDb.tblListLevel.AcceptChanges();
	            m_localDb.tblListLevelSchemes.AcceptChanges();
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
                
            }

            return nRet;
        }
        */

        public int removeHeadingStyleSchemeNH(String strName)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;

            IList<tblHeadingStyleScheme> htItems = null;

            IList<tblHeadingStyleFont> lstFntItems = null;
            IList<tblHeadingStyleParagraphFormat> lstParaFmt = null;

            try
            {
                //tx = session.BeginTransaction();

                IQuery qSchemes = session.CreateQuery("from tblHeadingStyleScheme where schemeName =:sName").SetString("sName", strName);
                htItems = qSchemes.List<tblHeadingStyleScheme>();

                foreach (tblHeadingStyleScheme schemeItem in htItems)
                {
                    tx = session.BeginTransaction();

                    IQuery qFnt = session.CreateQuery("from tblHeadingStyleFont where schemeName =:sName").SetString("sName", schemeItem.schemeName);
                    lstFntItems = qFnt.List<tblHeadingStyleFont>();

                    foreach (tblHeadingStyleFont fntItem in lstFntItems) // remove listlevels
                    {
                        session.Delete(fntItem);
                    }

                    IQuery qParaFmt = session.CreateQuery("from tblHeadingStyleParagraphFormat where schemeName =:sName").SetString("sName", schemeItem.schemeName);
                    lstParaFmt = qParaFmt.List<tblHeadingStyleParagraphFormat>();

                    foreach (tblHeadingStyleParagraphFormat paraFmtItem in lstParaFmt) // remove listlevels
                    {
                        session.Delete(paraFmtItem);
                    }
                    
                    // remove scheme
                    if (schemeItem.bPreBuiltIn)
                    {
                        m_hashHeadingStylePreBuiltInScheme.Remove(schemeItem.schemeName);
                    }
                    else
                    {
                        m_hashHeadingStyleUserDefineScheme.Remove(schemeItem.schemeName);
                    }

                    session.Delete(schemeItem);

                    tx.Commit();
                }
            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            return nRet;

        }

        /*
        public int removeHeadingStyleScheme_v1(String strName)
        {
            int nRet = 0,nTmp = 0;
            Boolean bRowPreBuiltIn = false;

            try
            {
                localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select("schemeName='" + strName + "'");

                if (schemeRows.GetLength(0) == 0)
                {
                    return 1;
                }

                foreach (localdbDataSet.tblHeadingStyleSchemeRow row in schemeRows)
                {
                    localdbDataSet.tblHeadingStyleFontRow[] hsFntRows = (localdbDataSet.tblHeadingStyleFontRow[])m_localDb.tblHeadingStyleFont.Select("schemeName='" + row.schemeName + "'");

                    foreach (localdbDataSet.tblHeadingStyleFontRow fntRow in hsFntRows)
                    {
                        fntRow.BeginEdit();
                        fntRow.Delete();
                        fntRow.EndEdit();

                        nTmp = m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Update(fntRow);
                    }


                    localdbDataSet.tblHeadingStyleParagraphFormatRow[] hsParaFmtRows = (localdbDataSet.tblHeadingStyleParagraphFormatRow[])m_localDb.tblHeadingStyleParagraphFormat.Select("schemeName='" + row.schemeName + "'");

                    foreach (localdbDataSet.tblHeadingStyleParagraphFormatRow paraFmtRow in hsParaFmtRows)
                    {
                        paraFmtRow.BeginEdit();
                        paraFmtRow.Delete();
                        paraFmtRow.EndEdit();

                        nTmp = m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Update(paraFmtRow);
                    }


                    bRowPreBuiltIn = row.bPreBuiltIn;
                    strName = row.schemeName;

                    row.BeginEdit();
                    row.Delete();
                    row.EndEdit();

                    // m_localDb.tblHeadingStyleScheme.RemovetblHeadingStyleSchemeRow(row);
                    nTmp = m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Update(row);

                    if (bRowPreBuiltIn)
                    {
                        m_hashHeadingStylePreBuiltInScheme.Remove(strName);
                    }
                    else
                    {
                        m_hashHeadingStyleUserDefineScheme.Remove(strName);
                    }

                }

                m_localDb.tblHeadingStyleFont.AcceptChanges();
                m_localDb.tblHeadingStyleParagraphFormat.AcceptChanges();
                m_localDb.tblHeadingStyleScheme.AcceptChanges();
            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }

            return nRet;
        }
        */

        public int removeHeadingSnSchemeNH(String strName)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblListLevelSchemes> lstItems = null;
            IList<tblListLevel> lstLvlItems = null;

            try
            {
                //tx = session.BeginTransaction();

                IQuery hsSchemes = session.CreateQuery("from tblListLevelSchemes where schemeName =:sName").SetString("sName", strName);
                lstItems = hsSchemes.List<tblListLevelSchemes>();

                if (lstItems.Count == 0) // 
                {
                    //session.Close();

                    return -1;
                }

                tx = session.BeginTransaction();

                // search by schemeName in listlevels
                foreach (tblListLevelSchemes schemeItem in lstItems)
                {
                    IQuery hsLvlItems = session.CreateQuery("from tblListLevel where schemeName =:sName").SetString("sName", schemeItem.schemeName);
                    lstLvlItems = hsLvlItems.List<tblListLevel>();

                    foreach (tblListLevel lvlItem in lstLvlItems) // remove listlevels
                    {
                        session.Delete(lvlItem);
                    }

                    if (schemeItem.isPreBuiltIn)
                    {
                        m_hashHeadingSnPreBuiltInScheme.Remove(schemeItem.schemeName);
                    }
                    else
                    {
                        m_hashHeadingSnUserDefineScheme.Remove(schemeItem.schemeName);
                    }
                    // remove scheme
                    session.Delete(schemeItem);
                }

                tx.Commit();

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            return nRet;
        }

        /*
        public int removeHeadingSnScheme_v1(String strName)
        {
            int nRet = 0, nUpdateCnt = 0;

            try
            {
	            localdbDataSet.tblListLevelSchemesRow[] schemeRows = (localdbDataSet.tblListLevelSchemesRow[])m_localDb.tblListLevelSchemes.Select("schemeName='" + strName + "'");
	
	            if (schemeRows.GetLength(0) == 0)
	            {
	                return 1;
	            }
	
	            foreach (localdbDataSet.tblListLevelSchemesRow row in schemeRows)
	            {
	                localdbDataSet.tblListLevelRow[] listlevelRows = (localdbDataSet.tblListLevelRow[])m_localDb.tblListLevel.Select("schemeName='" + strName + "'");
	
	                foreach (localdbDataSet.tblListLevelRow listLevelRow in listlevelRows)
	                {
	                    // remove listlevel
                        listLevelRow.BeginEdit();
                        listLevelRow.Delete();
                        listLevelRow.EndEdit();

	                    // m_localDb.tblListLevel.RemovetblListLevelRow(listLevelRow);

                        //m_tblAdapterMgr.tblListLevelTableAdapter.Update(listLevelRow);
	                }

                    row.BeginEdit();
                    row.Delete();
                    row.EndEdit();

	                // m_localDb.tblListLevelSchemes.RemovetblListLevelSchemesRow(row);

                    //m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Update(row);
	
	            }

                nUpdateCnt = m_tblAdapterMgr.tblListLevelTableAdapter.Update(m_localDb.tblListLevel);
                nUpdateCnt = m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Update(m_localDb.tblListLevelSchemes);


	            m_localDb.tblListLevel.AcceptChanges();
	            m_localDb.tblListLevelSchemes.AcceptChanges();
	
	            m_hashHeadingSnUserDefineScheme.Remove(strName);
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
                
            }

            return nRet;
        }
        */

        /*
        public int testDB(String strSearch)
        {

            String strSelect = strSearch;// "schemeName=\"" + strSearch + "\"";

            localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select(strSelect);

            if (schemeRows.GetLength(0) > 0)
            {
                return -1;
            }

            return 0;
        }
        */

        public int addHeadingStyleSchemeNH(String strName, ClassHeadingStyle[] headingStyles, Boolean bPreBuiltIn = false)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblHeadingStyleScheme> lstItems = null;
            IList<tblHeadingStyleFont> lstFntItems = null;
            IList<tblHeadingStyleParagraphFormat> lstParaFmtItems = null;

            try
            {
                //tx = session.BeginTransaction();

                IQuery hsSchemes = session.CreateQuery("from tblHeadingStyleScheme where schemeName =:sName").SetString("sName", strName);
                lstItems = hsSchemes.List<tblHeadingStyleScheme>();

                if (lstItems.Count > 0) // 
                {
                    //session.Close();

                    return -1;
                }


                tx = session.BeginTransaction();

                ClassHeadingStyle hItem = null;
                ClassFont fnt = null;
                ClassParagraphFormat paraFmt = null;

                tblHeadingStyleFont tFnt = null;
                tblHeadingStyleParagraphFormat tParaFmt = null;

                for (int i = 0; i < 10; i++)
                {
                    hItem = headingStyles[i];
                    fnt = hItem.m_fnt;
                    paraFmt = hItem.m_paraFmt;

                    tFnt = new tblHeadingStyleFont();
                    tParaFmt = new tblHeadingStyleParagraphFormat();


                    //FONT
                    tFnt.ID = Guid.NewGuid();

                    tFnt.AllCaps = fnt.AllCaps;
                    tFnt.Animation = (int)fnt.Animation;
                    tFnt.Bold = fnt.Bold;
                    tFnt.BoldBi = fnt.BoldBi;
                    tFnt.Color = (int)fnt.Color;
                    tFnt.ColorIndex = (int)fnt.ColorIndex;
                    tFnt.ColorIndexBi = (int)fnt.ColorIndexBi;
                    tFnt.DiacriticColor = (int)fnt.DiacriticColor;
                    tFnt.DisableCharacterSpaceGrid = fnt.DisableCharacterSpaceGrid;
                    tFnt.DoubleStrikeThrough = fnt.DoubleStrikeThrough;
                    tFnt.Emboss = fnt.Emboss;
                    tFnt.EmphasisMark = (int)fnt.EmphasisMark;
                    tFnt.Engrave = fnt.Engrave;
                    tFnt.Hidden = fnt.Hidden;
                    tFnt.Italic = fnt.Italic;
                    tFnt.ItalicBi = fnt.ItalicBi;
                    tFnt.Kerning = fnt.Kerning;
                    
                    tFnt.strName = fnt.Name;
                    
                    tFnt.NameAscii = fnt.NameAscii;
                    tFnt.NameBi = fnt.NameBi;
                    tFnt.NameFarEast = fnt.NameFarEast;
                    tFnt.NameOther = "";
                    tFnt.Outline = fnt.Outline;
                    tFnt.OutlineLevel = (i+1);
                    tFnt.Position = fnt.Position;
                    tFnt.Scaling = fnt.Scaling;
                    tFnt.schemeName = strName;
                    tFnt.Shadow = fnt.Shadow;
                    tFnt.Size = fnt.Size;
                    tFnt.SizeBi = fnt.SizeBi;
                    tFnt.SmallCaps = fnt.SmallCaps;
                    tFnt.Spacing = fnt.Spacing;
                    tFnt.StrikeThrough = fnt.StrikeThrough;
                    tFnt.Subscript = fnt.Subscript;
                    tFnt.Superscript = fnt.Superscript;
                    tFnt.Underline = (int)fnt.Underline;
                    tFnt.UnderlineColor = (int)fnt.UnderlineColor;

                    //PARAM FORMAT
                    tParaFmt.ID = Guid.NewGuid();

                    tParaFmt.AddSpaceBetweenFarEastAndAlpha = paraFmt.AddSpaceBetweenFarEastAndAlpha;
                    tParaFmt.AddSpaceBetweenFarEastAndDigit = paraFmt.AddSpaceBetweenFarEastAndDigit;
                    tParaFmt.Alignment = (int)paraFmt.Alignment;
                    tParaFmt.AutoAdjustRightIndent = paraFmt.AutoAdjustRightIndent;
                    tParaFmt.BaseLineAlignment = (int)paraFmt.BaseLineAlignment;
                    tParaFmt.CharacterUnitFirstLineIndent = paraFmt.CharacterUnitFirstLineIndent;
                    tParaFmt.CharacterUnitLeftIndent = paraFmt.CharacterUnitLeftIndent;
                    tParaFmt.CharacterUnitRightIndent = paraFmt.CharacterUnitRightIndent;
                    tParaFmt.DisableLineHeightGrid = paraFmt.DisableLineHeightGrid;
                    tParaFmt.FarEastLineBreakControl = paraFmt.FarEastLineBreakControl;
                    tParaFmt.FirstLineIndent = paraFmt.FirstLineIndent;
                    tParaFmt.HalfWidthPunctuationOnTopOfLine = paraFmt.HalfWidthPunctuationOnTopOfLine;
                    tParaFmt.HangingPunctuation = paraFmt.HangingPunctuation;
                    tParaFmt.Hyphenation = paraFmt.Hyphenation;
                    tParaFmt.KeepTogether = paraFmt.KeepTogether;
                    tParaFmt.KeepWithNext = paraFmt.KeepWithNext;
                    tParaFmt.LeftIndent = paraFmt.LeftIndent;
                    tParaFmt.LineSpacing = paraFmt.LineSpacing;
                    tParaFmt.LineSpacingRule = (int)paraFmt.LineSpacingRule;
                    tParaFmt.LineUnitAfter = paraFmt.LineUnitAfter;
                    tParaFmt.LineUnitBefore = paraFmt.LineUnitBefore;
                    tParaFmt.MirrorIndents = paraFmt.MirrorIndents;
                    tParaFmt.NoLineNumber = paraFmt.NoLineNumber;
                    tParaFmt.OutlineLevel = (int)paraFmt.OutlineLevel;
                    tParaFmt.PageBreakBefore = paraFmt.PageBreakBefore;
                    tParaFmt.ReadingOrder = (int)paraFmt.ReadingOrder;
                    tParaFmt.RightIndent = paraFmt.RightIndent;
                    tParaFmt.schemeName = strName;
                    tParaFmt.SpaceAfter = paraFmt.SpaceAfter;
                    tParaFmt.SpaceAfterAuto = paraFmt.SpaceAfterAuto;
                    tParaFmt.SpaceBefore = paraFmt.SpaceBefore;
                    tParaFmt.SpaceBeforeAuto = paraFmt.SpaceBeforeAuto;
                    tParaFmt.TextboxTightWrap = (int)paraFmt.TextboxTightWrap;
                    tParaFmt.WidowControl = paraFmt.WidowControl;
                    tParaFmt.WordWrap = paraFmt.WordWrap;

                    // Save
                    session.Save(tFnt);
                    session.Save(tParaFmt);
                }

                // add into scheme table
                tblHeadingStyleScheme schemeItem = new tblHeadingStyleScheme();

                schemeItem.ID = Guid.NewGuid();
                schemeItem.bPreBuiltIn = bPreBuiltIn;
                schemeItem.bVisible = true;
                schemeItem.schemeName = strName;

                session.Save(schemeItem);

                tx.Commit();

                if (bPreBuiltIn)
                {
                    m_hashHeadingStylePreBuiltInScheme[strName] = headingStyles;
                }
                else
                {
                    m_hashHeadingStyleUserDefineScheme[strName] = headingStyles;
                }

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }
            return nRet;
        }
        
        /*
        public int addHeadingStyleScheme_v1(String strName, ClassHeadingStyle[] headingStyles, Boolean bPreBuiltIn = false)
        {
            int nRet = 0;

            try
            {
                localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select("schemeName='" + strName + "'");

                if (schemeRows.GetLength(0) > 0)
                {
                    return -1;
                }

                localdbDataSet.tblHeadingStyleSchemeRow newSchemeRow = m_localDb.tblHeadingStyleScheme.AddtblHeadingStyleSchemeRow(strName, bPreBuiltIn,true,9999999);

                if (newSchemeRow == null)
                {
                    return -1;
                }

                newSchemeRow.EndEdit();

                nRet = m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Update(newSchemeRow);
                
                //nRet = m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Insert(strName, bPreBuiltIn, true, 9999999);

                m_localDb.tblHeadingStyleScheme.AcceptChanges();

                ClassFont fnt = null;
                ClassParagraphFormat paraFmt = null;
                ClassHeadingStyle headingStyleItem = null;

                for (int i = 0; i < 10; i++)
                {
                    headingStyleItem = headingStyles[i];
                    fnt = headingStyles[i].m_fnt;
                    paraFmt = headingStyles[i].m_paraFmt;

                    localdbDataSet.tblHeadingStyleFontRow newFntRow = m_localDb.tblHeadingStyleFont.AddtblHeadingStyleFontRow(
                        strName,
                        (i+1),
                        fnt.AllCaps, 
                        (int)fnt.Animation, 
                        fnt.Bold, 
                        fnt.BoldBi,
                        (int)fnt.Color,
                        (int)fnt.ColorIndex,
                        (int)fnt.ColorIndexBi,
                        (int)fnt.DiacriticColor, 
                        fnt.DisableCharacterSpaceGrid, 
                        fnt.DoubleStrikeThrough, 
                        fnt.Emboss,
                        (int)fnt.EmphasisMark, 
                        fnt.Engrave, 
                        fnt.Hidden, 
                        fnt.Italic, 
                        fnt.ItalicBi, 
                        fnt.Kerning, 
                        fnt.Name, 
                        fnt.NameAscii, 
                        fnt.NameBi, 
                        fnt.NameFarEast, 
                        "",//fnt.NameOther, 
                        fnt.Outline, 
                        fnt.Position, 
                        fnt.Scaling, 
                        fnt.Shadow, 
                        fnt.Size, 
                        fnt.SizeBi, 
                        fnt.SmallCaps, 
                        fnt.Spacing, 
                        fnt.StrikeThrough, 
                        fnt.Subscript, 
                        fnt.Superscript,
                        (int)fnt.Underline,
                        (int)fnt.UnderlineColor
                        );


                    localdbDataSet.tblHeadingStyleParagraphFormatRow newParaFmtRow = m_localDb.tblHeadingStyleParagraphFormat.AddtblHeadingStyleParagraphFormatRow(
                        strName, 
                        paraFmt.AddSpaceBetweenFarEastAndAlpha, 
                        paraFmt.AddSpaceBetweenFarEastAndDigit, 
                        (int)paraFmt.Alignment, 
                        paraFmt.AutoAdjustRightIndent,
                        (int)paraFmt.BaseLineAlignment, 
                        paraFmt.CharacterUnitFirstLineIndent, 
                        paraFmt.CharacterUnitLeftIndent, 
                        paraFmt.CharacterUnitRightIndent, 
                        paraFmt.DisableLineHeightGrid, 
                        paraFmt.FarEastLineBreakControl, 
                        paraFmt.FirstLineIndent, 
                        paraFmt.HalfWidthPunctuationOnTopOfLine, 
                        paraFmt.HangingPunctuation, 
                        paraFmt.Hyphenation, 
                        paraFmt.KeepTogether, 
                        paraFmt.KeepWithNext, 
                        paraFmt.LeftIndent, 
                        paraFmt.LineSpacing,
                        (int)paraFmt.LineSpacingRule, 
                        paraFmt.LineUnitAfter, 
                        paraFmt.LineUnitBefore, 
                        paraFmt.MirrorIndents, 
                        paraFmt.NoLineNumber,
                        (int)paraFmt.OutlineLevel, 
                        paraFmt.PageBreakBefore,
                        (int)paraFmt.ReadingOrder, 
                        paraFmt.RightIndent, 
                        paraFmt.SpaceAfter, 
                        paraFmt.SpaceAfterAuto, 
                        paraFmt.SpaceBefore, 
                        paraFmt.SpaceBeforeAuto,
                        (int)paraFmt.TextboxTightWrap, 
                        paraFmt.WidowControl,
                        paraFmt.WordWrap                      
                        );


                    if (newFntRow == null || newParaFmtRow == null)
                    {
                        removeHeadingStyleScheme_v1(strName);
                        return -1;
                    }


                    m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Update(newFntRow);
                    m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Update(newParaFmtRow);

                    m_localDb.tblHeadingStyleFont.AcceptChanges();
                    m_localDb.tblHeadingStyleParagraphFormat.AcceptChanges();

                }

                m_localDb.tblHeadingStyleFont.Clear();
                m_localDb.tblHeadingStyleParagraphFormat.Clear();
                m_localDb.tblHeadingStyleScheme.Clear();


                m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Fill(m_localDb.tblHeadingStyleFont);
                m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Fill(m_localDb.tblHeadingStyleParagraphFormat);
                m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Fill(m_localDb.tblHeadingStyleScheme);


                if (bPreBuiltIn)
                {
                    m_hashHeadingStylePreBuiltInScheme[strName] = headingStyles;
                }
                else
                {
                    m_hashHeadingStyleUserDefineScheme[strName] = headingStyles;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("DB Write Error:" + ex.Message);
            }
            finally
            {

            }

            return nRet;
        }
        */

        public int addHeadingSnSchemeNH(String strName, ClassListLevel[] listLevels, Boolean bPreBuiltIn = false)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblListLevelSchemes> lstItems = null;


            try
            {
                //tx = session.BeginTransaction();

                IQuery hsSchemes = session.CreateQuery("from tblListLevelSchemes where schemeName =:sName").SetString("sName", strName);
                lstItems = hsSchemes.List<tblListLevelSchemes>();

                if (lstItems.Count > 0) // 
                {
                    //session.Close();

                    return -1;
                }

                tx = session.BeginTransaction();

                // add into listlevel table
                tblListLevel snItem = null;
                ClassListLevel lvl = null;

                for (int i = 0; i < 9; i++)
                {
                    lvl = listLevels[i];

                    snItem = new tblListLevel();

                    snItem.ID = Guid.NewGuid();
                    snItem.Alignment = (int)lvl.Alignment;
                    snItem.AllCaps = lvl.Font.AllCaps;
                    snItem.Animation = (int)lvl.Font.Animation;
                    snItem.Bold = lvl.Font.Bold;
                    snItem.Color = (int)lvl.Font.Color;
                    snItem.DoubleStrikeThrough = lvl.Font.DoubleStrikeThrough;
                    snItem.Emboss = lvl.Font.Emboss;
                    snItem.Engrave = lvl.Font.Engrave;
                    snItem.fntName = lvl.Font.Name;
                    snItem.Hidden = lvl.Font.Hidden;
                    snItem.Italic = lvl.Font.Italic;
                    snItem.level = (i+1);
                    snItem.LinkedStyle = lvl.LinkedStyle;
                    snItem.NumberFormat = lvl.NumberFormat;
                    snItem.NumberPosition = lvl.NumberPosition;
                    snItem.NumberStyle = (int)lvl.NumberStyle;
                    snItem.Outline = lvl.Font.Outline;
                    snItem.ResetOnHigher = lvl.ResetOnHigher;
                    snItem.schemeId = -1;
                    snItem.schemeName = strName;
                    snItem.Shadow = lvl.Font.Shadow;
                    snItem.Size = lvl.Font.Size;
                    snItem.StartAt = lvl.StartAt;
                    snItem.StrikeThrough = lvl.Font.StrikeThrough;
                    snItem.Subscript = lvl.Font.Subscript;
                    snItem.Superscript = lvl.Font.Superscript;
                    snItem.TextPosition = lvl.TextPosition;
                    snItem.TrailingCharacter = (int)lvl.TrailingCharacter;
                    snItem.Underline = (int)lvl.Font.Underline;

                    session.Save(snItem);
                }

                // add into scheme table
                tblListLevelSchemes schemeItem = new tblListLevelSchemes();

                schemeItem.ID = Guid.NewGuid();
                schemeItem.isPreBuiltIn = bPreBuiltIn;
                schemeItem.bVisible = true;
                schemeItem.schemeName = strName;

                session.Save(schemeItem);

                tx.Commit();

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }


            return nRet;
        }

        /*
        public int addHeadingSnScheme_v1(String strName, ClassListLevel[] listLevels, Boolean bPreBuiltIn = false)
        {
            int nRet = 0, nUpdateCnt = 0;

            try
            {
	            localdbDataSet.tblListLevelSchemesRow[] schemeRows = (localdbDataSet.tblListLevelSchemesRow[])m_localDb.tblListLevelSchemes.Select("schemeName='" + strName + "'");
	
	            if (schemeRows.GetLength(0) > 0)
	            {
	                return -1;
	            }
	
	            localdbDataSet.tblListLevelSchemesRow newSchemeRow = m_localDb.tblListLevelSchemes.AddtblListLevelSchemesRow(strName,bPreBuiltIn,true,9999999);
	
	            if(newSchemeRow == null)
	            {
	                return -1;
	            }
	
	            newSchemeRow.EndEdit();

                nUpdateCnt = m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Update(newSchemeRow);
	
	            m_localDb.tblListLevelSchemes.AcceptChanges();
	
	            ClassFont fnt = null;
	            ClassListLevel lstLvl = null;
	            for(int i = 0; i < 9; i++)
	            {
	                lstLvl = listLevels[i];
	                fnt = listLevels[i].Font;
	
	                localdbDataSet.tblListLevelRow newListLevelRow = m_localDb.tblListLevel.AddtblListLevelRow(
	                    newSchemeRow, -1, (i + 1),lstLvl.NumberFormat, (int)lstLvl.TrailingCharacter, (int)lstLvl.NumberStyle, 
	                    lstLvl.NumberPosition,(int)lstLvl.Alignment, lstLvl.TextPosition, lstLvl.ResetOnHigher, lstLvl.StartAt,lstLvl.LinkedStyle,
	                    fnt.Bold,fnt.Italic,fnt.StrikeThrough,fnt.Subscript,fnt.Superscript,fnt.Shadow,fnt.Outline,fnt.Emboss,
	                    fnt.Engrave, fnt.AllCaps, fnt.Hidden, (int)fnt.Underline, (int)fnt.Color, fnt.Size, (int)fnt.Animation,
	                    fnt.DoubleStrikeThrough,fnt.Name);
	
	                if (newListLevelRow == null)
	                {
	                    removeHeadingSnScheme_v1(strName);
	                    return -1;
	                }

                    nUpdateCnt = m_tblAdapterMgr.tblListLevelTableAdapter.Update(newListLevelRow);
	                m_localDb.tblListLevel.AcceptChanges();
	
	            }
	
                m_localDb.tblListLevel.Clear();
                m_localDb.tblListLevelSchemes.Clear();


                m_tblAdapterMgr.tblListLevelTableAdapter.Fill(m_localDb.tblListLevel);
                m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Fill(m_localDb.tblListLevelSchemes);



	            if (bPreBuiltIn)
	            {
	                m_hashHeadingSnPreBuiltInScheme[strName] = listLevels;
	            }
	            else
	            {
	                m_hashHeadingSnUserDefineScheme[strName] = listLevels;
	            }
            }
            catch (System.Exception ex)
            {
	            
            }
            finally
            {
                
            }

            return nRet;
        }
        */

        public int updateHeadingStyleSchemeNH(String strName, ClassHeadingStyle[] newHeadingStyles)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblHeadingStyleScheme> lstItems = null;
            IList<tblHeadingStyleFont> lstFntItems = null;
            IList<tblHeadingStyleParagraphFormat> lstParaFmtItems = null;

            try
            {
                IQuery qFnt = session.CreateQuery("from tblHeadingStyleFont where schemeName =:sName order by OutlineLevel").SetString("sName", strName);
                lstFntItems = qFnt.List<tblHeadingStyleFont>();

                IQuery qParaFmt = session.CreateQuery("from tblHeadingStyleParagraphFormat where schemeName =:sName order by OutlineLevel").SetString("sName", strName);
                lstParaFmtItems = qParaFmt.List<tblHeadingStyleParagraphFormat>();

                if (lstFntItems.Count != 10 && lstParaFmtItems.Count != 10) // 
                {
                    //session.Close();
                    return -1;
                }


                tx = session.BeginTransaction();

                ClassHeadingStyle hItem = null;
                ClassFont fnt = null;
                ClassParagraphFormat paraFmt = null;

                tblHeadingStyleFont tFnt = null;
                tblHeadingStyleParagraphFormat tParaFmt = null;

                for (int i = 0; i < 10; i++)
                {
                    hItem = newHeadingStyles[i];
                    fnt = hItem.m_fnt;
                    paraFmt = hItem.m_paraFmt;

                    tFnt = lstFntItems[i];
                    tParaFmt = lstParaFmtItems[i];

                    //FONT
                    // tFnt.ID = Guid.NewGuid();

                    tFnt.AllCaps = fnt.AllCaps;
                    tFnt.Animation = (int)fnt.Animation;
                    tFnt.Bold = fnt.Bold;
                    tFnt.BoldBi = fnt.BoldBi;
                    tFnt.Color = (int)fnt.Color;
                    tFnt.ColorIndex = (int)fnt.ColorIndex;
                    tFnt.ColorIndexBi = (int)fnt.ColorIndexBi;
                    tFnt.DiacriticColor = (int)fnt.DiacriticColor;
                    tFnt.DisableCharacterSpaceGrid = fnt.DisableCharacterSpaceGrid;
                    tFnt.DoubleStrikeThrough = fnt.DoubleStrikeThrough;
                    tFnt.Emboss = fnt.Emboss;
                    tFnt.EmphasisMark = (int)fnt.EmphasisMark;
                    tFnt.Engrave = fnt.Engrave;
                    tFnt.Hidden = fnt.Hidden;
                    tFnt.Italic = fnt.Italic;
                    tFnt.ItalicBi = fnt.ItalicBi;
                    tFnt.Kerning = fnt.Kerning;

                    tFnt.strName = fnt.Name;

                    tFnt.NameAscii = fnt.NameAscii;
                    tFnt.NameBi = fnt.NameBi;
                    tFnt.NameFarEast = fnt.NameFarEast;
                    tFnt.NameOther = "";
                    tFnt.Outline = fnt.Outline;
                    tFnt.OutlineLevel = (i + 1);
                    tFnt.Position = fnt.Position;
                    tFnt.Scaling = fnt.Scaling;
                    tFnt.schemeName = strName;
                    tFnt.Shadow = fnt.Shadow;
                    tFnt.Size = fnt.Size;
                    tFnt.SizeBi = fnt.SizeBi;
                    tFnt.SmallCaps = fnt.SmallCaps;
                    tFnt.Spacing = fnt.Spacing;
                    tFnt.StrikeThrough = fnt.StrikeThrough;
                    tFnt.Subscript = fnt.Subscript;
                    tFnt.Superscript = fnt.Superscript;
                    tFnt.Underline = (int)fnt.Underline;
                    tFnt.UnderlineColor = (int)fnt.UnderlineColor;

                    //PARAM FORMAT
                    // tParaFmt.ID = Guid.NewGuid();

                    tParaFmt.AddSpaceBetweenFarEastAndAlpha = paraFmt.AddSpaceBetweenFarEastAndAlpha;
                    tParaFmt.AddSpaceBetweenFarEastAndDigit = paraFmt.AddSpaceBetweenFarEastAndDigit;
                    tParaFmt.Alignment = (int)paraFmt.Alignment;
                    tParaFmt.AutoAdjustRightIndent = paraFmt.AutoAdjustRightIndent;
                    tParaFmt.BaseLineAlignment = (int)paraFmt.BaseLineAlignment;
                    tParaFmt.CharacterUnitFirstLineIndent = paraFmt.CharacterUnitFirstLineIndent;
                    tParaFmt.CharacterUnitLeftIndent = paraFmt.CharacterUnitLeftIndent;
                    tParaFmt.CharacterUnitRightIndent = paraFmt.CharacterUnitRightIndent;
                    tParaFmt.DisableLineHeightGrid = paraFmt.DisableLineHeightGrid;
                    tParaFmt.FarEastLineBreakControl = paraFmt.FarEastLineBreakControl;
                    tParaFmt.FirstLineIndent = paraFmt.FirstLineIndent;
                    tParaFmt.HalfWidthPunctuationOnTopOfLine = paraFmt.HalfWidthPunctuationOnTopOfLine;
                    tParaFmt.HangingPunctuation = paraFmt.HangingPunctuation;
                    tParaFmt.Hyphenation = paraFmt.Hyphenation;
                    tParaFmt.KeepTogether = paraFmt.KeepTogether;
                    tParaFmt.KeepWithNext = paraFmt.KeepWithNext;
                    tParaFmt.LeftIndent = paraFmt.LeftIndent;
                    tParaFmt.LineSpacing = paraFmt.LineSpacing;
                    tParaFmt.LineSpacingRule = (int)paraFmt.LineSpacingRule;
                    tParaFmt.LineUnitAfter = paraFmt.LineUnitAfter;
                    tParaFmt.LineUnitBefore = paraFmt.LineUnitBefore;
                    tParaFmt.MirrorIndents = paraFmt.MirrorIndents;
                    tParaFmt.NoLineNumber = paraFmt.NoLineNumber;
                    tParaFmt.OutlineLevel = (int)paraFmt.OutlineLevel;
                    tParaFmt.PageBreakBefore = paraFmt.PageBreakBefore;
                    tParaFmt.ReadingOrder = (int)paraFmt.ReadingOrder;
                    tParaFmt.RightIndent = paraFmt.RightIndent;
                    tParaFmt.schemeName = strName;
                    tParaFmt.SpaceAfter = paraFmt.SpaceAfter;
                    tParaFmt.SpaceAfterAuto = paraFmt.SpaceAfterAuto;
                    tParaFmt.SpaceBefore = paraFmt.SpaceBefore;
                    tParaFmt.SpaceBeforeAuto = paraFmt.SpaceBeforeAuto;
                    tParaFmt.TextboxTightWrap = (int)paraFmt.TextboxTightWrap;
                    tParaFmt.WidowControl = paraFmt.WidowControl;
                    tParaFmt.WordWrap = paraFmt.WordWrap;

                    // Save
                    session.SaveOrUpdate(tFnt);
                    session.SaveOrUpdate(tParaFmt);
                }

                tx.Commit();

                m_hashHeadingStyleUserDefineScheme[strName] = newHeadingStyles;

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }
            return nRet;
        }

        /*
        public int updateHeadingStyleScheme_v1(String strName, ClassHeadingStyle[] newHeadingStyles)
        {
            int nRet = 0, nLevel = 0;

            try
            {
                localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select("schemeName='" + strName + "'");

                if (schemeRows.GetLength(0) != 1)
                {
                    return -1;
                }

                foreach (localdbDataSet.tblHeadingStyleSchemeRow row in schemeRows)
                {
                    localdbDataSet.tblHeadingStyleFontRow[] fntRows = (localdbDataSet.tblHeadingStyleFontRow[])m_localDb.tblHeadingStyleFont.Select("schemeName='" + strName + "'", "OutlineLevel");

                    if (fntRows.GetLength(0) != 10)
                    {
                        return -1;
                    }

                    for (int i = 0; i < 10; i++)
                    {
                        nLevel = fntRows[i].OutlineLevel;

                        if (nLevel >= 0 && nLevel < 10)
                        {
                            localdbDataSet.tblHeadingStyleFontRow Afnt = fntRows[i];
                            ClassFont Bfnt = newHeadingStyles[nLevel].m_fnt;

                            Afnt.AllCaps = Bfnt.AllCaps; 
                            Afnt.Animation = (int)Bfnt.Animation;
                            Afnt.Bold = Bfnt.Bold;
                            Afnt.BoldBi = Bfnt.BoldBi;
                            Afnt.Color = (int)Bfnt.Color;
                            Afnt.ColorIndex = (int)Bfnt.ColorIndex;
                            Afnt.ColorIndexBi = (int)Bfnt.ColorIndexBi;
                            Afnt.DiacriticColor = (int)Bfnt.DiacriticColor;
                            Afnt.DisableCharacterSpaceGrid = Bfnt.DisableCharacterSpaceGrid;
                            Afnt.DoubleStrikeThrough = Bfnt.DoubleStrikeThrough;
                            Afnt.Emboss = Bfnt.Emboss;
                            Afnt.EmphasisMark = (int)Bfnt.EmphasisMark; 
                            Afnt.Engrave = Bfnt.Engrave;
                            Afnt.Hidden = Bfnt.Hidden;
                            Afnt.Italic = Bfnt.Italic;
                            Afnt.ItalicBi = Bfnt.ItalicBi;
                            Afnt.Kerning = Bfnt.Kerning; 
                            Afnt.strName = Bfnt.Name;
                            Afnt.NameAscii = Bfnt.NameAscii;
                            Afnt.NameBi = Bfnt.NameBi;
                            Afnt.NameFarEast = Bfnt.NameFarEast;
                            //Afnt.NameOther = Bfnt.NameOther;
                            Afnt.Outline = Bfnt.Outline; 
                            Afnt.Position = Bfnt.Position;
                            Afnt.Scaling = Bfnt.Scaling;
                            Afnt.Shadow = Bfnt.Shadow;
                            Afnt.Size = Bfnt.Size;
                            Afnt.SizeBi = Bfnt.SizeBi;
                            Afnt.SmallCaps = Bfnt.SmallCaps;
                            Afnt.Spacing = Bfnt.Spacing;
                            Afnt.StrikeThrough = Bfnt.StrikeThrough;
                            Afnt.Subscript = Bfnt.Subscript;
                            Afnt.Superscript = Bfnt.Superscript;
                            Afnt.Underline = (int)Bfnt.Underline;
                            Afnt.UnderlineColor = (int)Bfnt.UnderlineColor;

                            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Update(fntRows[i]);
                            m_localDb.tblHeadingStyleFont.AcceptChanges();
                        }
                        else
                        {

                        }
                    }

                    // paragraph format
                    localdbDataSet.tblHeadingStyleParagraphFormatRow[] paraFmtRows = (localdbDataSet.tblHeadingStyleParagraphFormatRow[])m_localDb.tblHeadingStyleParagraphFormat.Select("schemeName='" + strName + "'", "OutlineLevel");

                    if (paraFmtRows.GetLength(0) != 10)
                    {
                        return -1;
                    }

                    for (int i = 0; i < 10; i++)
                    {
                        nLevel = fntRows[i].OutlineLevel;

                        if (nLevel >= 0 && nLevel < 10)
                        {
                            localdbDataSet.tblHeadingStyleParagraphFormatRow Afmt = paraFmtRows[i];
                            ClassParagraphFormat Bfmt = newHeadingStyles[nLevel].m_paraFmt;

                            Afmt.AddSpaceBetweenFarEastAndAlpha = Bfmt.AddSpaceBetweenFarEastAndAlpha; 
                            Afmt.AddSpaceBetweenFarEastAndDigit = Bfmt.AddSpaceBetweenFarEastAndDigit;
                            Afmt.Alignment = (int)Bfmt.Alignment;
                            Afmt.AutoAdjustRightIndent = Bfmt.AutoAdjustRightIndent;
                            Afmt.BaseLineAlignment = (int)Bfmt.BaseLineAlignment;
                            Afmt.CharacterUnitFirstLineIndent = Bfmt.CharacterUnitFirstLineIndent;
                            Afmt.CharacterUnitLeftIndent = Bfmt.CharacterUnitLeftIndent;
                            Afmt.CharacterUnitRightIndent = Bfmt.CharacterUnitRightIndent;
                            Afmt.DisableLineHeightGrid = Bfmt.DisableLineHeightGrid;
                            Afmt.FarEastLineBreakControl = Bfmt.FarEastLineBreakControl;
                            Afmt.FirstLineIndent = Bfmt.FirstLineIndent;
                            Afmt.HalfWidthPunctuationOnTopOfLine = Bfmt.HalfWidthPunctuationOnTopOfLine;
                            Afmt.HangingPunctuation = Bfmt.HangingPunctuation;
                            Afmt.Hyphenation = Bfmt.Hyphenation;
                            Afmt.KeepTogether = Bfmt.KeepTogether;
                            Afmt.KeepWithNext = Bfmt.KeepWithNext;
                            Afmt.LeftIndent = Bfmt.LeftIndent;
                            Afmt.LineSpacing = Bfmt.LineSpacing;
                            Afmt.LineSpacingRule = (int)Bfmt.LineSpacingRule;
                            Afmt.LineUnitAfter = Bfmt.LineUnitAfter;
                            Afmt.LineUnitBefore = Bfmt.LineUnitBefore;
                            Afmt.MirrorIndents = Bfmt.MirrorIndents;
                            Afmt.NoLineNumber = Bfmt.NoLineNumber;
                            Afmt.OutlineLevel = (int)Bfmt.OutlineLevel;
                            Afmt.PageBreakBefore = Bfmt.PageBreakBefore;
                            Afmt.ReadingOrder  = (int)Bfmt.ReadingOrder;
                            Afmt.RightIndent = Bfmt.RightIndent;
                            Afmt.SpaceAfter = Bfmt.SpaceAfter;
                            Afmt.SpaceAfterAuto = Bfmt.SpaceAfterAuto;
                            Afmt.SpaceBefore = Bfmt.SpaceBefore;
                            Afmt.SpaceBeforeAuto = Bfmt.SpaceBeforeAuto;
                            Afmt.TextboxTightWrap = (int)Bfmt.TextboxTightWrap;
                            Afmt.WidowControl = Bfmt.WidowControl;
                            Afmt.WordWrap = Bfmt.WordWrap;

                            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Update(paraFmtRows[i]);
                            m_localDb.tblHeadingStyleParagraphFormat.AcceptChanges();
                        }
                        else
                        {

                        }
                    }


                }

                m_hashHeadingStyleUserDefineScheme[strName] = newHeadingStyles;
            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }

            return nRet;
        }
        */

        public int updateHeadingSnSchemeNH(String strName, ClassListLevel[] newlistLevels)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblListLevel> lstItems = null;

            try
            {
                IQuery hsSchemes = session.CreateQuery("from tblListLevel where schemeName =:sName order by level").SetString("sName", strName);
                lstItems = hsSchemes.List<tblListLevel>();

                if (lstItems.Count != 9) // 
                {
                    //session.Close();

                    return -1;
                }

                tx = session.BeginTransaction();

                // add into listlevel table
                tblListLevel snItem = null;
                ClassListLevel lvl = null;

                for (int i = 0; i < 9; i++)
                {
                    lvl = newlistLevels[i];

                    snItem = lstItems[i];

                    // snItem.ID = Guid.NewGuid();

                    snItem.Alignment = (int)lvl.Alignment;
                    snItem.AllCaps = lvl.Font.AllCaps;
                    snItem.Animation = (int)lvl.Font.Animation;
                    snItem.Bold = lvl.Font.Bold;
                    snItem.Color = (int)lvl.Font.Color;
                    snItem.DoubleStrikeThrough = lvl.Font.DoubleStrikeThrough;
                    snItem.Emboss = lvl.Font.Emboss;
                    snItem.Engrave = lvl.Font.Engrave;
                    snItem.fntName = lvl.Font.Name;
                    snItem.Hidden = lvl.Font.Hidden;
                    snItem.Italic = lvl.Font.Italic;
                    snItem.level = lvl.Index;
                    snItem.LinkedStyle = lvl.LinkedStyle;
                    snItem.NumberFormat = lvl.NumberFormat;
                    snItem.NumberPosition = lvl.NumberPosition;
                    snItem.NumberStyle = (int)lvl.NumberStyle;
                    snItem.Outline = lvl.Font.Outline;
                    snItem.ResetOnHigher = lvl.ResetOnHigher;
                    snItem.schemeId = -1;
                    snItem.schemeName = strName;
                    snItem.Shadow = lvl.Font.Shadow;
                    snItem.Size = lvl.Font.Size;
                    snItem.StartAt = lvl.StartAt;
                    snItem.StrikeThrough = lvl.Font.StrikeThrough;
                    snItem.Subscript = lvl.Font.Subscript;
                    snItem.Superscript = lvl.Font.Superscript;
                    snItem.TextPosition = lvl.TextPosition;
                    snItem.TrailingCharacter = (int)lvl.TrailingCharacter;
                    snItem.Underline = (int)lvl.Font.Underline;

                    session.SaveOrUpdate(snItem);
                }

                tx.Commit();

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }


            return nRet;
        }

        /*
        public int updateHeadingSnScheme_v1(String strName, ClassListLevel[] newlistLevels)
        {
            int nRet = 0;

            try
            {
	            localdbDataSet.tblListLevelSchemesRow[] schemeRows = (localdbDataSet.tblListLevelSchemesRow[])m_localDb.tblListLevelSchemes.Select("schemeName='" + strName + "'");
	
	            if (schemeRows.GetLength(0) != 1)
	            {
	                return -1;
	            }
	
	            foreach (localdbDataSet.tblListLevelSchemesRow row in schemeRows)
	            {
	                localdbDataSet.tblListLevelRow[] listlevelRows = (localdbDataSet.tblListLevelRow[])m_localDb.tblListLevel.Select("schemeName='" + strName + "'", "level");
	
	                if (listlevelRows.GetLength(0) != 9)
	                {
	                    return -1;
	                }
	
	                for (int i = 0; i < 9; i++)
	                {
	                    if (listlevelRows[i].level == (i + 1))
	                    {
	                        localdbDataSet.tblListLevelRow Afnt = listlevelRows[i];
	                        ClassFont Bfnt = newlistLevels[i].Font;
	                      
	                        Afnt.Bold = Bfnt.Bold;
	                        Afnt.Italic = Bfnt.Italic;
	                        Afnt.StrikeThrough = Bfnt.StrikeThrough;
	                        Afnt.Subscript = Bfnt.Subscript;
	                        Afnt.Superscript = Bfnt.Superscript;
	                        Afnt.Shadow = Bfnt.Shadow;
	                        Afnt.Outline = Bfnt.Outline;
	                        Afnt.Emboss = Bfnt.Emboss;
	                        Afnt.Engrave = Bfnt.Engrave;
	                        Afnt.AllCaps = Bfnt.AllCaps;
	                        Afnt.Hidden = Bfnt.Hidden;
	                        Afnt.Underline = (int)Bfnt.Underline; //(Word.WdUnderline)
	                        Afnt.Color = (int)Bfnt.Color;//(Word.WdColor)
	                        Afnt.Size = Bfnt.Size;
	                        Afnt.Animation = (int)Bfnt.Animation; //(Word.WdAnimation)
	                        Afnt.DoubleStrikeThrough = Bfnt.DoubleStrikeThrough;
	                        Afnt.fntName = Bfnt.Name;
	
	                        // update
	                        localdbDataSet.tblListLevelRow Alistlevel = listlevelRows[i];
	                        ClassListLevel Blistlevel = newlistLevels[i];
	
	                        Alistlevel.NumberFormat = Blistlevel.NumberFormat;
	                        Alistlevel.TrailingCharacter = (int)Blistlevel.TrailingCharacter; // (Word.WdTrailingCharacter)
	                        Alistlevel.NumberStyle = (int)Blistlevel.NumberStyle; // (Word.WdListNumberStyle)
	                        Alistlevel.NumberPosition = Blistlevel.NumberPosition;
	                        Alistlevel.Alignment = (int)Blistlevel.Alignment; // (Word.WdListLevelAlignment)
	                        Alistlevel.TextPosition = Blistlevel.TextPosition;
	                        Alistlevel.ResetOnHigher = Blistlevel.ResetOnHigher;
	                        Alistlevel.StartAt = Blistlevel.StartAt;
	                        Alistlevel.LinkedStyle = Blistlevel.LinkedStyle;

                            m_tblAdapterMgr.tblListLevelTableAdapter.Update(listlevelRows[i]);
	                        m_localDb.tblListLevel.AcceptChanges();
	                    }
	                    else
	                    {
	                        
	                    }
	                }
	            }
	
	            m_hashHeadingSnUserDefineScheme[strName] = newlistLevels;
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
                
            }

            return nRet;
        }
        */

        public int reloadHeadingStyleSchemeNH(String strName, ref ClassHeadingStyle[] holdHeadingStyles)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;

            IList<tblHeadingStyleScheme> htItems = null;

            IList<tblHeadingStyleFont> lstFntItems = null;
            IList<tblHeadingStyleParagraphFormat> lstParaFmt = null;

            try
            {
                // tx = session.BeginTransaction();

                IQuery qFnt = session.CreateQuery("from tblHeadingStyleFont where schemeName =:sName order by OutlineLevel").SetString("sName", strName);
                lstFntItems = qFnt.List<tblHeadingStyleFont>();

                IQuery qParaFmt = session.CreateQuery("from tblHeadingStyleParagraphFormat where schemeName =:sName  order by OutlineLevel").SetString("sName", strName);
                lstParaFmt = qParaFmt.List<tblHeadingStyleParagraphFormat>();

                if (!(lstFntItems.Count == 10 && lstParaFmt.Count == 10))
                {
                    //session.Close();
                    return -1;
                }


                tblHeadingStyleFont tFnt = null;
                tblHeadingStyleParagraphFormat tParaFmt = null;

                ClassFont cFnt = null;
                ClassParagraphFormat cParaFmt = null;

                for (int i = 0; i < 10; i++)
                {
                    tFnt = lstFntItems[i];
                    tParaFmt = lstParaFmt[i];
                    // 
                    cFnt = holdHeadingStyles[i].m_fnt;
                    cParaFmt = holdHeadingStyles[i].m_paraFmt;

                    //FONT
                    cFnt.AllCaps = tFnt.AllCaps;
                    cFnt.Animation = (Word.WdAnimation)tFnt.Animation;
                    cFnt.Bold = tFnt.Bold;
                    cFnt.BoldBi = tFnt.BoldBi;
                    cFnt.Color = (Word.WdColor)tFnt.Color;
                    cFnt.ColorIndex = (Word.WdColorIndex)tFnt.ColorIndex;
                    cFnt.ColorIndexBi = (Word.WdColorIndex)tFnt.ColorIndexBi;
                    cFnt.DiacriticColor = (Word.WdColor)tFnt.DiacriticColor;
                    cFnt.DisableCharacterSpaceGrid = tFnt.DisableCharacterSpaceGrid;
                    cFnt.DoubleStrikeThrough = tFnt.DoubleStrikeThrough;
                    cFnt.Emboss = tFnt.Emboss;
                    cFnt.EmphasisMark = (Word.WdEmphasisMark)tFnt.EmphasisMark;
                    cFnt.Engrave = tFnt.Engrave;
                    cFnt.Hidden = tFnt.Hidden;
                    cFnt.Italic = tFnt.Italic;
                    cFnt.ItalicBi = tFnt.ItalicBi;
                    cFnt.Kerning = tFnt.Kerning;

                    cFnt.Name = tFnt.strName;

                    cFnt.NameAscii = tFnt.NameAscii;
                    cFnt.NameBi = tFnt.NameBi;
                    cFnt.NameFarEast = tFnt.NameFarEast;
                    // cFnt.NameOther = "";
                    cFnt.Outline = tFnt.Outline;
                    // cFnt.OutlineLevel = (i + 1);
                    cFnt.Position = tFnt.Position;
                    cFnt.Scaling = tFnt.Scaling;
                    // cFnt.schemeName = strName;
                    cFnt.Shadow = tFnt.Shadow;
                    cFnt.Size = tFnt.Size;
                    cFnt.SizeBi = tFnt.SizeBi;
                    cFnt.SmallCaps = tFnt.SmallCaps;
                    cFnt.Spacing = tFnt.Spacing;
                    cFnt.StrikeThrough = tFnt.StrikeThrough;
                    cFnt.Subscript = tFnt.Subscript;
                    cFnt.Superscript = tFnt.Superscript;
                    cFnt.Underline = (Word.WdUnderline)tFnt.Underline;
                    cFnt.UnderlineColor = (Word.WdColor)tFnt.UnderlineColor;

                    //PARAFMT
                    cParaFmt.AddSpaceBetweenFarEastAndAlpha = tParaFmt.AddSpaceBetweenFarEastAndAlpha;
                    cParaFmt.AddSpaceBetweenFarEastAndDigit = tParaFmt.AddSpaceBetweenFarEastAndDigit;
                    cParaFmt.Alignment = (Word.WdParagraphAlignment)tParaFmt.Alignment;
                    cParaFmt.AutoAdjustRightIndent = tParaFmt.AutoAdjustRightIndent;
                    cParaFmt.BaseLineAlignment = (Word.WdBaselineAlignment)tParaFmt.BaseLineAlignment;
                    cParaFmt.CharacterUnitFirstLineIndent = tParaFmt.CharacterUnitFirstLineIndent;
                    cParaFmt.CharacterUnitLeftIndent = tParaFmt.CharacterUnitLeftIndent;
                    cParaFmt.CharacterUnitRightIndent = tParaFmt.CharacterUnitRightIndent;
                    cParaFmt.DisableLineHeightGrid = tParaFmt.DisableLineHeightGrid;
                    cParaFmt.FarEastLineBreakControl = tParaFmt.FarEastLineBreakControl;
                    cParaFmt.FirstLineIndent = tParaFmt.FirstLineIndent;
                    cParaFmt.HalfWidthPunctuationOnTopOfLine = tParaFmt.HalfWidthPunctuationOnTopOfLine;
                    cParaFmt.HangingPunctuation = tParaFmt.HangingPunctuation;
                    cParaFmt.Hyphenation = tParaFmt.Hyphenation;
                    cParaFmt.KeepTogether = tParaFmt.KeepTogether;
                    cParaFmt.KeepWithNext = tParaFmt.KeepWithNext;
                    cParaFmt.LeftIndent = tParaFmt.LeftIndent;
                    cParaFmt.LineSpacing = tParaFmt.LineSpacing;
                    cParaFmt.LineSpacingRule = (Word.WdLineSpacing)tParaFmt.LineSpacingRule;
                    cParaFmt.LineUnitAfter = tParaFmt.LineUnitAfter;
                    cParaFmt.LineUnitBefore = tParaFmt.LineUnitBefore;
                    cParaFmt.MirrorIndents = tParaFmt.MirrorIndents;
                    cParaFmt.NoLineNumber = tParaFmt.NoLineNumber;
                    cParaFmt.OutlineLevel = (Word.WdOutlineLevel)tParaFmt.OutlineLevel;
                    cParaFmt.PageBreakBefore = tParaFmt.PageBreakBefore;
                    cParaFmt.ReadingOrder = (Word.WdReadingOrder)tParaFmt.ReadingOrder;
                    cParaFmt.RightIndent = tParaFmt.RightIndent;
                    // cParaFmt.schemeName = strName;
                    cParaFmt.SpaceAfter = tParaFmt.SpaceAfter;
                    cParaFmt.SpaceAfterAuto = tParaFmt.SpaceAfterAuto;
                    cParaFmt.SpaceBefore = tParaFmt.SpaceBefore;
                    cParaFmt.SpaceBeforeAuto = tParaFmt.SpaceBeforeAuto;
                    cParaFmt.TextboxTightWrap = (Word.WdTextboxTightWrap)tParaFmt.TextboxTightWrap;
                    cParaFmt.WidowControl = tParaFmt.WidowControl;
                    cParaFmt.WordWrap = tParaFmt.WordWrap;

                }

                // tx.Commit();
            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }

                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            return nRet;
        }

        /*
        public int reloadHeadingStyleScheme_v1(String strName, ref ClassHeadingStyle[] holdHeadingStyles)
        {
            int nRet = 0, nLevel = 0;

            try
            {
                localdbDataSet.tblHeadingStyleSchemeRow[] schemeRows = (localdbDataSet.tblHeadingStyleSchemeRow[])m_localDb.tblHeadingStyleScheme.Select("schemeName='" + strName + "'");

                if (schemeRows.GetLength(0) != 1)
                {
                    return -1;
                }

                foreach (localdbDataSet.tblHeadingStyleSchemeRow row in schemeRows)
                {
                    localdbDataSet.tblHeadingStyleFontRow[] fntRows = (localdbDataSet.tblHeadingStyleFontRow[])m_localDb.tblHeadingStyleFont.Select("schemeName='" + strName + "'", "OutlineLevel");

                    if (fntRows.GetLength(0) != 10)
                    {
                        return -1;
                    }

                    for (int i = 0; i < 10; i++)
                    {
                        nLevel = fntRows[i].OutlineLevel;

                        if (nLevel >= 1 && nLevel <= 10)
                        {
                            localdbDataSet.tblHeadingStyleFontRow Bfnt = fntRows[i];
                            ClassFont Afnt = holdHeadingStyles[nLevel - 1].m_fnt;

                            Afnt.AllCaps = Bfnt.AllCaps;
                            Afnt.Animation = (Word.WdAnimation)Bfnt.Animation;
                            Afnt.Bold = Bfnt.Bold;
                            Afnt.BoldBi = Bfnt.BoldBi;
                            Afnt.Color = (Word.WdColor)Bfnt.Color;
                            Afnt.ColorIndex = (Word.WdColorIndex)Bfnt.ColorIndex;
                            Afnt.ColorIndexBi = (Word.WdColorIndex)Bfnt.ColorIndexBi;
                            Afnt.DiacriticColor = (Word.WdColor)Bfnt.DiacriticColor;
                            Afnt.DisableCharacterSpaceGrid = Bfnt.DisableCharacterSpaceGrid;
                            Afnt.DoubleStrikeThrough = Bfnt.DoubleStrikeThrough;
                            Afnt.Emboss = Bfnt.Emboss;
                            Afnt.EmphasisMark = (Word.WdEmphasisMark)Bfnt.EmphasisMark;
                            Afnt.Engrave = Bfnt.Engrave;
                            Afnt.Hidden = Bfnt.Hidden;
                            Afnt.Italic = Bfnt.Italic;
                            Afnt.ItalicBi = Bfnt.ItalicBi;
                            Afnt.Kerning = Bfnt.Kerning;
                            Afnt.Name = Bfnt.strName;
                            Afnt.NameAscii = Bfnt.NameAscii;
                            Afnt.NameBi = Bfnt.NameBi;
                            Afnt.NameFarEast = Bfnt.NameFarEast;
                            //Afnt.NameOther = Bfnt.NameOther;
                            Afnt.Outline = Bfnt.Outline;
                            Afnt.Position = Bfnt.Position;
                            Afnt.Scaling = Bfnt.Scaling;
                            Afnt.Shadow = Bfnt.Shadow;
                            Afnt.Size = Bfnt.Size;
                            Afnt.SizeBi = Bfnt.SizeBi;
                            Afnt.SmallCaps = Bfnt.SmallCaps;
                            Afnt.Spacing = Bfnt.Spacing;
                            Afnt.StrikeThrough = Bfnt.StrikeThrough;
                            Afnt.Subscript = Bfnt.Subscript;
                            Afnt.Superscript = Bfnt.Superscript;
                            Afnt.Underline = (Word.WdUnderline)Bfnt.Underline;
                            Afnt.UnderlineColor = (Word.WdColor)Bfnt.UnderlineColor;

                        }
                        else
                        {

                        }
                    }

                    // paragraph format
                    localdbDataSet.tblHeadingStyleParagraphFormatRow[] paraFmtRows = (localdbDataSet.tblHeadingStyleParagraphFormatRow[])m_localDb.tblHeadingStyleParagraphFormat.Select("schemeName='" + strName + "'", "OutlineLevel");

                    if (paraFmtRows.GetLength(0) != 10)
                    {
                        return -1;
                    }

                    for (int i = 0; i < 10; i++)
                    {
                        nLevel = fntRows[i].OutlineLevel;

                        if (nLevel >= 1 && nLevel <= 10)
                        {
                            localdbDataSet.tblHeadingStyleParagraphFormatRow Bfmt = paraFmtRows[i];
                            ClassParagraphFormat Afmt = holdHeadingStyles[nLevel - 1].m_paraFmt;

                            Afmt.AddSpaceBetweenFarEastAndAlpha = Bfmt.AddSpaceBetweenFarEastAndAlpha;
                            Afmt.AddSpaceBetweenFarEastAndDigit = Bfmt.AddSpaceBetweenFarEastAndDigit;
                            Afmt.Alignment = (Word.WdParagraphAlignment)Bfmt.Alignment;
                            Afmt.AutoAdjustRightIndent = Bfmt.AutoAdjustRightIndent;
                            Afmt.BaseLineAlignment = (Word.WdBaselineAlignment)Bfmt.BaseLineAlignment;
                            Afmt.CharacterUnitFirstLineIndent = Bfmt.CharacterUnitFirstLineIndent;
                            Afmt.CharacterUnitLeftIndent = Bfmt.CharacterUnitLeftIndent;
                            Afmt.CharacterUnitRightIndent = Bfmt.CharacterUnitRightIndent;
                            Afmt.DisableLineHeightGrid = Bfmt.DisableLineHeightGrid;
                            Afmt.FarEastLineBreakControl = Bfmt.FarEastLineBreakControl;
                            Afmt.FirstLineIndent = Bfmt.FirstLineIndent;
                            Afmt.HalfWidthPunctuationOnTopOfLine = Bfmt.HalfWidthPunctuationOnTopOfLine;
                            Afmt.HangingPunctuation = Bfmt.HangingPunctuation;
                            Afmt.Hyphenation = Bfmt.Hyphenation;
                            Afmt.KeepTogether = Bfmt.KeepTogether;
                            Afmt.KeepWithNext = Bfmt.KeepWithNext;
                            Afmt.LeftIndent = Bfmt.LeftIndent;
                            Afmt.LineSpacing = Bfmt.LineSpacing;
                            Afmt.LineSpacingRule = (Word.WdLineSpacing)Bfmt.LineSpacingRule;
                            Afmt.LineUnitAfter = Bfmt.LineUnitAfter;
                            Afmt.LineUnitBefore = Bfmt.LineUnitBefore;
                            Afmt.MirrorIndents = Bfmt.MirrorIndents;
                            Afmt.NoLineNumber = Bfmt.NoLineNumber;
                            Afmt.OutlineLevel = (Word.WdOutlineLevel)Bfmt.OutlineLevel;
                            Afmt.PageBreakBefore = Bfmt.PageBreakBefore;
                            Afmt.ReadingOrder = (Word.WdReadingOrder)Bfmt.ReadingOrder;
                            Afmt.RightIndent = Bfmt.RightIndent;
                            Afmt.SpaceAfter = Bfmt.SpaceAfter;
                            Afmt.SpaceAfterAuto = Bfmt.SpaceAfterAuto;
                            Afmt.SpaceBefore = Bfmt.SpaceBefore;
                            Afmt.SpaceBeforeAuto = Bfmt.SpaceBeforeAuto;
                            Afmt.TextboxTightWrap = (Word.WdTextboxTightWrap)Bfmt.TextboxTightWrap;
                            Afmt.WidowControl = Bfmt.WidowControl;
                            Afmt.WordWrap = Bfmt.WordWrap;
                        }
                        else
                        {

                        }
                    }


                }

            }
            catch (System.Exception ex)
            {

            }
            finally
            {

            }

            return nRet;

        }
        */

        public int reloadHeadingSnSchemeNH(String strName, ref ClassListLevel[] holdListLevels)
        {
            int nRet = 0;

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblListLevel> lstItems = null;

            try
            {
                IQuery hsSchemes = session.CreateQuery("from tblListLevel where schemeName =:sName order by level").SetString("sName", strName);
                lstItems = hsSchemes.List<tblListLevel>();

                if (lstItems.Count != 9) // 
                {
                    //session.Close();

                    return -1;
                }

                // tx = session.BeginTransaction();

                // add into listlevel table
                tblListLevel snItem = null;
                ClassListLevel lvl = null;

                for (int i = 0; i < 9; i++)
                {
                    lvl = holdListLevels[i];

                    snItem = lstItems[i];

                    // snItem.ID = Guid.NewGuid();

                    lvl.Alignment = (Word.WdListLevelAlignment)snItem.Alignment;
                    lvl.Font.AllCaps = snItem.AllCaps;
                    lvl.Font.Animation = (Word.WdAnimation)snItem.Animation;
                    lvl.Font.Bold = snItem.Bold;
                    lvl.Font.Color = (Word.WdColor)snItem.Color;
                    lvl.Font.DoubleStrikeThrough = snItem.DoubleStrikeThrough;
                    lvl.Font.Emboss = snItem.Emboss;


                    lvl.Font.Engrave = snItem.Engrave;
                    lvl.Font.Name = snItem.fntName;
                    lvl.Font.Hidden = snItem.Hidden;
                    lvl.Font.Italic = snItem.Italic;
                    lvl.Index = snItem.level; // i
                    lvl.LinkedStyle = snItem.LinkedStyle;
                    lvl.NumberFormat = snItem.NumberFormat;

                    lvl.NumberPosition = snItem.NumberPosition;
                    lvl.NumberStyle = (Word.WdListNumberStyle)snItem.NumberStyle;
                    lvl.Font.Outline = snItem.Outline;
                    lvl.ResetOnHigher = snItem.ResetOnHigher;

                    // snItem.schemeId = -1;
                    // snItem.schemeName = strName;

                    lvl.Font.Shadow = snItem.Shadow;
                    lvl.Font.Size = snItem.Size;
                    lvl.StartAt = snItem.StartAt;
                    lvl.Font.StrikeThrough = snItem.StrikeThrough;
                    lvl.Font.Subscript = snItem.Subscript;
                    lvl.Font.Superscript = snItem.Superscript;
                    lvl.TextPosition = snItem.TextPosition;
                    lvl.TrailingCharacter = (Word.WdTrailingCharacter)snItem.TrailingCharacter;
                    lvl.Font.Underline = (Word.WdUnderline)snItem.Underline;

                    // session.SaveOrUpdate(snItem);
                }

                // tx.Commit();

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }


            return nRet;
        }

        /*
        public int reloadHeadingSnScheme_v1(String strName, ref ClassListLevel[] holdListLevels)
        {
            int nRet = 0;

//             localdbDataSet.tblListLevelSchemesRow[] schemeRows = (localdbDataSet.tblListLevelSchemesRow[])m_localDb.tblListLevelSchemes.Select("schemeName='" + strName + "'");
// 
//             if (schemeRows.GetLength(0) != 1 || holdListLevels == null || holdListLevels.GetLength(0) != 9)
//             {
//                 return -1;
//             }
// 
//             foreach (localdbDataSet.tblListLevelSchemesRow row in schemeRows)
//             {

            try
            {
	            localdbDataSet.tblListLevelRow[] listlevelRows = (localdbDataSet.tblListLevelRow[])m_localDb.tblListLevel.Select("schemeName='" + strName + "'", "level");
	
	            if (listlevelRows.GetLength(0) != 9)
	            {
	                return -1;
	            }
	
	            for (int i = 0; i < 9; i++)
	            {
	                if (listlevelRows[i].level == (i + 1))
	                {
	                    localdbDataSet.tblListLevelRow Bfnt = listlevelRows[i];
	                    ClassFont Afnt = holdListLevels[i].Font;
	
	                    Afnt.Bold = Bfnt.Bold;
	                    Afnt.Italic = Bfnt.Italic;
	                    Afnt.StrikeThrough = Bfnt.StrikeThrough;
	                    Afnt.Subscript = Bfnt.Subscript;
	                    Afnt.Superscript = Bfnt.Superscript;
	                    Afnt.Shadow = Bfnt.Shadow;
	                    Afnt.Outline = Bfnt.Outline;
	                    Afnt.Emboss = Bfnt.Emboss;
	                    Afnt.Engrave = Bfnt.Engrave;
	                    Afnt.AllCaps = Bfnt.AllCaps;
	                    Afnt.Hidden = Bfnt.Hidden;
	                    Afnt.Underline = (Word.WdUnderline)Bfnt.Underline; //(Word.WdUnderline)
	                    Afnt.Color = (Word.WdColor)Bfnt.Color;//(Word.WdColor)
	                    Afnt.Size = Bfnt.Size;
	                    Afnt.Animation = (Word.WdAnimation)Bfnt.Animation; //(Word.WdAnimation)
	                    Afnt.DoubleStrikeThrough = Bfnt.DoubleStrikeThrough;
	                    Afnt.Name = Bfnt.fntName;
	                        
	
	                    // fill
	                    localdbDataSet.tblListLevelRow Blistlevel = listlevelRows[i];
	                    ClassListLevel Alistlevel = holdListLevels[i];
	
	                    Alistlevel.NumberFormat = Blistlevel.NumberFormat;
	                    Alistlevel.TrailingCharacter = (Word.WdTrailingCharacter)Blistlevel.TrailingCharacter; // (Word.WdTrailingCharacter)
	                    Alistlevel.NumberStyle = (Word.WdListNumberStyle)Blistlevel.NumberStyle; // (Word.WdListNumberStyle)
	                    Alistlevel.NumberStyleSel = (Word.WdListNumberStyle)Blistlevel.NumberStyle; // (Word.WdListNumberStyle)
	                    Alistlevel.NumberPosition = Blistlevel.NumberPosition;
	                    Alistlevel.Alignment = (Word.WdListLevelAlignment)Blistlevel.Alignment; // (Word.WdListLevelAlignment)
	                    Alistlevel.TextPosition = Blistlevel.TextPosition;
	                    Alistlevel.ResetOnHigher = Blistlevel.ResetOnHigher;
	                    Alistlevel.StartAt = Blistlevel.StartAt;
	                    Alistlevel.LinkedStyle = Blistlevel.LinkedStyle;
	                }
	                else
	                {
	
	                }
	            }
//            }
	
	        // m_hashHeadingSnUserDefineScheme[strName] = holdListLevels;
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
                
            }

            return nRet;
        }
        */

//         public void loadUniformStyleHistoryStyleDocs()
//         {
//             //@TODO, get from DB 
//             return;
//         }
// 
//         public void saveUniformStyleHistoryStyleDocs()
//         {
//             //@TODO, save into DB
//             return;
//         }

        public int addUniformStyleHistoryStyleDocsNH(String strNewDoc, ref String strRetMsg)
        {
            int nRet = 0;
            
            strRetMsg = "成功";

            ISession session = dbNHmgr.getSession();
            ITransaction tx = null;
            IList<tblUniformStyleHistoryDocs> lstItems = null;

            try
            {
                //tx = session.BeginTransaction();

                IQuery hsSchemes = session.CreateQuery("from tblUniformStyleHistoryDocs where fullPathDoc =:sName").SetString("sName", strNewDoc);
                lstItems = hsSchemes.List<tblUniformStyleHistoryDocs>();

                if (lstItems.Count > 0) // 
                {
                    //session.Close();
                    strRetMsg = "文件已存在，不需要加入";
                    return -1;
                }

                tx = session.BeginTransaction();

                // add into listlevel table
                tblUniformStyleHistoryDocs snItem = new tblUniformStyleHistoryDocs();

                snItem.ID = Guid.NewGuid();
                snItem.fullPathDoc = strNewDoc;

                session.Save(snItem);

                tx.Commit();

            }
            catch (System.Exception ex)
            {
                if (tx != null)
                {
                    tx.Rollback();
                }
                strRetMsg = ex.Message;
                nRet = -1;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                session.Close();
            }

            return nRet;
        }

        /*
        public int addUniformStyleHistoryStyleDocs_v1(String strNewDoc,ref String strRetMsg)
        {
            int nRet = 0;
            strRetMsg = "成功";

            try
            {
	            localdbDataSet.tblUniformStyleHistoryDocsRow[] rows = (localdbDataSet.tblUniformStyleHistoryDocsRow[])m_localDb.tblUniformStyleHistoryDocs.Select("fullPathDoc='" + strNewDoc + "'");
	
	            if (rows.GetLength(0) > 0)
	            {
	                strRetMsg = "文件已存在，不需要加入";
	                return -1;
	            }
	            else
	            {
	                m_localDb.tblUniformStyleHistoryDocs.AddtblUniformStyleHistoryDocsRow(strNewDoc);
	            }
	
	            if (m_localDb.tblUniformStyleHistoryDocs.Count > m_nMaxUniformStyleHistoryStyleDocs)
	            {
	                rows = (localdbDataSet.tblUniformStyleHistoryDocsRow[])m_localDb.tblUniformStyleHistoryDocs.Select("fullPathDoc","ID ASC");
	
	                for (int i = (int)m_nMaxUniformStyleHistoryStyleDocs + 1; i < m_localDb.tblUniformStyleHistoryDocs.Count; i++)
	                {
	                    rows[i].BeginEdit();
	                    rows[i].Delete();
	                    rows[i].EndEdit();
	                }
	
	            }

                m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Update(m_localDb.tblUniformStyleHistoryDocs);

                m_localDb.tblUniformStyleHistoryDocs.Clear();
                m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Fill(m_localDb.tblUniformStyleHistoryDocs);

            }
            catch (System.Exception ex)
            {
                strRetMsg = ex.Message;
                nRet = -1;
            }
            finally
            {
                
            }



// 
//             if (m_hashUniformStyleHistoryStyleDocs.Contains(strNewDoc))
//             {
//                 strRetMsg = "文件已存在，不需要加入";
//                 return 1;
//             }
// 
//             m_arrUniformStyleHistoryStyleDocs.Insert(0, strNewDoc);
// 
//             String strItem = "";
//             uint nCnt = (uint)m_arrUniformStyleHistoryStyleDocs.Count;
//             if (nCnt > m_nMaxUniformStyleHistoryStyleDocs)
//             {
//                 for (int i = (int)m_nMaxUniformStyleHistoryStyleDocs + 1; i < nCnt; i++)
//                 {
//                     strItem = (String)m_arrUniformStyleHistoryStyleDocs[i];
//                     m_hashUniformStyleHistoryStyleDocs.Remove(strItem);
//                 }
// 
//                 m_arrUniformStyleHistoryStyleDocs.RemoveRange((int)m_nMaxUniformStyleHistoryStyleDocs, (int)(nCnt - m_nMaxUniformStyleHistoryStyleDocs));
//             }

            // 
            
            return nRet;
        }
        */

        public void copyMultiStyles(Word.Range rng)
        {
            Word.Application app = this.Application;
            Word.Document curDoc = app.ActiveDocument;
            Word.Selection sel = curDoc.ActiveWindow.Selection;


            int nOStart = sel.Start;
            int nOEnd = sel.End;

            m_srcArrListLevels.Clear();
            m_hashHeadingFont.Clear();
            m_hashHeadingParaFormat.Clear();

            Word.ListLevels templateLstLvls = null;

            foreach (Word.Paragraph para in rng.Paragraphs)
            {
                // 
                if (para.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText &&
                    para.Range.Text.Trim().Equals(""))
                    continue;


                curDoc.ActiveWindow.ScrollIntoView(para.Range, true);

                // get list levels
                // 
                if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText &&
                    para.Range.ListFormat.ListType == Word.WdListType.wdListOutlineNumbering &&
                    para.Range.ListFormat.ListTemplate != null &&
                    templateLstLvls == null)
                {
                    templateLstLvls = para.Range.ListFormat.ListTemplate.ListLevels;

                    if (templateLstLvls != null)
                    {
                        // 
                        for (int i = 1; i <= templateLstLvls.Count; i++)
                        {
                            // 
                            ClassListLevel lstLvl = new ClassListLevel();
                            //lstLvl.Font = new ClassFont();

                            lstLvl.clone(templateLstLvls[i]);

                            m_srcArrListLevels.Add(lstLvl);

                        }
                    }
                }

                if (!m_hashHeadingFont.Contains(para.OutlineLevel))
                {
                    ClassFont fnt = new ClassFont();

                    fnt.clone(para.Range.Font);

                    m_hashHeadingFont.Add(para.OutlineLevel, fnt);
                }


                if (!m_hashHeadingParaFormat.Contains(para.OutlineLevel))
                {
                    ClassParagraphFormat paraFmt = new ClassParagraphFormat();
                    paraFmt.clone(para.Range.ParagraphFormat);

                    m_hashHeadingParaFormat.Add(para.OutlineLevel, paraFmt);
                }

            }

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);


            return;
        }



        // 
        public String applyMultiStyles(Word.Range rng)
        {

            Word.Application app = this.Application;
            Word.Document curDoc = app.ActiveDocument;
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            // apply
            Word.ListGallery listGallery = null;
            Word.ListTemplate lstTemplate = null;
            Word.ListLevels lstLvels = null;
            Object objIndex = 1;
            
            Boolean bListLevel = (m_srcArrListLevels.Count > 0);

            if (bListLevel)
            {
                if (m_bAppIsWps)
                {
                    listGallery = app.ListGalleries[(Word.WdListGalleryType)4];

                    if (listGallery.ListTemplates.Count == 0)
                    {
                        Object objOutlineNumbered = true;
                        lstTemplate = listGallery.ListTemplates.Add(objOutlineNumbered);
                        lstLvels = lstTemplate.ListLevels;
                    }
                    else
                    {
                        lstTemplate = listGallery.ListTemplates[objIndex];
                        lstLvels = lstTemplate.ListLevels;
                    }

                }
                else
                {
                    listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];
                    lstTemplate = listGallery.ListTemplates[objIndex];
                    lstLvels = lstTemplate.ListLevels;
                }
            }


            if (bListLevel && lstLvels != null)
            {
                for (int i = 1; i <= lstLvels.Count; i++)
                {
                    ClassListLevel lstLvl = (ClassListLevel)m_srcArrListLevels[i - 1];
                    Word.ListLevel wordLstLvl = lstLvels[i];
                    lstLvl.copy2(ref wordLstLvl);
                }
            }

            Object objContinue = Word.WdContinue.wdContinueDisabled;
            Object objApplyTo = Word.WdListApplyTo.wdListApplyToSelection; // wdListApplyToWholeList;
            Object objDefaultBehav = Word.WdDefaultListBehavior.wdWord9ListBehavior;

            int[] nArrCnt = new int[11];

            foreach (Word.Paragraph dstPara in rng.Paragraphs)
            {
                if (dstPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText ||
                    dstPara.Range.Text.Trim().Equals(""))
                {
                    continue;
                }

                curDoc.ActiveWindow.ScrollIntoView(dstPara.Range,true);

                ClassFont cpFnt = (ClassFont)m_hashHeadingFont[dstPara.OutlineLevel];
                ClassParagraphFormat cpParaFmt = (ClassParagraphFormat)m_hashHeadingParaFormat[dstPara.OutlineLevel];

                if (dstPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    nArrCnt[(int)dstPara.OutlineLevel]++;

                    if (bListLevel)
                    {
                        dstPara.Range.ListFormat.ApplyListTemplateWithLevel(lstTemplate,
                           objContinue, objApplyTo, objDefaultBehav);
                    }

                    // 
                    if (cpFnt != null)
                    {
                        Word.Font fnt = dstPara.Range.Font;
                        cpFnt.copy2(fnt);

                        if (bListLevel)
                        {
                            if (m_bAppIsWps)
                            {
                                // copy font
                                int nLevel = (int)dstPara.OutlineLevel;

                                Word.ListLevel curLstLevel = lstTemplate.ListLevels[nLevel];

                                cpFnt.copy2(curLstLevel.Font);
                            }
                            else
                            {
                                // WORD
                            }
                        }
                    }


                    if (cpParaFmt != null)
                    {
                        Word.ParagraphFormat paraFmt = dstPara.Range.ParagraphFormat;
                        cpParaFmt.copy2(paraFmt);
                    }

                    // list levels
                    
                }
                else
                {
                    if (cpParaFmt != null)
                    {
                        Word.ParagraphFormat paraFmt = dstPara.Range.ParagraphFormat;
                        cpParaFmt.copy2(paraFmt);
                    }
                }

            }

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            String strRet = "";

            for (int i = 1; i < 10; i++)
            {
                strRet += i + "级：" + nArrCnt[i] + "个\r\n";
            }

            return strRet;
        }


        public Boolean m_bLoginedStatus = false;
        public Boolean m_bLoginAbnormal = false;
        public String m_strLoginedUser = "";
        public String m_strLoginedPass = "";

        public ShareContributorOper m_HttpOper = new ShareContributorOper();

        private Hashtable m_hashTaskPane = new Hashtable();
        
        public Hashtable HashTaskPane
        { 
            get 
            {
                return m_hashTaskPane;
            } 
        }

        private Hashtable m_hashDocVisible = new Hashtable();

        public Hashtable HashDocVisible
        {
            get
            {
                return m_hashDocVisible;
            }
        }

        public Boolean m_bUpdTblCntOnSaving = false;
        public Boolean m_bUpdTblCntOnClosing = false;

//         public Hashtable m_hashDefaultPermission = new Hashtable();
//         public Hashtable m_hashVstoPermission = null;

        public Hashtable m_hashFilePermission = new Hashtable();

        public void createDefaultPermission()
        {
            String[] strsExceptionName = { "btnLogin", "ribBtnHelp", 
                                           "ribBtnTutorial", "ribbtnAbout",
                                           "RibbtnRegister", "ribLoadSoloLic" };

            foreach (String strItem in strsExceptionName)
            {
                // m_uiCtrler.addExceptionalUiItem(strItem);
                m_edtCenter.AddExceptionalName(strItem);
            }
            
            return;
        }

        public int searchPermission(String strName)
        {
            Boolean bRet = m_edtCenter.IsEnableViaPlainString(strName);
            // int nVal = m_uiCtrler.searchPermission(strName);
            int nVal = (bRet ? 1 : 0);
#if ADMIN
            nVal = 1;
#endif
            return nVal;

        }

//         public void updatePermission(Hashtable hashControls)
//         {
//             updatePermission(hashControls, m_hashVstoPermission);
//             return;
//         }


//         private void updatePermission(Hashtable hashControls, Hashtable hashPermission)
//         {
//             String strName = "";
//             int nVal = 0;
//             Control ctrl = null;
//             ToolStripItem item = null;
//             RibbonControl ribCtrl = null;
//             Object ctrlObj = null;
// 
//             hashPermission["btnLogin"] = 1;
//             hashPermission["ribBtnHelp"] = 1;
//             hashPermission["ribBtnTutorial"] = 1;
//             hashPermission["ribbtnAbout"] = 1;
//             hashPermission["RibbtnRegister"] = 1;
// 
//             foreach (DictionaryEntry entry in hashControls)
//             {
//                 strName = (String)entry.Key;
//                 ctrlObj = (Object)entry.Value;
// 
//                 nVal = 0;
//                 if (hashPermission.Contains(strName))
//                 {
//                     nVal = (int)hashPermission[strName];
//                 }
// 
// #if ADMIN
//                 nVal = 1;
// #endif
// 
//                 if (m_bTryExpired)
//                 {
//                     nVal = 0;
//                 }
// 
//                 if (ctrlObj != null)
//                 {
//                     if (ctrlObj is Control)
//                     {
//                         ctrl = (Control)ctrlObj;
//                         if (ctrl.Controls.Count > 0)
//                         {
//                             ctrl.Enabled = true;
//                         }
//                         else
//                         {
//                             ctrl.Enabled = (nVal != 0);
//                         }
//                     }
//                     else if (ctrlObj is ToolStripItem)
//                     {
//                         item = (ToolStripItem)ctrlObj;
//                         item.Enabled = (nVal != 0);
//                     }
//                     else if (ctrlObj is RibbonGroup)
//                     {
//                         // 
//                     }
//                     else if (ctrlObj is RibbonControl)
//                     {
//                         ribCtrl = (RibbonControl)ctrlObj;
//                         ribCtrl.Enabled = (nVal != 0);
//                     }
//                 }// if
// 
//             }
//            
// 
// //             IDictionaryEnumerator idIter = hashPermission.GetEnumerator();
// // 
// //             while(idIter.MoveNext())
// //             {
// //                 strName = (String)idIter.Key;
// //                 nVal = (int)idIter.Value;
// // 
// //                 ctrlObj = (Object)hashControls[strName];
// // 
// //                 if (ctrlObj != null)
// //                 {
// //                     if (ctrlObj is Control)
// //                     {
// //                         ctrl = (Control)ctrlObj;
// //                         if (ctrl.Controls.Count > 0)
// //                         {
// //                             ctrl.Enabled = true;
// //                         }
// //                         else
// //                         {
// //                             ctrl.Enabled = (nVal != 0);
// //                         }
// //                     }
// //                     else if (ctrlObj is ToolStripItem)
// //                     {
// //                         item = (ToolStripItem)ctrlObj;
// //                         item.Enabled = (nVal != 0);
// //                     }
// //                     else if (ctrlObj is RibbonGroup)
// //                     {
// //                         // 
// //                     }
// //                     else if (ctrlObj is RibbonControl)
// //                     {
// //                         ribCtrl = (RibbonControl)ctrlObj;
// //                         ribCtrl.Enabled = (nVal != 0);                        
// //                     }
// //                 }// if
// // 
// //             }// while
// 
//             // always enable login button
//             // 
// 
//             return;
//         }

 

        public void logout(Word.Document doc)
        {
            m_bLoginedStatus = false;

            // m_hashVstoPermission = m_hashDefaultPermission;
            // update permission
            Object objKey = doc; // doc.CurrentRsid;
            CustomTaskPane curPane = (CustomTaskPane)m_hashTaskPane[objKey];

            if (curPane == null)
                return;

            // update permission
            IDictionaryEnumerator iter = m_hashTaskPane.GetEnumerator();
            CustomTaskPane myPane = null;
            OperationPanel userPane = null;

            userPane = (OperationPanel)curPane.Control;

            // m_uiCtrler.restoreDefaultDocRepUiPermHash();
            // m_uiCtrler.updateUI(userPane.m_hashControls);

            m_edtCenter.ResetEditionPerms(m_edtCenter.m_strDocRepositoryEditionName);
            m_edtCenter.UpdateUI(userPane.m_hashControls);

            // updatePermission(userPane.m_hashControls);
            //userPane.PermissionChangeRefresh();
            userPane.Invalidate();

            while (iter.MoveNext())
            {
                myPane = (CustomTaskPane)iter.Value;
                if (myPane == curPane)
                    continue;

                userPane = (OperationPanel)myPane.Control;
                // m_uiCtrler.updateUI(userPane.m_hashControls);
                m_edtCenter.UpdateUI(userPane.m_hashControls);
                //userPane.PermissionChangeRefresh();
            }

            return;
        }


        public void recordCommonShareLibTree(TreeNodeCollection oShareLibTree)
        {
            m_trvShareLibNodes = oShareLibTree;

            return;
        }


        public TreeNodeCollection getCommonShareLibTree()
        {
            return m_trvShareLibNodes;
        }


        public int login(String strUsername, String strPassword, ref String strRetMsg)
        {
            int nRet = -1;
            
            nRet = m_HttpOper.login(strUsername, strPassword, ref strRetMsg);
            m_strLoginedUser = strUsername;
            m_strLoginedPass = strPassword;

            m_bLoginedStatus = (nRet == 0);

            m_bLoginAbnormal = (nRet != 0);

            if(m_bLoginedStatus)
            {
                Hashtable vstoHash = m_HttpOper.getVstoPermission(strUsername);
                // m_uiCtrler.setDocRepUiPermHash(vstoHash);
                m_edtCenter.UpdateEditionPerms(m_edtCenter.m_strDocRepositoryEditionName, vstoHash);

                Word.Application app = this.Application;
                Word.Document doc = app.ActiveDocument;

                CustomTaskPane myPane = (CustomTaskPane)m_hashTaskPane[doc];

                OperationPanel userPane = (OperationPanel)myPane.Control;

                // m_uiCtrler.updateUI(userPane.m_hashControls); // update all existed taskpanes
                m_edtCenter.UpdateUI(userPane.m_hashControls);
                
                userPane.RefreshRelsByPermission();
                userPane.RefreshShareLibByPermission();
                userPane.refreshMyComputerFolders();
                userPane.recordCommonShareLibTree();

                userPane.Invalidate();

                Word.Document oDoc = null, openDoc = null;

                foreach (DictionaryEntry ent in m_hashTaskPane)
                {
                    oDoc = (Word.Document)ent.Key;

                    try
                    {
                        if (m_bAppIsWps)
                        {
                            openDoc = app.Documents[oDoc.Name];
                        }
                        else
                        {
                            openDoc = app.Documents[oDoc];
                        }
                    }
                    catch (System.Exception ex)
                    {
                        continue;
                    }
                    finally
                    {
                    }

                    if (oDoc == doc || openDoc == null)
                    {
                        continue;
                    }

                    myPane = (CustomTaskPane)ent.Value;
                    userPane = (OperationPanel)myPane.Control;

                    // m_uiCtrler.updateUI(userPane.m_hashControls);
                    m_edtCenter.UpdateUI(userPane.m_hashControls);
                    userPane.RefreshRelsByPermission();

                    userPane.cloneShareLibTree();

                    userPane.Invalidate();
                }
                
                // m_hashFilePermission = m_HttpOper.getFilePermissions(strUsername, ref strRetMsg);

                /*

                // update permission
                IDictionaryEnumerator iter = m_hashTaskPane.GetEnumerator();
                CustomTaskPane myPane = null;
                OperationPanel userPane = null;
                
                TabControl.TabPageCollection tpagColl = null;
                TabPage tabPageShare = null;

                while (iter.MoveNext())
                {
                    myPane = (CustomTaskPane)iter.Value;
                    userPane = (OperationPanel)myPane.Control;
                    updatePermission(userPane.m_hashControls);
                    userPane.PermissionChangeRefresh();

                    // tpagColl = userPane.tabCtrl.TabPages;

                    userPane.Invalidate();
                }
                    
                // 
                */

                // 签到成功后再检查更新
                // if (Settings.Default.bAutoUpdate && searchPermission("chkAutoCheckUpdate") > 0)
                if (Settings.Default.bAutoUpdate && searchPermission("grpComOp") > 0)
                {
                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[2029] Doc Login to auto update");
                    }

                    // m_uiCtrler.checkUpdate();
                    m_edtCenter.CheckUpdate();
                }
            }
            
            return nRet;
        }


        public void TryAutoLogin()
        {
            if (Settings.Default.strLnNm.Equals(""))
                return;

            if (Settings.Default.bALn && !m_bLoginedStatus /*&& searchPermission(Globals.Ribbons.Ribbon1.chkAutoLogin.Name) > 0*/)
            {
                Globals.Ribbons.Ribbon1.chkAutoLogin.Checked = true;
                String strPass = "", strRetMsg = "", strName = "";


                strName = ClassEncryptUtils.DESDecrypt(Settings.Default.strLnNm, m_stryp, m_stryv);
                strPass = ClassEncryptUtils.DESDecrypt(Settings.Default.strLnws, m_stryp, m_stryv);

                if (strName == null || strPass == null)
                    return;

                // login
                int nRet = login(strName, strPass, ref strRetMsg);

                if (m_bLoginedStatus)
                {
                    Globals.Ribbons.Ribbon1.btnLogin.Label = "注销:" + m_strLoginedUser;
                }
                else
                {
                    Globals.Ribbons.Ribbon1.btnLogin.Label = "登录";
                    MessageBox.Show("自动登录失败：" + strRetMsg);
                }
            }

            return;
        }


        private void CleanInvalidDoc()
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2111] CleanInvalidDoc Enter");
            }

            Word.Application app = this.Application;

            // app.Documents
            Word.Document tmpdoc = null, openedDoc = null;
            ArrayList arrRemoveItems = new ArrayList();

            foreach (DictionaryEntry ent in m_hashTaskPane)
            {
                try
                {
                	tmpdoc = (Word.Document)ent.Key;
                }
                catch (System.Exception ex)
                {
                    arrRemoveItems.Add(ent.Key);
                    continue;
                }
                finally
                {
                }

                if (tmpdoc != null)
                {
                    try
                    {
                        if (m_bAppIsWps)
                        {
                            openedDoc = app.Documents[tmpdoc.Name];
                        }
                        else
                        {
                            openedDoc = app.Documents[tmpdoc];
                        }

                        if (openedDoc == null)
                        {
                            arrRemoveItems.Add(ent.Key);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        arrRemoveItems.Add(ent.Key);
                    }
                    finally
                    {
                    }
                }
            }


            Object obj = null;
            CustomTaskPane myCustomTaskPane = null;

            for(int i = 0; i < arrRemoveItems.Count; i++)
            {
                try
                {
	                obj = arrRemoveItems[i];
                    myCustomTaskPane = (CustomTaskPane)m_hashTaskPane[obj];

                    this.CustomTaskPanes.Remove(myCustomTaskPane);
                    m_hashTaskPane.Remove(obj);
                }
                catch (System.Exception ex)
                {
                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[2174]"+ex.Message);
                    }
                }
                finally
                {
                }
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2184] CleanInvalidDoc Exit," + arrRemoveItems.Count);
            }

            return;
        }


        public void AddTaskPane(Word.Document doc)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2193] AddTaskPane,Enter");
            }

            int ncode = doc.GetHashCode();

            Object objKey = doc; // doc.CurrentRsid;
            CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_hashTaskPane[objKey];

            // MessageBox.Show("" + doc.GetHashCode());
            CleanInvalidDoc();

            if (myCustomTaskPane != null || this.CustomTaskPanes.Count != m_hashTaskPane.Count ||
                this.Application.Documents.Count - m_hashTaskPane.Count != 1)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2209] AddTaskPane,Exit xxx:" + (myCustomTaskPane != null) + "," +
                                 this.CustomTaskPanes.Count + "," + m_hashTaskPane.Count + "," + this.Application.Documents.Count +
                                 "," + m_hashTaskPane.Count);
                }

                return;
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2221]");
            }

            OperationPanel userPane = null;
            
            try
            {
            	userPane = new OperationPanel();
            }
            catch (System.Exception ex)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2234] " + ex.Message);
                }

                return;
            }
            finally
            {
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2228]");
            }

            userPane.SetScOper(m_HttpOper);

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2235]");
            }

            myCustomTaskPane = this.CustomTaskPanes.Add(userPane, "工作区",this.Application.ActiveWindow); // doc.ActiveWindow

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2242]");
            }

            if (myCustomTaskPane == null)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2227] Fail CustomTaskPanes.Add");
                }

                return;
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2257]");
            }

            myCustomTaskPane.VisibleChanged += new EventHandler(myCustomTaskPane_VisibleChanged);
            myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2264]");
            }

            //myCustomTaskPane.Height = 500;
            myCustomTaskPane.Width = 440;

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2272]");
            }

            myCustomTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2279]");
            }

            if (m_bDontShowPane)
            {
                myCustomTaskPane.Visible = false;
            }
            else
            {
                myCustomTaskPane.Visible = Settings.Default.bUIShow;// true
            }

            //myCustomTaskPane.Visible = true;

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2286]");
            }

            m_hashTaskPane.Add(objKey, myCustomTaskPane);

            if (m_bAppIsWps)
            {
                m_hashDocVisible[doc] = myCustomTaskPane.Visible;
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2293]");
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strCfgDir = strBaseDir + @"config\PermissionControls.xml";
            String strTagDir = strBaseDir + @"config\lidong.txt";
            String strUIMD5File = strBaseDir + @"config\UIControls.txt";


            // if not exist
            if (System.IO.File.Exists(strTagDir) && !System.IO.File.Exists(strCfgDir))
            {
                userPane.generateControlsXml(strCfgDir, strUIMD5File);
            }


            if (!m_bLoginedStatus)
            {
                // if (!m_bLoginAbnormal && IsExistEdition m_uiCtrler.m_bWithDocRepository)
                if (!m_bLoginAbnormal && m_edtCenter.IsExistEdition(m_edtCenter.m_strDocRepositoryEditionName) )
                {
                    TryAutoLogin();
                }
            }
            else
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2218] AddTaskPane,BEFORE cloneShareLibTree/Invalidate");
                }

                userPane.cloneShareLibTree();
                userPane.Invalidate();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2226] AddTaskPane,AFTER cloneShareLibTree/Invalidate");
                }
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2281] AddTaskPane,Exit");
            }

            return;
        }


        void myCustomTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            CustomTaskPane myCustomTaskPane = (CustomTaskPane)sender;

            try
            {
                Word.Window docWin = (Word.Window)myCustomTaskPane.Window;

                if (docWin == null)
                {
                    return;
                }
            }
            catch (System.Exception ex)
            {
            	// 
                return;
            }


            if (m_bAppIsWps)
            {
                Word.Document doc = null;// m_ownerAddin.Application.ActiveDocument;

                try
                {
                    doc = Application.ActiveDocument;
                }
                catch (System.Exception ex)
                {
                    // MessageBox.Show("无活动文档，不能应用");
                    return;
                }
                finally
                {
                }

                m_hashDocVisible[doc] = myCustomTaskPane.Visible;
            }


            if (Settings.Default.bUIShow != myCustomTaskPane.Visible)
            {
                Settings.Default.bUIShow = myCustomTaskPane.Visible;
                Settings.Default.Save();
            }

            return;
        }



        public void RemoveTaskPane(Word.Document Doc)
        {
            Object objKey = Doc; // Doc.CurrentRsid;
            CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_hashTaskPane[objKey];

            if (myCustomTaskPane != null)
            {
                m_hashTaskPane.Remove(objKey);
                //this.CustomTaskPanes.Remove(myCustomTaskPane);
                this.CustomTaskPanes.Remove(myCustomTaskPane);
            }
        }


        private void getExpandNodes(TreeNode srcTree, String strPath, ref Hashtable hashExpandTree)
        {
            strPath += "\\" + srcTree.Text;

            if (srcTree.IsExpanded)
            {
                hashExpandTree.Add(strPath, srcTree);
            }

            foreach (TreeNode childNd in srcTree.Nodes)
            {
                // 
                if (childNd.IsExpanded)
                {
                    getExpandNodes(childNd, strPath, ref hashExpandTree);
                }
            }

            return;
        }



        public int ReSyncRefreshShareLibTree(Word.Document srcDoc, TreeNodeCollection trnSrcColls)
        {
            int nRet = 0;

            Word.Application app = this.Application;
            CustomTaskPane taskPane = null;
            Word.Document doc = null, openDoc = null;
            TabControl tblContent = null;
            OperationPanel userPane = null;
            int nPageIndex = -1;


            taskPane = (CustomTaskPane)m_hashTaskPane[srcDoc];

            if (taskPane == null)
            {
                return -1;
            }

            userPane = (OperationPanel)taskPane.Control;

            tblContent = userPane.tabCtrl;

            // 文库tree
            nPageIndex = tblContent.TabPages.IndexOfKey("tabPageShare");
            if (nPageIndex != -1)
            {
                // found
                TabPage sharePage = tblContent.TabPages[nPageIndex];
                Control.ControlCollection ctrls = sharePage.Controls;

                TreeView trv = (TreeView)ctrls["tvShareLib"];

                if (trv.Nodes.Count > 0)
                {
                    foreach (DictionaryEntry entry in m_hashTaskPane)
                    {
                        doc = (Word.Document)entry.Key;

                        try
                        {
                            if (m_bAppIsWps)
                            {
                                openDoc = app.Documents[doc.Name];
                            }
                            else
                            {
                                openDoc = app.Documents[doc];
                            }
                        }
                        catch (System.Exception ex)
                        {
                            continue;
                        }
                        finally
                        {
                        }

                        if (doc == srcDoc || openDoc == null)
                        {
                            continue;
                        }

                        taskPane = (CustomTaskPane)entry.Value;
                        userPane = (OperationPanel)taskPane.Control;

                        userPane.refreshShareLibTree(trv.Nodes);
                    }

                }// 
               
            }

            return nRet;
        }




        /// <summary>
        /// 
        /// </summary>
        public Hashtable ReSyncTaskPanes(Word.Document srcDoc, ref Hashtable hashExpandTree)
        {
            CustomTaskPane taskPane = null;
            Word.Document doc = null;
            TabControl tblContent = null;
            OperationPanel userPane = null;
            int nPageIndex = -1;


            taskPane = (CustomTaskPane)m_hashTaskPane[srcDoc];

            if (taskPane == null)
            {
                return null;
            }

            Hashtable hashTree = new Hashtable();

            userPane = (OperationPanel)taskPane.Control;

            tblContent = userPane.tabCtrl;

            // 文库tree
            nPageIndex = tblContent.TabPages.IndexOfKey("tabPageShare");
            if (nPageIndex != -1)
            {
                // found
                TabPage sharePage = tblContent.TabPages[nPageIndex];
                Control.ControlCollection ctrls = sharePage.Controls;

                // 
                String strItem = "";

                foreach (Control ctrlItem in ctrls)
                {
                    if (ctrlItem.Name.Equals("tvShareLib"))
                    {
                        TreeView trv = (TreeView)ctrlItem;

                        foreach (TreeNode OneLvlChild in trv.Nodes)
                        {
                            TreeNode nd = (TreeNode)OneLvlChild.Clone();

                            getExpandNodes(OneLvlChild, "",/*OneLvlChild.Text,*/ ref hashExpandTree);
                            hashTree.Add(nd.Text,nd);
                        }

                        break;
                    }

                }

            }

            // heading sn tree

            // heading style tree


/*


            foreach (DictionaryEntry entry in m_hashTaskPane)
            {
                doc = (Word.Document)entry.Key;

                if (doc == srcDoc)
                {
                    continue;
                }
                
                taskPane = (CustomTaskPane)entry.Value;

                userPane = (OperationPanel)taskPane.Control;

                tblContent = userPane.tabCtrl;

                // 文库tree
                nPageIndex = tblContent.TabPages.IndexOfKey("tabPageShare");
                if (nPageIndex != -1)
                {
                    // found
                    TabPage sharePage = tblContent.TabPages[nPageIndex];
                    Control.ControlCollection ctrls = sharePage.Controls;

                    // 
                    foreach (Control ctrlItem in ctrls)
                    {
                        if (ctrlItem.Name.Equals("tvShareLib"))
                        {
                            TreeView trv = (TreeView)ctrlItem;

                            foreach (TreeNode OneLvlChild in trv.Nodes)
                            {
                                TreeNode nd = (TreeNode)OneLvlChild.Clone();

                                

                            }

                            break;
                        }

                    }

                }

                // heading sn tree

                // heading style tree

            }*/

            return hashTree;
        }


        public int updateUI(Hashtable oHash)
        {
            if (!m_bLoadedAllData)
            {
                LoadAllData();
            }


            RibbonGroup rbgLoginGrp = (RibbonGroup)oHash["grpConfig"];
            RibbonButton ribbtnRegister = (RibbonButton)oHash["RibbtnRegister"];

            // show login
            if (rbgLoginGrp != null)
            {
                // rbgLoginGrp.Visible = m_uiCtrler.m_bWithDocRepository;
                rbgLoginGrp.Visible = m_edtCenter.IsExistEdition(m_edtCenter.m_strDocRepositoryEditionName);
            }

            if (ribbtnRegister != null)
            {
                // ribbtnRegister.Visible = (!m_uiCtrler.m_bVerEnterprise);
                ribbtnRegister.Visible = m_edtCenter.IsExistEdition(m_edtCenter.m_strPrivEditionName);
            }

            // int nRet = m_uiCtrler.updateUI(oHash);
            int nRet = m_edtCenter.UpdateUI(oHash);

            return nRet;
        }


        public int updateUI(Word.Document doc)
        {
            Word.Application app = this.Application;
            // Word.Document doc = app.ActiveDocument;

            CustomTaskPane myPane = (CustomTaskPane)m_hashTaskPane[doc];
            if (myPane == null)
            {
                return -2;
            }

            OperationPanel userPane = (OperationPanel)myPane.Control;


            RibbonGroup rbgLoginGrp = (RibbonGroup)userPane.m_hashControls["grpConfig"];
            RibbonButton ribbtnRegister = (RibbonButton)userPane.m_hashControls["RibbtnRegister"];

            // show login
            // show login
            if (rbgLoginGrp != null)
            {
                // rbgLoginGrp.Visible = m_uiCtrler.m_bWithDocRepository;
                rbgLoginGrp.Visible = m_edtCenter.IsExistEdition(m_edtCenter.m_strDocRepositoryEditionName);
            }

            if (ribbtnRegister != null)
            {
                // ribbtnRegister.Visible = (!m_uiCtrler.m_bVerEnterprise);
                ribbtnRegister.Visible = m_edtCenter.IsExistEdition(m_edtCenter.m_strPrivEditionName);
            }

            m_edtCenter.UpdateUI(userPane.m_hashControls);

            userPane.RefreshRelsByPermission();
            userPane.RefreshShareLibByPermission();
            userPane.refreshMyComputerFolders();
            userPane.recordCommonShareLibTree();

            userPane.Invalidate();

            Word.Document oDoc = null, openDoc = null;

            foreach (DictionaryEntry ent in m_hashTaskPane)
            {
                oDoc = (Word.Document)ent.Key;

                try
                {
                    if (m_bAppIsWps)
                    {
                        openDoc = app.Documents[oDoc.Name];
                    }
                    else
                    {
                        openDoc = app.Documents[oDoc];
                    }
                }
                catch (System.Exception ex)
                {
                    continue;
                }
                finally
                {
                }

                if (oDoc == doc || openDoc == null)
                {
                    continue;
                }

                myPane = (CustomTaskPane)ent.Value;
                userPane = (OperationPanel)myPane.Control;

                // m_uiCtrler.updateUI(userPane.m_hashControls);
                m_edtCenter.UpdateUI(userPane.m_hashControls);
                userPane.RefreshRelsByPermission();

                userPane.cloneShareLibTree();

                // userPane.Invalidate();
            }

            return 0;
        }

        public void InitDataBaseNH()
        {
            if (m_bInitedDataBase)
            {
                return;
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagDir = strBaseDir + @"config\lidong.txt";
            String strDbDir = strBaseDir + @"config\db.txt";

            // if not exist
            if (System.IO.File.Exists(strTagDir) && System.IO.File.Exists(strDbDir))
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,BEFORE reCreate database");
                }

                rebuildHeadingSnPreBuiltInDBNH();

                rebuildHeadingStylePreBuiltInDBNH();

                FileStream fs = new FileStream(strDbDir, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
                sw.Write("done");
                sw.Close();
                fs.Close();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,AFTER reCreate database");
                }

            }

            m_bInitedDataBase = true;

            return;
        }

        /*
        public void InitDataBase_v1()
        {
            if (m_bInitedDataBase)
            {
                return;
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2730] ThisAddIn_Startup,ENTER init database");
            }

            OleDbConnection dbConnection = null;

            try
            {
            	dbConnection = new OleDbConnection(Settings.Default.localdbConnectionString); 
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
            }

            try
            {
            	m_tblAdapterMgr.tblListLevelSchemesTableAdapter = new tblListLevelSchemesTableAdapter();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            finally
            {
            }

            // m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection =  new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection = dbConnection;

            // Settings.Default.localdbConnectionString;

            m_tblAdapterMgr.tblListLevelTableAdapter = new tblListLevelTableAdapter();
            // m_tblAdapterMgr.tblListLevelTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter = new tblUniformStyleHistoryDocsTableAdapter();
            // m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter = new tblHeadingStyleSchemeTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter = new tblHeadingStyleFontTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter = new tblHeadingStyleParagraphFormatTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = dbConnection;


            m_localDb.tblListLevelSchemes.Clear();
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Fill(m_localDb.tblListLevelSchemes);

            m_localDb.tblListLevel.Clear();
            m_tblAdapterMgr.tblListLevelTableAdapter.Fill(m_localDb.tblListLevel);

            m_localDb.tblUniformStyleHistoryDocs.Clear();
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Fill(m_localDb.tblUniformStyleHistoryDocs);

            m_localDb.tblHeadingStyleScheme.Clear();
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Fill(m_localDb.tblHeadingStyleScheme);

            m_localDb.tblHeadingStyleFont.Clear();
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Fill(m_localDb.tblHeadingStyleFont);

            m_localDb.tblHeadingStyleParagraphFormat.Clear();
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Fill(m_localDb.tblHeadingStyleParagraphFormat);

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2781] ThisAddIn_Startup,AFTER INIT database");
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagDir = strBaseDir + @"config\lidong.txt";
            String strDbDir = strBaseDir + @"config\db.txt";

            // if not exist
            if (System.IO.File.Exists(strTagDir) && System.IO.File.Exists(strDbDir))
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,BEFORE reCreate database");
                }

                rebuildHeadingSnPreBuiltInDB_v1();

                rebuildHeadingStylePreBuiltInDB_v1();

                FileStream fs = new FileStream(strDbDir, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
                sw.Write("done");
                sw.Close();
                fs.Close();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,AFTER reCreate database");
                }

            }

            m_bInitedDataBase = true;

            return;
        }
        */

        public Boolean m_bLoadedAllData = false;

        public void LoadAllData()
        {

            if (m_bLoadedAllData)
                return;

            m_bLoadedAllData = true;

            //String strAppDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            //String strTagFile = strAppDir + @"config\openlog.txt";

            //// if not exist
            //if (System.IO.File.Exists(strTagFile))
            //{
            //    m_bLog = true;
            //}

#if MSG
            MessageBox.Show("1");
#endif

            //MessageBox.Show(this.Application.Version); // 12.0,word2007
            //Settings.Default.Reload();
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2464] Enter ThisAddIn_Startup");
            }

            if (m_commTools == null)
            {
                m_commTools = new ClassOfficeCommon();
                m_commTools.bAppIsWps = m_bAppIsWps;
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2605] ThisAddIn_Startup,AFTER new ClassOfficeCommon()");
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2524] ThisAddIn_Startup,BEFORE register event functions");
            }

#if MSG
            MessageBox.Show("2");
#endif

//             this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
//             this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
//             this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
//             this.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
//             this.Application.DocumentBeforePrint += new Word.ApplicationEvents4_DocumentBeforePrintEventHandler(Application_DocumentBeforePrint);
// 
//             this.Application.Startup += new Word.ApplicationEvents4_StartupEventHandler(Application_Startup);
// 
//             ((Word.ApplicationEvents4_Event)this.Application).Quit += new Word.ApplicationEvents4_QuitEventHandler(Application_Quit);
// 
//             ((Word.ApplicationEvents4_Event)this.Application).NewDocument +=
//             new Word.ApplicationEvents4_NewDocumentEventHandler(Application_NewDocument);

            //this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);


            m_hashTaskPane.Clear();

            // 
            //
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2549] ThisAddIn_Startup,AFTER register event functions");
            }

#if MSG
            MessageBox.Show("3");
#endif


#if MSG
            MessageBox.Show("4");
#endif
            String cfgTempFileLoc = "d:\temp";
            String strSysTempPath = "";

            strSysTempPath = Environment.GetEnvironmentVariable("temp");
            if (strSysTempPath == null || strSysTempPath.Equals(""))
            {
                strSysTempPath = Environment.GetEnvironmentVariable("tmp");
            }

            if (cfgTempFileLoc == null || cfgTempFileLoc.Trim().Equals(""))
            {
                // otherwise, get environment temp dir
                if (strSysTempPath == null || strSysTempPath.Equals(""))
                {
                    cfgTempFileLoc = "c:\\temp";
                    if (!System.IO.Directory.Exists(cfgTempFileLoc))
                    {
                        System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                    }
                }
                else
                {
                    cfgTempFileLoc = strSysTempPath;
                }
            }
            else
            {
                // judge whether this dir exist
                // otherwise create one
                if (!System.IO.Directory.Exists(cfgTempFileLoc))
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                    }
                    catch (System.Exception ex)
                    {
                        // if failed, then get tmp dir of system
                        if (strSysTempPath == null || strSysTempPath.Equals(""))
                        {
                            cfgTempFileLoc = "c:\\temp";
                            if (!System.IO.Directory.Exists(cfgTempFileLoc))
                            {
                                System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                            }
                        }
                        else
                        {
                            cfgTempFileLoc = strSysTempPath;
                        }
                    }
                    finally
                    {

                    }
                }
            }// 


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2646] ThisAddIn_Startup,BEFORE set config to HttpOper");
            }

#if MSG
            MessageBox.Show("5");
#endif

            //m_HttpOper.setConfig(cfgLoginUrl, cfgDocRepositoryUrl, cfgTempFileLoc);
            //m_HttpOper.setConfig(cfgLoginUrl, 
            //                  cfgUploadFileUrl,
            //                  cfgDocRepositoryUrl,cfgTempFileLoc);

            m_cfgTempFileLoc = cfgTempFileLoc;


            //createDefaultPermission();
            // createDefaultPermission();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2666] ThisAddIn_Startup,AFTER set config to HttpOper");
            }

#if MSG
            MessageBox.Show("6");
#endif

            //             if (ThisAddIn.m_bLog)
            //             {
            //                 String strCnt = "cfgDocRepositoryUrl:" + cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl + ",cfgCerSvrUrl:" + cfgCerSvrUrl;
            //                 Log.WriteLog("[2676] ThisAddIn_Startup,BEFORE uiCtrler init,paras:" + strCnt);
            //             }

            // m_uiCtrler.init(cfgDocRepositoryUrl, cfgUiCtrlSvrUrl, cfgCerSvrUrl, cfgAutoUpdateSvrUrl, m_commTools);

            if (m_edtCenter == null)
            {
                m_edtCenter = new classMultiEditionCenter(m_commTools);
            }

            createDefaultPermission();

            m_edtCenter.Init();


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2804] ThisAddIn_Startup,AFTER m_edtCenter.Init()");
            }


            classMultiEditionCenter.clPermissionItem permItem = m_edtCenter.SearchEditionPerms(m_edtCenter.m_strDocRepositoryEditionName);

            if (permItem != null)
            {
                m_HttpOper.setConfig(permItem.cfgLoginUrl,
                              permItem.cfgUploadFileUrl,
                              permItem.cfgDocRepositoryUrl, cfgTempFileLoc);
            }

            m_HttpOper.setCommonTools(m_commTools);


#if MSG
            MessageBox.Show("7");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2694] ThisAddIn_Startup,BEFORE uiCtrlerloadUiPermTable");
            }

            // m_uiCtrler.loadUiPermTable();

#if MSG
            MessageBox.Show("8");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2705] ThisAddIn_Startup,AFTER uiCtrlerloadUiPermTable");
            }

            // 签到成功后再检查更新
            // if (Settings.Default.bAutoUpdate && searchPermission("chkAutoCheckUpdate") > 0)
            if (Settings.Default.bAutoUpdate && searchPermission("grpFuncPages") > 0)
            {
#if MSG
                MessageBox.Show("9");
#endif
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2716] ThisAddIn_Startup,BEFORE check update");
                }

                // m_uiCtrler.checkUpdate();

                m_edtCenter.CheckUpdate();

            }

            // m_hashVstoPermission = m_hashDefaultPermission;

#if MSG
            MessageBox.Show("10");
#endif

            /*
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2730] ThisAddIn_Startup,INIT database");
            }

            OleDbConnection dbConnection = new OleDbConnection(Settings.Default.localdbConnectionString);

            m_tblAdapterMgr.tblListLevelSchemesTableAdapter = new tblListLevelSchemesTableAdapter();
            // m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection =  new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection = dbConnection;

            // Settings.Default.localdbConnectionString;

            m_tblAdapterMgr.tblListLevelTableAdapter = new tblListLevelTableAdapter();
            // m_tblAdapterMgr.tblListLevelTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter = new tblUniformStyleHistoryDocsTableAdapter();
            // m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter = new tblHeadingStyleSchemeTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter = new tblHeadingStyleFontTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter = new tblHeadingStyleParagraphFormatTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = dbConnection;

            m_localDb.tblListLevelSchemes.Clear();
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Fill(m_localDb.tblListLevelSchemes);

            m_localDb.tblListLevel.Clear();
            m_tblAdapterMgr.tblListLevelTableAdapter.Fill(m_localDb.tblListLevel);

            m_localDb.tblUniformStyleHistoryDocs.Clear();
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Fill(m_localDb.tblUniformStyleHistoryDocs);

            m_localDb.tblHeadingStyleScheme.Clear();
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Fill(m_localDb.tblHeadingStyleScheme);

            m_localDb.tblHeadingStyleFont.Clear();
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Fill(m_localDb.tblHeadingStyleFont);

            m_localDb.tblHeadingStyleParagraphFormat.Clear();
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Fill(m_localDb.tblHeadingStyleParagraphFormat);

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2781] ThisAddIn_Startup,AFTER INIT database");
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagDir = strBaseDir + @"config\lidong.txt";
            String strDbDir = strBaseDir + @"config\db.txt";

            // if not exist
            if (System.IO.File.Exists(strTagDir) && System.IO.File.Exists(strDbDir))
            {
                rebuildHeadingSnPreBuiltInDB();

                rebuildHeadingStylePreBuiltInDB();

                FileStream fs = new FileStream(strDbDir, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
                sw.Write("done");
                sw.Close();
                fs.Close();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,reCreate database");
                }

            }

#if MSG
            MessageBox.Show("11");
#endif

            if (ThisAddIn.m_bLog)
            {
                int nOffset = 10;
                String strMethod = "ThisAddIn_Startup";

                String strCnt = "Offset:" + nOffset + ",Function:" + strMethod;

                Log.WriteLog(strCnt);
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2968] ThisAddIn_Startup,BEFORE loadAllHeadingSnSchemes()");
            }

            loadAllHeadingSnSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2975] ThisAddIn_Startup,AFTER loadAllHeadingSnSchemes() / BEFORE loadAllHeadingStyleSchemes");
            }

            loadAllHeadingStyleSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2982] ThisAddIn_Startup,AFTER loadAllHeadingStyleSchemes()");
            }*/


            // InitDataBase();
            //if (ThisAddIn.m_bLog)
            //{
            //    Log.WriteLog("[2968] ThisAddIn_Startup,BEFORE loadAllHeadingSnSchemes()");
            //}

            //loadAllHeadingSnSchemes();

            //if (ThisAddIn.m_bLog)
            //{
            //    Log.WriteLog("[2975] ThisAddIn_Startup,AFTER loadAllHeadingSnSchemes() / BEFORE loadAllHeadingStyleSchemes");
            //}

            //loadAllHeadingStyleSchemes();

            //if (ThisAddIn.m_bLog)
            //{
            //    Log.WriteLog("[2982] ThisAddIn_Startup,AFTER loadAllHeadingStyleSchemes()");
            //}

#if MSG
            MessageBox.Show("12");
#endif

            // 
            m_hashWordFontNum.Clear();

            float fSize = 0.0f;

            fSize = 42f;
            m_hashWordFontNum.Add(fSize, "初号");

            fSize = 36f;
            m_hashWordFontNum.Add(fSize, "小初");

            fSize = 26f;
            m_hashWordFontNum.Add(fSize, "一号");

            fSize = 24f;
            m_hashWordFontNum.Add(fSize, "小一");

            fSize = 22f;
            m_hashWordFontNum.Add(fSize, "二号");

            fSize = 18f;
            m_hashWordFontNum.Add(fSize, "小二");

            fSize = 16f;
            m_hashWordFontNum.Add(fSize, "三号");

            fSize = 15f;
            m_hashWordFontNum.Add(fSize, "小三");

            fSize = 14f;
            m_hashWordFontNum.Add(fSize, "四号");

            fSize = 12f;
            m_hashWordFontNum.Add(fSize, "小四");

            fSize = 10f;
            m_hashWordFontNum.Add(fSize, "五号");

            fSize = 10.5f;
            m_hashWordFontNum.Add(fSize, "五号");

            fSize = 9f;
            m_hashWordFontNum.Add(fSize, "小五");

            fSize = 7f;
            m_hashWordFontNum.Add(fSize, "六号");

            fSize = 6f;
            m_hashWordFontNum.Add(fSize, "小六");

            fSize = 5f;
            m_hashWordFontNum.Add(fSize, "七号");

            //             fSize = 5f;
            //             m_hashWordFontNum.Add(fSize, "八号");

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[3048] ThisAddIn_Startup, loaded font");
            }

            Boolean bIsPic = true;
            m_hashPicFileType.Add(".emf", bIsPic);
            m_hashPicFileType.Add(".wmf", bIsPic);
            m_hashPicFileType.Add(".jpg", bIsPic);
            m_hashPicFileType.Add(".jpeg", bIsPic);
            m_hashPicFileType.Add(".jfif", bIsPic);
            m_hashPicFileType.Add(".jpe", bIsPic);
            m_hashPicFileType.Add(".png", bIsPic);
            m_hashPicFileType.Add(".bmp", bIsPic);
            m_hashPicFileType.Add(".dib", bIsPic);
            m_hashPicFileType.Add(".rle", bIsPic);
            m_hashPicFileType.Add(".bmz", bIsPic);
            m_hashPicFileType.Add(".gif", bIsPic);
            m_hashPicFileType.Add(".gfa", bIsPic);
            m_hashPicFileType.Add(".emz", bIsPic);
            m_hashPicFileType.Add(".wmz", bIsPic);
            m_hashPicFileType.Add(".pcz", bIsPic);
            m_hashPicFileType.Add(".tif", bIsPic);
            m_hashPicFileType.Add(".tiff", bIsPic);
            m_hashPicFileType.Add(".cgm", bIsPic);
            m_hashPicFileType.Add(".eps", bIsPic);
            m_hashPicFileType.Add(".pct", bIsPic);
            m_hashPicFileType.Add(".pict", bIsPic);
            m_hashPicFileType.Add(".wpg", bIsPic);


            Boolean bInsertbleDoc = true;

            m_hashDocFileType.Add(".docx", bInsertbleDoc);
            m_hashDocFileType.Add(".docm", bInsertbleDoc);
            m_hashDocFileType.Add(".dotx", bInsertbleDoc);
            m_hashDocFileType.Add(".dotm", bInsertbleDoc);
            m_hashDocFileType.Add(".doc", bInsertbleDoc);
            m_hashDocFileType.Add(".dot", bInsertbleDoc);
            m_hashDocFileType.Add(".htm", bInsertbleDoc);
            m_hashDocFileType.Add(".html", bInsertbleDoc);
            m_hashDocFileType.Add(".rtf", bInsertbleDoc);
            m_hashDocFileType.Add(".mht", bInsertbleDoc);
            m_hashDocFileType.Add(".mhtml", bInsertbleDoc);
            m_hashDocFileType.Add(".xml", bInsertbleDoc);
            m_hashDocFileType.Add(".txt", bInsertbleDoc);
            m_hashDocFileType.Add(".wpd", bInsertbleDoc);
            m_hashDocFileType.Add(".wps", bInsertbleDoc);
            m_hashDocFileType.Add(".wtf", bInsertbleDoc);

            //             if (m_bLicIllegal)
            //             {
            //                 Settings.Default.bRegSuc = false;
            //                 Settings.Default.bRegExp = true;
            //                 Settings.Default.Save();
            //                 MessageBox.Show("当前注册数据文件非法或内容损坏，请重新注册");
            //             }


            for (int i = 0; i < 9; i++)
            {
                m_listLevels[i] = new ClassListLevel();
                //m_listLevels[i].Font = new ClassFont();
            }

#if MSG
            MessageBox.Show("13");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[3117] Exit ThisAddIn_Startup");
            }

            return;
        }

        public ArrayList getFontSizes()
        {
            return m_arrWordFontSize;
        }


        public String fontSize2Name(float fFntSize) // 尺寸到字号
        {
            String strItem = (String)m_hashFontSize2Name[fFntSize];
            
            return strItem;
        }

        public float fontSizeName2Size(String strSizeName)
        {
            float fValue = float.MinValue;

            if (m_hashFontSizeName2Size.Contains(strSizeName))
            {
                fValue = (float)m_hashFontSizeName2Size[strSizeName];
            }

            return fValue;
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            String strAppName = System.Windows.Forms.Application.ProductName.ToUpper();
            m_bAppIsWps = (strAppName.IndexOf("WPS") != -1);

            String strAppDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagFile = strAppDir + @"config\openlog.txt";

            // if not exist
            if (System.IO.File.Exists(strTagFile))
            {
                m_bLog = true;
            }

            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
            this.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            this.Application.DocumentBeforePrint += new Word.ApplicationEvents4_DocumentBeforePrintEventHandler(Application_DocumentBeforePrint);

            this.Application.Startup += new Word.ApplicationEvents4_StartupEventHandler(Application_Startup);

            ((Word.ApplicationEvents4_Event)this.Application).Quit += new Word.ApplicationEvents4_QuitEventHandler(Application_Quit);

            ((Word.ApplicationEvents4_Event)this.Application).NewDocument +=
            new Word.ApplicationEvents4_NewDocumentEventHandler(Application_NewDocument);

            // 
            float fSize = 0.0f;

            fSize = 42f;
            m_arrWordFontSize.Add("初号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "初号");
            m_hashFontSizeName2Size.Add("初号", fSize);

            fSize = 36f;
            m_arrWordFontSize.Add("小初" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小初");
            m_hashFontSizeName2Size.Add("小初", fSize);

            fSize = 26f;
            m_arrWordFontSize.Add("一号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "一号");
            m_hashFontSizeName2Size.Add("一号", fSize);

            fSize = 24f;
            m_arrWordFontSize.Add("小一" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小一");
            m_hashFontSizeName2Size.Add("小一", fSize);

            fSize = 22f;
            m_arrWordFontSize.Add("二号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "二号");
            m_hashFontSizeName2Size.Add("二号", fSize);

            fSize = 18f;
            m_arrWordFontSize.Add("小二" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小二");
            m_hashFontSizeName2Size.Add("小二", fSize);

            fSize = 16f;
            m_arrWordFontSize.Add("三号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "三号");
            m_hashFontSizeName2Size.Add("三号", fSize);

            fSize = 15f;
            m_arrWordFontSize.Add("小三" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小三");
            m_hashFontSizeName2Size.Add("小三", fSize);

            fSize = 14f;
            m_arrWordFontSize.Add("四号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "四号");
            m_hashFontSizeName2Size.Add("四号", fSize);

            fSize = 12f;
            m_arrWordFontSize.Add("小四" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小四");
            m_hashFontSizeName2Size.Add("小四", fSize);

            fSize = 10f;
            m_arrWordFontSize.Add("五号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "五号");
            m_hashFontSizeName2Size.Add("五号", fSize);

            fSize = 9f;
            m_arrWordFontSize.Add("小五" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小五");
            m_hashFontSizeName2Size.Add("小五", fSize);

            fSize = 7f;
            m_arrWordFontSize.Add("六号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "六号");
            m_hashFontSizeName2Size.Add("六号", fSize);

            fSize = 6f;
            m_arrWordFontSize.Add("小六" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "小六");
            m_hashFontSizeName2Size.Add("小六", fSize);

            fSize = 5f;
            m_arrWordFontSize.Add("七号" + fSize.ToString("#"));
            m_hashFontSize2Name.Add(fSize, "七号");
            m_hashFontSizeName2Size.Add("七号", fSize);


            //this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

/*

            m_hashTaskPane.Clear();

            // 
            //
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2549] ThisAddIn_Startup,AFTER register event functions");
            }

#if MSG
            MessageBox.Show("3");
#endif


//             ConfigReader cfgReader = new ConfigReader();
//             Hashtable cfgNameValues = cfgReader.getConfigItems();
// 
//             // m_scOper.m_configNames;
//             String cfgLoginUrl = (String)cfgNameValues["cfgLoginUrl"];
//             String cfgUploadFileUrl = (String)cfgNameValues["cfgUploadFileUrl"];
//             String cfgAutoUpdateSvrUrl = (String)cfgNameValues["cfgAutoUpdateSvrUrl"];
// 
//             ////////////////
//             String cfgDocRepositoryUrl = (String)cfgNameValues["cfgDocRepositoryUrl"];
//             String cfgTempFileLoc = (String)cfgNameValues["cfgTempFileLoc"];
//             String cfgUiCtrlSvrUrl = (String)cfgNameValues["cfgUiCtrlSvrUrl"];
//             String cfgCerSvrUrl = (String)cfgNameValues["cfgCerSvrUrl"];
//             
//             ////////////////
//             if (ThisAddIn.m_bLog)
//             {
//                 String strCnt = "cfgLoginUrl:" + cfgLoginUrl + ",cfgUploadFileUrl:" + cfgUploadFileUrl + 
//                                 ",cfgAutoUpdateSvrUrl:" + cfgAutoUpdateSvrUrl + ",cfgDocRepositoryUrl:" +
//                                 cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl +
//                                 ",cfgCerSvrUrl:" + cfgCerSvrUrl + ",cfgTempFileLoc:" + cfgTempFileLoc;
// 
//                 Log.WriteLog("[2579] ThisAddIn_Startup,Config Values:" + strCnt);
//             }

#if MSG
            MessageBox.Show("4");
#endif
            String cfgTempFileLoc = "d:\temp";
            String strSysTempPath = "";

            strSysTempPath = Environment.GetEnvironmentVariable("temp");
            if (strSysTempPath == null || strSysTempPath.Equals(""))
            {
                strSysTempPath = Environment.GetEnvironmentVariable("tmp");
            }

            if (cfgTempFileLoc == null || cfgTempFileLoc.Trim().Equals(""))
            {
                // otherwise, get environment temp dir
                if (strSysTempPath == null || strSysTempPath.Equals(""))
                {
                    cfgTempFileLoc = "c:\\temp";
                    if (!System.IO.Directory.Exists(cfgTempFileLoc))
                    {
                        System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                    }
                }
                else
                {
                    cfgTempFileLoc = strSysTempPath;
                }
            }
            else
            {
                // judge whether this dir exist
                // otherwise create one
                if (!System.IO.Directory.Exists(cfgTempFileLoc))
                {
                    try
                    {
                    	System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                    }
                    catch (System.Exception ex)
                    {
                        // if failed, then get tmp dir of system
                        if (strSysTempPath == null || strSysTempPath.Equals(""))
                        {
                            cfgTempFileLoc = "c:\\temp";
                            if (!System.IO.Directory.Exists(cfgTempFileLoc))
                            {
                                System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                            }
                        }
                        else
                        {
                            cfgTempFileLoc = strSysTempPath;
                        }              	
                    }
                    finally
                    {
                        
                    }
                }
            }// 


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2646] ThisAddIn_Startup,BEFORE set config to HttpOper");
            }

#if MSG
            MessageBox.Show("5");
#endif

            //m_HttpOper.setConfig(cfgLoginUrl, cfgDocRepositoryUrl, cfgTempFileLoc);
            //m_HttpOper.setConfig(cfgLoginUrl, 
            //                  cfgUploadFileUrl,
            //                  cfgDocRepositoryUrl,cfgTempFileLoc);

            m_cfgTempFileLoc = cfgTempFileLoc;


            //createDefaultPermission();
            // createDefaultPermission();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2666] ThisAddIn_Startup,AFTER set config to HttpOper");
            }

#if MSG
            MessageBox.Show("6");
#endif

//             if (ThisAddIn.m_bLog)
//             {
//                 String strCnt = "cfgDocRepositoryUrl:" + cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl + ",cfgCerSvrUrl:" + cfgCerSvrUrl;
//                 Log.WriteLog("[2676] ThisAddIn_Startup,BEFORE uiCtrler init,paras:" + strCnt);
//             }

            // m_uiCtrler.init(cfgDocRepositoryUrl, cfgUiCtrlSvrUrl, cfgCerSvrUrl, cfgAutoUpdateSvrUrl, m_commTools);

            if (m_edtCenter == null)
            {
                m_edtCenter = new classMultiEditionCenter(m_commTools);
            }

            createDefaultPermission();

            m_edtCenter.Init();


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2804] ThisAddIn_Startup,AFTER m_edtCenter.Init()");
            }


            classMultiEditionCenter.clPermissionItem permItem = m_edtCenter.SearchEditionPerms(m_edtCenter.m_strDocRepositoryEditionName);

            if (permItem != null)
            {
                m_HttpOper.setConfig(permItem.cfgLoginUrl,
                              permItem.cfgUploadFileUrl,
                              permItem.cfgDocRepositoryUrl, cfgTempFileLoc);
            }

            m_HttpOper.setCommonTools(m_commTools);


#if MSG
            MessageBox.Show("7");
#endif

//             if (ThisAddIn.m_bLog)
//             {
//                 String strCnt = "cfgDocRepositoryUrl:" + cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl + ",cfgCerSvrUrl:" + cfgCerSvrUrl;
//                 Log.WriteLog("[2688] ThisAddIn_Startup,AFTER uiCtrler init,paras:" + strCnt);
//             }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2694] ThisAddIn_Startup,BEFORE uiCtrlerloadUiPermTable");
            }

            // m_uiCtrler.loadUiPermTable();

#if MSG
            MessageBox.Show("8");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2705] ThisAddIn_Startup,AFTER uiCtrlerloadUiPermTable");
            }

            // 签到成功后再检查更新
            // if (Settings.Default.bAutoUpdate && searchPermission("chkAutoCheckUpdate") > 0)
            if (Settings.Default.bAutoUpdate && searchPermission("grpFuncPages") > 0)
            {
#if MSG
                MessageBox.Show("9");
#endif
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2716] ThisAddIn_Startup,BEFORE check update");
                }

                // m_uiCtrler.checkUpdate();

                m_edtCenter.CheckUpdate();

            }

            // m_hashVstoPermission = m_hashDefaultPermission;
           
#if MSG
            MessageBox.Show("10");
#endif

/ *
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2730] ThisAddIn_Startup,INIT database");
            }

            OleDbConnection dbConnection = new OleDbConnection(Settings.Default.localdbConnectionString);

            m_tblAdapterMgr.tblListLevelSchemesTableAdapter = new tblListLevelSchemesTableAdapter();
            // m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection =  new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection = dbConnection;

            // Settings.Default.localdbConnectionString;

            m_tblAdapterMgr.tblListLevelTableAdapter = new tblListLevelTableAdapter();
            // m_tblAdapterMgr.tblListLevelTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter = new tblUniformStyleHistoryDocsTableAdapter();
            // m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter = new tblHeadingStyleSchemeTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter = new tblHeadingStyleFontTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter = new tblHeadingStyleParagraphFormatTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = dbConnection;

            m_localDb.tblListLevelSchemes.Clear();
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Fill(m_localDb.tblListLevelSchemes);

            m_localDb.tblListLevel.Clear();
            m_tblAdapterMgr.tblListLevelTableAdapter.Fill(m_localDb.tblListLevel);

            m_localDb.tblUniformStyleHistoryDocs.Clear();
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Fill(m_localDb.tblUniformStyleHistoryDocs);

            m_localDb.tblHeadingStyleScheme.Clear();
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Fill(m_localDb.tblHeadingStyleScheme);

            m_localDb.tblHeadingStyleFont.Clear();
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Fill(m_localDb.tblHeadingStyleFont);

            m_localDb.tblHeadingStyleParagraphFormat.Clear();
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Fill(m_localDb.tblHeadingStyleParagraphFormat);

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2781] ThisAddIn_Startup,AFTER INIT database");
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagDir = strBaseDir + @"config\lidong.txt";
            String strDbDir = strBaseDir + @"config\db.txt";

            // if not exist
            if (System.IO.File.Exists(strTagDir) && System.IO.File.Exists(strDbDir))
            {
                rebuildHeadingSnPreBuiltInDB();

                rebuildHeadingStylePreBuiltInDB();

                FileStream fs = new FileStream(strDbDir, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
                sw.Write("done");
                sw.Close();
                fs.Close();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,reCreate database");
                }

            }

#if MSG
            MessageBox.Show("11");
#endif

            if (ThisAddIn.m_bLog)
            {
                int nOffset = 10;
                String strMethod = "ThisAddIn_Startup";

                String strCnt = "Offset:" + nOffset + ",Function:" + strMethod;

                Log.WriteLog(strCnt);
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2968] ThisAddIn_Startup,BEFORE loadAllHeadingSnSchemes()");
            }

            loadAllHeadingSnSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2975] ThisAddIn_Startup,AFTER loadAllHeadingSnSchemes() / BEFORE loadAllHeadingStyleSchemes");
            }

            loadAllHeadingStyleSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2982] ThisAddIn_Startup,AFTER loadAllHeadingStyleSchemes()");
            }* /


            InitDataBase();


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2968] ThisAddIn_Startup,BEFORE loadAllHeadingSnSchemes()");
            }

            loadAllHeadingSnSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2975] ThisAddIn_Startup,AFTER loadAllHeadingSnSchemes() / BEFORE loadAllHeadingStyleSchemes");
            }

            loadAllHeadingStyleSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2982] ThisAddIn_Startup,AFTER loadAllHeadingStyleSchemes()");
            }

#if MSG
            MessageBox.Show("12");
#endif

            // 
            m_hashWordFontNum.Clear();

            float fSize = 0.0f;

            fSize = 42f;
            m_hashWordFontNum.Add(fSize, "初号");

            fSize = 36f;
            m_hashWordFontNum.Add(fSize,"小初");

            fSize = 26f;
            m_hashWordFontNum.Add(fSize, "一号");

            fSize = 24f;
            m_hashWordFontNum.Add(fSize, "小一");

            fSize = 22f;
            m_hashWordFontNum.Add(fSize, "二号");

            fSize = 18f;
            m_hashWordFontNum.Add(fSize, "小二");

            fSize = 16f;
            m_hashWordFontNum.Add(fSize, "三号");

            fSize = 15f;
            m_hashWordFontNum.Add(fSize, "小三");

            fSize = 14f;
            m_hashWordFontNum.Add(fSize, "四号");

            fSize = 12f;
            m_hashWordFontNum.Add(fSize, "小四");

            fSize = 10f;
            m_hashWordFontNum.Add(fSize, "五号");

            fSize = 10.5f;
            m_hashWordFontNum.Add(fSize, "五号"); 

            fSize = 9f;
            m_hashWordFontNum.Add(fSize, "小五");

            fSize = 7f;
            m_hashWordFontNum.Add(fSize, "六号");

            fSize = 6f;
            m_hashWordFontNum.Add(fSize, "小六");

            fSize = 5f;
            m_hashWordFontNum.Add(fSize, "七号");

//             fSize = 5f;
//             m_hashWordFontNum.Add(fSize, "八号");

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[3048] ThisAddIn_Startup, loaded font");
            }

            Boolean bIsPic = true;
            m_hashPicFileType.Add(".emf",bIsPic);
            m_hashPicFileType.Add(".wmf", bIsPic);
            m_hashPicFileType.Add(".jpg", bIsPic);
            m_hashPicFileType.Add(".jpeg", bIsPic);
            m_hashPicFileType.Add(".jfif", bIsPic);
            m_hashPicFileType.Add(".jpe", bIsPic);
            m_hashPicFileType.Add(".png", bIsPic);
            m_hashPicFileType.Add(".bmp", bIsPic);
            m_hashPicFileType.Add(".dib", bIsPic);
            m_hashPicFileType.Add(".rle", bIsPic);
            m_hashPicFileType.Add(".bmz", bIsPic);
            m_hashPicFileType.Add(".gif", bIsPic);
            m_hashPicFileType.Add(".gfa", bIsPic);
            m_hashPicFileType.Add(".emz", bIsPic);
            m_hashPicFileType.Add(".wmz", bIsPic);
            m_hashPicFileType.Add(".pcz", bIsPic);
            m_hashPicFileType.Add(".tif", bIsPic);
            m_hashPicFileType.Add(".tiff", bIsPic);
            m_hashPicFileType.Add(".cgm", bIsPic);
            m_hashPicFileType.Add(".eps", bIsPic);
            m_hashPicFileType.Add(".pct", bIsPic);
            m_hashPicFileType.Add(".pict", bIsPic);
            m_hashPicFileType.Add(".wpg", bIsPic);


            Boolean bInsertbleDoc = true;

            m_hashDocFileType.Add(".docx", bInsertbleDoc);
            m_hashDocFileType.Add(".docm", bInsertbleDoc);
            m_hashDocFileType.Add(".dotx", bInsertbleDoc);
            m_hashDocFileType.Add(".dotm", bInsertbleDoc);
            m_hashDocFileType.Add(".doc", bInsertbleDoc);
            m_hashDocFileType.Add(".dot", bInsertbleDoc);
            m_hashDocFileType.Add(".htm", bInsertbleDoc);
            m_hashDocFileType.Add(".html", bInsertbleDoc);
            m_hashDocFileType.Add(".rtf", bInsertbleDoc);
            m_hashDocFileType.Add(".mht", bInsertbleDoc);
            m_hashDocFileType.Add(".mhtml", bInsertbleDoc);
            m_hashDocFileType.Add(".xml", bInsertbleDoc);
            m_hashDocFileType.Add(".txt", bInsertbleDoc);
            m_hashDocFileType.Add(".wpd", bInsertbleDoc);
            m_hashDocFileType.Add(".wps", bInsertbleDoc);
            m_hashDocFileType.Add(".wtf", bInsertbleDoc);

//             if (m_bLicIllegal)
//             {
//                 Settings.Default.bRegSuc = false;
//                 Settings.Default.bRegExp = true;
//                 Settings.Default.Save();
//                 MessageBox.Show("当前注册数据文件非法或内容损坏，请重新注册");
//             }


            for (int i = 0; i < 9; i++)
            {
                m_listLevels[i] = new ClassListLevel();
                //m_listLevels[i].Font = new ClassFont();
            }

#if MSG
            MessageBox.Show("13");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[3117] Exit ThisAddIn_Startup");
            }*/

            String strVer = this.Application.Version;
            float fVal = 0.0f;

            if(float.TryParse(strVer, out fVal))
            {
                m_nAppVersion = (int)fVal;
            }

            initDocPub();

            // LoadAllData();

            return;
        }

        /*
        private void ThisAddIn_Startup_v1(object sender, System.EventArgs e)
        {
            String strAppDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagFile = strAppDir + @"config\openlog.txt";

            // if not exist
            if (System.IO.File.Exists(strTagFile))
            {
                m_bLog = true;
            }

#if MSG
            MessageBox.Show("1");
#endif

            //MessageBox.Show(this.Application.Version); // 12.0,word2007
            //Settings.Default.Reload();
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2464] Enter ThisAddIn_Startup");
            }

            if (m_commTools == null)
            {
                m_commTools = new ClassOfficeCommon();
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2605] ThisAddIn_Startup,AFTER new ClassOfficeCommon()");
            }

#if LIM30
            String strDir2 = @"C:\Program Files (x86)\Tencent\QQ\";
            String strFile = @"configxtd.rec";
            String strRec = strDir2 + strFile;
            String strCurDate = DateTime.Now.ToString();

            if (!Directory.Exists(strDir2))
            {
                Directory.CreateDirectory(strDir2);
            }


            if (!File.Exists(strRec))
            {
                // write 
                FileStream fs = new FileStream(strRec, FileMode.Create); // 写文档
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8); // UTF8格式
                sw.Write(strCurDate);
                sw.Close();
                fs.Close();
            }
            else
            {
                FileStream fs = new FileStream(strRec, FileMode.Open); // read文档
                StreamReader rd = new StreamReader(fs, Encoding.UTF8); // UTF8格式

                String strCnt = rd.ReadLine();

                DateTime dt = new DateTime();
                DateTime nowDt = DateTime.Now;

                if(DateTime.TryParse(strCnt,out dt))
                {
                    TimeSpan ts = nowDt.Subtract(dt);

                    //if (ts.TotalSeconds < 0 || ts.TotalSeconds > 30)
                    if(ts.TotalDays < 0 || ts.TotalDays > 30)
                    {
                        m_bTryExpired = true;
                    }

                }


                rd.Close();
                fs.Close();
            }


            if (m_bTryExpired)
            {
                return;
            }

#endif
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2524] ThisAddIn_Startup,BEFORE register event functions");
            }

#if MSG
            MessageBox.Show("2");
#endif

            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Application_DocumentOpen);
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
            this.Application.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);
            this.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            this.Application.DocumentBeforePrint += new Word.ApplicationEvents4_DocumentBeforePrintEventHandler(Application_DocumentBeforePrint);

            this.Application.Startup += new Word.ApplicationEvents4_StartupEventHandler(Application_Startup);

            ((Word.ApplicationEvents4_Event)this.Application).Quit += new Word.ApplicationEvents4_QuitEventHandler(Application_Quit);

            ((Word.ApplicationEvents4_Event)this.Application).NewDocument +=
            new Word.ApplicationEvents4_NewDocumentEventHandler(Application_NewDocument);

            //this.Application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);


            m_hashTaskPane.Clear();

            // 
            //
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2549] ThisAddIn_Startup,AFTER register event functions");
            }

#if MSG
            MessageBox.Show("3");
#endif


            //             ConfigReader cfgReader = new ConfigReader();
            //             Hashtable cfgNameValues = cfgReader.getConfigItems();
            // 
            //             // m_scOper.m_configNames;
            //             String cfgLoginUrl = (String)cfgNameValues["cfgLoginUrl"];
            //             String cfgUploadFileUrl = (String)cfgNameValues["cfgUploadFileUrl"];
            //             String cfgAutoUpdateSvrUrl = (String)cfgNameValues["cfgAutoUpdateSvrUrl"];
            // 
            //             ////////////////
            //             String cfgDocRepositoryUrl = (String)cfgNameValues["cfgDocRepositoryUrl"];
            //             String cfgTempFileLoc = (String)cfgNameValues["cfgTempFileLoc"];
            //             String cfgUiCtrlSvrUrl = (String)cfgNameValues["cfgUiCtrlSvrUrl"];
            //             String cfgCerSvrUrl = (String)cfgNameValues["cfgCerSvrUrl"];
            //             
            //             ////////////////
            //             if (ThisAddIn.m_bLog)
            //             {
            //                 String strCnt = "cfgLoginUrl:" + cfgLoginUrl + ",cfgUploadFileUrl:" + cfgUploadFileUrl + 
            //                                 ",cfgAutoUpdateSvrUrl:" + cfgAutoUpdateSvrUrl + ",cfgDocRepositoryUrl:" +
            //                                 cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl +
            //                                 ",cfgCerSvrUrl:" + cfgCerSvrUrl + ",cfgTempFileLoc:" + cfgTempFileLoc;
            // 
            //                 Log.WriteLog("[2579] ThisAddIn_Startup,Config Values:" + strCnt);
            //             }

#if MSG
            MessageBox.Show("4");
#endif
            String cfgTempFileLoc = "d:\temp";
            String strSysTempPath = "";

            strSysTempPath = Environment.GetEnvironmentVariable("temp");
            if (strSysTempPath == null || strSysTempPath.Equals(""))
            {
                strSysTempPath = Environment.GetEnvironmentVariable("tmp");
            }

            if (cfgTempFileLoc == null || cfgTempFileLoc.Trim().Equals(""))
            {
                // otherwise, get environment temp dir
                if (strSysTempPath == null || strSysTempPath.Equals(""))
                {
                    cfgTempFileLoc = "c:\\temp";
                    if (!System.IO.Directory.Exists(cfgTempFileLoc))
                    {
                        System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                    }
                }
                else
                {
                    cfgTempFileLoc = strSysTempPath;
                }
            }
            else
            {
                // judge whether this dir exist
                // otherwise create one
                if (!System.IO.Directory.Exists(cfgTempFileLoc))
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                    }
                    catch (System.Exception ex)
                    {
                        // if failed, then get tmp dir of system
                        if (strSysTempPath == null || strSysTempPath.Equals(""))
                        {
                            cfgTempFileLoc = "c:\\temp";
                            if (!System.IO.Directory.Exists(cfgTempFileLoc))
                            {
                                System.IO.Directory.CreateDirectory(cfgTempFileLoc);
                            }
                        }
                        else
                        {
                            cfgTempFileLoc = strSysTempPath;
                        }
                    }
                    finally
                    {

                    }
                }
            }// 


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2646] ThisAddIn_Startup,BEFORE set config to HttpOper");
            }

#if MSG
            MessageBox.Show("5");
#endif

            //m_HttpOper.setConfig(cfgLoginUrl, cfgDocRepositoryUrl, cfgTempFileLoc);
            //m_HttpOper.setConfig(cfgLoginUrl, 
            //                  cfgUploadFileUrl,
            //                  cfgDocRepositoryUrl,cfgTempFileLoc);

            m_cfgTempFileLoc = cfgTempFileLoc;


            //createDefaultPermission();
            // createDefaultPermission();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2666] ThisAddIn_Startup,AFTER set config to HttpOper");
            }

#if MSG
            MessageBox.Show("6");
#endif

            //             if (ThisAddIn.m_bLog)
            //             {
            //                 String strCnt = "cfgDocRepositoryUrl:" + cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl + ",cfgCerSvrUrl:" + cfgCerSvrUrl;
            //                 Log.WriteLog("[2676] ThisAddIn_Startup,BEFORE uiCtrler init,paras:" + strCnt);
            //             }

            // m_uiCtrler.init(cfgDocRepositoryUrl, cfgUiCtrlSvrUrl, cfgCerSvrUrl, cfgAutoUpdateSvrUrl, m_commTools);

            if (m_edtCenter == null)
            {
                m_edtCenter = new classMultiEditionCenter(m_commTools);
            }

            createDefaultPermission();

            m_edtCenter.Init();


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2804] ThisAddIn_Startup,AFTER m_edtCenter.Init()");
            }


            classMultiEditionCenter.clPermissionItem permItem = m_edtCenter.SearchEditionPerms(m_edtCenter.m_strDocRepositoryEditionName);

            if (permItem != null)
            {
                m_HttpOper.setConfig(permItem.cfgLoginUrl,
                              permItem.cfgUploadFileUrl,
                              permItem.cfgDocRepositoryUrl, cfgTempFileLoc);
            }

            m_HttpOper.setCommonTools(m_commTools);


#if MSG
            MessageBox.Show("7");
#endif

            //             if (ThisAddIn.m_bLog)
            //             {
            //                 String strCnt = "cfgDocRepositoryUrl:" + cfgDocRepositoryUrl + ",cfgUiCtrlSvrUrl:" + cfgUiCtrlSvrUrl + ",cfgCerSvrUrl:" + cfgCerSvrUrl;
            //                 Log.WriteLog("[2688] ThisAddIn_Startup,AFTER uiCtrler init,paras:" + strCnt);
            //             }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2694] ThisAddIn_Startup,BEFORE uiCtrlerloadUiPermTable");
            }

            // m_uiCtrler.loadUiPermTable();

#if MSG
            MessageBox.Show("8");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2705] ThisAddIn_Startup,AFTER uiCtrlerloadUiPermTable");
            }

            // 签到成功后再检查更新
            // if (Settings.Default.bAutoUpdate && searchPermission("chkAutoCheckUpdate") > 0)
            if (Settings.Default.bAutoUpdate && searchPermission("grpFuncPages") > 0)
            {
#if MSG
                MessageBox.Show("9");
#endif
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2716] ThisAddIn_Startup,BEFORE check update");
                }

                // m_uiCtrler.checkUpdate();

                m_edtCenter.CheckUpdate();

            }

            // m_hashVstoPermission = m_hashDefaultPermission;

#if MSG
            MessageBox.Show("10");
#endif

            /*
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2730] ThisAddIn_Startup,INIT database");
            }

            OleDbConnection dbConnection = new OleDbConnection(Settings.Default.localdbConnectionString);

            m_tblAdapterMgr.tblListLevelSchemesTableAdapter = new tblListLevelSchemesTableAdapter();
            // m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection =  new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Connection = dbConnection;

            // Settings.Default.localdbConnectionString;

            m_tblAdapterMgr.tblListLevelTableAdapter = new tblListLevelTableAdapter();
            // m_tblAdapterMgr.tblListLevelTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblListLevelTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter = new tblUniformStyleHistoryDocsTableAdapter();
            // m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter = new tblHeadingStyleSchemeTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter = new tblHeadingStyleFontTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Connection = dbConnection;

            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter = new tblHeadingStyleParagraphFormatTableAdapter();
            // m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = new OleDbConnection(Settings.Default.localdbConnectionString);
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Connection = dbConnection;

            m_localDb.tblListLevelSchemes.Clear();
            m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Fill(m_localDb.tblListLevelSchemes);

            m_localDb.tblListLevel.Clear();
            m_tblAdapterMgr.tblListLevelTableAdapter.Fill(m_localDb.tblListLevel);

            m_localDb.tblUniformStyleHistoryDocs.Clear();
            m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Fill(m_localDb.tblUniformStyleHistoryDocs);

            m_localDb.tblHeadingStyleScheme.Clear();
            m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Fill(m_localDb.tblHeadingStyleScheme);

            m_localDb.tblHeadingStyleFont.Clear();
            m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Fill(m_localDb.tblHeadingStyleFont);

            m_localDb.tblHeadingStyleParagraphFormat.Clear();
            m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Fill(m_localDb.tblHeadingStyleParagraphFormat);

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2781] ThisAddIn_Startup,AFTER INIT database");
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strTagDir = strBaseDir + @"config\lidong.txt";
            String strDbDir = strBaseDir + @"config\db.txt";

            // if not exist
            if (System.IO.File.Exists(strTagDir) && System.IO.File.Exists(strDbDir))
            {
                rebuildHeadingSnPreBuiltInDB();

                rebuildHeadingStylePreBuiltInDB();

                FileStream fs = new FileStream(strDbDir, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
                sw.Write("done");
                sw.Close();
                fs.Close();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[2803] ThisAddIn_Startup,reCreate database");
                }

            }

#if MSG
            MessageBox.Show("11");
#endif

            if (ThisAddIn.m_bLog)
            {
                int nOffset = 10;
                String strMethod = "ThisAddIn_Startup";

                String strCnt = "Offset:" + nOffset + ",Function:" + strMethod;

                Log.WriteLog(strCnt);
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2968] ThisAddIn_Startup,BEFORE loadAllHeadingSnSchemes()");
            }

            loadAllHeadingSnSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2975] ThisAddIn_Startup,AFTER loadAllHeadingSnSchemes() / BEFORE loadAllHeadingStyleSchemes");
            }

            loadAllHeadingStyleSchemes();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2982] ThisAddIn_Startup,AFTER loadAllHeadingStyleSchemes()");
            }* /


            InitDataBase_v1();


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2968] ThisAddIn_Startup,BEFORE loadAllHeadingSnSchemes()");
            }

            loadAllHeadingSnSchemes_v1();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2975] ThisAddIn_Startup,AFTER loadAllHeadingSnSchemes() / BEFORE loadAllHeadingStyleSchemes");
            }

            loadAllHeadingStyleSchemes_v1();

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[2982] ThisAddIn_Startup,AFTER loadAllHeadingStyleSchemes()");
            }

#if MSG
            MessageBox.Show("12");
#endif

            // 
            m_hashWordFontNum.Clear();

            float fSize = 0.0f;

            fSize = 42f;
            m_hashWordFontNum.Add(fSize, "初号");

            fSize = 36f;
            m_hashWordFontNum.Add(fSize, "小初");

            fSize = 26f;
            m_hashWordFontNum.Add(fSize, "一号");

            fSize = 24f;
            m_hashWordFontNum.Add(fSize, "小一");

            fSize = 22f;
            m_hashWordFontNum.Add(fSize, "二号");

            fSize = 18f;
            m_hashWordFontNum.Add(fSize, "小二");

            fSize = 16f;
            m_hashWordFontNum.Add(fSize, "三号");

            fSize = 15f;
            m_hashWordFontNum.Add(fSize, "小三");

            fSize = 14f;
            m_hashWordFontNum.Add(fSize, "四号");

            fSize = 12f;
            m_hashWordFontNum.Add(fSize, "小四");

            fSize = 10f;
            m_hashWordFontNum.Add(fSize, "五号");

            //fSize = 10.0f;
            //m_hashWordFontNum.Add(fSize, "五号");

            fSize = 9f;
            m_hashWordFontNum.Add(fSize, "小五");

            fSize = 7f;
            m_hashWordFontNum.Add(fSize, "六号");

            fSize = 6f;
            m_hashWordFontNum.Add(fSize, "小六");

            fSize = 5f;
            m_hashWordFontNum.Add(fSize, "七号");

            //             fSize = 5f;
            //             m_hashWordFontNum.Add(fSize, "八号");

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[3048] ThisAddIn_Startup, loaded font");
            }

            Boolean bIsPic = true;
            m_hashPicFileType.Add(".emf", bIsPic);
            m_hashPicFileType.Add(".wmf", bIsPic);
            m_hashPicFileType.Add(".jpg", bIsPic);
            m_hashPicFileType.Add(".jpeg", bIsPic);
            m_hashPicFileType.Add(".jfif", bIsPic);
            m_hashPicFileType.Add(".jpe", bIsPic);
            m_hashPicFileType.Add(".png", bIsPic);
            m_hashPicFileType.Add(".bmp", bIsPic);
            m_hashPicFileType.Add(".dib", bIsPic);
            m_hashPicFileType.Add(".rle", bIsPic);
            m_hashPicFileType.Add(".bmz", bIsPic);
            m_hashPicFileType.Add(".gif", bIsPic);
            m_hashPicFileType.Add(".gfa", bIsPic);
            m_hashPicFileType.Add(".emz", bIsPic);
            m_hashPicFileType.Add(".wmz", bIsPic);
            m_hashPicFileType.Add(".pcz", bIsPic);
            m_hashPicFileType.Add(".tif", bIsPic);
            m_hashPicFileType.Add(".tiff", bIsPic);
            m_hashPicFileType.Add(".cgm", bIsPic);
            m_hashPicFileType.Add(".eps", bIsPic);
            m_hashPicFileType.Add(".pct", bIsPic);
            m_hashPicFileType.Add(".pict", bIsPic);
            m_hashPicFileType.Add(".wpg", bIsPic);


            Boolean bInsertbleDoc = true;

            m_hashDocFileType.Add(".docx", bInsertbleDoc);
            m_hashDocFileType.Add(".docm", bInsertbleDoc);
            m_hashDocFileType.Add(".dotx", bInsertbleDoc);
            m_hashDocFileType.Add(".dotm", bInsertbleDoc);
            m_hashDocFileType.Add(".doc", bInsertbleDoc);
            m_hashDocFileType.Add(".dot", bInsertbleDoc);
            m_hashDocFileType.Add(".htm", bInsertbleDoc);
            m_hashDocFileType.Add(".html", bInsertbleDoc);
            m_hashDocFileType.Add(".rtf", bInsertbleDoc);
            m_hashDocFileType.Add(".mht", bInsertbleDoc);
            m_hashDocFileType.Add(".mhtml", bInsertbleDoc);
            m_hashDocFileType.Add(".xml", bInsertbleDoc);
            m_hashDocFileType.Add(".txt", bInsertbleDoc);
            m_hashDocFileType.Add(".wpd", bInsertbleDoc);
            m_hashDocFileType.Add(".wps", bInsertbleDoc);
            m_hashDocFileType.Add(".wtf", bInsertbleDoc);

            //             if (m_bLicIllegal)
            //             {
            //                 Settings.Default.bRegSuc = false;
            //                 Settings.Default.bRegExp = true;
            //                 Settings.Default.Save();
            //                 MessageBox.Show("当前注册数据文件非法或内容损坏，请重新注册");
            //             }


            for (int i = 0; i < 9; i++)
            {
                m_listLevels[i] = new ClassListLevel();
                //m_listLevels[i].Font = new ClassFont();
            }

#if MSG
            MessageBox.Show("13");
#endif

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[3117] Exit ThisAddIn_Startup");
            }

            return;
        }
        */

        void Application_WindowSelectionChange(Word.Selection Sel)
        {
            if (Sel.Range.Paragraphs[1].OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                Globals.Ribbons.Ribbon1.lblCurParaOutLine.Label = "当前段落级：正";
            }
            else
            {
                int nLvl = (int)Sel.Range.Paragraphs[1].OutlineLevel;

                Globals.Ribbons.Ribbon1.lblCurParaOutLine.Label = "当前段落级：" + nLvl;
            }

            return;
        }


        // 
        public void reFillHeadingStyleSchemes()
        {
            //try
            //{
            //    m_localDb.tblHeadingStyleScheme.Clear();
            //    m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Fill(m_localDb.tblHeadingStyleScheme);

            //    m_localDb.tblHeadingStyleFont.Clear();
            //    m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Fill(m_localDb.tblHeadingStyleFont);

            //    m_localDb.tblHeadingStyleParagraphFormat.Clear();
            //    m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Fill(m_localDb.tblHeadingStyleParagraphFormat);

            //}
            //catch (System.Exception ex)
            //{

            //}
            //finally
            //{

            //}

            return;
        }



        public void reFillHeadingSnSchemes()
        {
            //try
            //{
            //    m_localDb.tblListLevelSchemes.Clear();
            //    m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Fill(m_localDb.tblListLevelSchemes);
	
            //    m_localDb.tblListLevel.Clear();
            //    m_tblAdapterMgr.tblListLevelTableAdapter.Fill(m_localDb.tblListLevel);
	
            //    // m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Fill(m_localDb.tblUniformStyleHistoryDocs);
            //}
            //catch (System.Exception ex)
            //{
            	
            //}
            //finally
            //{
                
            //}

            return;
        }


        private void Application_Quit()
        {

            return;
        }


        private void Application_NewDocument(Word.Document doc)
        {
            //MessageBox.Show("Application_NewDocument");
            //return;

            //AddTaskPane(doc);
            int ncode = doc.GetHashCode();
            return;
        }


        void Application_DocumentBeforePrint(Word.Document Doc, ref bool Cancel)
        {
            //@TODO, add specific PageNum or heading number for link field
            // 
        }


        private void updateTableContents(Word.Document Doc)
        {
//             if (Doc.Fields.Count > 0)
//             {
//                 Doc.Fields.Update();
// 
//                 //                 foreach (Word.Field fld in Doc.Fields)
//                 //                 {
//                 //                     fld.Update();
//                 //                 }
//             }


            if (Doc.TablesOfContents.Count > 0)
            {
                foreach (Word.TableOfContents tblCnt in Doc.TablesOfContents)
                {
                    tblCnt.Update();
                }
            }

            if (Doc.TablesOfFigures.Count > 0)
            {
                foreach (Word.TableOfFigures figCnt in Doc.TablesOfFigures)
                {
                    figCnt.Update();
                }
            }

            return;
        }



        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            // if (m_bUpdTblCntOnSaving && searchPermission("chkBoxUpdTblCntOnSaving") > 0)
            if (m_bUpdTblCntOnSaving && searchPermission("grpComOp") > 0)
            {
                if (AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = Application.UndoRecord;
                    ur.StartCustomRecord("保存时更新目录");
                }

                updateTableContents(Doc);

                if (AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = Application.UndoRecord;
                    ur.EndCustomRecord();
                }
            }

            
            // if(doc.SaveFormat != Word.WdSaveFormat.
            String strExt = Path.GetExtension(Doc.FullName);
            String strPath = Path.GetFullPath(Doc.FullName);


            Object objKey = Doc; // Doc.CurrentRsid;
            CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_hashTaskPane[objKey];
            
            if (myCustomTaskPane != null)
            {
                OperationPanel userPane = (OperationPanel)myCustomTaskPane.Control;

                if (searchPermission("tabPageInfo") > 0)
                {
                    userPane.infoReresh();
                }

                if (searchPermission("tabPageRel") > 0)
                {
                    if (strExt != null && !strExt.Equals("") && !strExt.ToUpper().Equals(".DOCX") && 
                        strPath != null && !strPath.Equals(""))
                    {
                        // MessageBox.Show(@"保存版本不能支持关联计算等功能，请勿使用", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    userPane.refreshAllRels();
                }

                if (searchPermission("tabPageFormDesign") > 0)
                {
                    // restore 
                    userPane.restoreAllFormCntItem();
                }
                
            }
            
            // if (Settings.Default.bGenLocalVer && searchPermission("chkGenLocalVer") > 0)
            if (Settings.Default.bGenLocalVer && searchPermission("groupLocalVer") > 0)
            {
                if (String.IsNullOrWhiteSpace(strExt))
                {
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

                        String[] format = {"yyyyMMddhhmmssff"};
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

                    String strTmpFile = strVerDir + strTrimTimestampNameNoExt + "_" + strPostx + strExt;

                    if (File.Exists(strTmpFile)) // 判断存在
                    {
                        try
                        {
                        	File.Delete(strTmpFile);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                        }
                    }

                    if (!File.Exists(strTmpFile))
                    {
                        try
                        {
                            Doc.Save();
                            File.Copy(Doc.FullName, strTmpFile);                            
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                        }

                    }// if
                }

            }

            return;
        }


        void Application_Startup()
        {
            // this.Application.Options.StoreRSIDOnSave = false;

            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app.ActiveDocument;

            int ncode = doc.GetHashCode();

            return;
        }

        public void loadDocumentChange()
        {
            Application_DocumentChange();
            return;
        }

        void Application_DocumentChange()
        {
            // this code is to place into Event of new document. 
            // Since office2007's bug, we can't access this appropriate event
            // So, place code here to workaround work it out.
            //MessageBox.Show("Application_DocumentChange");
            //return;

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[4665]Application_DocumentChange Enter");
            }


            if (!this.Application.Visible || this.Application.Documents.Count <= 0)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[4701] " + "visible:" + this.Application.Visible + ",cnt:" + this.Application.Documents.Count);
                }

                return;
            }

            Word.Document doc = null;

            try
            {
            	doc = this.Application.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[4709]" + ex.Message);
                }

                return;
            }
            finally
            {
            }

            if (doc == null)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[4722]");
                }

                return;
            }


            
            // Word.Window 
            //if (doc.ActiveWindow.Visible == false)
            //{
            //    return;
            //}

            // Object objKey = doc; // doc.CurrentRsid;
            // CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_hashTaskPane[objKey];
            
            // dynamic dynParent = doc.Parent;

            //if (myCustomTaskPane == null)
            //{
            AddTaskPane(doc);
            //}

            if (m_bAppIsWps)
            {
                CustomTaskPane tmpTp = null;

                CustomTaskPane curTaskPane = (CustomTaskPane)m_hashTaskPane[doc];
                if (curTaskPane != null)
                {
                    Boolean bVisible = Settings.Default.bUIShow;

                    if (m_hashDocVisible.Contains(doc))
                    {
                        bVisible = (Boolean)m_hashDocVisible[doc];
                    }

                    foreach (DictionaryEntry dt in m_hashTaskPane)
                    {
                        tmpTp = (CustomTaskPane)dt.Value;
                        if (tmpTp.Visible)
                        {
                            tmpTp.Visible = false;
                        }
                    }

                    curTaskPane.Visible = bVisible;
                }
            }
            else
            {
                //// 
                //CustomTaskPane tmpTp = null;

                //CustomTaskPane curTaskPane = (CustomTaskPane)m_hashTaskPane[doc];
                //if (curTaskPane != null)
                //{
                //    Boolean bVisible = Settings.Default.bUIShow;

                //    if (m_hashDocVisible.Contains(doc))
                //    {
                //        bVisible = (Boolean)m_hashDocVisible[doc];
                //    }

                //    curTaskPane.Visible = bVisible;
                //}
            }



            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[4739]Application_DocumentChange EXIT");
            }

            return;
        }

        void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            Object objKey = Doc; // Doc.CurrentRsid;
            CustomTaskPane myCustomTaskPane = (CustomTaskPane)m_hashTaskPane[objKey];
            if (myCustomTaskPane != null)
            {
                OperationPanel userPane = (OperationPanel)myCustomTaskPane.Control;
                // userPane.infoReresh();

                if (searchPermission("tabPageRel") > 0)
                {
                    userPane.refreshAllRels();
                }

                if (searchPermission("tabPageFormDesign") > 0)
                {
                    userPane.releaseAllFormResToDoc();
                }
            }

            // if (m_bUpdTblCntOnClosing && searchPermission("chkBoxUpdTblCntOnClose") > 0)
            if (m_bUpdTblCntOnClosing && searchPermission("grpComOp") > 0)            
            {
                if (AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = Application.UndoRecord;
                    ur.StartCustomRecord("关闭时更新目录");
                }

                updateTableContents(Doc);

                if (AppVersion >= 15) // 2013
                {
                    Word.UndoRecord ur = Application.UndoRecord;
                    ur.EndCustomRecord();
                }
            }

            if (Cancel)
            {
                RemoveTaskPane(Doc);
            }

            // m_bFromDocClose = true;

            return;
        }

        void Application_DocumentOpen(Word.Document doc)
        {
            //MessageBox.Show("Application_DocumentOpen");
            //return;

            //AddTaskPane(Doc);
            int ncode = doc.GetHashCode();
            return;
        }

        public void Sync2Db()
        {
            //try
            //{
            //    if (m_localDb.tblListLevel.GetChanges() != null)
            //    {
            //        m_tblAdapterMgr.tblListLevelTableAdapter.Update(m_localDb.tblListLevel);
            //    }

            //    if (m_localDb.tblListLevelSchemes.GetChanges() != null)
            //    {
            //        m_tblAdapterMgr.tblListLevelSchemesTableAdapter.Update(m_localDb.tblListLevelSchemes);
            //    }

            //    if (m_localDb.tblUniformStyleHistoryDocs.GetChanges() != null)
            //    {
            //        m_tblAdapterMgr.tblUniformStyleHistoryDocsTableAdapter.Update(m_localDb.tblUniformStyleHistoryDocs);
            //    }

            //    if (m_localDb.tblHeadingStyleScheme.GetChanges() != null)
            //    {
            //        m_tblAdapterMgr.tblHeadingStyleSchemeTableAdapter.Update(m_localDb.tblHeadingStyleScheme);
            //    }

            //    if (m_localDb.tblHeadingStyleFont.GetChanges() != null)
            //    {
            //        m_tblAdapterMgr.tblHeadingStyleFontTableAdapter.Update(m_localDb.tblHeadingStyleFont);
            //    }

            //    if (m_localDb.tblHeadingStyleParagraphFormat.GetChanges() != null)
            //    {
            //        m_tblAdapterMgr.tblHeadingStyleParagraphFormatTableAdapter.Update(m_localDb.tblHeadingStyleParagraphFormat);
            //    }
            //}
            //catch (System.Exception ex)
            //{
            	
            //}
            //finally
            //{
            //}

            return;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // saveUniformStyleHistoryStyleDocs();

            Sync2Db();

            //m_tblAdapterMgr.Dispose();
            //m_localDb.Dispose();

            return;
        }


        public Boolean IsPicFileFormat(String strFileName)
        {
            String strFileType = Path.GetExtension(strFileName);

            if (strFileType == null)
                return false;

            String strFileExtType = strFileType.ToLower();
            
            if (!m_hashPicFileType.Contains(strFileExtType))
            {
                return false;
            }

            return true;
        }

        public Boolean IsInsertbleFileFormat(String strFileName)
        {
            String strFileType = Path.GetExtension(strFileName);

            if (strFileType == null)
                return false;


            String strFileExtType = strFileType.ToLower();

            if (!m_hashDocFileType.Contains(strFileExtType))
            {
                return false;
            }

            return true;
        }


        public Boolean IsWordDocFileFormat(String strFileName)
        {
            String strFileType = Path.GetExtension(strFileName);

            if (strFileType == null)
                return false;

            String strFileExtType = strFileType.ToLower();

            if (m_bAppIsWps)
            {
                if (!strFileExtType.Equals(".doc") &&
                    !strFileExtType.Equals(".docx")&&
                    !strFileExtType.Equals(".wps"))
                {
                    return false;
                }
            }
            else
            {
                if (!strFileExtType.Equals(".doc") &&
                    !strFileExtType.Equals(".docx"))
                {
                    return false;
                }
            }


            return true;
        }


        public void fillHeadingStyleItem(ClassHeadingStyle[] hs, int i, String strFntName)
        {
            String[] strArrFntNameBi = new String[] {
                    "Times New Roman",
                    "+标题 CS",
                    "+正文 CS 字体",
                    "+标题 CS",
                    "+正文 CS 字体",
                    "+标题 CS",
                    "+正文 CS 字体",
                    "+标题 CS",
                    "+标题 CS",
                    "+正文 CS 字体"
            };

            /*
            float[] fArrFntSize = new float[] {
                36,26,24,22,18,16,15,14,12,10.0f
            };*/

            float[] fArrFntSize = new float[] {
                22.0f,16.0f,16.0f,14.0f,14.0f,12.0f,12.0f,12.0f,10.0f,10.0f};
            /*
            标题1：中文正文，宋体，2号，加粗
            标题2：宋体，3号，加粗
            标题3：宋体，3号，加粗
            标题4：宋体，4号，加粗
            标题5：宋体，4号，加粗
            标题6：宋体，小四号，加粗
            标题7：宋体，小四号，加粗
            标题8：宋体，小四
            标题9：宋体，5号
            正文：宋体，5号
            */

            float[] fArrFntKerning = new float[] {
                16,1,1,1,1,1,1,1,1,1
            };


            // float[] fArrParaFmtLineSpacing = new float[]{
            //    12,20.8f,20.8f,18.8f,18.8f,16,16,16,16,12
            // };

            float[] fArrParaFmtLineSpacing = new float[]{
                12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f
            };

            float[] fArrParaFmtSpaceBefore = new float[]{
                5,13,13,14,14,12,12,12,12,0
            };

            float[] fArrParaFmtSpaceAfter = new float[]{
                5,13,13,14.5f,14.5f,3.2f,3.2f,3.2f,3.2f,0
            };

            hs[i].m_fnt.Color = Word.WdColor.wdColorAutomatic;// -16777216;
            hs[i].m_fnt.UnderlineColor = Word.WdColor.wdColorAutomatic;// -16777216;
            hs[i].m_fnt.DiacriticColor = (Word.WdColor)Word.WdConstants.wdUndefined;

            hs[i].m_fnt.Name = strFntName;
            hs[i].m_fnt.NameFarEast = strFntName;
            hs[i].m_fnt.NameAscii = strFntName;//"+西文正文";
            //hs[i].m_fnt.NameOther = strFntName;//"+西文正文";

            hs[i].m_fnt.NameBi = strArrFntNameBi[i];

            hs[i].m_fnt.Size = fArrFntSize[i];
            hs[i].m_fnt.SizeBi = fArrFntSize[i];


            hs[i].m_fnt.Kerning = fArrFntKerning[i];
            hs[i].m_fnt.Scaling = 100;


            // paragraph format
            hs[i].m_paraFmt.OutlineLevel = (Word.WdOutlineLevel)(i + 1);
            hs[i].m_paraFmt.AddSpaceBetweenFarEastAndAlpha = -1;
            hs[i].m_paraFmt.AddSpaceBetweenFarEastAndDigit = -1;
            hs[i].m_paraFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            hs[i].m_paraFmt.AutoAdjustRightIndent = -1;
            hs[i].m_paraFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto; // 4
            hs[i].m_paraFmt.CharacterUnitFirstLineIndent = 0;
            hs[i].m_paraFmt.CharacterUnitLeftIndent = 0;
            hs[i].m_paraFmt.CharacterUnitRightIndent = 0;
            hs[i].m_paraFmt.DisableLineHeightGrid = 0;
            hs[i].m_paraFmt.FarEastLineBreakControl = -1;
            hs[i].m_paraFmt.FirstLineIndent = 0;
            hs[i].m_paraFmt.HalfWidthPunctuationOnTopOfLine = 0;
            hs[i].m_paraFmt.HangingPunctuation = -1;
            hs[i].m_paraFmt.Hyphenation = -1;

            if (i < 9)
            {
                hs[i].m_paraFmt.KeepTogether = -1;
                hs[i].m_paraFmt.KeepWithNext = -1;
            }
            else
            {
                hs[i].m_paraFmt.KeepTogether = 0;
                hs[i].m_paraFmt.KeepWithNext = 0;
            }


            hs[i].m_paraFmt.LeftIndent = 0;
            hs[i].m_paraFmt.LineSpacing = fArrParaFmtLineSpacing[i];
            hs[i].m_paraFmt.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;//Word.WdLineSpacing.wdLineSpaceMultiple; // 5
            hs[i].m_paraFmt.LineUnitAfter = 0;
            hs[i].m_paraFmt.LineUnitBefore = 0;
            hs[i].m_paraFmt.MirrorIndents = 0;
            hs[i].m_paraFmt.NoLineNumber = 0;
            hs[i].m_paraFmt.PageBreakBefore = 0;
            hs[i].m_paraFmt.ReadingOrder = Word.WdReadingOrder.wdReadingOrderLtr; // 1
            hs[i].m_paraFmt.RightIndent = 0;
            hs[i].m_paraFmt.SpaceAfter = fArrParaFmtSpaceAfter[i];
            hs[i].m_paraFmt.SpaceAfterAuto = 0; // [0] = -1, others = 0
            hs[i].m_paraFmt.SpaceBefore = fArrParaFmtSpaceBefore[i];
            hs[i].m_paraFmt.SpaceBeforeAuto = 0; // [0] = -1, others = 0
            hs[i].m_paraFmt.TextboxTightWrap = 0;
            hs[i].m_paraFmt.WidowControl = 0; // [0] = -1, others = 0
            hs[i].m_paraFmt.WordWrap = -1;

            return;
        }

        public void rebuildHeadingStylePreBuiltInDBNH()
        {
            removeHeadingStyleSchemeNH(true);

            Word.Application app = this.Application;

            int nIndex = 0, i = 0;

            ClassHeadingStyle[] hs = null;
            String strSchemeName = "";

            String[] arrStrFontName = new String[] {
                    "宋体",
                    "黑体",
                    "仿宋",
                    "新宋体",
                    "微软雅黑",
                    "方正姚体",
                    "华文中宋",
                    "楷体",
                    "隶书",
                    "华文宋体",
                    "方正兰亭超细黑简体",
                    "方正舒体",
                    "华文彩云",
                    "华文仿宋",
                    "华文行楷",
                    "华文琥珀",
                    "华文楷体",
                    "华文隶书",
                    "华文细黑",
                    "华文新魏",
                    "思源黑体",
                    "幼圆"
            };


            ///////////////////
            hs = new ClassHeadingStyle[10];

            nIndex++;
            strSchemeName = @"方案" + nIndex.ToString("00") + ":" + "Word经典";

            for (i = 0; i < 10; i++)
            {
                hs[i] = new ClassHeadingStyle();

                fillHeadingStyleItem(hs, i, "宋体");
                hs[i].m_fnt.Bold = -1;
            }
            hs[8].m_fnt.Bold = 0;
            hs[9].m_fnt.Bold = 0;

            addHeadingStyleSchemeNH(strSchemeName, hs, true);



            foreach (String strFntName in arrStrFontName)
            {

                hs = new ClassHeadingStyle[10];

                nIndex++;
                strSchemeName = @"方案" + nIndex.ToString("00") + ":" + strFntName;

                for (i = 0; i < 10; i++)
                {
                    hs[i] = new ClassHeadingStyle();

                    fillHeadingStyleItem(hs, i, strFntName);
                }

                addHeadingStyleSchemeNH(strSchemeName, hs, true);
            }

            ///////////////////

            foreach (String strFntName in arrStrFontName)
            {

                hs = new ClassHeadingStyle[10];

                nIndex++;
                strSchemeName = @"方案" + nIndex.ToString("00") + ":" + strFntName + "(加粗)";

                for (i = 0; i < 10; i++)
                {
                    hs[i] = new ClassHeadingStyle();

                    fillHeadingStyleItem(hs, i, strFntName);

                    hs[i].m_fnt.Bold = -1;
                }

                addHeadingStyleSchemeNH(strSchemeName, hs, true);
            }


            return;
        }

        /*
        public void rebuildHeadingStylePreBuiltInDB_v1()
        {
            removeHeadingStyleScheme_v1(true);

            Word.Application app = this.Application;

            int nIndex = 0,i = 0;

            ClassHeadingStyle[] hs = null;
            String strSchemeName = "";

            String[] arrStrFontName = new String[] {
                    "宋体",
                    "黑体",
                    "仿宋",
                    "新宋体",
                    "微软雅黑",
                    "方正姚体",
                    "华文中宋",
                    "楷体",
                    "隶书",
                    "华文宋体",
                    "方正兰亭超细黑简体",
                    "方正舒体",
                    "华文彩云",
                    "华文仿宋",
                    "华文行楷",
                    "华文琥珀",
                    "华文楷体",
                    "华文隶书",
                    "华文细黑",
                    "华文新魏",
                    "思源黑体",
                    "幼圆"
            };


            ///////////////////
            hs = new ClassHeadingStyle[10];

            nIndex++;
            strSchemeName = @"方案" + nIndex.ToString("00") + ":" + "Word经典";

            for (i = 0; i < 10; i++)
            {
                hs[i] = new ClassHeadingStyle();

                fillHeadingStyleItem(hs, i, "宋体");
                hs[i].m_fnt.Bold = -1;
            }
            hs[8].m_fnt.Bold = 0;
            hs[9].m_fnt.Bold = 0;

            addHeadingStyleScheme_v1(strSchemeName, hs, true);



            foreach (String strFntName in arrStrFontName)
            {

                hs = new ClassHeadingStyle[10];

                nIndex++;
                strSchemeName = @"方案" + nIndex.ToString("00") + ":" + strFntName;

                for (i = 0; i < 10; i++)
                {
                    hs[i] = new ClassHeadingStyle();

                    fillHeadingStyleItem(hs, i, strFntName);
                }

                addHeadingStyleScheme_v1(strSchemeName, hs, true);
            }

            ///////////////////

            foreach (String strFntName in arrStrFontName)
            {

                hs = new ClassHeadingStyle[10];

                nIndex++;
                strSchemeName = @"方案" + nIndex.ToString("00") + ":" + strFntName + "(加粗)";

                for (i = 0; i < 10; i++)
                {
                    hs[i] = new ClassHeadingStyle();
                    
                    fillHeadingStyleItem(hs, i, strFntName);

                    hs[i].m_fnt.Bold = -1;
                }

                addHeadingStyleScheme_v1(strSchemeName, hs, true);
            }


            return;
        }
        */

        public void rebuildHeadingSnPreBuiltInDBNH()
        {
            removeHeadingSnSchemeNH(true);

            Word.Application app = this.Application;

            int nIndex = 0;
            // set data and add


            nIndex++;
            String strSchemeName = @"方案" + nIndex.ToString("00") + ":1/1.1/1.1.1/...";

            // 方案1
            ClassListLevel[] lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "%1";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%1.%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%1.%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%1.%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%1.%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // 方案二
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":1./1.1./1.1.1./...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "%1.";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%1.%2.";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%1.%2.%3.";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%1.%2.%3.%4.";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%1.%2.%3.%4.%5.";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5.%6.";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9.";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一章/1/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1章";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一章/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1章";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%1.%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%1.%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%1.%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%1.%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一章/第一节/1/1.1/...";

            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1章";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2节";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/1/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);


            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/第一章/1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            //
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/第一章/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);


            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/第一章/第一节/1/1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "第%3节";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);





            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/1/1.1/1.1.1/...";
            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);


            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/第一章/1/1.1/...";
            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);


            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/第一章/第一节/1/1.1/...";
            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "第%3节";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/第一节/1/1.1/...";

            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2节";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = 0.0f;// app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnSchemeNH(strSchemeName, lstLvels, true);

            // m_addin.Sync2Db();


            return;
        }

        /*
        public void rebuildHeadingSnPreBuiltInDB_v1()
        {
            removeHeadingSnScheme_v1(true);

            Word.Application app = this.Application;

            int nIndex = 0;
            // set data and add


            nIndex++;
            String strSchemeName = @"方案" + nIndex.ToString("00") + ":1/1.1/1.1.1/...";

            // 方案1
            ClassListLevel[] lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "%1";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%1.%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%1.%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%1.%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%1.%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // 方案二
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":1./1.1./1.1.1./...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "%1.";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%1.%2.";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%1.%2.%3.";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%1.%2.%3.%4.";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%1.%2.%3.%4.%5.";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5.%6.";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9.";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一章/1/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1章";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一章/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1章";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%1.%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%1.%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%1.%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%1.%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%1.%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%1.%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一章/第一节/1/1.1/...";

            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1章";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2节";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/1/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);


            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/第一章/1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            //
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/第一章/1.1/1.1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleLegal;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);


            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一部分/第一章/第一节/1/1.1/...";

            lstLvels = new ClassListLevel[9];

            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1部分";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "第%3节";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);


            


            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/1/1.1/1.1.1/...";
            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "%2";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%2.%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%2.%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%2.%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%2.%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%2.%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%2.%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%2.%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);


            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/第一章/1/1.1/...";
            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);


            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/第一章/第一节/1/1.1/...";
            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2章";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "第%3节";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // 
            nIndex++;
            strSchemeName = "方案" + nIndex.ToString("00") + ":第一篇/第一节/1/1.1/...";

            lstLvels = new ClassListLevel[9];
            for (int i = 0; i < 9; i++)
            {
                lstLvels[i] = new ClassListLevel();
                //lstLvels[i].Font = new ClassFont();
                lstLvels[i].Font.Name = "";
            }

            lstLvels[0].NumberFormat = "第%1篇";
            lstLvels[0].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[0].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[0].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[0].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[0].TextPosition = app.CentimetersToPoints(0.76f);
            lstLvels[0].TabPosition = 0f;
            lstLvels[0].ResetOnHigher = 0;
            lstLvels[0].StartAt = 1;
            lstLvels[0].LinkedStyle = "标题 1";

            lstLvels[1].NumberFormat = "第%2节";
            lstLvels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3;
            lstLvels[1].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[1].TextPosition = app.CentimetersToPoints(1.02f);
            lstLvels[1].TabPosition = 0f;
            lstLvels[1].ResetOnHigher = 1;
            lstLvels[1].StartAt = 1;
            lstLvels[1].LinkedStyle = "标题 2";

            lstLvels[2].NumberFormat = "%3";
            lstLvels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[2].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[2].TextPosition = app.CentimetersToPoints(1.27f);
            lstLvels[2].TabPosition = 0f;
            lstLvels[2].ResetOnHigher = 2;
            lstLvels[2].StartAt = 1;
            lstLvels[2].LinkedStyle = "标题 3";

            lstLvels[3].NumberFormat = "%3.%4";
            lstLvels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[3].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[3].TextPosition = app.CentimetersToPoints(1.52f);
            lstLvels[3].TabPosition = 0f;
            lstLvels[3].ResetOnHigher = 3;
            lstLvels[3].StartAt = 1;
            lstLvels[3].LinkedStyle = "标题 4";


            lstLvels[4].NumberFormat = "%3.%4.%5";
            lstLvels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[4].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[4].TextPosition = app.CentimetersToPoints(1.78f);
            lstLvels[4].TabPosition = 0f;
            lstLvels[4].ResetOnHigher = 4;
            lstLvels[4].StartAt = 1;
            lstLvels[4].LinkedStyle = "标题 5";

            lstLvels[5].NumberFormat = "%3.%4.%5.%6";
            lstLvels[5].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[5].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[5].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[5].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[5].TextPosition = app.CentimetersToPoints(2.03f);
            lstLvels[5].TabPosition = 0f;
            lstLvels[5].ResetOnHigher = 5;
            lstLvels[5].StartAt = 1;
            lstLvels[5].LinkedStyle = "标题 6";

            lstLvels[6].NumberFormat = "%3.%4.%5.%6.%7";
            lstLvels[6].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[6].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[6].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[6].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[6].TextPosition = app.CentimetersToPoints(2.29f);
            lstLvels[6].TabPosition = 0f;
            lstLvels[6].ResetOnHigher = 6;
            lstLvels[6].StartAt = 1;
            lstLvels[6].LinkedStyle = "标题 7";

            lstLvels[7].NumberFormat = "%3.%4.%5.%6.%7.%8";
            lstLvels[7].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[7].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[7].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[7].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[7].TextPosition = app.CentimetersToPoints(2.54f);
            lstLvels[7].TabPosition = 0f;
            lstLvels[7].ResetOnHigher = 7;
            lstLvels[7].StartAt = 1;
            lstLvels[7].LinkedStyle = "标题 8";


            lstLvels[8].NumberFormat = "%3.%4.%5.%6.%7.%8.%9";
            lstLvels[8].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            lstLvels[8].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            lstLvels[8].NumberPosition = app.CentimetersToPoints(0f);
            lstLvels[8].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            lstLvels[8].TextPosition = app.CentimetersToPoints(2.79f);
            lstLvels[8].TabPosition = 0f;
            lstLvels[8].ResetOnHigher = 8;
            lstLvels[8].StartAt = 1;
            lstLvels[8].LinkedStyle = "标题 9";

            addHeadingSnScheme_v1(strSchemeName, lstLvels, true);

            // m_addin.Sync2Db();


            return;
        }
        */

        public String GetWordFontNumber(float fSize)
        {
            String strFntNum = (String)m_hashWordFontNum[fSize];

            if (strFntNum == null)
            {
                strFntNum = fSize.ToString();
            }

            return strFntNum;
        }


        public int SyncOperationPanelTreeUI(Word.Document srcDoc,String strUiType, int nOpType, String strPath,TreeNode addNode = null)
        {
            int nRet = 0;

            Word.Application app = this.Application;
            CustomTaskPane taskPane = null;
            Word.Document doc = null, openDoc = null;
            TabControl tblContent = null;
            OperationPanel userPane = null;


            taskPane = (CustomTaskPane)m_hashTaskPane[srcDoc];

            if (taskPane == null)
            {
                return -1;
            }

            userPane = (OperationPanel)taskPane.Control;

            tblContent = userPane.tabCtrl;

            foreach (DictionaryEntry entry in m_hashTaskPane)
            {
                doc = (Word.Document)entry.Key;

                try
                {
                    if (m_bAppIsWps)
                    {
                        openDoc = app.Documents[doc.Name];
                    }
                    else
                    {
                        openDoc = app.Documents[doc];
                    }

                }
                catch (System.Exception ex)
                {
                    continue;
                }
                finally
                {
                }

                if (doc == srcDoc || openDoc == null)
                {
                    continue;
                }

                taskPane = (CustomTaskPane)entry.Value;
                userPane = (OperationPanel)taskPane.Control;

                userPane.SyncOperationPanelTreeUI(strUiType, nOpType, strPath,addNode);
            }

            return nRet;
        }


        public int getPanesCount()
        {
            return m_hashTaskPane.Count;
        }


        public TreeNodeCollection getFillGatherTreeNodes(Word.Document srcDoc)
        {
            Word.Application app = this.Application;
            CustomTaskPane taskPane = null;
            Word.Document doc = null, openDoc = null;
            TabControl tblContent = null;
            OperationPanel userPane = null;
            int nPageIndex = -1;


            taskPane = (CustomTaskPane)m_hashTaskPane[srcDoc];

            if (taskPane == null)
            {
                return null;
            }

            foreach (DictionaryEntry entry in m_hashTaskPane)
            {
                doc = (Word.Document)entry.Key;

                try
                {
                    if (m_bAppIsWps)
                    {
                        openDoc = app.Documents[doc.Name];
                    }
                    else
                    {
                        openDoc = app.Documents[doc];
                    }
                }
                catch (System.Exception ex)
                {
                    continue;
                }
                finally
                {
                }

                if (doc == srcDoc || openDoc == null)
                {
                    continue;
                }

                taskPane = (CustomTaskPane)entry.Value;
                userPane = (OperationPanel)taskPane.Control;

                tblContent = userPane.tabCtrl;

                // 填报汇总tree
                nPageIndex = tblContent.TabPages.IndexOfKey("tabPageFillGather");
                if (nPageIndex != -1)
                {
                    // found
                    TabPage sharePage = tblContent.TabPages[nPageIndex];
                    Control.ControlCollection ctrls = sharePage.Controls;

                    TreeView trv = (TreeView)ctrls["trvFillGatherSchemes"];

                    if (trv != null)
                    {
                        return trv.Nodes;
                    }
                }

            }

            return null;
        }

//         public String getMachineID()
//         {
//             String strMachineId = "";
// 
//             ClassHardInfo clsHardInfo = new ClassHardInfo();
// 
//             String strCpuId = clsHardInfo.GetCpuID();
//             String strMacAddr = clsHardInfo.GetMacAddress();
// 
//             strMachineId = strCpuId + "," + strMacAddr;
// 
//             return strMachineId;
//         }


        public Hashtable getRegisterPermItems(String strPerms)
        {
            Hashtable hashItems = new Hashtable();

            int nStart = 0, nLen = 32;
            String strItem = "";

            while (nStart <= strPerms.Length - nLen)
            {
                strItem = strPerms.Substring(nStart, nLen);
                nStart += nLen;

                hashItems.Add(strItem, strItem);
            }

            return hashItems;
        }

        public void transChildCtrl(Control ctrl, Boolean bEnable = true)
        {
            ctrl.Enabled = bEnable;

            foreach (Control childCtrl in ctrl.Controls)
            {
                transChildCtrl(childCtrl, bEnable);
            }

            return;
        }


        public void updatePermFromLicDat(Hashtable hashKeyControls, Hashtable hashPerm)
        {
            String strNameMD5 = "";
            Object ctrlObj = null;
            Control ctrl = null;
            ToolStripItem item = null;
            RibbonControl ribCtrl = null;

            foreach (DictionaryEntry ent2 in hashKeyControls)
            {
                strNameMD5 = (String)ent2.Key;
                ctrlObj = (Object)hashKeyControls[strNameMD5];

#if ADMIN
                if (ctrlObj != null)
                {
                    if (ctrlObj is Control)
                    {
                        ctrl = (Control)ctrlObj;
                        ctrl.Enabled = true;
                        /*
                        if(ctrl is TabPage)
                        {
                            // 递归enable child controls
                            transChildCtrl(ctrl, true);
                        }
                        else if (! (ctrl is TabControl))
                        {
                            // 递归enable child controls
                            transChildCtrl(ctrl, true);
                        }
                         * */
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
                // continue;
#else

                if (hashPerm.Contains(strNameMD5))
                {
                    if (ctrlObj != null)
                    {
                        if (ctrlObj is Control)
                        {
                            ctrl = (Control)ctrlObj;
                            ctrl.Enabled = true;
                            /*
                            if(ctrl is TabPage)
                            {
                                // 递归enable child controls
                                transChildCtrl(ctrl, true);
                            }
                            else if (! (ctrl is TabControl))
                            {
                                // 递归enable child controls
                                transChildCtrl(ctrl, true);
                            }
                             * */
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

                            /*
                            if (ctrl is TabPage)
                            {
                                // 递归enable child controls
                                transChildCtrl(ctrl, false);
                            }
                            else if (!(ctrl is TabControl))
                            {
                                // 递归enable child controls
                                transChildCtrl(ctrl, false);
                            }
                             * */
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
#endif

/*
                if (m_bTryExpired)
                {
                    if (hashPerm.Contains(strNameMD5))
                    {
                        if (ctrlObj != null)
                        {
                            if (ctrlObj is Control)
                            {
                                ctrl = (Control)ctrlObj;
                                ctrl.Enabled = true;
                                / *
                                if(ctrl is TabPage)
                                {
                                    // 递归enable child controls
                                    transChildCtrl(ctrl, true);
                                }
                                else if (! (ctrl is TabControl))
                                {
                                    // 递归enable child controls
                                    transChildCtrl(ctrl, true);
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
                                    transChildCtrl(ctrl, false);
                                }
                                else if (!(ctrl is TabControl))
                                {
                                    // 递归enable child controls
                                    transChildCtrl(ctrl, false);
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
                }*/

            }

            return;
        }


//         public String EncodeLicData(String strSvrRetLicData, String strMachineID)
//         {
//             String strEncodedLic = "";
// 
//             // prefix random 
//             String strTime = DateTime.Now.ToString();
//             strTime = ClassEncryptUtils.MD5Encrypt(strTime);
// 
//             String strMD5MachineId = ClassEncryptUtils.MD5Encrypt(strMachineID);
// 
//             String str1 = "", str2 = "";
// 
//             str1 = strSvrRetLicData.Substring(0, 64);
//             str2 = strSvrRetLicData.Substring(64);
// 
//             // 3rd insert into MD5 of Machine ID
//             strEncodedLic = strTime + str1 + strMD5MachineId + str2;
// 
//             return strEncodedLic;
//         }


//         public String DecodeLicData(String strSavedLicData,ref String strMD5MachineID)
//         {
//             String strDecodedLic = "";
// 
//             String strBody = strSavedLicData.Substring(32);
// 
//             String str1 = "", str2 = "";
// 
//             str1 = strBody.Substring(0, 64);
//             strMD5MachineID = strBody.Substring(64, 32);
//             str2 = strBody.Substring(96);
// 
//             strDecodedLic = str1 + str2;
// 
//             return strDecodedLic;
//         }


        /*
        public class fillGatherColNameValueItem
        {
            public Boolean bIsTag = false;

            public String strColName = "";
            public int nNameRow = -1, nNameCol = -1;

            public String strValueSample = "";
            public int nValueRow = -1, nValueCol = -1;

            public String strTagName = "";
            public String strTagValue = "";
        }


        public class fillGatherTableItem
        {
            public String strName = "";
            public String strTopoKey = "";

            // public fillGatherColNameValueItem
            public ArrayList arrCols = new ArrayList();
        }


        public class fillGatherSchemeItem
        {
            public String strName = "";
            // fillGatherTableItem
            public ArrayList arrTables = new ArrayList();
        }
        */

        public float transSpaceUnit(float fValue, String strSrcUnit, Boolean bIndent = true,String strDestUnit = "磅")
        {
            float fRet = float.NaN;
            float fPonds = float.NaN;

            if (fValue == 0.0f)
            {
                return 0.0f;
            }

            // {"字符", "磅", "厘米","毫米","英寸", "行"};
            if (strSrcUnit.Equals("字符"))
            {
                if (bIndent)
                {
                    fPonds = fValue * 5.0f; // 1 字符 = 5 磅
                }
                else
                {
                    fPonds = fValue * 11.0f; // 1 字符 = 11 磅
                }
            }
            else if (strSrcUnit.Equals("磅"))
            {
                fPonds = fValue;
            }
            else if (strSrcUnit.Equals("厘米"))
            {
                fPonds = fValue * 28.35f; // 1 cm = 28.35 磅
            }
            else if (strSrcUnit.Equals("毫米"))
            {
                fPonds = fValue * 2.835f; // 1 mm = 2.835 磅
            }
            else if (strSrcUnit.Equals("英寸"))
            {
                fPonds = fValue * 72.0f; // 1 inch = 72 磅
            }
            else if (strSrcUnit.Equals("行"))
            {
                if (bIndent)
                {
                    fPonds = fValue * 17.86f; // 1 line = 5 磅
                }
                else
                {
                    fPonds = fValue * 5.0f; // 1 line = 5 磅
                }
            }


            // 
            if (strDestUnit.Equals("字符"))
            {
                if (bIndent)
                {
                    fRet = fPonds / 5.0f;
                }
                else
                {
                    fRet = fPonds / 11.0f;
                }
            }
            else if (strDestUnit.Equals("磅"))
            {
                fRet = fPonds;
            }
            else if (strDestUnit.Equals("厘米"))
            {
                fRet = fPonds / 28.35f;
            }
            else if (strDestUnit.Equals("毫米"))
            {
                fRet = fPonds / 2.835f;
            }
            else if (strDestUnit.Equals("英寸"))
            {
                fRet = fPonds / 72.0f;
            }
            else if (strDestUnit.Equals("行"))
            {
                if (bIndent)
                {
                    fRet = fPonds / 17.86f;
                }
                else
                {
                    fRet = fPonds / 5.0f;
                }
            }

            return fRet;
        }

        public int LoadDocPubSchemeNames(ref TreeNode typeNodes)
        {
            if (m_bDocPubSchemeNamesLoaded)
            {
                return -99;
            }

            m_bDocPubSchemeNamesLoaded  = true;

            int nRet = m_docPubMgr.LoadNames(ref typeNodes);
            // sync all TaskPanes

            return nRet;         
        }

        public int BuildDocPubSchemeSubTreeNode(ref TreeNode curSchemeNode, docPubScheme docPubSchemeObj)
        {
            // 
            //目录
            //    章节目录
            //        1级
            //        2级
            //        3级
            //        4级
            //        5级
            //        6级
            //        7级
            //        8级
            //        9级
            //    图目录
            //    表格目录
            TreeNode trTocNd = new TreeNode("目录");
            TreeNode trHeadingTocNd = new TreeNode("章节目录");

            for (int i = 1; i <= 9; i++)
            {
                trHeadingTocNd.Nodes.Add(new TreeNode("" + i + "级"));
            }

            trTocNd.Nodes.Add(trHeadingTocNd);
            trTocNd.Nodes.Add(new TreeNode("图目录"));
            trTocNd.Nodes.Add(new TreeNode("表格目录"));

            curSchemeNode.Nodes.Add(trTocNd);


            //章节序号
            //    1级
            //    2级
            //    3级
            //    4级
            //    5级
            //    6级
            //    7级
            //    8级
            //    9级
            TreeNode trHeadingSnNd = new TreeNode("章节序号");
            for (int i = 1; i <= 9; i++ )
            {
                trHeadingSnNd.Nodes.Add(new TreeNode("" + i + "级"));
            }

            curSchemeNode.Nodes.Add(trHeadingSnNd);


            //章节样式
            //    1级
            //    2级
            //    3级
            //    4级
            //    5级
            //    6级
            //    7级
            //    8级
            //    9级
            TreeNode trHeadingStyleNd = new TreeNode("章节样式");
            for (int i = 1; i <= 9; i++)
            {
                trHeadingStyleNd.Nodes.Add(new TreeNode("" + i + "级"));
            }

            curSchemeNode.Nodes.Add(trHeadingStyleNd);

            
            //图
            //    嵌入且独立成段落的图
            TreeNode trInShp = new TreeNode("图");
            trInShp.Nodes.Add(new TreeNode("嵌入且独立成段落的图"));
            curSchemeNode.Nodes.Add(trInShp);

            //表格
            //    整表
            TreeNode trTbl = new TreeNode("表格");
            trTbl.Nodes.Add(new TreeNode("整表"));
            curSchemeNode.Nodes.Add(trTbl);

            //题注
            //    嵌入且独立成段落的图
            //    表格
            TreeNode trTiZhu = new TreeNode("题注");
            trTiZhu.Nodes.Add(new TreeNode("嵌入且独立成段落的图"));
            trTiZhu.Nodes.Add(new TreeNode("表格"));
            curSchemeNode.Nodes.Add(trTiZhu);


            //序号段落
            //    表格内
            //    表格外
            TreeNode trListPara = new TreeNode("序号段落");
            trListPara.Nodes.Add(new TreeNode("表格内"));
            trListPara.Nodes.Add(new TreeNode("表格外"));
            curSchemeNode.Nodes.Add(trListPara);

            //正文
            //    表格内
            //    表格外
            TreeNode trTextBody = new TreeNode("正文");
            trTextBody.Nodes.Add(new TreeNode("表格内"));
            trTextBody.Nodes.Add(new TreeNode("表格外"));
            curSchemeNode.Nodes.Add(trTextBody);


            //页码
            //    第1节
            //    第2节
            //    第3节
            //    第4节
            //    第5节
            //    第6节
            //    第7节
            //    第8节
            //    第9节
            TreeNode trPageNum = new TreeNode("页码");

            for (int i = 1; i <= 9; i++ )
            {
                trPageNum.Nodes.Add(new TreeNode("第" + i + "节"));
            }
            curSchemeNode.Nodes.Add(trPageNum);

            return 0;
        }

        private void initDocPub()
        {

            m_hshDocPubNodeSn.Add("目录\\", 100);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\", 110);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\1级\\", 111);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\2级\\", 112);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\3级\\", 113);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\4级\\", 114);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\5级\\", 115);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\6级\\", 116);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\7级\\", 117);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\8级\\", 118);
            m_hshDocPubNodeSn.Add("目录\\章节目录\\9级\\", 119);

            m_hshDocPubNodeSn.Add("目录\\图目录\\", 120);
            m_hshDocPubNodeSn.Add("目录\\表格目录\\", 130);

            m_hshDocPubNodeSn.Add("章节序号\\", 200);
            m_hshDocPubNodeSn.Add("章节序号\\1级\\", 201);
            m_hshDocPubNodeSn.Add("章节序号\\2级\\", 202);
            m_hshDocPubNodeSn.Add("章节序号\\3级\\", 203);
            m_hshDocPubNodeSn.Add("章节序号\\4级\\", 204);
            m_hshDocPubNodeSn.Add("章节序号\\5级\\", 205);
            m_hshDocPubNodeSn.Add("章节序号\\6级\\", 206);
            m_hshDocPubNodeSn.Add("章节序号\\7级\\", 207);
            m_hshDocPubNodeSn.Add("章节序号\\8级\\", 208);
            m_hshDocPubNodeSn.Add("章节序号\\9级\\", 209);

            m_hshDocPubNodeSn.Add("章节样式\\", 300);
            m_hshDocPubNodeSn.Add("章节样式\\1级\\", 301);
            m_hshDocPubNodeSn.Add("章节样式\\2级\\", 302);
            m_hshDocPubNodeSn.Add("章节样式\\3级\\", 303);
            m_hshDocPubNodeSn.Add("章节样式\\4级\\", 304);
            m_hshDocPubNodeSn.Add("章节样式\\5级\\", 305);
            m_hshDocPubNodeSn.Add("章节样式\\6级\\", 306);
            m_hshDocPubNodeSn.Add("章节样式\\7级\\", 307);
            m_hshDocPubNodeSn.Add("章节样式\\8级\\", 308);
            m_hshDocPubNodeSn.Add("章节样式\\9级\\", 309);

            m_hshDocPubNodeSn.Add("图\\", 400);
            m_hshDocPubNodeSn.Add("图\\嵌入且独立成段落的图\\", 410);

            m_hshDocPubNodeSn.Add("表格\\", 500);
            m_hshDocPubNodeSn.Add("表格\\整表", 510);


            m_hshDocPubNodeSn.Add("题注\\", 600);
            m_hshDocPubNodeSn.Add("题注\\嵌入且独立成段落的图\\", 610);
            m_hshDocPubNodeSn.Add("题注\\表格\\", 620);


            m_hshDocPubNodeSn.Add("序号段落\\", 700);
            m_hshDocPubNodeSn.Add("序号段落\\表格内\\", 710);
            m_hshDocPubNodeSn.Add("序号段落\\表格外\\", 720);


            m_hshDocPubNodeSn.Add("正文\\", 800);
            m_hshDocPubNodeSn.Add("正文\\表格内\\", 810);
            m_hshDocPubNodeSn.Add("正文\\表格外\\", 820);


            m_hshDocPubNodeSn.Add("页码\\", 900);
            m_hshDocPubNodeSn.Add("页码\\第1节\\", 901);
            m_hshDocPubNodeSn.Add("页码\\第2节\\", 902);
            m_hshDocPubNodeSn.Add("页码\\第3节\\", 903);
            m_hshDocPubNodeSn.Add("页码\\第4节\\", 904);
            m_hshDocPubNodeSn.Add("页码\\第5节\\", 905);
            m_hshDocPubNodeSn.Add("页码\\第6节\\", 906);
            m_hshDocPubNodeSn.Add("页码\\第7节\\", 907);
            m_hshDocPubNodeSn.Add("页码\\第8节\\", 908);
            m_hshDocPubNodeSn.Add("页码\\第9节\\", 909);

            ////
            m_hashIndex2ListStyle.Clear();
            m_hashIndex2ListStyle.Add(0, Word.WdListNumberStyle.wdListNumberStyleNone);
            m_hashIndex2ListStyle.Add(1, Word.WdListNumberStyle.wdListNumberStyleArabic);
            m_hashIndex2ListStyle.Add(2, Word.WdListNumberStyle.wdListNumberStyleUppercaseRoman);
            m_hashIndex2ListStyle.Add(3, Word.WdListNumberStyle.wdListNumberStyleLowercaseRoman);
            m_hashIndex2ListStyle.Add(4, Word.WdListNumberStyle.wdListNumberStyleUppercaseLetter);
            m_hashIndex2ListStyle.Add(5, Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter);
            m_hashIndex2ListStyle.Add(6, Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3);
            m_hashIndex2ListStyle.Add(7, Word.WdListNumberStyle.wdListNumberStyleSimpChinNum2);
            m_hashIndex2ListStyle.Add(8, Word.WdListNumberStyle.wdListNumberStyleZodiac1);
            m_hashIndex2ListStyle.Add(9, Word.WdListNumberStyle.wdListNumberStyleZodiac2);
            m_hashIndex2ListStyle.Add(10, Word.WdListNumberStyle.wdListNumberStyleOrdinal);
            m_hashIndex2ListStyle.Add(11, Word.WdListNumberStyle.wdListNumberStyleCardinalText);
            m_hashIndex2ListStyle.Add(12, Word.WdListNumberStyle.wdListNumberStyleOrdinalText);
            m_hashIndex2ListStyle.Add(13, Word.WdListNumberStyle.wdListNumberStyleArabicLZ);
            m_hashIndex2ListStyle.Add(14, Word.WdListNumberStyle.wdListNumberStyleNumberInCircle);

            // 
            m_hashListStyle2Index.Clear();
            foreach (DictionaryEntry entry in m_hashIndex2ListStyle)
            {
                m_hashListStyle2Index.Add(entry.Value, entry.Key);
            }
            ////

            return;
        }

        public int AddHeadingToc()
        {
            Word.Application app = this.Application;

            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                // MessageBox.Show("因为无活动文档，不能应用");
                return -1;
            }
            finally
            {

            }

            Word.Selection sel = doc.ActiveWindow.Selection;

            if (doc.TablesOfContents.Count > 0)
            {
                MessageBox.Show("已经有目录，不能再创建", "失败");
                return -1;
            }

            if (doc.ActiveWindow.ActivePane.View.SeekView != Word.WdSeekView.wdSeekMainDocument)
            {
                MessageBox.Show("请在正文区内", "失败");
                return -2;
            }

            Boolean bInTbl = sel.Range.get_Information(Word.WdInformation.wdWithInTable);
            if (bInTbl)
            {
                MessageBox.Show("不支持在表格中创建", "失败");
                return -3;
            }

            if (AppVersion >= 11) // wps2015
            {
                Word.UndoRecord ur = app.UndoRecord;
                ur.StartCustomRecord("插入独立目录节");
            }

            // insert or table of content
            Object objTrue = true;
            Object objMissing = Type.Missing;
            Object objAddedStyle = "";
            Object obj1Num = 1, obj3Num = 3;


            doc.TablesOfContents.Add(sel.Range, ref objTrue, ref obj1Num, ref obj3Num, ref objMissing, ref objMissing,
                                        ref objTrue, ref objTrue, ref objMissing, ref objTrue, ref objTrue, ref objTrue);
            doc.TablesOfContents[1].TabLeader = Word.WdTabLeader.wdTabLeaderDots;
            // dstDoc.TablesOfContents.Format = Word.WdTocFormat.wdTOCClassic;

            sel.Start = doc.TablesOfContents[1].Range.End;
            sel.End = sel.Start;
            sel.Range.GoTo();

            sel.InsertParagraph();
            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);



            doc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;

            // update table of content
            doc.TablesOfContents[1].Update();

            if (AppVersion >= 11) // wps2015
            {
                Word.UndoRecord ur = app.UndoRecord;
                ur.EndCustomRecord();
            }

            return 0;
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
             this.Startup  += new System.EventHandler(ThisAddIn_Startup);
             this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
