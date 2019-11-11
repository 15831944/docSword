using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using AutoUpdate;
using OfficeTools.Common;
using System.Collections;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OfficeAssist.Properties;


namespace OfficeAssist
{
    public class ClassUIController
    {
        private String m_privVerSvrUrl = "http://10.115.246.179:30000/WordAddinService.asmx";  // 互联网公共地址

        public readonly String m_strUiRecordDatFileName = @"UiRec.dll";

        // private String m_strUiRecFilePath = "";

        // private Boolean bAvaiable = false; 
        public Boolean m_bVerEnterprise = false; // false -- individual ver, true -- enterprise ver.
        public Boolean m_bWithDocRepository = false; // false -- no doc repository, true -- with doc repository

        public String m_strAccount = ""; // 账户名称
        public String m_strActiveSn = "";// 激活码
        public String m_strVerName = ""; // 版本名称

        public Boolean m_bValid = true;
        public String m_strInvalidMessage = "";
        public DateTime m_dtExpireDate = DateTime.Now; 


        private String m_docRepositorySvrUrl = "";
        private String m_uiCtrlSvrUrl = "";
        private String m_licSvrUrl = "";
        private String m_updateSvrUrl = "";

        
        private String m_strUiData = "";
        
        private ClassOfficeCommon m_commTools = null;

        public String m_strMachineId = "";
        public String m_strMD5MachineId = "";

        private Hashtable m_hashUiItems = null;    // pure NAME
        private Hashtable m_hashDefaultUiItems = new Hashtable();

        private Hashtable m_hashUiMD5Items = new Hashtable(); // NAME MD5ized
        private Hashtable m_hashExceptionalUiItems = new Hashtable(); // pure NAME

        private Hashtable m_hashUiName2MD5Name = new Hashtable();

        private AutoUpdate.AutoUpdateClass m_licSvrIntf = null;

        private Hashtable m_hashMD5toNum = new Hashtable();
        private ArrayList m_arrNum2MD5 = new ArrayList();

        


        public ClassUIController()
        {
#if MSG
            MessageBox.Show("p1");
#endif

            ClassHardInfo clsHardInfo = new ClassHardInfo();

            String strCpuId = clsHardInfo.GetCpuID();
            // String strMacAddr = clsHardInfo.GetMacAddress();

            m_strMachineId = strCpuId;// +strMacAddr;
            m_strMachineId = m_strMachineId.ToUpper() ;

            m_strMD5MachineId = ClassEncryptUtils.MD5Encrypt(m_strMachineId);
            m_strMD5MachineId = m_strMD5MachineId.ToUpper();


            String strMD5 = "";

            for (int i = 0; i < 10; i++)
            {
                strMD5 = ClassEncryptUtils.MD5Encrypt(i.ToString());
                strMD5 = strMD5.ToUpper();

                m_arrNum2MD5.Add(strMD5);
                m_hashMD5toNum[strMD5] = i.ToString();
            }

#if MSG
            MessageBox.Show("p2");
#endif

            return;
        }



        public int init(String docRepositorySvrUrl, String uiCtrlSvrUrl,
                        String licSvrUrl,           String updateSvrUrl,
                        ClassOfficeCommon cmnTools)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[111] Enter ClassUIController::init");
            }

//             m_strUiRecFilePath = Environment.GetEnvironmentVariable("APPDATA");
// 
//             m_strUiRecFilePath += "\\docSword\\";
// 
//             if (!String.IsNullOrWhiteSpace(m_strUiRecFilePath) && !File.Exists(m_strUiRecFilePath))
//             {
//                 try
//                 {
//                 	Directory.CreateDirectory(m_strUiRecFilePath);
//                 }
//                 catch (System.Exception ex)
//                 {
//                    
//                 }
//                 finally
//                 {
//                 }
//             }


            String strDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strPrivUrlFile = strDir + "sysUrl.dll";

#if MSG
            MessageBox.Show("a1");
#endif

            if (File.Exists(strPrivUrlFile))
            {
#if MSG
                MessageBox.Show("a2");
#endif

                String strCnt = "";

                try
                {
	                StreamReader rd = new StreamReader(strPrivUrlFile);
	                strCnt = rd.ReadToEnd();
	                rd.Close();
                }
                catch (System.Exception ex)
                {
#if MSG
                    MessageBox.Show("a3,"+ex.Message);
#endif
                    strCnt = "";

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[165] ClassUIController::init, Exception in reading privVerSvrUrl:" + ex.ToString());
                    }

                }
                finally
                {
                }

                if (!String.IsNullOrWhiteSpace(strCnt))
                {
                    m_privVerSvrUrl = strCnt.Trim();

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[179] ClassUIController::init, read privVerSvrUrl:" + m_privVerSvrUrl);
                    }

                }
            }

#if MSG
            MessageBox.Show("a4");
#endif

            restoreDefaultDocRepUiPermHash();

            m_commTools = cmnTools;
           

            // mandatory
            if (String.IsNullOrWhiteSpace(licSvrUrl) || 
                String.IsNullOrWhiteSpace(updateSvrUrl) ||
                String.IsNullOrWhiteSpace(uiCtrlSvrUrl))
            {
                m_bVerEnterprise = false;
                m_licSvrUrl = m_privVerSvrUrl;  // 使用互联网公共地址
                m_updateSvrUrl = m_privVerSvrUrl;
                m_uiCtrlSvrUrl = m_privVerSvrUrl;
            }
            else
            {
                m_bVerEnterprise = true;
                m_licSvrUrl = licSvrUrl; // 使用配置文件中的地址
                m_updateSvrUrl = updateSvrUrl;

                if (String.IsNullOrWhiteSpace(docRepositorySvrUrl))
                {
                    m_bWithDocRepository = false;
                    m_uiCtrlSvrUrl = uiCtrlSvrUrl;     // 使用配置中填写的地址
                }
                else
                {
                    m_bVerEnterprise = true;
                    m_bWithDocRepository = true;

                    m_docRepositorySvrUrl = docRepositorySvrUrl;    // 使用文库地址
                    m_uiCtrlSvrUrl = docRepositorySvrUrl;           // 使用文库地址
                }
            }

#if MSG
            MessageBox.Show("a5");
#endif
            if (m_licSvrIntf == null)
            {
                try
                {
                	m_licSvrIntf = new AutoUpdateClass(m_licSvrUrl);
                }
                catch (System.Exception ex)
                {
#if MSG
                    MessageBox.Show("a6,"+ex.Message);
#endif
                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[241] ClassUIController::init,Create AutoUpdate Exception:" + ex.ToString());
                    }
                }
                finally
                {
                }
            }

            if (ThisAddIn.m_bLog)
            {
                String strCnt = "m_bVerEnterprise:" + m_bVerEnterprise + ",m_bWithDocRepository:" + m_bWithDocRepository +
                                ",m_licSvrUrl:" + m_licSvrUrl + ",m_updateSvrUrl:" + m_updateSvrUrl + ",m_uiCtrlSvrUrl:" + m_uiCtrlSvrUrl;

                Log.WriteLog("[254] Exit ClassUIController::init, paras:" + strCnt);
            }

#if MSG
            MessageBox.Show("a7");
#endif

            return 0;
        }


        public int activateSoft(String strUserName, String strActCode, ref String strRetMsg)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[269] Enter ClassUIController::activateSoft,Paras: UserName:" + strUserName + ",ActCode:" + strActCode);
            }

            int nRet = 0;

            if (ThisAddIn.m_bLog)
            {
                String strCnt = "m_bVerEnterprise:" + m_bVerEnterprise + ",m_bWithDocRepository:" + m_bWithDocRepository +
                                ",m_licSvrUrl:" + m_licSvrUrl + ",m_updateSvrUrl:" + m_updateSvrUrl + ",m_uiCtrlSvrUrl:" + m_uiCtrlSvrUrl;

                Log.WriteLog("[279] ClassUIController::activateSoft, members:" + strCnt);
            }

            if(m_bVerEnterprise)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[286] Exit ClassUIController::activateSoft");
                }

                return -1;
            }

            QueryResult qRet = null;

            try
            {
                if (m_licSvrIntf != null)
                {
                    qRet = m_licSvrIntf.ActiveProject(strUserName, strActCode, m_strMachineId);
                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[301] ClassUIController::activateSoft,call ActiveProject, Paras: UserName:" + strUserName + ",ActCode:" + strActCode + ",MachineId:" + m_strMachineId);
                    }
                }
                else
                {
                    qRet = null;

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[310] ClassUIController::activateSoft,LicSvrIntf is NULL");
                    }
                }
            }
            catch (System.Exception ex)
            {
                qRet = null;

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[320] ClassUIController::activateSoft Exception:" + ex.ToString());
                }
            }
            finally
            {
            }

            if (qRet != null && qRet.IsSuccess)
            {
                if (qRet.Data3 != null)
                {
                    String[] strs = qRet.Data3.Split('_');

                    if (strs != null && strs.GetLength(0) >= 2)
                    {
                        m_strAccount = strs[0];
                        m_strActiveSn = strs[1];
                    }

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[341] ClassUIController::activateSoft ret Data3:" + qRet.Data3);
                    }

                }

                if (qRet.Data2 != null)
                {
                    m_strVerName = qRet.Data2;
                    Settings.Default.strVerName = m_strVerName;
                    Settings.Default.Save();

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[354] ClassUIController::activateSoft ret Data2:" + qRet.Data2);
                    }

                }

                m_dtExpireDate = qRet.Date;

                m_strUiData = qRet.Data1;
                strRetMsg = qRet.ErrorInfo;
                nRet = 0;

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[367] ClassUIController::activateSoft ExpData:" + m_dtExpireDate.ToString() + ",ret Data1:" + qRet.Data1);
                }


                // 判断是否过期超限
                double dbDays = m_commTools.DateDiff(DateTime.Now, qRet.Date);
                int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                if (nDays > 0)
                {
                    m_bValid = false;
                }
                else
                {
                    m_bValid = true;
                }

                m_strInvalidMessage = "到期";

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[388] ClassUIController::activateSoft, members:" + "m_bValid:" + m_bValid + ",m_strInvalidMessage:" + m_strInvalidMessage);
                }

            }
            else
            {
                m_strInvalidMessage = qRet.ErrorInfo;
                m_bValid = false;
                nRet = -1;

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[400] ClassUIController::activateSoft, members:" + "m_bValid:" + m_bValid + ",m_strInvalidMessage:" + m_strInvalidMessage);
                }
            }


            if (!m_bVerEnterprise) // 个人版
            {
                if (m_bValid)
                {
                    if (qRet.Data1 != null)
                    {
                        // 有更新的UI data
                        decdUiData(qRet.Data1);

                        // update UI data file
                        updateUiDataFile(qRet.Data1, qRet.Date);
                    }
                    else
                    {
                        updateUiDataFile("", DateTime.Now.AddDays(-1));
                    }
                }
                else
                {
                    updateUiDataFile("", DateTime.Now.AddDays(-1));
                }
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[430] Exit ClassUIController::activateSoft");
            }

            return nRet;
        }


        // 到Lic Svr 签到
        public int loadUiPermTable()
        {
            //MessageBox.Show("10");

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[445] Enter ClassUIController::loadUiPermTable");
            }


            if (String.IsNullOrWhiteSpace(m_licSvrUrl))
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[453] Exit ClassUIController::loadUiPermTable");
                }

                m_bValid = false;
                return -1;
            }

            String strDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strUpdatedTag = strDir + @"\Compare\UpdateMark.txt";
            String strUiDataFile = strDir + m_strUiRecordDatFileName;
            
            Boolean bExistUpdateFlag = false, bNoExistLocalUiDataFile = false, bLocalUiDataFileTooOld = false;


            bExistUpdateFlag        = File.Exists(strUpdatedTag);
            bNoExistLocalUiDataFile = !File.Exists(strUiDataFile);
            
            if (File.Exists(strUiDataFile))
            {
                DateTime lastWriteTime = File.GetLastWriteTime(strUiDataFile);
                double dbDays = m_commTools.DateDiff(DateTime.Now, lastWriteTime);
                if (dbDays > 15.0)
                {
                    bLocalUiDataFileTooOld = true; // 
                }
            }

            QueryResult qRet = null;
            DateTime dt = DateTime.MinValue;

            m_dtExpireDate = dt;
            m_strVerName = Settings.Default.strVerName;

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[489] ClassUIController::loadUiPermTable,members:bExistUpdateFlag:" + bExistUpdateFlag + ",bNoExistLocalUiDataFile:" + bNoExistLocalUiDataFile + ",bLocalUiDataFileTooOld:" + bLocalUiDataFileTooOld);
            }

            if (m_bVerEnterprise)
            {
                try
                {
                    if (m_licSvrIntf != null)
                    {
                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[500] ClassUIController::loadUiPermTable");
                        }

                        qRet = m_licSvrIntf.SignForEntire(m_strMachineId, true);

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[507] ClassUIController::loadUiPermTable");
                        }
                    }
                    else
                    {
                        qRet = null;

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[516] ClassUIController::loadUiPermTable");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    qRet = null;

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[526] ClassUIController::loadUiPermTable,Exception:" + ex.ToString());
                    }
                }
                finally
                {
                }

                if (!m_bWithDocRepository) // 非文库版
                {
                    if (qRet == null)
                    {
                        // 从本地中取
                        // get from UI data file
                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[541] ClassUIController::loadUiPermTable");
                        }
                        int nRet = loadUiDataFromFile(ref dt);

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[547] ClassUIController::loadUiPermTable");
                        }
                        //MessageBox.Show("16");

                        //m_strVerName = Settings.Default.strVerName;

                        if (nRet != 0)
                        {
                            m_dtExpireDate = dt;
                            m_strInvalidMessage = "获取许可失败，请确保网络畅通";
                            m_bValid = false;

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[561] ClassUIController::loadUiPermTable");
                            }
                        }
                        else
                        {
                            m_dtExpireDate = dt;

                            // 判断是否过期超限
                            double dbDays = m_commTools.DateDiff(DateTime.Now, dt);
                            int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                            if (nDays < 0 && nDays >= -30 && (nDays % 5 == 0))
                            {
                                //MessageBox.Show("doc利器还余" + (-1 * nDays) + "天到期(" + dt.ToString("yyyy年M月d日") + ")", "提醒");
                            }

                            if (nDays > 0)
                            {
                                m_bValid = false;
                            }
                            else
                            {
                                m_bValid = true;
                                m_strInvalidMessage = "到期";
                            }

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[589] ClassUIController::loadUiPermTable");
                            }
                        }

                    }
                    else if (qRet.IsSuccess)
                    {
                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[598] ClassUIController::loadUiPermTable");
                        }

                        if (qRet.Data2 != null)
                        {
                            m_strVerName = qRet.Data2;
                            Settings.Default.strVerName = m_strVerName;
                            Settings.Default.Save();

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[609] ClassUIController::loadUiPermTable, Data2:" + qRet.Data2);
                            }

                        }

                        m_dtExpireDate = qRet.Date;

                        if (qRet.Data1 != null)
                        {
                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[620] ClassUIController::loadUiPermTable, Data1:" + qRet.Data1);
                            }

                            decdUiData(qRet.Data1);

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[627] ClassUIController::loadUiPermTable");
                            }

                            // update UI data file
                            if (bExistUpdateFlag || bNoExistLocalUiDataFile || bLocalUiDataFileTooOld)
                            {
                                //MessageBox.Show("19");
                                //MessageBox.Show(qRet.Data1);

                                if (ThisAddIn.m_bLog)
                                {
                                    Log.WriteLog("[638] ClassUIController::loadUiPermTable,updateUiDataFile paras:Data1:" + qRet.Data1 + ",ExpDate:" + qRet.Date.ToString());
                                }

                                int rt = updateUiDataFile(qRet.Data1, qRet.Date);
                                if (rt == 0)
                                {
                                    if (bExistUpdateFlag)
                                    {
                                        File.Delete(strUpdatedTag);
                                    }
                                }
                                //MessageBox.Show("20");
                            }
                            
                        }
                        else
                        {
                            m_bValid = false;
                        }
                    }
                    else
                    {
                        if (qRet.Data2 != null)
                        {
                            m_strVerName = qRet.Data2;
                            Settings.Default.strVerName = m_strVerName;
                            Settings.Default.Save();
                        }

                        m_dtExpireDate = qRet.Date;
                        m_strInvalidMessage = qRet.ErrorInfo;
                        m_bValid = false;

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[673] ClassUIController::loadUiPermTable,updateUiDataFile paras:Data2:" + qRet.Data2 + ",ExpDate:" + qRet.Date.ToString() + ",ErrInfo:" + m_strInvalidMessage);
                        }

                    }
                }

            }
            else // 个人版
            {
                // 优先从本地取
                // 从本地中取
                // get from UI data file
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[687] ClassUIController::loadUiPermTable");
                }

                int nRet = loadUiDataFromFile(ref dt);
                //MessageBox.Show("22");

                if (nRet != 0)
                {
                    // 本地没有则从网上取

                    try
                    {
                        if (m_licSvrIntf != null)
                        {
                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[703] ClassUIController::loadUiPermTable");
                            }

                            qRet = m_licSvrIntf.SignForPerson(m_strMachineId, (bExistUpdateFlag || bNoExistLocalUiDataFile || bLocalUiDataFileTooOld));
                            
                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[710] ClassUIController::loadUiPermTable");
                            }
                        }
                        else
                        {
                            qRet = null;
                            //MessageBox.Show("25");
                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[719] ClassUIController::loadUiPermTable");
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        qRet = null;
                        //MessageBox.Show("26," + ex.Message);

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[730] ClassUIController::loadUiPermTable, Exception:" + ex.ToString());
                        }
                    }
                    finally
                    {
                    }

                    // 网上失败则禁用
                    if (qRet == null)
                    {
                        m_dtExpireDate = dt;
                        m_strVerName = Settings.Default.strVerName;
                        m_strInvalidMessage = "获取许可失败，请确保网络畅通";
                        m_bValid = false;

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[747] ClassUIController::loadUiPermTable");
                        }

                    }
                    else if (qRet.IsSuccess)
                    {
                        if (qRet.Data3 != null)
                        {
                            String[] strs = qRet.Data3.Split('_');

                            if (strs != null && strs.GetLength(0) >= 2)
                            {
                                m_strAccount = strs[0];
                                m_strActiveSn = strs[1];
                            }

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[765] ClassUIController::loadUiPermTable, Data3:" + qRet.Data3);
                            }

                        }

                        if (qRet.Data2 != null)
                        {
                            m_strVerName = qRet.Data2;
                            Settings.Default.strVerName = m_strVerName;
                            Settings.Default.Save();

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[778] ClassUIController::loadUiPermTable, Data2:" + qRet.Data2);
                            }
                        }

                        m_dtExpireDate = qRet.Date;

                        m_strUiData = qRet.Data1;

                        // 判断是否过期超限
                        double dbDays = m_commTools.DateDiff(DateTime.Now, qRet.Date);
                        int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                        if (nDays < 0 && nDays >= -30 && (nDays % 5 == 0))
                        {
                            //MessageBox.Show("doc利器还余" + (-1 * nDays) + "天到期(" + qRet.Date.ToString("yyyy年M月d日") + ")", "提醒");
                        }

                        if (nDays > 0)
                        {
                            m_bValid = false;
                        }

                        if (qRet.Data1 != null)
                        {
                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[804] ClassUIController::loadUiPermTable, ExpDate:" + m_dtExpireDate.ToString() + ",Data1:" + qRet.Data1);
                            }

                            decdUiData(qRet.Data1);

                            // update UI data file
                            if (bExistUpdateFlag || bNoExistLocalUiDataFile || bLocalUiDataFileTooOld)
                            {
                                if (ThisAddIn.m_bLog)
                                {
                                    Log.WriteLog("[814] ClassUIController::loadUiPermTable");
                                }

                                int rt = updateUiDataFile(qRet.Data1, qRet.Date);
                                if (rt == 0)
                                {
                                    if (bExistUpdateFlag)
                                    {
                                        File.Delete(strUpdatedTag);
                                        if (ThisAddIn.m_bLog)
                                        {
                                            Log.WriteLog("[825] ClassUIController::loadUiPermTable");
                                        }
                                    }
                                }
                                //MessageBox.Show("30");
                            }
                        }
                        else
                        {
                            m_bValid = false;
                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[837] ClassUIController::loadUiPermTable");
                            }
                        }
                    }
                    else
                    {
                        if (qRet.Data2 != null)
                        {
                            m_strVerName = qRet.Data2;
                            Settings.Default.strVerName = m_strVerName;
                            Settings.Default.Save();

                            if (ThisAddIn.m_bLog)
                            {
                                Log.WriteLog("[851] ClassUIController::loadUiPermTable, Data2:" + qRet.Data2);
                            }
                        }

                        m_dtExpireDate = qRet.Date;
                        m_strInvalidMessage = qRet.ErrorInfo;
                        m_bValid = false;

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[861] ClassUIController::loadUiPermTable, memebers:ExpDate:" + m_dtExpireDate.ToString() + ",ErrInfo:" + qRet.ErrorInfo);
                        }

                    }
                }
                else
                {
                    // 判断是否过期超限
                    double dbDays = m_commTools.DateDiff(DateTime.Now, dt);
                    int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                    if (nDays < 0 && nDays >= -30 && (nDays % 5 == 0))
                    {
                        //MessageBox.Show("doc利器还余" + (-1 * nDays) + "天到期(" + dt.ToString("yyyy年M月d日") + ")", "提醒");
                    }

                    m_dtExpireDate = dt;
                    m_strVerName = Settings.Default.strVerName;

                    if (nDays > 0)
                    {
                        m_strInvalidMessage = "到期";
                        m_bValid = false;
                    }
                    else
                    {
                        m_bValid = true;
                    }

                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[892] ClassUIController::loadUiPermTable");
                    }

                }
               
            }

            //MessageBox.Show("31");

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[903] Exit ClassUIController::loadUiPermTable");
            }

            return 0;
        }


        private int decdUiData(String strMD5UiData)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[914] Enter ClassUIController::decdUiData");
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[920] ClassUIController::decdUiData, UI MD5:" + strMD5UiData);
            }


            if (String.IsNullOrWhiteSpace(strMD5UiData))
            {
                return -1;
            }

            int nLen = strMD5UiData.Length;

            if ((nLen % 32) != 0)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[935] Exit ClassUIController::decdUiData, UI MD5 Len:" + nLen);
                }

                return -2;
            }

            int nCnt = (nLen / 32);

            String strItem = "";

            m_hashUiMD5Items.Clear();

            for (int i = 0; i < nCnt; i++)
            {
                strItem = strMD5UiData.Substring(i*32,32);

                m_hashUiMD5Items[strItem] = 1;
            }


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[957] Exit ClassUIController::decdUiData");
            }

            return 0;
        }


        private String DateTime2StrMD5(DateTime dt)
        {
            String strDt = dt.ToString("yyyyMMdd");
            char[] chs = strDt.ToCharArray();

            String strMD5 = "";
            int nVal = 0;

            for (int i = 0; i < 8; i++)
            {
                if (int.TryParse(chs[i].ToString(), out nVal))
                {
                    strMD5 += (String)m_arrNum2MD5[nVal];
                }
            }

            return strMD5;
        }


        private int StrMD5toDateTime(String strMD5, ref DateTime dt)
        {
            if ((strMD5.Length % 32) != 0)
            {
                return -1;
            }

            int nCnt = (strMD5.Length / 32);

            if (nCnt != 8)
            {
                return -2;
            }


            String strMD5Item = "", strItem = "";
            String strDt = "";

            for (int i = 0; i < nCnt; i++)
            {
                strMD5Item = strMD5.Substring(i * 32, 32);

                if (m_hashMD5toNum.Contains(strMD5Item))
                {
                    strItem = (String)m_hashMD5toNum[strMD5Item];
                    strDt += strItem;
                }
                else
                {
                    return -3;
                }
            }


            if (DateTime.TryParseExact(strDt,
                                   "yyyyMMdd",
                                   System.Globalization.CultureInfo.InvariantCulture,
                                   System.Globalization.DateTimeStyles.None,
                                   out dt))
            {

            }
            else
            {
                return -4;
            }

            return 0;
        }




        private int updateUiDataFile(String strMD5UiData, DateTime dtExp)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1041] Enter ClassUIController::updateUiDataFile");
            }

            String strDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strUiDataFile = strDir + m_strUiRecordDatFileName;
            //String strUiDataFile = m_strUiRecFilePath + m_strUiRecordDatFileName;

            String strWriteMD5UiDate = strMD5UiData;

            String strMD5DateTime = DateTime2StrMD5(dtExp);
            String strMD5Year = strMD5DateTime.Substring(0, 4 * 32);
            String strMD5MonDay = strMD5DateTime.Substring(4 * 32,4*32);


            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1058] ClassUIController::updateUiDataFile, dtExp:" + dtExp.ToString() + "MD5(Year):" + strMD5Year + ",MD5(MonDay):" + strMD5MonDay);
            }

            if (String.IsNullOrWhiteSpace(strMD5UiData))
            {
                strWriteMD5UiDate = "";
            }

            int nLen = strMD5UiData.Length;

            if ((nLen % 32) != 0)
            {
                strWriteMD5UiDate = "";
            }

            nLen = strWriteMD5UiDate.Length;

            int nCnt = (nLen / 32);

            String strItem = "";

            ArrayList arrItems = new ArrayList();

            for (int i = 0; i < nCnt; i++)
            {
                strItem = strWriteMD5UiDate.Substring(i * 32, 32);
                arrItems.Add(strItem);
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1089] ClassUIController::updateUiDataFile");
            }
            //MessageBox.Show("100");

            StreamWriter sw = null;
            
            try
            {
            	sw = new StreamWriter(strUiDataFile);

                sw.Write(strMD5Year);
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1098] ClassUIController::updateUiDataFile, Write: MD5(y)" + strMD5Year);
                }

                int nPos = Math.Min(2, arrItems.Count);

                for (int i = 0; i < nPos; i++)
                {
                    strItem = (String)arrItems[i];
                    sw.Write(strItem);
                }
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1114] ClassUIController::updateUiDataFile, Write: MD5(UI)");
                }

                sw.Write(m_strMD5MachineId);
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1120] ClassUIController::updateUiDataFile, Write: MachineID(MD5):" + m_strMD5MachineId);
                }

                for (int i = nPos; i < arrItems.Count; i++)
                {
                    strItem = (String)arrItems[i];
                    sw.Write(strItem);
                }
                
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1131] ClassUIController::updateUiDataFile, Write: MD5(UI)");
                }

                //MessageBox.Show("105");
                sw.Write(strMD5MonDay);
                
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1139] ClassUIController::updateUiDataFile, Write: MD5(md)" + strMD5MonDay);
                }

                //MessageBox.Show("106");
                sw.Close();
                //MessageBox.Show("107");
            }
            catch (System.Exception ex)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1150] Enter ClassUIController::updateUiDataFile, Exception:" + ex.ToString());
                }

                MessageBox.Show(ex.Message);
                return -1;
            }
            finally
            {
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1162] Exit ClassUIController::updateUiDataFile");
            }

            return 0;
        }


        private int loadUiDataFromFile(ref DateTime dt, Boolean b2Validate = true)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1173] Enter ClassUIController::loadUiDataFromFile");
            }

            String strDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strUiDataFile = strDir + m_strUiRecordDatFileName;
            // String strUiDataFile = m_strUiRecFilePath + m_strUiRecordDatFileName;

            m_hashUiMD5Items.Clear();

            if (!File.Exists(strUiDataFile))
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1187] Exit ClassUIController::loadUiDataFromFile");
                }

                return -1;
            }

            String strCnt = "";

            try
            {
	            StreamReader rd = new StreamReader(strUiDataFile);
	            strCnt = rd.ReadToEnd();
	            rd.Close();

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1203] ClassUIController::loadUiDataFromFile");
                }

            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
            }

            int nLen = strCnt.Length;

            if ((nLen % 32) != 0)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1221] Exit ClassUIController::loadUiDataFromFile, Content:" + strCnt + ",Len:" + nLen);
                }

                return -1;
            }

            int nCnt = (nLen / 32);

            if (nCnt < 9)
            {
                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1233] Exit ClassUIController::loadUiDataFromFile, Content:" + strCnt + ",Len:" + nLen);
                }

                return -1;
            }

            String strItem = "";
            ArrayList arrItems = new ArrayList();

            for (int i = 0; i < nCnt; i++)
            {
                strItem = strCnt.Substring(i * 32, 32);

                arrItems.Add(strItem);
            }

            String strMD5Year = "", strMD5MonDay = "";
            for (int i = 0; i < 4; i++)
            {
                strMD5Year += (String)arrItems[i];
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1257] ClassUIController::loadUiDataFromFile, MD(y):" + strMD5Year);
            }

            for (int i = arrItems.Count - 4; i < arrItems.Count; i++)
            {
                strMD5MonDay += (String)arrItems[i];
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1267] ClassUIController::loadUiDataFromFile, MD(md):" + strMD5MonDay);
            }


            for (int i = 0; i < 4; i++)
            {
                arrItems.RemoveAt(0);
            }

            for (int i = 0; i < 4; i++)
            {
                arrItems.RemoveAt(arrItems.Count - 1);
            }

            // get machine id
            String strMD5MachineID = "";
            int nMachineIdPos = Math.Min(2, arrItems.Count);

            if (b2Validate)
            {
                strMD5MachineID = (String)arrItems[nMachineIdPos];

                if (!strMD5MachineID.Equals(m_strMD5MachineId))
                {
                    if (ThisAddIn.m_bLog)
                    {
                        Log.WriteLog("[1293] Exit ClassUIController::loadUiDataFromFile, FileMachineID:" + strMD5MachineID + ",MachineID:" + m_strMD5MachineId);
                    }

                    return -1;
                }
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1302] ClassUIController::loadUiDataFromFile, FileMachineID:" + strMD5MachineID);
            }

            String strMD5DateTime = strMD5Year + strMD5MonDay;
            int nRet = StrMD5toDateTime(strMD5DateTime,ref dt);

            if (nRet != 0)
            {
                return nRet;
            }
            
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1315] ClassUIController::loadUiDataFromFile, Dt:" + dt.ToString());
            }

            arrItems.RemoveAt(nMachineIdPos);
            for (int i = 0; i < arrItems.Count; i++)
            {
                strItem = (String)arrItems[i];
                m_hashUiMD5Items[strItem] = 1;
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1327] Exit ClassUIController::loadUiDataFromFile");
            }

            return 0;
        }


        public int addExceptionalUiItem(String strCtrlItemName)
        {
            m_hashExceptionalUiItems[strCtrlItemName] = 1;
            m_hashDefaultUiItems[strCtrlItemName] = 1;

            return 0;
        }


        public int searchPermission(String strCtrlName)
        {
            if (!m_bValid)
            {
                if (m_hashExceptionalUiItems.Contains(strCtrlName))
                {
                    return 1;
                }

                return 0;
            }

            // private Hashtable m_hashUiItems = new Hashtable();    // pure NAME
            // private Hashtable m_hashUiMD5Items = new Hashtable(); // NAME MD5ized
            // private Hashtable m_hashExceptionalUiItems = new Hashtable(); // pure NAME


            // with doc repository
            // search hash (not MD5)
            int nVal = 0;

            if (m_bWithDocRepository)
            {
                if (m_hashUiItems.Contains(strCtrlName))
                {
                    nVal = 1;
                }
            }
            else
            {
                String strMD5Name = ClassEncryptUtils.MD5Encrypt(strCtrlName);
                strMD5Name = strMD5Name.ToUpper();
                if (m_hashUiMD5Items.Contains(strMD5Name))
                {
                    nVal = 1;
                }
            }

            return nVal;
        }


        public int updateUI(Hashtable hashControls)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1390] Enter ClassUIController::updateUI");
            }

            String strName = "";
            int nVal = 0;
            Control ctrl = null;
            ToolStripItem item = null;
            RibbonControl ribCtrl = null;
            Object ctrlObj = null;

            if (!m_bValid)
            {
                foreach (DictionaryEntry ent in hashControls)
                {
                    strName = (String)ent.Key;
                    ctrlObj = (Object)ent.Value;

                    nVal = 0;

                    if(m_hashExceptionalUiItems.Contains(strName))
                    {
                        nVal = 1;
                    }
#if ADMIN
                    nVal = 1;
#endif

                    if (ctrlObj != null)
                    {
                        if (ctrlObj is Control)
                        {
                            ctrl = (Control)ctrlObj;
                            if (ctrl.Controls.Count > 0)
                            {
                                ctrl.Enabled = true;
                            }
                            else
                            {
                                ctrl.Enabled = (nVal != 0);
                            }
                        }
                        else if (ctrlObj is ToolStripItem)
                        {
                            item = (ToolStripItem)ctrlObj;
                            item.Enabled = (nVal != 0);
                        }
                        else if (ctrlObj is RibbonGroup)
                        {
                            // 
                        }
                        else if (ctrlObj is RibbonControl)
                        {
                            ribCtrl = (RibbonControl)ctrlObj;
                            ribCtrl.Enabled = (nVal != 0);
                        }
                    }// if
                 }

                if (ThisAddIn.m_bLog)
                {
                    Log.WriteLog("[1449] Exit ClassUIController::updateUI");
                }

                return 0;
            }


            // with doc repository
            // search hash (not MD5)

            if (m_bWithDocRepository)
            {
                foreach (DictionaryEntry ent in hashControls)
                {
                    strName = (String)ent.Key;
                    ctrlObj = (Object)ent.Value;

                    nVal = 0;
                    if (m_hashUiItems.Contains(strName))
                    {
                        nVal = (int)m_hashUiItems[strName];
                    }
                    else
                    {
                        if (m_hashExceptionalUiItems.Contains(strName))
                        {
                            nVal = 1;
                        }
                    }
#if ADMIN
                    nVal = 1;
#endif
                    if (ctrlObj != null)
                    {
                        if (ctrlObj is Control)
                        {
                            ctrl = (Control)ctrlObj;
                            if (ctrl.Controls.Count > 0)
                            {
                                ctrl.Enabled = true;
                            }
                            else
                            {
                                ctrl.Enabled = (nVal != 0);
                            }
                        }
                        else if (ctrlObj is ToolStripItem)
                        {
                            item = (ToolStripItem)ctrlObj;
                            item.Enabled = (nVal != 0);
                        }
                        else if (ctrlObj is RibbonGroup)
                        {
                            // 
                        }
                        else if (ctrlObj is RibbonControl)
                        {
                            ribCtrl = (RibbonControl)ctrlObj;
                            ribCtrl.Enabled = (nVal != 0);
                        }
                    }// if
                }
            }
            else
            {
                String strMD5Name = "";

                foreach (DictionaryEntry ent in hashControls)
                {
                    strName = (String)ent.Key;
                    ctrlObj = (Object)ent.Value;

                    nVal = 0;

                    if (m_hashUiName2MD5Name.Contains(strName))
                    {
                        strMD5Name = (String)m_hashUiName2MD5Name[strName];
                    }
                    else
                    {
                        strMD5Name = ClassEncryptUtils.MD5Encrypt(strName);
                        strMD5Name = strMD5Name.ToUpper();
                        m_hashUiName2MD5Name[strName] = strMD5Name;
                    }

                    if (m_hashExceptionalUiItems.Contains(strName))
                    {
                        nVal = 1;
                    }
                    else
                    {
                        if (m_hashUiMD5Items.Contains(strMD5Name))
                        {
                            nVal = (int)m_hashUiMD5Items[strMD5Name];
                        }
                    }
#if ADMIN
                    nVal = 1;
#endif
                    if (ctrlObj != null)
                    {
                        if (ctrlObj is Control)
                        {
                            ctrl = (Control)ctrlObj;
                            if (ctrl.Controls.Count > 0)
                            {
                                ctrl.Enabled = true;
                            }
                            else
                            {
                                ctrl.Enabled = (nVal != 0);
                            }
                        }
                        else if (ctrlObj is ToolStripItem)
                        {
                            item = (ToolStripItem)ctrlObj;
                            item.Enabled = (nVal != 0);
                        }
                        else if (ctrlObj is RibbonGroup)
                        {
                            // 
                        }
                        else if (ctrlObj is RibbonControl)
                        {
                            ribCtrl = (RibbonControl)ctrlObj;
                            ribCtrl.Enabled = (nVal != 0);
                        }
                    }// if

                }
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1583] Exit ClassUIController::updateUI");
            }

            return 0;
        }



        public String getMachineID()
        {
            return m_strMachineId;
        }


        public String getMD5MachineID()
        {
            return m_strMD5MachineId;
        }


        public void setDocRepUiPermHash(Hashtable oHash)
        {
            m_hashUiItems = oHash;
            return;
        }

        public void restoreDefaultDocRepUiPermHash()
        {
            m_hashUiItems = m_hashDefaultUiItems;
            return;
        }

        public String checkUpdate()
        {
            String strRet = "";

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1621] Enter ClassUIController::checkUpdate");
            }

            // 检查更新
            //MessageBox.Show(m_updateSvrUrl);

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1629] ClassUIController::checkUpdate, members:m_updateSvrUrl:" + m_updateSvrUrl + ",m_licSvrUrl:" + m_licSvrUrl);
            }

            if (!String.IsNullOrWhiteSpace(m_updateSvrUrl) && m_updateSvrUrl.Equals(m_licSvrUrl))
            {
                //MessageBox.Show("2");

                if (m_licSvrIntf != null)
                {
                    try
                    {
                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[1642] ClassUIController::checkUpdate");
                        }
                    	strRet = m_licSvrIntf.CheckUpdate();

                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[1648] ClassUIController::checkUpdate");
                        }

                        //MessageBox.Show("4");
                    }
                    catch (System.Exception ex)
                    {
                        if (ThisAddIn.m_bLog)
                        {
                            Log.WriteLog("[1657] ClassUIController::checkUpdate, Exception:" + ex.ToString());
                        }

                        MessageBox.Show(ex.Message);
                        strRet = "-3";
                    }
                    finally
                    {
                    }
                }
            }

            //MessageBox.Show("5,RetMessage:" + strRet);
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[1672] Exit ClassUIController::checkUpdate");
            }

            return strRet;
        }


    }
}
