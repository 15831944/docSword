using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using AutoUpdate;
using OfficeTools.Common;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OfficeAssist.Properties;
using System.Collections.Specialized;


namespace OfficeAssist
{
    public class classMultiEditionCenter
    {
        public class clPermissionItem
        {
            public Boolean   bOpen = true;
            public DateTime  expDate = DateTime.MinValue;
            public Hashtable hashPermission = new Hashtable();  // control name(MD5), val(1--enable,0--disable)
            public String    strEditionName = "";
            public String    strStatus = "";                    // 状态：启用、禁用、激活/未激活、超期（到期日）、超限
            public String    strErrInfo = "";
            public String    strAccount = "";
            public String    strActCode = "";


            public String    cfgAutoUpdateSvrUrl = ""; // Lic Svr
            public String    cfgCerSvrUrl = "";

            public String    cfgDocRepositoryUrl = ""; // 文库 Svr
            public String    cfgLoginUrl = "";
            public String    cfgUploadFileUrl = "";

        }

        private ClassOfficeCommon m_commTools = null;

        private Hashtable m_hashExceptional = new Hashtable();
        private Hashtable m_hashPermTables = new Hashtable();

        private Hashtable m_hashStr2MD5 = new Hashtable();
        private Hashtable m_hashMD5toStr = new Hashtable();


        private String m_strAppDir = "";
        //public String m_strMachineId = "", m_strMD5MachineId = "", m_strMD5MD5MachineId = "";
        //private readonly String m_strSpecialMachineId = "AFBECD12AFBECD12AFBECD12AFBECD12";

        //private Hashtable m_hashMD5toNum = new Hashtable();
        //private ArrayList m_arrNum2MD5 = new ArrayList();


        // private AutoUpdate.AutoUpdateClass m_licSvrIntf = null;


        // 单机版（脱机）
        public readonly String m_strSoloEditionName = "单机版";
        // lic
        private readonly String m_strLicSoloEdition = @"slUiRec.txt";


        // 个人版（互联网，激活注册）
        public readonly String m_strPrivEditionName = "个人版";
        private readonly String m_strCfgPrivEdition = @"\config\pvCfg.xml";
        private readonly String m_strCfgHidePrivEdition = @"pvSysUrl.txt";
        // lic
        private readonly String m_strLicPrivEdition = @"pvUiRec.txt";


        // 企业版
        public readonly String m_strEntEditionName = "企业版";
        private readonly String m_strCfgEntEdition = @"\config\etCfg.xml";
        // lic
        private readonly String m_strLicEntEdition = @"etUiRec.txt";

        // 文库版
        public readonly String m_strDocRepositoryEditionName = "文库版";
        private readonly String m_strCfgDocRepoEdition = @"\config\drCfg.xml";


        public classMultiEditionCenter(ClassOfficeCommon mnTools)
        {
            m_commTools = mnTools;

            // 
            //ClassHardInfo clsHardInfo = new ClassHardInfo();
            //String strCpuId = clsHardInfo.GetCpuID();

            //m_strMachineId = strCpuId;// +strMacAddr;
            //m_strMachineId = m_strMachineId.ToUpper();

            //m_strMD5MachineId = ClassEncryptUtils.MD5Encrypt(m_strMachineId);
            //m_strMD5MachineId = m_strMD5MachineId.ToUpper();

            //m_strMD5MD5MachineId = ClassEncryptUtils.MD5Encrypt(m_strMD5MachineId);
            //m_strMD5MD5MachineId = m_strMD5MD5MachineId.ToUpper();


            // 
            m_strAppDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;


            // 
            //String strMD5 = "";

            //for (int i = 0; i < 10; i++)
            //{
            //    strMD5 = ClassEncryptUtils.MD5Encrypt("li" + i.ToString() + "dong");
            //    strMD5 = strMD5.ToUpper();

            //    m_arrNum2MD5.Add(strMD5);
            //    m_hashMD5toNum[strMD5] = i.ToString();
            //}


            return;
        }



        // 函数,添加例外的名称
        public void AddExceptionalName(String strName)
        {
            //String strNameMD5 = ClassEncryptUtils.MD5Encrypt(strName);
            //strNameMD5 = strNameMD5.ToUpper();

            // m_hashExceptional[strNameMD5] = strName;
            m_hashExceptional[strName] = strName;

            return;
        }


        public void AddMD5Str(String strItem)
        {
            String strMD5 = ClassEncryptUtils.MD5Encrypt(strItem);
            strMD5 = strMD5.ToUpper();

            m_hashStr2MD5[strItem] = strMD5;
            m_hashMD5toStr[strMD5] = strItem;

            return;
        }


/*
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

            DateTime tmpDt = DateTime.MinValue;

            if (DateTime.TryParseExact(strDt,
                                   "yyyyMMdd",
                                   System.Globalization.CultureInfo.InvariantCulture,
                                   System.Globalization.DateTimeStyles.None,
                                   out tmpDt))
            {
                dt = tmpDt.AddHours(23).AddMinutes(59).AddSeconds(59);
            }
            else
            {
                return -4;
            }

            return 0;
        }*/


/*
        private int DecodeLic(String strCnt,ref clPermissionItem permItem, String strLicFile = "")
        {
            int nRet = -1;

            if((strCnt.Length % 32) != 0)
            {
                // LOG

                return nRet;
            }

            int nCnt = (strCnt.Length / 32);

            if (nCnt < 10)
            {
                // LOG

                return nRet;
            }

            
            // CRC 1
            String strCRC = strCnt.Substring(0, 32);

            // machine id (MD5 MD5) 1
            String strMacineIdMD5MD5 = strCnt.Substring(32, 32);

            // Expire date(yyyyMMdd) 8
            String strExpDt = strCnt.Substring(64, 8*32);

            // UI data(MD5, Val = 1)
            String strUI = "";

            if (strCnt.Length > 320)
            {
                strUI = strCnt.Substring(320);
            }

            // 

            String strCrcCnt = ClassEncryptUtils.MD5Encrypt(strMacineIdMD5MD5 + strExpDt + strUI);
            strCrcCnt = strCrcCnt.ToUpper();


            // CRC
            if (!strCrcCnt.Equals(strCRC))
            {
                // LOG

                return nRet;
            }


            Boolean bSpecialMachineId = false;

            // machine id validate
            if (strMacineIdMD5MD5.Equals(m_strSpecialMachineId))
            {
                // left
                bSpecialMachineId = true;
            }
            else
            {
                // 
                if (!strMacineIdMD5MD5.Equals(m_strMD5MD5MachineId))
                {
                    // LOG

                    return nRet;
                }
            }


            if (StrMD5toDateTime(strExpDt,ref permItem.expDate) != 0)
            {
                return nRet;
            }


            int nUiCnt = (strUI.Length / 32);
            String strItem = "";

            for (int i = 0; i < nUiCnt; i++)
            {
                strItem = strUI.Substring(i*32, 32);
                permItem.hashPermission[strItem] = (int)1;
            }

            permItem.bOpen = true;


            if (bSpecialMachineId && !String.IsNullOrWhiteSpace(strLicFile))
            {
                // rewrite
                StreamWriter sw = new StreamWriter(strLicFile);

                // machine id (MD5 MD5) 1
                strMacineIdMD5MD5 = m_strMD5MD5MachineId;

                strCRC = ClassEncryptUtils.MD5Encrypt(strMacineIdMD5MD5 + strExpDt + strUI);
                strCRC = strCRC.ToUpper();

                sw.Write(strCRC + strMacineIdMD5MD5 + strExpDt + strUI);

                sw.Flush();
                sw.Close();

            }

            return 0;
        }*/


/*
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
        }*/


/*
        private int EncodeLic(DateTime expDt,String strUI, ref String strEncodedCnt, Boolean bSpecialMachineID = false)
        {
            // CRC 1
            String strCRC = "";

            // machine id (MD5 MD5) 1
            String strMacineIdMD5MD5 = m_strMD5MD5MachineId;

            if (bSpecialMachineID)
            {
                strMacineIdMD5MD5 = m_strSpecialMachineId;
            }

            // Expire date(yyyyMMdd) 8
            String strExpDt = "";

            // UI data(MD5, Val = 1)

            strExpDt = DateTime2StrMD5(expDt);

            strCRC = ClassEncryptUtils.MD5Encrypt(strMacineIdMD5MD5 + strExpDt + strUI);
            strCRC = strCRC.ToUpper();

            strEncodedCnt = strCRC + strMacineIdMD5MD5 + strExpDt + strUI;

            return 0;
        }*/


/*
        private int EncodeLic(clPermissionItem permItem, ref String strEncodedCnt,Boolean bSpecialMachineID= false)
        {
            // CRC 1
            String strCRC = "";

            // machine id (MD5 MD5) 1
            String strMacineIdMD5MD5 = m_strMD5MD5MachineId;

            if(bSpecialMachineID)
            {
                strMacineIdMD5MD5 = m_strSpecialMachineId;
            }

            // Expire date(yyyyMMdd) 8
            String strExpDt = "";

            // UI data(MD5, Val = 1)
            String strUI = "";


            strExpDt = DateTime2StrMD5(permItem.expDate);

            String strName = "";
            int nVal = 0;

            foreach(DictionaryEntry ent in permItem.hashPermission)
            {
                strName = (String)ent.Key;
                nVal = (int)ent.Value;

                if (nVal > 0)
                {
                    strUI += strName;
                }
            }

            strCRC = ClassEncryptUtils.MD5Encrypt(strMacineIdMD5MD5 + strExpDt + strUI);
            strCRC = strCRC.ToUpper();

            strEncodedCnt = strCRC + strMacineIdMD5MD5 + strExpDt + strUI;

            return 0;
        }*/

        public Boolean IsExistCurSoloLic()
        {
            String strLicFile = m_strAppDir + m_strLicSoloEdition;

            Boolean bRet = File.Exists(strLicFile);

            return bRet;
        }


        public int BackupCurSoloLic(String strSavedLicFile)
        {
            String strLicFile = m_strAppDir + m_strLicSoloEdition;

            if (!File.Exists(strLicFile))
            {
                return -1;
            }

            if (File.Exists(strSavedLicFile))
            {
                try
                {
                	File.Delete(strSavedLicFile);
                }
                catch (System.Exception ex)
                {
                    return -2;
                }
                finally
                {
                }
            }

            try
            {
                File.Copy(strLicFile, strSavedLicFile);
            }
            catch (System.Exception ex)
            {
                return -3;
            }
            finally
            {
            }

            return 0;
        }


        public int LoadSoloLic(String strNewLicFile)
        {
            if (!File.Exists(strNewLicFile))
            {
                return -1;
            }

            String strLicFile = m_strAppDir + m_strLicSoloEdition;

            if (File.Exists(strLicFile))
            {
                try
                {
                	File.Delete(strLicFile);
                }
                catch (System.Exception ex)
                {
                    return -2;
                }
                finally
                {
                }
            }

            try
            {
                File.Copy(strNewLicFile, strLicFile);
            }
            catch (System.Exception ex)
            {
                return -3;
            }
            finally
            {
            }

            BuildSoloEdition();
            CheckValid();

            return 0;
        }


        public int BuildSoloEdition()
        {
            int nRet = -1;

            String strLicFile = m_strAppDir + m_strLicSoloEdition;

            if (!File.Exists(strLicFile))
            {
                return -1;
            }

            // read
            String strCnt = "";

            try
            {
                StreamReader rd = new StreamReader(strLicFile);
                strCnt = rd.ReadToEnd();
                rd.Close();
            }
            catch (System.Exception ex)
            {
                strCnt = "";
            }
            finally
            {
            }

            if (String.IsNullOrWhiteSpace(strCnt))
            {
                return -1;
            }

            // validate
            clPermissionItem permItem = new clPermissionItem();
            String strMachineID2 = "";
            nRet = m_commTools.DecodeLic(strCnt,ref strMachineID2,ref permItem.hashPermission,ref permItem.expDate,true, strLicFile);

            if (nRet != 0)
            {
                // LOG
                return nRet;
            }

            // add this edition perms
            permItem.bOpen = true;
            permItem.strStatus = "启用";
            m_hashPermTables[m_strSoloEditionName] = permItem;

            return 0;
        }


        private int BuildPrivateEdition()
        {
            int nRet = -1;

            String strCfgFile = m_strAppDir + m_strCfgPrivEdition;
            String strLicFile = m_strAppDir + m_strLicPrivEdition;

            clPermissionItem permItem = new clPermissionItem();
            String strCfgAutoUpdateSvrUrl = "";
            String strCfgCerSvrUrl = "";

            // if config exist, read config to get URLs
            Boolean bGetLicData = false;

            if (File.Exists(strCfgFile))
            {
                ConfigReader cfgReader = new ConfigReader();
                Hashtable cfgNameValues = cfgReader.getConfigItems(strCfgFile);

                if (cfgNameValues.Contains("cfgAutoUpdateSvrUrl"))
                {
                    strCfgAutoUpdateSvrUrl = (String)cfgNameValues["cfgAutoUpdateSvrUrl"];
                }

                if (cfgNameValues.Contains("cfgCerSvrUrl"))
                {
                    strCfgCerSvrUrl = (String)cfgNameValues["cfgCerSvrUrl"];
                }


                String strPrivHideUrlFile = m_strAppDir + m_strCfgHidePrivEdition;

                String strUrl = "";

                if (String.IsNullOrWhiteSpace(strCfgCerSvrUrl) && File.Exists(strPrivHideUrlFile))
                {
                    try
                    {
                        StreamReader rd = new StreamReader(strPrivHideUrlFile);
                        strUrl = rd.ReadToEnd();
                        rd.Close();

                        strCfgCerSvrUrl = strUrl;
                        strCfgAutoUpdateSvrUrl = strUrl;
                    }
                    catch (System.Exception ex)
                    {
                        strUrl = "";
                    }
                    finally
                    {
                    }
                }


                // access URL to get lic data
                if (!String.IsNullOrWhiteSpace(strCfgCerSvrUrl))
                {
                    AutoUpdateClass licSvrIntf = null;// new AutoUpdateClass(strCfgCerSvrUrl);

                    try
                    {
                        licSvrIntf = new AutoUpdateClass(strCfgCerSvrUrl);
                    }
                    catch (System.Exception ex)
                    {
                        // LOG
                        licSvrIntf = null;
                    }
                    finally
                    {
                    }

                    if (licSvrIntf != null)
                    {
                        QueryResult qRet = null;

                        try
                        {
                            qRet = licSvrIntf.SignForPerson(m_commTools.MachineId, true);

                            // parse 
                            if (qRet != null)
                            {
                                if (qRet.Data2 != null)
                                {
                                    permItem.strEditionName = qRet.Data2;

                                    Settings.Default.strVerName = qRet.Data2;
                                    Settings.Default.Save();
                                }

                                if (qRet.IsSuccess)
                                {
                                    permItem.expDate = qRet.Date;

                                    if (qRet.Data1 != null)
                                    {
                                        int nCnt = (qRet.Data1.Length / 32);
                                        String strItem = "";
                                        int nVal = 1;
                                        for (int i = 0; i < nCnt; i++)
                                        {
                                            strItem = qRet.Data1.Substring(i * 32, 32);
                                            permItem.hashPermission[strItem] = nVal;
                                        }
                                    }

                                    if (qRet.Data3 != null)
                                    {
                                        String[] strs = qRet.Data3.Split('_');

                                        if (strs != null && strs.GetLength(0) >= 2)
                                        {
                                            permItem.strAccount = strs[0];
                                            permItem.strActCode = strs[1];

                                            Settings.Default.strRegAcnt = strs[0];
                                            Settings.Default.strRegActSn = strs[1];
                                            Settings.Default.Save();
                                        }
                                        // LOG
                                    }

                                    // write into lic file
                                    //String strLicFile = m_strAppDir + m_strLicPrivEdition;

                                    String strEncodeLic = "";
                                    // int nrt = EncodeLic(permItem, ref strEncodeLic);
                                    int nrt = m_commTools.EncodeLic(m_commTools.MachineIdMD5MD5, permItem.expDate, permItem.hashPermission, ref strEncodeLic);
                                    if (nrt == 0)
                                    {
                                        StreamWriter sw = new StreamWriter(strLicFile);
                                        sw.Write(strEncodeLic);
                                        sw.Close();
                                    }
                                }
                                else
                                {
                                    permItem.strErrInfo = qRet.ErrorInfo;
                                }

                                permItem.bOpen = qRet.IsSuccess;
                                // permItem.strStatus = "启用";
                                permItem.cfgAutoUpdateSvrUrl = strCfgAutoUpdateSvrUrl;
                                permItem.cfgCerSvrUrl = strCfgCerSvrUrl;

                                m_hashPermTables[m_strPrivEditionName] = permItem;

                                bGetLicData = true;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            qRet = null;
                            // LOG
                        }
                        finally
                        {

                        }
                    }
                }
            }


            if (!bGetLicData && File.Exists(strLicFile))
            {
                // read
                String strCnt = "";

                try
                {
                    StreamReader rd = new StreamReader(strLicFile);
                    strCnt = rd.ReadToEnd();
                    rd.Close();
                }
                catch (System.Exception ex)
                {
                    strCnt = "";
                }
                finally
                {

                }

                if (!String.IsNullOrWhiteSpace(strCnt))
                {
                    // validate
                    // nRet = DecodeLic(strCnt, ref permItem);
                    String strMachineID2 = "";
                    nRet = m_commTools.DecodeLic(strCnt,ref strMachineID2,ref permItem.hashPermission, ref permItem.expDate);

                    if (nRet == 0)
                    {
                        // add this edition perms
                        if (permItem.expDate != DateTime.MinValue)
                        {
                            // 判断是否过期超限
                            double dbDays = m_commTools.DateDiff(DateTime.Now, permItem.expDate);
                            int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                            if (nDays > 0)
                            {
                                permItem.bOpen = false;
                            }
                            else
                            {
                                permItem.bOpen = true;
                            }
                        }

                        permItem.cfgAutoUpdateSvrUrl = strCfgAutoUpdateSvrUrl;
                        permItem.cfgCerSvrUrl = strCfgCerSvrUrl;

                        m_hashPermTables[m_strPrivEditionName] = permItem;
                    }
                }
            }

            return 0;
        }



        // 
        private int BuildEntEdition()
        {
            int nRet = -1;

            String strCfgFile = m_strAppDir + m_strCfgEntEdition;
            String strLicFile = m_strAppDir + m_strLicEntEdition;

            clPermissionItem permItem = new clPermissionItem();
            String strCfgAutoUpdateSvrUrl = "";
            String strCfgCerSvrUrl = "";

            // if config exist, read config to get URLs
            Boolean bGetLicData = false;

            if (File.Exists(strCfgFile))
            {
                ConfigReader cfgReader = new ConfigReader();
                Hashtable cfgNameValues = cfgReader.getConfigItems(strCfgFile);

                if (cfgNameValues.Contains("cfgAutoUpdateSvrUrl"))
                {
                    strCfgAutoUpdateSvrUrl = (String)cfgNameValues["cfgAutoUpdateSvrUrl"];
                }

                if (cfgNameValues.Contains("cfgCerSvrUrl"))
                {
                    strCfgCerSvrUrl = (String)cfgNameValues["cfgCerSvrUrl"];
                }

                // access URL to get lic data
                if (!String.IsNullOrWhiteSpace(strCfgCerSvrUrl))
                {
                    AutoUpdateClass licSvrIntf = null;// new AutoUpdateClass(strCfgCerSvrUrl);

                    try
                    {
                        licSvrIntf = new AutoUpdateClass(strCfgCerSvrUrl);
                    }
                    catch (System.Exception ex)
                    {
                        // LOG
                        licSvrIntf = null;
                    }
                    finally
                    {
                    }

                    if (licSvrIntf != null)
                    {
                        QueryResult qRet = null;

                        try
                        {
                            qRet = licSvrIntf.SignForEntire(m_commTools.MachineId, true);

                            // parse 
                            if(qRet != null)
                            {
                                if (qRet.Data2 != null)
                                {
                                    permItem.strEditionName = qRet.Data2;
                                }

                                if(qRet.IsSuccess)
                                {
                                    permItem.expDate = qRet.Date;

                                    if (qRet.Data1 != null)
                                    {
                                        int nCnt = (qRet.Data1.Length / 32);
                                        String strItem = "";
                                        int nVal = 1;
                                        for (int i = 0; i < nCnt; i++)
                                        {
                                            strItem = qRet.Data1.Substring(i*32,32);
                                            permItem.hashPermission[strItem] = nVal;
                                        }
                                    }

                                    // write into lic file
                                    // 
                                    String strEncodeLic = "";
                                    // int nrt = EncodeLic(permItem, ref strEncodeLic);
                                    int nrt = m_commTools.EncodeLic(m_commTools.MachineIdMD5MD5, permItem.expDate, permItem.hashPermission, ref strEncodeLic);
                                    if (nrt == 0)
                                    {
                                        StreamWriter sw = new StreamWriter(strLicFile);
                                        sw.Write(strEncodeLic);
                                        sw.Close();
                                    }
                                }

                                permItem.bOpen = qRet.IsSuccess;
                                permItem.strStatus = "启用";

                                permItem.cfgAutoUpdateSvrUrl = strCfgAutoUpdateSvrUrl;
                                permItem.cfgCerSvrUrl = strCfgCerSvrUrl;

                                m_hashPermTables[m_strEntEditionName] = permItem;

                                bGetLicData = true;
                            }
                        }
                        catch (System.Exception ex)
                        {
                            qRet = null;
                            // LOG
                        }
                        finally
                        {
                            
                        }
                    }
                }
            }


            if (!bGetLicData && File.Exists(strLicFile))
            {
                // read
                String strCnt = "";

                try
                {
                    StreamReader rd = new StreamReader(strLicFile);
                    strCnt = rd.ReadToEnd();
                    rd.Close();
                }
                catch (System.Exception ex)
                {
                    strCnt = "";
                }
                finally
                {
                    
                }

                if (!String.IsNullOrWhiteSpace(strCnt))
                {
                    // validate
                    // nRet = DecodeLic(strCnt, ref permItem);
                    String strMachineID2 = "";

                    nRet = m_commTools.DecodeLic(strCnt,ref strMachineID2, ref permItem.hashPermission, ref permItem.expDate);

                    if (nRet == 0)
                    {
                        // add this edition perms
                        permItem.bOpen = true;
                        permItem.strStatus = "启用";

                        permItem.cfgAutoUpdateSvrUrl = strCfgAutoUpdateSvrUrl;
                        permItem.cfgCerSvrUrl = strCfgCerSvrUrl;

                        m_hashPermTables[m_strEntEditionName] = permItem;
                    }
                }
            }

            return 0;
        }


        public int Init()
        {
            BuildSoloEdition();

            BuildPrivateEdition();

            BuildEntEdition();

            BuildDocRepEdition();

            CheckValid();

            return 0;
        }



        private int BuildDocRepEdition()
        {
            String strCfgFile = m_strAppDir + m_strCfgDocRepoEdition;

            clPermissionItem permItem = null;
            String strCfgAutoUpdateSvrUrl = "";
            String strCfgCerSvrUrl = "";
            String strCfgDocRepositoryUrl = "";
            String strCfgLoginUrl = "";
            String strCfgUploadFileUrl = "";

            if (File.Exists(strCfgFile))
            {
                ConfigReader cfgReader = new ConfigReader();
                Hashtable cfgNameValues = cfgReader.getConfigItems(strCfgFile);

                if (cfgNameValues.Contains("cfgAutoUpdateSvrUrl"))
                {
                    strCfgAutoUpdateSvrUrl = (String)cfgNameValues["cfgAutoUpdateSvrUrl"];
                }

                if (cfgNameValues.Contains("cfgCerSvrUrl"))
                {
                    strCfgCerSvrUrl = (String)cfgNameValues["cfgCerSvrUrl"];
                }

                if (cfgNameValues.Contains("cfgDocRepositoryUrl"))
                {
                    strCfgDocRepositoryUrl = (String)cfgNameValues["cfgDocRepositoryUrl"];
                }

                if (cfgNameValues.Contains("cfgLoginUrl"))
                {
                    strCfgLoginUrl = (String)cfgNameValues["cfgLoginUrl"];
                }

                if (cfgNameValues.Contains("cfgUploadFileUrl"))
                {
                    strCfgUploadFileUrl = (String)cfgNameValues["cfgUploadFileUrl"];
                }

                if (m_hashPermTables.Contains(m_strDocRepositoryEditionName))
                {
                    permItem = (clPermissionItem)m_hashPermTables[m_strDocRepositoryEditionName];
                    m_hashPermTables.Remove(m_strDocRepositoryEditionName);
                }

                permItem = new clPermissionItem();

                permItem.cfgAutoUpdateSvrUrl = strCfgAutoUpdateSvrUrl;
                permItem.cfgCerSvrUrl = strCfgCerSvrUrl;

                permItem.cfgDocRepositoryUrl = strCfgDocRepositoryUrl;
                permItem.cfgLoginUrl = strCfgLoginUrl;
                permItem.cfgUploadFileUrl = strCfgUploadFileUrl;

                m_hashPermTables[m_strDocRepositoryEditionName] = permItem;
            }

            return 0;
        }


        // 函数，遍历各个版本，以确定是否超期，以关闭相应版本的权限
        // 
        public int CheckValid()
        {

            String strItem = "";
            clPermissionItem permItem = null;

            DateTime dtNow = DateTime.Now;

            foreach (DictionaryEntry ent in m_hashPermTables)
            {
                strItem = (String)ent.Key;
                permItem = (clPermissionItem)ent.Value;

                if (permItem.expDate != DateTime.MinValue)
                {
                    // 判断是否过期超限
                    double dbDays = m_commTools.DateDiff(DateTime.Now, permItem.expDate);
                    int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                    if (nDays > 0)
                    {
                        permItem.bOpen = false;
                    }
                    else
                    {
                        permItem.bOpen = true;
                    }
                }


            }


            return 0;
        }



        // 函数，激活/注册
        // 
        public int ActivateSoft(String strUserName, String strActCode ,ref String strRetMsg)
        {
            clPermissionItem permItem = (clPermissionItem)m_hashPermTables[m_strPrivEditionName];

            int nRet = -1;

            if (permItem == null)
            {
                // LOG

                return -1;
            }

            String strCfgCerSvrUrl = permItem.cfgCerSvrUrl;

            if (String.IsNullOrWhiteSpace(strCfgCerSvrUrl))
            {
                // LOG

                return -1;
            }

            AutoUpdateClass licSvrIntf = null;

            try
            {
                licSvrIntf = new AutoUpdateClass(strCfgCerSvrUrl);
            }
            catch (System.Exception ex)
            {
                // LOG
                licSvrIntf = null;
            }
            finally
            {
            }

            if (licSvrIntf != null)
            {
                QueryResult qRet = null;

                try
                {
                    qRet = licSvrIntf.ActiveProject(strUserName, strActCode, m_commTools.MachineId);

                    // parse 
                    if (qRet != null)
                    {
                        if (qRet.IsSuccess)
                        {
                            if (qRet.Data2 != null)
                            {
                                permItem.strEditionName = qRet.Data2;

                                Settings.Default.strVerName = qRet.Data2;
                                Settings.Default.Save();
                            }

                            permItem.expDate = qRet.Date;

                            if (qRet.Data1 != null)
                            {
                                int nCnt = (qRet.Data1.Length / 32);
                                String strItem = "";
                                int nVal = 1;
                                for (int i = 0; i < nCnt; i++)
                                {
                                    strItem = qRet.Data1.Substring(i * 32, 32);
                                    permItem.hashPermission[strItem] = nVal;
                                }
                            }

                            if (qRet.Data3 != null)
                            {
                                String[] strs = qRet.Data3.Split('_');

                                if (strs != null && strs.GetLength(0) >= 2)
                                {
                                    permItem.strAccount = strs[0];
                                    permItem.strActCode = strs[1];

                                    Settings.Default.strRegAcnt  = strs[0];
                                    Settings.Default.strRegActSn = strs[1];
                                    Settings.Default.Save();
                                }
                                // LOG
                            }

                            // write into lic file
                            String strLicFile = m_strAppDir + m_strLicPrivEdition;

                            String strEncodeLic = "";
                            // int nrt = EncodeLic(permItem, ref strEncodeLic);
                            int nrt = m_commTools.EncodeLic(m_commTools.MachineIdMD5MD5,permItem.expDate, permItem.hashPermission, ref strEncodeLic);
                            if (nrt == 0)
                            {
                                StreamWriter sw = new StreamWriter(strLicFile);
                                sw.Write(strEncodeLic);
                                sw.Close();
                            }

                            nRet = 0;
                        }
                        else
                        {
                            permItem.strErrInfo = qRet.ErrorInfo;
                            strRetMsg = qRet.ErrorInfo;
                        }

                        permItem.bOpen = qRet.IsSuccess;
                        // permItem.strStatus = "启用";
                        m_hashPermTables[m_strEntEditionName] = permItem;
                       
                    }
                }
                catch (System.Exception ex)
                {
                    qRet = null;
                    // LOG
                    nRet = -1;
                }
                finally
                {

                }
            }

            return nRet;
        }

        // 函数，不同版本要显示/不显示相应的UI按钮
        // 



        // 函数，启用某版本的权限、Lic数据(login）
        //       停用某版本的权限、Lic数据(logout)
        public int ToggleEditionStatus(String strEdition, Boolean bOpen, String strStatus = "")
        {
            int nRet = -1;

            if (!m_hashPermTables.Contains(strEdition))
            {
                return nRet;
            }

            clPermissionItem item = (clPermissionItem)m_hashPermTables[strEdition];

            item.bOpen = bOpen;

            if (!String.IsNullOrWhiteSpace(strStatus))
            {
                item.strStatus = strStatus;
            }

            return nRet;
        }


        // 函数，添加获取的权限
        public int AddEditionPerms(String strEdition, Hashtable hashPerms,Boolean bOpen, DateTime expDt)
        {
            clPermissionItem item = null;

            if (!m_hashPermTables.Contains(strEdition))
            {
                item = new clPermissionItem();
                m_hashPermTables[strEdition] = item;
            }

            item = (clPermissionItem)m_hashPermTables[strEdition];

            item.bOpen = bOpen;
            item.expDate = expDt;

            String strItem = "", strMD5 = "";
            int nVal = 0;

            foreach (DictionaryEntry ent in hashPerms)
            {
                strItem = (String)ent.Key;
                nVal = (int)ent.Value;

                strMD5 = ClassEncryptUtils.MD5Encrypt(strItem);
                strMD5 = strMD5.ToUpper();

                item.hashPermission[strMD5] = nVal;
            }

            return 0;
        }


        // 
        public int UpdateEditionPerms(String strEdition, Hashtable hashPlainNamePerms)
        {

            if (!m_hashPermTables.Contains(strEdition))
            {
                return -1;
            }

            clPermissionItem permItem = (clPermissionItem)m_hashPermTables[strEdition];

            permItem.hashPermission.Clear();

            String strName = "";
            int nVal = 0;

            foreach(DictionaryEntry ent in hashPlainNamePerms)
            {
                strName = (String)ent.Key;
                nVal = (int)ent.Value;

                if (nVal != 0)
                {
                    String strMD5 = "";
                    
                    if(m_hashStr2MD5.Contains(strName))
                    {
                        strMD5 = (String)m_hashStr2MD5[strName];
                    }
                    else
                    {
                        strMD5 = ClassEncryptUtils.MD5Encrypt(strName);
                        strMD5 = strMD5.ToUpper();

                        m_hashStr2MD5[strName] = strMD5;
                    }

                    permItem.hashPermission[strMD5] = nVal;
                }
            }

            return 0;
        }


        // 
        public int RemoveEditionPerms(String strEdition)
        {
            if (!m_hashPermTables.Contains(strEdition))
            {
                return -1;
            }

            m_hashPermTables.Remove(strEdition);
            
            return 0;
        }

        public int ResetEditionPerms(String strEdition)
        {
            if (!m_hashPermTables.Contains(strEdition))
            {
                return -1;
            }

            clPermissionItem permItem = (clPermissionItem)m_hashPermTables[strEdition];

            permItem.hashPermission.Clear();

            return 0;
        }


        public Boolean IsExistEdition(String strEdition)
        {
            Boolean bRet = m_hashPermTables.Contains(strEdition);
            return bRet;
        }

        // 
        public clPermissionItem SearchEditionPerms(String strEdition)
        {
            return (clPermissionItem)m_hashPermTables[strEdition];
        }


        // 函数，多版本的权限数据、Lic数据的合并、查询
        public Boolean IsEnableViaPlainString(String strItem)
        {
            String strMD5 = "";
            if (m_hashStr2MD5.Contains(strItem))
            {
                strMD5 = (String)m_hashStr2MD5[strItem];
            }
            else
            {
                strMD5 = ClassEncryptUtils.MD5Encrypt(strItem);
                strMD5 = strMD5.ToUpper();

                m_hashStr2MD5[strItem] = strMD5;
            }

            return IsEnable(strMD5);
        }


        public Boolean IsEnable(String strItemMD5)
        {
            Boolean bRet = false;

            String strItem = "";
            clPermissionItem val = null;
            int nVal = 0;

            //if (m_hashExceptional.Contains(strItemMD5))
            //{
            //    return true;
            //}

            foreach (DictionaryEntry ent in m_hashPermTables)
            {
                strItem = (String)ent.Key;
                val = (clPermissionItem)ent.Value;

                if (val.bOpen)
                {
                    if (val.hashPermission.Contains(strItemMD5))
                    {
                        nVal = (int)val.hashPermission[strItemMD5];
                        if (nVal > 0)
                        {
                            return true;
                        }
                    }
                }

            }// foreach

            return bRet;
        }


        public void UpdateSubRibbonUI(RibbonBox ribBox, int nVal)
        {
            foreach(RibbonControl subCtrl in ribBox.Items)
            {
                if (m_hashExceptional.Contains(subCtrl.Name))
                {
                    subCtrl.Enabled = true;
                }
                else
                {
                    subCtrl.Enabled = (nVal != 0); 
                }

                if (subCtrl is RibbonBox)
                {
                    UpdateSubRibbonUI(subCtrl as RibbonBox, nVal);
                }

            }

            return;
        }


        public int UpdateUI(Hashtable hashControls)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[178] Enter classMultiEditionCenter::updateUI");
            }

            String strName = "";
            int nVal = 0;
            Control ctrl = null;
            ToolStripItem item = null;
            RibbonGroup ribGrp = null;
            RibbonControl ribCtrl = null;
            RibbonBox ribBox = null;

            Object ctrlObj = null;

            foreach (DictionaryEntry ent in hashControls)
            {
                strName = (String)ent.Key;
                ctrlObj = (Object)ent.Value;

                nVal = 0;

                if (IsEnableViaPlainString(strName))
                {
                    nVal = 1;
                }
#if ADMIN
                nVal = 1;
#endif

                if (ctrlObj != null)
                {
                    if (ctrlObj is TabPage /*|| ctrlObj is TabControl*/)
                    {
                        ctrl = (Control)ctrlObj;
                        ctrl.Enabled = (nVal != 0);
                    }
                    else if (ctrlObj is ToolStripItem)
                    {
                        item = (ToolStripItem)ctrlObj;
                        item.Enabled = (nVal != 0);
                    }
                    else if (ctrlObj is RibbonGroup)
                    {
                        ribGrp = (RibbonGroup)ctrlObj;
                        
                        foreach(RibbonControl subCtrl1 in ribGrp.Items )
                        {
                            if (m_hashExceptional.Contains(subCtrl1.Name))
                            {
                                subCtrl1.Enabled = true;
                            }
                            else
                            {
                                subCtrl1.Enabled = (nVal != 0); 
                            }

                            if (subCtrl1 is RibbonBox)
                            {
                                ribBox = (RibbonBox)subCtrl1;
                                UpdateSubRibbonUI(ribBox,nVal);
                            }
                        }
                    }
/*
                    else if (ctrlObj is RibbonControl)
                    {
                        ribCtrl = (RibbonControl)ctrlObj;
                        ribCtrl.Enabled = (nVal != 0);
                    }*/
                }// if
            }

            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[236] Exit classMultiEditionCenter::updateUI");
            }

            return 0;
        }


        // 函数，UpdateUi，即根据权限、Lic数据来决定界面是否可用或禁用
        public int UpdateUI_v1(Hashtable hashControls)
        {
            if (ThisAddIn.m_bLog)
            {
                Log.WriteLog("[178] Enter classMultiEditionCenter::updateUI");
            }

            String strName = "";
            int nVal = 0;
            Control ctrl = null;
            ToolStripItem item = null;
            RibbonControl ribCtrl = null;
            Object ctrlObj = null;

            foreach (DictionaryEntry ent in hashControls)
            {
                strName = (String)ent.Key;
                ctrlObj = (Object)ent.Value;

                nVal = 0;

                if (IsEnableViaPlainString(strName))
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
                Log.WriteLog("[236] Exit classMultiEditionCenter::updateUI");
            }

            return 0;
        }


        public String GetPrivEditionInfo(ref String strAccount, ref String strActCode)
        {
            String strInfo = "";

            if (!m_hashPermTables.Contains(m_strPrivEditionName))
            {
                return strInfo;
            }

            clPermissionItem permItem = (clPermissionItem)m_hashPermTables[m_strPrivEditionName];

            if (String.IsNullOrWhiteSpace(permItem.strAccount) || 
                String.IsNullOrWhiteSpace(permItem.strActCode) ||
                permItem.expDate == DateTime.MinValue)
            {
                strAccount = "";
                strActCode = "";
                return strInfo;
            }

            strAccount = permItem.strAccount;
            strActCode = permItem.strActCode;

            // 判断是否过期超限
            String strExpDate = permItem.expDate.ToString("yyyy年M月d日");

            double dbDays = m_commTools.DateDiff(DateTime.Now, permItem.expDate);
            int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

            if (nDays > 0)
            {
                strInfo = "状态：过期" + nDays + "天(到期日期：" + strExpDate + ")\r\n";
                    
                // 过期(nDays)
                // permItem.bOpen = false;
            }
            else
            {
                strInfo ="状态:正常，还余" + (-1 * nDays) + "天到期(到期日期：" + strExpDate + ")\r\n";

                // 正常:还余(-1*nDays）天到期
                // permItem.bOpen = true;
            }


            return strInfo;
        }


        // 函数，反应出当前版本名称的数据、超限数据（如超期）、已超数量限
        // 
        public String GetInfo()
        {
            String strItem = "", strInfo = "", strExpDate = "";
            clPermissionItem permItem = null;

            DateTime dtNow = DateTime.Now;

            foreach (DictionaryEntry ent in m_hashPermTables)
            {
                strItem = (String)ent.Key;
                permItem = (clPermissionItem)ent.Value;

                if (permItem.expDate != DateTime.MinValue)
                {
                    // 判断是否过期超限
                    strExpDate = permItem.expDate.ToString("yyyy年M月d日");

                    double dbDays = m_commTools.DateDiff(DateTime.Now, permItem.expDate);
                    int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                    if (nDays > 0)
                    {
                        if (!strItem.Equals(m_strPrivEditionName))
                        {
                            strInfo += strItem + ":过期" + nDays + "天(到期日期：" + strExpDate + ")\r\n";
                        }
                        else
                        {
                            strInfo += Settings.Default.strVerName + ":过期" + nDays + "天(到期日期：" + strExpDate + ")\r\n";
                        }
                        // 过期(nDays)
                        // permItem.bOpen = false;
                    }
                    else
                    {
                        if (!strItem.Equals(m_strPrivEditionName))
                        {
                            strInfo += strItem + ":正常，还余" + (-1 * nDays) + "天到期(到期日期：" + strExpDate + ")\r\n";
                        }
                        else
                        {
                            strInfo += Settings.Default.strVerName + ":正常，还余" + (-1 * nDays) + "天到期(到期日期：" + strExpDate + ")\r\n";
                        }

                        // 正常:还余(-1*nDays）天到期
                        // permItem.bOpen = true;
                    }
                }

            }


            return strInfo;
        }

        
        // 函数，auto update
        // 
        public String CheckUpdate()
        {
            String strRet = "";


            String strAutoUpdateUrl = "";
            clPermissionItem permItem = null;

            if (m_hashPermTables.Contains(m_strDocRepositoryEditionName))
            {
                permItem = (clPermissionItem)m_hashPermTables[m_strDocRepositoryEditionName];
                strAutoUpdateUrl = permItem.cfgAutoUpdateSvrUrl;
            }
            else if (m_hashPermTables.Contains(m_strEntEditionName))
            {
                permItem = (clPermissionItem)m_hashPermTables[m_strEntEditionName];
                strAutoUpdateUrl = permItem.cfgAutoUpdateSvrUrl;
            }
            else if (m_hashPermTables.Contains(m_strPrivEditionName))
            {
                permItem = (clPermissionItem)m_hashPermTables[m_strPrivEditionName];
                strAutoUpdateUrl = permItem.cfgAutoUpdateSvrUrl;
            }
            else
            {
                
            }


            if(String.IsNullOrWhiteSpace(strAutoUpdateUrl))
            {
                return "-4";
            }


            AutoUpdateClass licSvrIntf = null;

            try
            {
                licSvrIntf = new AutoUpdateClass(strAutoUpdateUrl);
            }
            catch (System.Exception ex)
            {
                // LOG
                licSvrIntf = null;
            }
            finally
            {
            }

            if (licSvrIntf != null)
            {
                try
                {
                    if (ThisAddIn.m_bLog)
                    {
                        // Log.WriteLog("[1642] ClassUIController::checkUpdate");
                    }

                    strRet = licSvrIntf.CheckUpdate();

                    if (ThisAddIn.m_bLog)
                    {
                        // Log.WriteLog("[1648] ClassUIController::checkUpdate");
                    }

                    //MessageBox.Show("4");
                }
                catch (System.Exception ex)
                {
                    if (ThisAddIn.m_bLog)
                    {
                        // Log.WriteLog("[1657] ClassUIController::checkUpdate, Exception:" + ex.ToString());
                    }

                    MessageBox.Show(ex.Message);
                    strRet = "-3";
                }
                finally
                {
                }
            }

            return strRet;
        }


    }
}
