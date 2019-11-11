﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Security.Cryptography;
using System.Reflection;
using System.Net;
using System.Web;
using System.Collections;
using System.Collections.Specialized;
using Microsoft.Office.Core;


namespace OfficeTools.Common
{
    public class ClassOfficeCommon
    {
        public enum TizhuScope
        {
            tizhuScopeAllDoc = 0,
            tizhuScopeAfterToc = 1,
            // tizhuScopeAfterHeading = 3,
            // tizhuScopeAfterPage = 4,
            // tizhuScopeAFterSection = 5,
        }


        public Boolean bAppIsWps = false;

        #region
        
        // 数字转换
        char[] m_digitChArabicNum = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.' };
        char[] m_digitChSimpChNum = { '〇', '一', '二', '三', '四', '五', '六', '七', '八', '九', '点' };
        char[] m_digitChBigSimpChNum = { '零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖', '点' };

        Hashtable m_digitHashArabicNum = new Hashtable();
        Hashtable m_digitHashSimpChNum = new Hashtable();
        Hashtable m_digitHashBigSimpChNum = new Hashtable();

        // 数值、金额转换
        String[] m_strArrArabicNum = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-" };
        String[] m_strArrSimpChNum = { "〇", "一", "二", "三", "四", "五", "六", "七", "八", "九", "负" };
        String[] m_strArrBigSimpChNum = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖", "负" };

        String[] m_strArrBigSimpLevel = { "拾", "佰", "仟", "万", "亿", "元", "点", "角", "分", "厘", "负" };
        double[] m_dbArrLevel = { 10.0, 100.0, 1000.0, 10000.0, 100000000.0, 1.0, 1.0, 0.1, 0.01, 0.001, -1.0 };
        double[] m_dbArrNum = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, -1 };

        Hashtable m_hashArabicNum = new Hashtable();
        Hashtable m_hashSimpChNum = new Hashtable();
        Hashtable m_hashBigSimpChNum = new Hashtable();
        Hashtable m_hashSimpChLevel = new Hashtable();

        Hashtable m_hashSimpLittle2Num = new Hashtable();
        Hashtable m_hashSimpBig2Num = new Hashtable();
        Hashtable m_hashSimp2NumLevel = new Hashtable();

        Hashtable m_hash4Check = new Hashtable();

        #endregion

        public char[] m_trimChars = new char[] { '　', ' ', '\t', '\r', '\n', '\a',(char)11, '\f',(char)14 };


        // lic code/decode
        private String m_strMachineId = "", m_strMD5MachineId = "", m_strMD5MD5MachineId = "";
        private readonly String m_strSpecialMachineId = "AFBECD12AFBECD12AFBECD12AFBECD12";

        private Hashtable m_hashMD5toNum = new Hashtable();
        private ArrayList m_arrNum2MD5 = new ArrayList();

        // 
        public String MachineId
        {
            get
            {
                if (String.IsNullOrWhiteSpace(m_strMachineId))
                {
                    ClassHardInfo clsHardInfo = new ClassHardInfo();
                    String strInfo = clsHardInfo.GetCpuID();

                    m_strMachineId = strInfo.ToUpper();
                }

                return m_strMachineId;
            }

        }


        public String MachineIdMD5
        {
            get
            {
                if (String.IsNullOrWhiteSpace(m_strMD5MachineId))
                {
                    String strInfo = ClassEncryptUtils.MD5Encrypt(MachineId);
                    m_strMD5MachineId = strInfo.ToUpper();
                }

                return m_strMD5MachineId;
            }
        }


        public String MachineIdMD5MD5
        {
            get
            {
                if (String.IsNullOrWhiteSpace(m_strMD5MD5MachineId))
                {
                    String strInfo = ClassEncryptUtils.MD5Encrypt(MachineIdMD5);
                    m_strMD5MD5MachineId = strInfo.ToUpper();
                }

                return m_strMD5MD5MachineId;
            }
        }


        public String SpecialMachineId
        {
            get
            {
                return m_strSpecialMachineId;
            }

        }

        public class TiZhuHeadingItem
        {
            public String strHeadingName;
            public int nRngStart;
            public int nRngEnd;
            public int nCoverRngStart;
            public int nCoverRngEnd;

            public Word.WdOutlineLevel wdLevel;

            public int nTblCnt;
            public int nInShpCnt;
        }

        public ClassOfficeCommon()
        {
            int i = 0;
            foreach (char ch in m_digitChArabicNum)
            {
                m_digitHashArabicNum.Add(ch, i);
                i++;
            }

            i = 0;
            foreach (char ch in m_digitChSimpChNum)
            {
                m_digitHashSimpChNum.Add(ch, i);
                i++;
            }

            i = 0;
            foreach (char ch in m_digitChBigSimpChNum)
            {
                m_digitHashBigSimpChNum.Add(ch, i);
                i++;
            }
           
            //
            foreach (String strItem in m_strArrSimpChNum)
            {
                m_hash4Check[strItem] = strItem;
            }

            foreach (String strItem in m_strArrBigSimpChNum)
            {
                m_hash4Check[strItem] = strItem;
            }

            foreach (String strItem in m_strArrBigSimpLevel)
            {
                m_hash4Check[strItem] = strItem;
            }


            i = 0;
            foreach (String strItem in m_strArrArabicNum)
            {
                m_hashArabicNum.Add(strItem, strItem);
                // i++;
            }

            i = 0;
            foreach (String strItem in m_strArrSimpChNum)
            {
                m_hashSimpChNum.Add(m_strArrArabicNum[i], strItem);
                i++;
            }

            i = 0;
            foreach (String strItem in m_strArrBigSimpChNum)
            {
                m_hashBigSimpChNum.Add(m_strArrArabicNum[i], strItem);
                i++;
            }

            m_hashSimpChLevel.Add(-10, "");
            m_hashSimpChLevel.Add(-9, "");
            m_hashSimpChLevel.Add(-8, "");
            m_hashSimpChLevel.Add(-7, "");
            m_hashSimpChLevel.Add(-6, "");
            m_hashSimpChLevel.Add(-5, "");
            m_hashSimpChLevel.Add(-4, "");
            m_hashSimpChLevel.Add(-3, "厘");
            m_hashSimpChLevel.Add(-2, "分");
            m_hashSimpChLevel.Add(-1, "角");
            m_hashSimpChLevel.Add(0, "");
            m_hashSimpChLevel.Add(1, "拾");
            m_hashSimpChLevel.Add(2, "佰");
            m_hashSimpChLevel.Add(3, "仟");
            m_hashSimpChLevel.Add(4, "万");
            m_hashSimpChLevel.Add(5, "拾");
            m_hashSimpChLevel.Add(6, "佰");
            m_hashSimpChLevel.Add(7, "仟");
            m_hashSimpChLevel.Add(8, "亿");
            m_hashSimpChLevel.Add(9, "拾");
            m_hashSimpChLevel.Add(10, "佰");
            m_hashSimpChLevel.Add(11, "仟");
            m_hashSimpChLevel.Add(12, "万");
            m_hashSimpChLevel.Add(13, "拾"); // 10万亿
            m_hashSimpChLevel.Add(14, "佰");
            m_hashSimpChLevel.Add(15, "仟");
            m_hashSimpChLevel.Add(16, "万");

            i = 0;
            foreach (String strItem in m_strArrSimpChNum)
            {
                m_hashSimpLittle2Num.Add(strItem, m_dbArrNum[i]);
                i++;
            }

            i = 0;
            foreach (String strItem in m_strArrBigSimpChNum)
            {
                m_hashSimpBig2Num.Add(strItem, m_dbArrNum[i]);
                i++;
            }

            i = 0;
            foreach (String strItem in m_strArrBigSimpLevel)
            {
                m_hashSimp2NumLevel.Add(strItem, m_dbArrLevel[i]);
                i++;
            }

            String strMD5 = "", strInfo = "";

            for (i = 0; i < 10; i++)
            {
                strInfo = ClassEncryptUtils.MD5Encrypt("li" + i + "dong");
                strMD5 = strInfo.ToUpper();

                m_arrNum2MD5.Add(strMD5);
                m_hashMD5toNum[strMD5] = "" + i;
            }
            //

            return;
        }


        #region Lic

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


        private int StrMD5toDateTime(String strMD5, ref String strDate)
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

            strDate = strDt;

            return 0;
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
        }


        // 
        public int DecodeLic(String strCnt,ref String strMachineID2, ref Hashtable hashPerm,ref DateTime dtExp,Boolean bValidateMachineID = true,String strLicFile = "")
        {
            int nRet = -1;

            if ((strCnt.Length % 32) != 0)
            {
                // LOG

                return -1;
            }

            int nCnt = (strCnt.Length / 32);

            if (nCnt < 10)
            {
                // LOG

                return -2;
            }


            // CRC 1
            String strCRC = strCnt.Substring(0, 32);

            // machine id (MD5 MD5) 1
            String strMacineIdMD5MD5 = strCnt.Substring(32, 32);

            strMachineID2 = strMacineIdMD5MD5;

            // Expire date(yyyyMMdd) 8
            String strExpDt = strCnt.Substring(64, 8 * 32);

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

                return -3;
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
                if (bValidateMachineID && !strMacineIdMD5MD5.Equals(MachineIdMD5MD5))
                {
                    // LOG

                    return -4;
                }
            }

            if (StrMD5toDateTime(strExpDt, ref dtExp) != 0)
            {
                return -5;
            }

            int nUiCnt = (strUI.Length / 32);
            String strItem = "";

            for (int i = 0; i < nUiCnt; i++)
            {
                strItem = strUI.Substring(i * 32, 32);
                hashPerm[strItem] = (int)1;
            }

            // permItem.bOpen = true;

            if (bSpecialMachineId && !String.IsNullOrWhiteSpace(strLicFile))
            {
                // rewrite
                StreamWriter sw = new StreamWriter(strLicFile);

                // machine id (MD5 MD5) 1
                strMacineIdMD5MD5 = MachineIdMD5MD5;

                strCRC = ClassEncryptUtils.MD5Encrypt(strMacineIdMD5MD5 + strExpDt + strUI);
                strCRC = strCRC.ToUpper();

                sw.Write(strCRC + strMacineIdMD5MD5 + strExpDt + strUI);

                sw.Flush();
                sw.Close();

            }

            return 0;
        }


        public int EncodeLic(String strMachineID2,DateTime expDt, String strUI, ref String strEncodedCnt, Boolean bSpecialMachineID = false)
        {
            // CRC 1
            String strCRC = "";

            // machine id (MD5 MD5) 1
            String strMacineIdMD5MD5 = strMachineID2;// m_strMD5MD5MachineId;

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
        }


        public int EncodeLic(String strMachineID2, DateTime expDt, Hashtable hashPerm, ref String strEncodedCnt, Boolean bSpecialMachineID = false)
        {

            String strUI = "";

            foreach (DictionaryEntry ent in hashPerm)
            {
                strUI += (String)ent.Key;
            }


            // CRC 1
            String strCRC = "";

            // machine id (MD5 MD5) 1
            String strMacineIdMD5MD5 = strMachineID2;// m_strMD5MD5MachineId;

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
        }

        #endregion




        #region 数字转换

        public void digitTranslate(double dbNum, out String strArabicNum,
                                   out String strSimpChNum, out String strBigSimpChNum)
        {
            String strNum = dbNum.ToString();

            digitTranslate(strNum, out strArabicNum, out strSimpChNum, out strBigSimpChNum);
            return;
        }


        public void digitTranslate(String strOrigText, out String strArabicNum,
                                   out String strSimpChNum, out String strBigSimpChNum)
        {
            //             // 数字转换
            //             // 
            //             char[] m_digitChArabicNum = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9','.','-'};
            //             char[] m_digitChSimpChNum = { '〇', '一', '二', '三', '四', '五', '六', '七', '八','九' ,'点','负'};
            //             char[] m_digitChBigSimpChNum = { '零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖' ,'点','负'};
            // 
            //             Hashtable m_digitHashArabicNum = new Hashtable();
            //             Hashtable m_digitHashSimpChNum = new Hashtable();
            //             Hashtable m_digitHashBigSimpChNum = new Hashtable();
            // 
            //             int i = 0;
            //             foreach (char ch in m_digitChArabicNum)
            //             {
            //                 m_digitHashArabicNum.Add(ch, i);
            //                 i++;
            //             }
            // 
            //             i = 0;
            //             foreach (char ch in m_digitChSimpChNum)
            //             {
            //                 m_digitHashSimpChNum.Add(ch, i);
            //                 i++;
            //             }
            // 
            //             i = 0;
            //             foreach (char ch in m_digitChBigSimpChNum)
            //             {
            //                 m_digitHashBigSimpChNum.Add(ch, i);
            //                 i++;
            //             }

            char[] chArr = strOrigText.ToCharArray();
            char[] chArrArabicNum = strOrigText.ToCharArray();
            char[] chArrSimpChNum = strOrigText.ToCharArray();
            char[] chArrBigSimpChNum = strOrigText.ToCharArray();

            char chTmp = ' ';
            int nIndex = 0;

            for (int n = 0; n < chArr.GetLength(0); n++)
            {
                chTmp = chArr[n];

                if (m_digitHashArabicNum.Contains(chTmp))
                {
                    nIndex = (int)m_digitHashArabicNum[chTmp];
                    chArrArabicNum[n] = m_digitChArabicNum[nIndex];
                    chArrSimpChNum[n] = m_digitChSimpChNum[nIndex];
                    chArrBigSimpChNum[n] = m_digitChBigSimpChNum[nIndex];
                }
                else if (m_digitHashSimpChNum.Contains(chTmp))
                {
                    nIndex = (int)m_digitHashSimpChNum[chTmp];
                    chArrArabicNum[n] = m_digitChArabicNum[nIndex];
                    chArrSimpChNum[n] = m_digitChSimpChNum[nIndex];
                    chArrBigSimpChNum[n] = m_digitChBigSimpChNum[nIndex];
                }
                else if (m_digitHashBigSimpChNum.Contains(chTmp))
                {
                    nIndex = (int)m_digitHashBigSimpChNum[chTmp];
                    chArrArabicNum[n] = m_digitChArabicNum[nIndex];
                    chArrSimpChNum[n] = m_digitChSimpChNum[nIndex];
                    chArrBigSimpChNum[n] = m_digitChBigSimpChNum[nIndex];
                }
                else
                {
                    chArrArabicNum[n] = chTmp;
                    chArrSimpChNum[n] = chTmp;
                    chArrBigSimpChNum[n] = chTmp;
                }

            }

            strArabicNum = new String(chArrArabicNum);
            strSimpChNum = new String(chArrSimpChNum);
            strBigSimpChNum = new String(chArrBigSimpChNum);

            return;
        }
        #endregion

        #region 数值转换
        Boolean IsValidSimpValueString(String strSimp, ref String strRetMsg)
        {
            String strItem = "";

            for (int i = 0; i < strSimp.Length; i++)
            {
                strItem = "";
                strItem += strSimp[i];

                if (strItem.Equals(""))
                {
                    continue;
                }

                if (!m_hash4Check.Contains(strItem))
                {
                    strRetMsg = "存在不符字符：" + strItem;
                    return false;
                }

                if (strItem.Equals("负") || strItem.Equals("-"))
                {
                    if (i != 0) // first
                    {
                        strRetMsg = "字符\'负\'不在首位:" + strSimp;
                        return false;
                    }
                }

            }

            return true;
        }


        public double SimpValueString2num(String strSimp, ref String strRetMsg)
        {
            if (!IsValidSimpValueString(strSimp, ref strRetMsg))
            {
                return double.NaN;
            }


            Stack stackPreNum = new Stack();

            double dbYiVal = 0.0, dbWangVal = 0.0;
            double dbRet = 0.0, dbNum = 0.0, dbLevel = 0.0;
            double dbMaxLevel = double.NaN, dbPosNeg = 1.0;
            Boolean bFloated = false;
            String strItem = "";
            int i = 0;

            for (i = 0; i < strSimp.Length; i++)
            {
                strItem = "";
                strItem += strSimp[i];

                if (strItem.Equals("负"))
                {
                    dbPosNeg = -1.0;
                    continue;
                }

                if (m_hashSimpLittle2Num.Contains(strItem))
                {
                    dbNum = (double)m_hashSimpLittle2Num[strItem];

                    stackPreNum.Push(dbNum);

                    if (bFloated)
                    {
                        dbLevel /= 10.0;
                    }
                    else
                    {
                        dbLevel = 0.0;
                    }
                }
                else if (m_hashSimpBig2Num.Contains(strItem))
                {
                    dbNum = (double)m_hashSimpBig2Num[strItem];

                    stackPreNum.Push(dbNum);

                    if (bFloated)
                    {
                        dbLevel /= 10.0;
                    }
                    else
                    {
                        dbLevel = 0.0;
                    }
                }
                else if (m_hashSimp2NumLevel.Contains(strItem))
                {
                    if (strItem.Equals("点"))
                    {
                        bFloated = true;
                        dbLevel = 1.0;

                    }
                    else
                    {
                        dbLevel = (double)m_hashSimp2NumLevel[strItem];
                        // continue;
                    }


                    if (double.IsNaN(dbMaxLevel))
                    {
                        dbMaxLevel = dbLevel;
                    }

                    if (stackPreNum.Count > 0)
                    {
                        dbNum = (double)stackPreNum.Pop();
                    }
                    else
                    {
                        dbNum = 0.0;
                    }

                    if (dbLevel == 10000.0) // 万
                    {
                        dbWangVal += dbNum;
                        dbWangVal *= dbLevel;

                        dbYiVal += dbWangVal;

                        dbWangVal = 0.0;
                    }
                    else if (dbLevel == 100000000.0) // 亿
                    {
                        dbYiVal += (dbNum + dbWangVal);
                        dbYiVal *= dbLevel;
                        dbWangVal = 0.0;
                    }
                    else
                    {
                        dbWangVal += (dbNum * dbLevel);
                        // dbYiVal += (dbNum * dbLevel);

                    }


                    /*
                    if (dbLevel > dbMaxLevel)
                    {
                        dbRet += dbNum;
                        // calc
                        dbRet *= dbLevel;

                        dbMaxLevel = dbLevel;
                    }
                    else
                    {
                        dbRet += (dbNum * dbLevel);
                        // dbLevel = 0.0;
                    }
                     */ 

                }

            } // for


            if (stackPreNum.Count > 0) // should be 1
            {
                for (i = 0; i < stackPreNum.Count; i++) // should
                {
                    dbNum = (double)stackPreNum.Pop();

                    if (bFloated)
                    {
                        dbRet += (dbNum * dbLevel);
                    }
                    else
                    {
                        dbRet += dbNum;
                    }

                    dbWangVal += dbRet;
                    //dbYiVal += dbRet;
                }
            }

            return ((dbYiVal + dbWangVal) * dbPosNeg);

            // return (dbRet * dbPosNeg);

        }




/*
        // 
        public double SimpValueString2num(String strSimp, ref String strRetMsg)
        {
            if (!IsValidSimpValueString(strSimp, ref strRetMsg))
            {
                return double.NaN;
            }

            Stack stackPreNum = new Stack();

            double dbRet = 0.0, dbNum = 0.0, dbLevel = 0.0;
            double dbMaxLevel = double.NaN, dbPosNeg = 1.0;
            Boolean bFloated = false;
            String strItem = "";
            int i = 0;

            for (i = 0; i < strSimp.Length; i++)
            {
                strItem = "";
                strItem += strSimp[i];

                if (strItem.Equals("负"))
                {
                    dbPosNeg = -1.0;
                    continue;
                }

                if (m_hashSimpLittle2Num.Contains(strItem))
                {
                    dbNum = (double)m_hashSimpLittle2Num[strItem];

                    stackPreNum.Push(dbNum);

                    if (bFloated)
                    {
                        dbLevel /= 10.0;
                    }
                    else
                    {
                        dbLevel = 0.0;
                    }
                }
                else if (m_hashSimpBig2Num.Contains(strItem))
                {
                    dbNum = (double)m_hashSimpBig2Num[strItem];

                    stackPreNum.Push(dbNum);

                    if (bFloated)
                    {
                        dbLevel /= 10.0;
                    }
                    else
                    {
                        dbLevel = 0.0;
                    }
                }
                else if (m_hashSimp2NumLevel.Contains(strItem))
                {

                    if (strItem.Equals("点"))
                    {
                        bFloated = true;
                        dbLevel = 1.0;

                    }
                    else
                    {
                        dbLevel = (double)m_hashSimp2NumLevel[strItem];
                        // continue;
                    }

                    if (double.IsNaN(dbMaxLevel))
                    {
                        dbMaxLevel = dbLevel;
                    }

                    dbNum = (double)stackPreNum.Pop();

                    if (dbLevel > dbMaxLevel)
                    {
                        dbRet += dbNum;
                        // calc
                        dbRet *= dbLevel;

                        dbMaxLevel = dbLevel;
                    }
                    else
                    {
                        dbRet += (dbNum * dbLevel);
                        // dbLevel = 0.0;
                    }

                }

            } // for

            if (stackPreNum.Count > 0) // should be 1
            {
                for (i = 0; i < stackPreNum.Count; i++) // should
                {
                    dbNum = (double)stackPreNum.Pop();

                    if (bFloated)
                    {
                        dbRet += (dbNum * dbLevel);
                    }
                    else
                    {
                        dbRet += dbNum;
                    }
                }
            }

            return (dbRet * dbPosNeg);
        }*/

        //
        public String num2SimpValueString(double dbNum, Boolean bBig = true, Boolean bMoney = false)
        {
            String strPreFix = "";

            String strNum = dbNum.ToString();//String.Format("{0:f}", dbNum); //

            if (dbNum < 0.0) // negative
            {
                strPreFix = "负";
                // strNum = (-1 * dbNum).ToString();
            }

            int i = 0;

            int nPointPos = strNum.IndexOf('.');

            String strIntegerPart = strNum, strFloatPart = "";
            double dbIntegerPart = 0.0, dbFloatPart = 0.0;

            if (nPointPos != -1)
            {
                strIntegerPart = strNum.Substring(0, nPointPos);
                strFloatPart = strNum.Substring(nPointPos + 1); // include point

                dbFloatPart = Convert.ToDouble(strFloatPart);
            }

            dbIntegerPart = Convert.ToDouble(strIntegerPart);


            int nLvl = 0;

            String strValue = "", strLevel = "", strIntegerResult = "", strFloatResult = "";
            String strResult = "";

            String strPreCh = "", strCurCh = "";
            String strScannedValue = "";
            double dbCurValue = 0.0;

            for (i = strIntegerPart.Length - 1, nLvl = 0; i >= 0; i--, nLvl++)
            {
                strValue = "";
                strLevel = "";

                if ((nLvl & 0x03) == 0)
                {
                    strScannedValue = "";
                }

                strCurCh = strIntegerPart[i] + "";

                strScannedValue = strCurCh + strScannedValue;

                dbCurValue = Convert.ToDouble(strScannedValue);

                if (strCurCh.Equals("0"))
                {
                    if (dbCurValue == 0.0)
                    {
                        strValue = "";
                        strLevel = "";
                    }
                    else if (!strPreCh.Equals(strCurCh))
                    {
                        if (!bBig)
                        {
                            strValue = (String)m_hashSimpChNum[strCurCh];
                        }
                        else
                        {
                            strValue = (String)m_hashBigSimpChNum[strCurCh];
                        }
                    }
                    else
                    {
                        strValue = "";
                        strLevel = "";
                    }

                    if ((nLvl & 0x03) == 0)
                    {
                        strLevel = (String)m_hashSimpChLevel[nLvl];
                    }
                }
                else
                {
                    if (!bBig)
                    {
                        strValue = (String)m_hashSimpChNum[strCurCh];
                    }
                    else
                    {
                        strValue = (String)m_hashBigSimpChNum[strCurCh];
                    }

                    strLevel = (String)m_hashSimpChLevel[nLvl];
                }

                strIntegerResult = strValue + strLevel + strIntegerResult;

                strPreCh = strCurCh;
            }



            // float
            strPreCh = "";
            strCurCh = "";
            strScannedValue = "";
            dbCurValue = 0.0;

            //for (i = strFloatPart.Length - 1, nLvl = -1 * strFloatPart.Length; i >= 0; i--, nLvl++)
            for (i = 0, nLvl = -1; i < strFloatPart.Length; i++, nLvl--)
            {
                strValue = "";
                strLevel = "";

                strCurCh = strFloatPart[i] + "";

                if (bMoney)
                {
                    if (strCurCh.Equals("0"))
                    {
                        if (!strPreCh.Equals(strCurCh))
                        {
                            if (!bBig)
                            {
                                strValue = (String)m_hashSimpChNum[strCurCh];
                            }
                            else
                            {
                                strValue = (String)m_hashBigSimpChNum[strCurCh];
                            }
                        }
                        else
                        {
                            strValue = "";
                            strLevel = "";
                        }
                    }
                    else
                    {
                        if (!bBig)
                        {
                            strValue = (String)m_hashSimpChNum[strCurCh];
                        }
                        else
                        {
                            strValue = (String)m_hashBigSimpChNum[strCurCh];
                        }

                        strLevel = (String)m_hashSimpChLevel[nLvl];

                    }
                }
                else
                {
                    if (!bBig)
                    {
                        strValue = (String)m_hashSimpChNum[strCurCh];
                    }
                    else
                    {
                        strValue = (String)m_hashBigSimpChNum[strCurCh];
                    }

                    // strLevel = (String)hashSimpChLevel[nLvl];
                }

                strFloatResult += strValue + strLevel;

                strPreCh = strCurCh;
            }


            if (Math.Abs(dbIntegerPart) >= 1.0)
            {
                if (bMoney)
                {
                    strIntegerResult += "元";
                }
            }
            else
            {
                if (!bMoney)
                {
                    if (!bBig)
                    {
                        strIntegerResult = (String)m_hashSimpChNum["0"];
                    }
                    else
                    {
                        strIntegerResult = (String)m_hashBigSimpChNum["0"];
                    }
                }
                else
                {
                    if (dbNum == 0.0)
                    {
                        strIntegerResult = (String)m_hashBigSimpChNum["0"];
                        strIntegerResult += "元";
                    }
                }
            }


            // if (Math.Abs(dbFloatPart) > 0.0)
            if (nPointPos != -1)
            {
                if (!bMoney)
                {
                    strFloatResult = "点" + strFloatResult;
                }

            }

            strResult = strPreFix + strIntegerResult + strFloatResult;

            return strResult;
        }

        // 
        public String num2SimpValueString1by1(double dbNum, Boolean bBig = true, Boolean bMoney = false)
        {
            String strNum = dbNum.ToString();

            int i = 0;

            int nPointPos = strNum.IndexOf('.');

            String strIntegerPart = strNum, strFloatPart = "";
            double dbIntegerPart = 0.0, dbFloatPart = 0.0;

            if (nPointPos != -1)
            {
                strIntegerPart = strNum.Substring(0, nPointPos);
                strFloatPart = strNum.Substring(nPointPos + 1); // include point

                dbFloatPart = Convert.ToDouble(strFloatPart);
            }

            dbIntegerPart = Convert.ToDouble(strIntegerPart);


            int nLvl = 0;

            String strValue = "", strLevel = "", strIntegerResult = "", strFloatResult = "";
            String strResult = "";

            String strPreCh = "", strCurCh = "";
            String strScannedValue = "";
            double dbCurValue = 0.0;

            for (i = strIntegerPart.Length - 1, nLvl = 0; i >= 0; i--, nLvl++)
            {
                strValue = "";
                strLevel = "";

                strCurCh = strIntegerPart[i] + "";

                if (!bBig)
                {
                    strValue = (String)m_hashSimpChNum[strCurCh];
                }
                else
                {
                    strValue = (String)m_hashBigSimpChNum[strCurCh];
                }

                strLevel = (String)m_hashSimpChLevel[nLvl];

                strIntegerResult = strValue + strLevel + strIntegerResult;

            }



            // float
            strPreCh = "";
            strCurCh = "";
            strScannedValue = "";
            dbCurValue = 0.0;

            //for (i = strFloatPart.Length - 1, nLvl = -1 * strFloatPart.Length; i >= 0; i--, nLvl++)
            for (i = 0, nLvl = -1; i < strFloatPart.Length; i++, nLvl--)
            {
                strValue = "";
                strLevel = "";

                strCurCh = strFloatPart[i] + "";

                if (!bBig)
                {
                    strValue = (String)m_hashSimpChNum[strCurCh];
                }
                else
                {
                    strValue = (String)m_hashBigSimpChNum[strCurCh];
                }

                if (bMoney)
                {
                    strLevel = (String)m_hashSimpChLevel[nLvl];
                }

                strFloatResult += strValue + strLevel;

                strPreCh = strCurCh;
            }


            if (Math.Abs(dbIntegerPart) >= 1.0)
            {
                if (bMoney)
                {
                    strIntegerResult += "元";
                }
            }
            else
            {
                if (!bMoney)
                {
                    if (!bBig)
                    {
                        strIntegerResult = (String)m_hashSimpChNum["0"];
                    }
                    else
                    {
                        strIntegerResult = (String)m_hashBigSimpChNum["0"];
                    }
                }
                else
                {
                    if (dbNum == 0.0)
                    {
                        strIntegerResult = (String)m_hashBigSimpChNum["0"];
                        strIntegerResult += "元";
                    }
                }
            }


            if (Math.Abs(dbFloatPart) > 0.0)
            {
                if (!bMoney)
                {
                    strFloatResult = "点" + strFloatResult;
                }

            }

            strResult = strIntegerResult + strFloatResult;

            return strResult;
        }

        #endregion


        private Word.WdBuiltinStyle OutlineLevel2BuiltinStyle(int nOutlineLevel)
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

        public void BulkPromote(Word.Document doc, Word.Paragraphs paras, Boolean bOnlyHeadings = false)
        {
            //             Word.Document doc = m_ownerAddin.Application.ActiveDocument;
            Word.Selection sel = doc.ActiveWindow.Selection;

            //             Word.Paragraphs paras = null;
            int nOStart = sel.Start;
            int nOEnd = sel.End;


            String strItem = "";

            if (bOnlyHeadings)
            {
                foreach (Word.Paragraph para in paras)
                {
                    doc.ActiveWindow.ScrollIntoView(para.Range);

                    strItem = para.Range.Text.Trim(m_trimChars);

                    if (String.IsNullOrWhiteSpace(strItem))
                    {
                        continue;
                    }

                    if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        para.OutlinePromote();
                    }
                }
            }
            else
            {
                foreach (Word.Paragraph para in paras)
                {
                    doc.ActiveWindow.ScrollIntoView(para.Range);

                    //if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    strItem = para.Range.Text.Trim(m_trimChars);

                    if (String.IsNullOrWhiteSpace(strItem))
                    {
                        continue;
                    }

                    if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        para.OutlinePromote();
                    }
                    else
                    {
                        //para.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel9;
                        Word.WdBuiltinStyle styleIndex = OutlineLevel2BuiltinStyle(9);
                        Object objStyle = doc.Styles[styleIndex];
                        para.set_Style(objStyle);
                    }
                }

                // paras.OutlinePromote();
            }

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return;

        }


        public void BulkDemote(Word.Document doc, Word.Paragraphs paras, Boolean bOnlyHeadings = false)
        {
            //             Word.Document doc = m_ownerAddin.Application.ActiveDocument;
            Word.Selection sel = doc.ActiveWindow.Selection;
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            String strItem = "";

            if (bOnlyHeadings)
            {
                foreach (Word.Paragraph para in paras)
                {
                    doc.ActiveWindow.ScrollIntoView(para.Range);

                    strItem = para.Range.Text.Trim(m_trimChars);

                    if (String.IsNullOrWhiteSpace(strItem))
                    {
                        continue;
                    }

                    if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        if (para.OutlineLevel < Word.WdOutlineLevel.wdOutlineLevel9)
                        {
                            para.OutlineDemote();
                        }
                        else
                        {
                            para.OutlineDemoteToBody();
                        }
                    }
                }
            }
            else
            {
                foreach (Word.Paragraph para in paras)
                {
                    doc.ActiveWindow.ScrollIntoView(para.Range);

                    strItem = para.Range.Text.Trim(m_trimChars);

                    if (String.IsNullOrWhiteSpace(strItem))
                    {
                        continue;
                    }

                    if (para.OutlineLevel < Word.WdOutlineLevel.wdOutlineLevel9)
                    {
                        para.OutlineDemote();
                    }
                    else
                    {
                        para.OutlineDemoteToBody();
                    }
                }
            }

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return;
        }

        public Boolean isIsolatePic(Word.Paragraph para)
        {
            Boolean bRet = false;

//             if (para.Range.InlineShapes.Count == 0)
//             {
//                 return bRet;
//             }

            // char[] trimChars = new char[] {'　',' ', '\t', '\r', '\n', '\a', '\f'};

            String strCnt = "";
            strCnt = para.Range.Text.Trim(m_trimChars);

//             strCnt = para.Range.Text.Replace(" ", "");
//             strCnt = strCnt.Replace("　", "");
//             strCnt = strCnt.Replace("\t", "");
//             strCnt = strCnt.Replace("\r", "");
//             strCnt = strCnt.Replace("\n", "");

            char[] chRes = strCnt.ToCharArray();
            int nCnt = 0;

            foreach (char ch in chRes)
            {
                if (ch != '/' && ch != 1 && ch != 21 && ch != ' ' &&
                    ch != '\t' && ch != '\f' && ch != '\a')
                {
                    return bRet;
                }
                else
                {
                    nCnt++;// 连续图片
                }
            }

            return true;
        }



        public int alignCurPic(Word.Paragraph para, Word.WdParagraphAlignment wAlign)
        {
            int nRet = -1;

            if (para.Range.InlineShapes.Count == 0)
            {
                return nRet;
            }

            String strCnt = "";

            strCnt = para.Range.Text.Trim(m_trimChars);

            //strCnt = para.Range.Text.Replace(" ", "");
            //strCnt = strCnt.Replace("　", "");
            //strCnt = strCnt.Replace("\t", "");
            //strCnt = strCnt.Replace("\r", "");
            //strCnt = strCnt.Replace("\n", "");

            char[] chRes = strCnt.ToCharArray();
            int nCnt = 0;

            foreach (char ch in chRes)
            {
                if (ch != '/' && ch != 1)
                {
                    return nRet;
                }
                else
                {
                    nCnt++;// 连续图片
                }
            }

            // 
            if (para.Range.InlineShapes.Count == nCnt)
            {
                para.Alignment = wAlign;// Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

            return nRet;
        }


        public int alignAllPicsInSel(Word.Paragraphs paras, Word.WdParagraphAlignment wAlign)
        {
            int nRet = -1;

            if (paras == null)
                return nRet;

            foreach (Word.Paragraph para in paras)
            {
                alignCurPic(para, wAlign);
            }
            nRet = 0;
            return nRet;
        }


        public int alignAllTablesInSel(Word.Tables tbls, Word.WdRowAlignment wAlign)
        {
            int nRet = -1;

            if (tbls == null)
                return nRet;

            foreach (Word.Table tbl in tbls)
            {
                tbl.Rows.Alignment = wAlign;
            }
            nRet = 0;
            return nRet;
        }


        public ArrayList get9HeadingParas(Word.Document curDoc)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null, nextPara = null;

            for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            {
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
                sel.Find.ClearFormatting();

                // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                sel.Find.ParagraphFormat.OutlineLevel = lvl;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";
                sel.Find.Forward = true;
                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                bool bRet = sel.Find.Execute();

                if(sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)) && para.OutlineLevel == lvl)
                        {
                            arrParas.Add(para);
                        }

                        nextPara = para.Next();

                        if (null == nextPara)
                        {
                            continue;
                        }
                    }

                    
                    // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    //if (nextPara != null)
                    //{
                    //    sel.Start = nextPara.Range.Start;
                    //    sel.End = sel.Start;
                    //    sel.Range.GoTo();
                    //}

                    // bRet = sel.Find.Execute();
                }

            }// for


            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            // ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            // arrParas.Sort(paraComparer);

            return arrParas;
        }


        public Word.Paragraph getOneHeadingPara(Word.Document curDoc,int nOutlineLevel, Boolean bForward = true,Boolean bCanEmpty = false)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            Word.Paragraph para = null;

            // sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            sel.Find.ClearFormatting();

            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nOutlineLevel;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = bForward;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            if (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    Word.Paragraph tmpPara = sel.Range.Paragraphs[1];

                    if (tmpPara.OutlineLevel == (Word.WdOutlineLevel)nOutlineLevel)
                    {
                        if (bCanEmpty)
                        {
                            para = sel.Range.Paragraphs[1];
                        }
                        else if (!String.IsNullOrWhiteSpace(tmpPara.Range.Text.Trim(m_trimChars)))
                        {
                            para = sel.Range.Paragraphs[1];
                        }
                    }
                }
            }

            sel.Find.ClearFormatting();

            return para;
        }


        public ArrayList getHeadingParas(Word.Document curDoc)
        {
            return getSpecificHeadingParasInScope(curDoc);
        }


        private ArrayList getHeadingParas_v1(Word.Document curDoc)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            // 切换到normal view
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null;

            for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            {
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
                sel.Find.ClearFormatting();

                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                sel.Find.ParagraphFormat.OutlineLevel = lvl;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";
                sel.Find.Forward = true;
                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                bool bRet = sel.Find.Execute();

                while (sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (para.OutlineLevel != lvl)
                        {
                            break;
                        }

                        if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                        {
                            arrParas.Add(para);
                        }

                        if (null == para.Next())
                        {
                            break;
                        }
                    }

                    /*
                    para = para.Next();
                    if (para != null)
                    {
                        //sel.Start = para.Range.Start;
                        //sel.End = sel.Start;
                        para.Range.Select();
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    }
                    else
                    {
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                    */

                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    bRet = sel.Find.Execute();
                }// while

            }// for


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            // 恢复特定view
            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            arrParas.Sort(paraComparer);

            return arrParas;
        }

        public ArrayList getTextBodyParas(Word.Document curDoc, Boolean bIgnoreTables = false, Boolean bIgnoreTocs = false)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            // 切换到normal view
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null, nextPara = null;
            Boolean bIgnored = false;
            String strTmp = "";
            Word.Range ignoreRng = null;

            sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            sel.Find.ClearFormatting();

            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];
                    strTmp = para.Range.Text.Trim(m_trimChars);

                    if (String.IsNullOrWhiteSpace(strTmp))
                    {
                        if (bIgnoreTables)
                        {
                            Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                            if (bInTable)
                            {
                                bIgnored = true;
                                ignoreRng = para.Range.Tables[1].Range;
                            }
                        }

                        nextPara = para.Next();
                        if (nextPara == null)
                        {
                            break;
                        }

                        if (bIgnored)
                        {
                            // ignoreRng.Select();
                            if (ignoreRng != null)
                            {
                                sel.Start = ignoreRng.End;
                                sel.End = sel.Start;
                                sel.Range.GoTo();
                            }
                        }
                        else
                        {
                            if (nextPara != null)
                            {
                                sel.Start = nextPara.Range.Start;
                                sel.End = sel.Start;
                                sel.Range.GoTo();
                            }
                        }
                        
                        bIgnored = false;
                        bRet = sel.Find.Execute();
                        continue;
                    }

                    bIgnored = false;

                    if (bIgnoreTables)
                    {
                        Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                        if (bInTable)
                        {
                            bIgnored = true;
                            ignoreRng = para.Range.Tables[1].Range;
                        }
                    }

                    if (bIgnoreTocs)
                    {
                        foreach (Word.TableOfContents toc in curDoc.TablesOfContents)
                        {
                            if (para.Range.InRange(toc.Range))
                            {
                                bIgnored = true;
                                ignoreRng = toc.Range;
                                break;
                            }
                        }

                        foreach (Word.TableOfFigures figs in curDoc.TablesOfFigures)
                        {
                            if (para.Range.InRange(figs.Range))
                            {
                                bIgnored = true;
                                ignoreRng = figs.Range;
                                break;
                            }
                        }
                    }

                    if (!bIgnored)
                    {
                        // judge isolate pics of this para
                        if (para.Range.InlineShapes.Count > 0)
                        {
                            if (!isIsolatePic(para))
                            {
                                if (!String.IsNullOrWhiteSpace(para.Range.Text))
                                {
                                    arrParas.Add(para);
                                }
                            }
                        }
                        else
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text))
                            {
                                arrParas.Add(para);
                            }
                        }
                    }

                    nextPara = para.Next();
                    if (nextPara == null)
                    {
                        break;
                    }

                    if (bIgnored)
                    {
                        // ignoreRng.Select();
                        if (ignoreRng != null)
                        {
                            sel.Start = ignoreRng.End;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                    else
                    {
                        if (nextPara != null)
                        {
                            sel.Start = nextPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                }
                else
                {
                    break;
                }


                bRet = sel.Find.Execute();
            }// while

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }

        /*
        public ArrayList getTextBodyParas(Word.Document curDoc,Boolean bIgnoreTables = false, Boolean bIgnoreTocs = false)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null;
            Boolean bIgnored = false;
            String strTmp = "";

            sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            sel.Find.ClearFormatting();

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];
                    strTmp = para.Range.Text.Trim(m_trimChars);

                    if (strTmp.Equals(""))
                    {
                        if (null == para.Next())
                        {
                            break;
                        }

                        sel.Start = para.Next().Range.Start;
                        sel.End = sel.Start;
                        // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        bRet = sel.Find.Execute();

                        continue;
                    }

                    bIgnored = false;

                    if (bIgnoreTables)
                    {
                        Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                        if (bInTable)
                        {
                            bIgnored = true;
                        }
                    }

                    if (bIgnoreTocs)
                    {
                        foreach(Word.TableOfContents toc in  curDoc.TablesOfContents)
                        {
                            if (para.Range.InRange(toc.Range))
                            {
                                bIgnored = true;
                                break;
                            }
                        }
                    }

                    if (!bIgnored)
                    {
                        // judge isolate pics of this para
                        if (para.Range.InlineShapes.Count > 0 )
                        {
                            if (!isIsolatePic(para))
                            {
                                arrParas.Add(para);
                            }
                        }
                        else
                        {
                            arrParas.Add(para);
                        }
                    }

                    if (null == para.Next())
                    {
                        break;
                    }
                }

                para = para.Next();
                if (para != null)
                {
                    sel.Start = para.Range.Start;
                    sel.End = sel.Start;
                }
                else
                {
                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                bRet = sel.Find.Execute();
            }// while

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }
        */


        public Boolean RangeOverlap(Word.Range rng1, Word.Range rng2)
        {
            Boolean bRet = false;

            if (rng1.Start <= rng2.Start && rng1.End >= rng2.End)
            {
                bRet = true;
            }
            else if(rng2.Start <= rng1.Start && rng2.End >= rng1.End)
            {
                bRet = true;
            }
            else if(rng1.Start <= rng2.Start && rng1.End > rng2.Start && rng1.End <= rng2.End)
            {
                bRet = true;
            }
            else if(rng2.Start <= rng1.Start && rng2.End > rng1.Start && rng2.End <= rng1.End)
            {
                bRet = true;
            }


            return bRet;
        }

        public ArrayList getSpecificTextBodyParasInScopeNoChangeView(Word.Document curDoc, Word.Range scopeRange = null, Boolean bIgnoreTables = false, Boolean bIgnoreTocs = false,Boolean bIgnorePages = false, uint nIgnorePages = 0)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;


            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null,nextPara = null;
            Word.Range ignoreRng = null;
            int nLvl = 0;
            Boolean bIgnored = false;
            String strTmp = "";
            Boolean bInToc = false, bInTables = false, bInIgnorePages = false;

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }


            nLvl = (int)Word.WdOutlineLevel.wdOutlineLevelBodyText;

            //if (scopeRange == null)
            //{
            //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            //}
            //else
            //{
            //    if (scopeRange.Paragraphs.Count > 0)
            //    {
            //        para = scopeRange.Paragraphs[1];

            //        sel.Start = para.Range.Start;
            //        sel.End = sel.Start;
            //    }
            //    else
            //    {
            //        sel.Start = nStartPos;
            //        sel.End = sel.Start;
            //    }

            //    //sel.HomeKey(Word.WdUnits.wdParagraph,Word.WdMovementType.wdMove);
            //    // sel.End = nEndPos;
            //    sel.Range.GoTo();
            //}

            sel.Range.GoTo();
            sel.Find.ClearFormatting();

            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];

                    if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                    //if (para.Range.InRange(scopeRange))
                    {
                        strTmp = para.Range.Text.Trim(m_trimChars);

                        // if (strTmp == "")
                        if (String.IsNullOrWhiteSpace(strTmp)) // ???
                        {
                            if (bIgnoreTables)
                            {
                                Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                                if (bInTable)
                                {
                                    bIgnored = true;
                                    ignoreRng = para.Range.Tables[1].Range;
                                }
                            }

                            nextPara = para.Next();
                            if (nextPara == null)
                            {
                                break;
                            }

                            if (bIgnored)
                            {
                                // ignoreRng.Select();
                                if (ignoreRng != null)
                                {
                                    sel.Start = ignoreRng.End;
                                    sel.End = sel.Start;
                                    sel.Range.GoTo();
                                }
                            }
                            else
                            {
                                if (nextPara != null)
                                {
                                    sel.Start = nextPara.Range.Start;
                                    sel.End = sel.Start;
                                    sel.Range.GoTo();
                                }
                            }

                            //if (bIgnored)
                            //{
                            //    ignoreRng.Select();
                            //}
                            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            bIgnored = false;
                            bRet = sel.Find.Execute();
                            continue;
                        }

                        bIgnored = false;
                        
                        bInToc = false;
                        bInTables = false;
                        bInIgnorePages = false;

                        if (bIgnoreTables)
                        {
                            Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                            if (bInTable)
                            {
                                bIgnored = true;
                                ignoreRng = para.Range.Tables[1].Range;
                            }
                        }

                        if (bIgnoreTocs)
                        {
                            foreach (Word.TableOfContents toc in curDoc.TablesOfContents)
                            {
                                if (para.Range.InRange(toc.Range))
                                {
                                    bIgnored = true;
                                    ignoreRng = toc.Range;
                                    break;
                                }
                            }
                        }

                        if (bIgnorePages)
                        {
                            int nPageSn = para.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                            bInIgnorePages = (nPageSn <= nIgnorePages);

                            ignoreRng = para.Range;
                        }

                        if ((bIgnoreTocs && bInToc) || (bIgnoreTables && bInTables) || (bIgnorePages && bInIgnorePages))
                        {
                            bIgnored = true;
                        }

                        if (!bIgnored)
                        {
                            // judge isolate pics of this para
                            if (para.Range.InlineShapes.Count > 0)
                            {
                                if (!isIsolatePic(para))
                                {
                                    if (!String.IsNullOrWhiteSpace(para.Range.Text))
                                    {
                                        arrParas.Add(para);
                                    }
                                }
                            }
                            else
                            {
                                if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                                {
                                    arrParas.Add(para);
                                }
                            }
                        }

                        nextPara = para.Next();
                        if (nextPara == null)
                        {
                            break;
                        }

                        if (bIgnored)
                        {
                            // ignoreRng.Select();
                            if (ignoreRng != null)
                            {
                                sel.Start = ignoreRng.End;
                                sel.End = sel.Start;
                                sel.Range.GoTo();
                            }
                        }
                        else
                        {
                            if (nextPara != null)
                            {
                                sel.Start = nextPara.Range.Start;
                                sel.End = sel.Start;
                                sel.Range.GoTo();
                            }
                        }

                        //if (bIgnored)
                        //{
                        //    ignoreRng.Select();
                        //}
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                    else
                    {
                        break;
                    }
                }
                else
                {
                    break;
                }

                bRet = sel.Find.Execute();
            }// while

            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }


        public ArrayList getSpecificTextBodyParasInScope(Word.Document curDoc, Word.Range scopeRange = null, Boolean bIgnoreTables = false, Boolean bIgnoreTocs = false)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null,nextPara = null;
            Word.Range ignoreRng = null;
            int nLvl = 0;
            Boolean bIgnored = false;
            String strTmp = "";

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)

            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            nLvl = (int)Word.WdOutlineLevel.wdOutlineLevelBodyText;

            //if (scopeRange == null)
            //{
            //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            //}
            //else
            //{
            //    if (scopeRange.Paragraphs.Count > 0)
            //    {
            //        para = scopeRange.Paragraphs[1];

            //        sel.Start = para.Range.Start;
            //        sel.End = sel.Start;
            //    }
            //    else
            //    {
            //        sel.Start = scopeRngStartPos;
            //        sel.End = sel.Start;
            //    }

            //    //sel.HomeKey(Word.WdUnits.wdParagraph,Word.WdMovementType.wdMove);
            //    // sel.End = nEndPos;
            //}

            sel.Start = scopeRngStartPos;
            sel.End = sel.Start;

            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];

                    if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                    //if (para.Range.InRange(scopeRange))
                    {
                        strTmp = para.Range.Text.Trim(m_trimChars);

                        // if (strTmp.Equals(""))
                        if(String.IsNullOrWhiteSpace(strTmp))
                        {
                            if (bIgnoreTables)
                            {
                                Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                                if (bInTable)
                                {
                                    bIgnored = true;
                                    ignoreRng = para.Range.Tables[1].Range;
                                }
                            }

                            nextPara = para.Next();
                            if (nextPara == null)
                            {
                                break;
                            }

                            //if (bIgnored)
                            //{
                            //    ignoreRng.Select();
                            //}

                            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            if (bIgnored)
                            {
                                // ignoreRng.Select();
                                if (ignoreRng != null)
                                {
                                    sel.Start = ignoreRng.End;
                                    sel.End = sel.Start;
                                    sel.Range.GoTo();
                                }
                            }
                            else
                            {
                                if (nextPara != null)
                                {
                                    sel.Start = nextPara.Range.Start;
                                    sel.End = sel.Start;
                                    sel.Range.GoTo();
                                }
                            }

                            bIgnored = false;
                            
                            bRet = sel.Find.Execute();
                            continue;
                        }

                        bIgnored = false;

                        if (bIgnoreTables)
                        {
                            Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                            if (bInTable)
                            {
                                bIgnored = true;
                                ignoreRng = para.Range.Tables[1].Range;
                            }
                        }

                        if (bIgnoreTocs)
                        {
                            foreach (Word.TableOfContents toc in curDoc.TablesOfContents)
                            {
                                if (para.Range.InRange(toc.Range))
                                {
                                    bIgnored = true;
                                    ignoreRng = toc.Range;
                                    break;
                                }
                            }
                        }

                        if (!bIgnored)
                        {
                            // judge isolate pics of this para
                            if (para.Range.InlineShapes.Count > 0)
                            {
                                if (!isIsolatePic(para))
                                {
                                    if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                                    {
                                        arrParas.Add(para);
                                    }
                                }
                            }
                            else
                            {
                                if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                                {
                                    arrParas.Add(para);
                                }
                            }
                        }

                        nextPara = para.Next();
                        if (nextPara == null)
                        {
                            break;
                        }

                        //if (bIgnored)
                        //{
                        //    ignoreRng.Select();
                        //}

                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        if (bIgnored)
                        {
                            // ignoreRng.Select();
                            if (ignoreRng != null)
                            {
                                sel.Start = ignoreRng.End;
                                sel.End = sel.Start;
                                sel.Range.GoTo();
                            }
                        }
                        else
                        {
                            if (nextPara != null)
                            {
                                sel.Start = nextPara.Range.Start;
                                sel.End = sel.Start;
                                sel.Range.GoTo();
                            }
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                else
                {
                    break;
                }

                bRet = sel.Find.Execute();
            }// while

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }

        /*
        public ArrayList getSpecificTextBodyParasInScope(Word.Document curDoc, Word.Range scopeRange = null, Boolean bIgnoreTables = false, Boolean bIgnoreTocs = false)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null;
            int nLvl = 0;
            Boolean bIgnored = false;
            String strTmp = "";

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            nLvl = (int)Word.WdOutlineLevel.wdOutlineLevelBodyText;

            if (scopeRange == null)
            {
                sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            }
            else
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];

                    sel.Start = para.Range.Start;
                    sel.End = sel.Start;
                }
                else
                {
                    sel.Start = nStartPos;
                    sel.End = sel.Start;
                }

                //sel.HomeKey(Word.WdUnits.wdParagraph,Word.WdMovementType.wdMove);
                // sel.End = nEndPos;
            }

            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];

                    if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                    //if (para.Range.InRange(scopeRange))
                    {
                        strTmp = para.Range.Text.Trim(m_trimChars);

                        if (strTmp.Equals(""))
                        {
                            if (null == para.Next())
                            {
                                break;
                            }

                            sel.Start = para.Next().Range.Start;
                            sel.End = sel.Start;
                            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            bRet = sel.Find.Execute();

                            continue;
                        }

                        bIgnored = false;

                        if (bIgnoreTables)
                        {
                            Boolean bInTable = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                            if (bInTable)
                            {
                                bIgnored = true;
                            }
                        }

                        if (bIgnoreTocs)
                        {
                            foreach (Word.TableOfContents toc in curDoc.TablesOfContents)
                            {
                                if (para.Range.InRange(toc.Range))
                                {
                                    bIgnored = true;
                                    break;
                                }
                            }
                        }

                        if (!bIgnored)
                        {
                            // judge isolate pics of this para
                            if (para.Range.InlineShapes.Count > 0)
                            {
                                if (!isIsolatePic(para))
                                {
                                    arrParas.Add(para);
                                }
                            }
                            else
                            {
                                arrParas.Add(para);
                            }
                        }
                            
                    }
                    else
                    {
                        break;
                    }

                    if (null == para.Next())
                    {
                        break;
                    }
                }

                para = para.Next();
                if (para != null)
                {
                    sel.Start = para.Range.Start;
                    sel.End = sel.Start;
                }
                else
                {
                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                bRet = sel.Find.Execute();
            }// while

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }
        */

        public ArrayList getSpecificHeadingParaArrsInScope(Word.Document curDoc, Word.Range scopeRange = null, 
                                                           int[] nArrOutlineLevels = null,
                                                           Boolean bIgnoreToc = false, Boolean bIgnoreTable = false,
                                                           Boolean bIgnorePages = false, uint nIgnorePages = 1)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;


            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Boolean bForward = true;

            ArrayList arrParas = null;
            ArrayList arrs = new ArrayList();

            for (int i = 0; i <= 9; i++)
            {
                arrs.Add(new ArrayList());
            }

            Word.Paragraph para = null,prevPara = null,nextPara = null;
            int nLvl = 0;

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            int[] nArrLocalOutlineLevels = new int[10];

            if (nArrOutlineLevels == null)
            {
                int j = 0;
                for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
                {
                    nArrLocalOutlineLevels[j] = (int)lvl;
                    j++;
                }
            }
            else
            {
                for (int i = 0; i < nArrOutlineLevels.GetLength(0); i++)
                {
                    nArrLocalOutlineLevels[i] = nArrOutlineLevels[i];
                }
            }


            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            Boolean bInToc = false;
            Boolean bInTables = false;
            Boolean bInIgnorePages = false;
            Boolean bRet = false;

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            // for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            for (int i = 0; i < nArrLocalOutlineLevels.GetLength(0); i++)
            {
                nLvl = nArrLocalOutlineLevels[i];
                if (nLvl < (int)Word.WdOutlineLevel.wdOutlineLevel1 || nLvl > (int)Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    //return null;
                    continue;
                }

                bInToc = false;
                bInTables = false;
                bInIgnorePages = false;

                arrParas = (ArrayList)arrs[nLvl];

                //if (scopeRange == null)
                //{
                //    if (bForward)
                //    {
                //        sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
                //    }
                //    else
                //    {
                //        sel.EndKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);
                //    }
                //}
                //else
                //{
                //    if (bForward)
                //    {
                //        //if (sel.Range.Paragraphs.Count > 0)
                //        if (scopeRange.Paragraphs.Count > 0)
                //        {
                //            // para = sel.Range.Paragraphs[1];
                //            para = scopeRange.Paragraphs[1];

                //            sel.Start = para.Range.Start;
                //            sel.End = sel.Start;
                //        }
                //        else
                //        {
                //            sel.Start = nStartPos;
                //            sel.End = sel.Start;
                //        }
                //    }
                //    else
                //    {
                //        if (scopeRange.Paragraphs.Count > 0)
                //        {
                //            // para = sel.Range.Paragraphs[1];
                //            para = scopeRange.Paragraphs[scopeRange.Paragraphs.Count];

                //            sel.Start = para.Range.End;
                //            sel.End = sel.Start;
                //        }
                //        else
                //        {
                //            sel.Start = nEndPos;
                //            sel.End = sel.Start;
                //        }
                //    }

                //    // sel.End = nEndPos;
                //    sel.Range.GoTo();
                //}

                if (bForward)
                {
                    sel.Start = scopeRngStartPos;
                }
                else
                {
                    sel.Start = scopeRngEndPos;
                }
                sel.End = sel.Start;
                sel.Range.GoTo();
                sel.Find.ClearFormatting();

                sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";

                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                sel.Find.Forward = bForward;

                bRet = sel.Find.Execute();

                while (sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (para.OutlineLevel != (Word.WdOutlineLevel)nLvl)
                        {
                            break;
                        }

                        if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                        //if (para.Range.InRange(scopeRange))
                        {
                            bInToc = false;
                            bInTables = false;
                            bInIgnorePages = false;

                            if (bIgnoreToc)
                            {
                                foreach (Word.TableOfContents toc in curDoc.TablesOfContents)
                                {
                                    if (para.Range.InRange(toc.Range))
                                    {
                                        bInToc = true;
                                        break;
                                    }
                                }

                            }//

                            if (bIgnoreTable)
                            {
                                bInTables = para.Range.get_Information(Word.WdInformation.wdWithInTable);
                            }


                            if (bIgnorePages)
                            {
                                int nPageSn = para.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                                bInIgnorePages = (nPageSn <= nIgnorePages);
                            }

                            if (!((bIgnoreToc && bInToc) || (bIgnoreTable && bInTables) || (bIgnorePages && bInIgnorePages)))
                            {
                                if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                                {
                                    arrParas.Add(para);
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nextPara = para.Next();
                            if (null == nextPara)
                            {
                                break;
                            }
                        }
                        else
                        {
                            prevPara = para.Previous();
                            if (null == prevPara)
                            {
                                break;
                            }
                        }
                    }

                    /*
                    para = para.Next();
                    if (para != null)
                    {
                        //sel.Start = para.Range.Start;
                        //sel.End = sel.Start;
                        // para.Range.Select();
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    }
                    else
                    {
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                    */

                    if (bForward)
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        //nextPara = para.Next();
                        if (nextPara != null)
                        {
                            sel.Start = nextPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                    else
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        //prevPara = para.Previous();
                        if (prevPara != null)
                        {
                            sel.Start = prevPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                    bRet = sel.Find.Execute();
                }// while

                if (!bAppIsWps)
                {
                    bForward = !bForward;
                }

                // bForward = true;

            }// for

            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            return arrs;
        }


        public ArrayList getSpecificHeadingParaArrsInScope(Word.Document curDoc, Word.Range scopeRange = null, int[] nArrOutlineLevels = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            //int nOStart = sel.Start;
            //int nOEnd = sel.End;

            //Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Boolean bForward = true;

            ArrayList arrParas = null;
            ArrayList arrs = new ArrayList();

            for (int i = 0; i <= 9; i++)
            {
                arrs.Add(new ArrayList());
            }

            Word.Paragraph para = null,prevPara = null, nextPara = null;
            int nLvl = 0;

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            int[] nArrLocalOutlineLevels = new int[10];

            if (nArrOutlineLevels == null)
            {
                int j = 0;
                for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
                {
                    nArrLocalOutlineLevels[j] = (int)lvl;
                    j++;
                }
            }
            else
            {
                for (int i = 0; i < nArrOutlineLevels.GetLength(0); i++)
                {
                    nArrLocalOutlineLevels[i] = nArrOutlineLevels[i];
                }
            }


            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            // for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            for (int i = 0; i < nArrLocalOutlineLevels.GetLength(0); i++)
            {
                nLvl = nArrLocalOutlineLevels[i];
                if (nLvl < (int)Word.WdOutlineLevel.wdOutlineLevel1 || nLvl > (int)Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    //return null;
                    continue;
                }

                arrParas = (ArrayList)arrs[nLvl];

                if (bForward)
                {
                    sel.Start = scopeRngStartPos;
                }
                else
                {
                    sel.Start = scopeRngEndPos;
                }
                sel.End = sel.Start;
                sel.Range.GoTo();
                                
                sel.Find.ClearFormatting();
                sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";

                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                sel.Find.Forward = bForward;

                bool bRet = sel.Find.Execute();

                while (sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (para.OutlineLevel != (Word.WdOutlineLevel)nLvl)
                        {
                            break;
                        }

                        if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                        //if (para.Range.InRange(scopeRange))
                        {
                            // ignore empty para
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                arrParas.Add(para);
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nextPara = para.Next();
                            if (null == nextPara)
                            {
                                break;
                            }
                        }
                        else
                        {
                            prevPara = para.Previous();
                            if (null == prevPara)
                            {
                                break;
                            }
                        }
                    }

                    if (bForward)
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        //nextPara = para.Next();
                        if (nextPara != null)
                        {
                            sel.Start = nextPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                    else
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        //prevPara = para.Previous();
                        if (prevPara != null)
                        {
                            sel.Start = prevPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                    bRet = sel.Find.Execute();
                }// while

                if (!bAppIsWps)
                {
                    bForward = !bForward;
                }
                // bForward = true;

            }// for


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            //curDoc.ActiveWindow.View.Type = oViewType;

            //sel.Start = nOStart;
            //sel.End = nOEnd;
            // sel.Range.Select();
            //sel.Range.GoTo();
            //curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrs;
        }



        public ArrayList getSpecificHeadingParasInScope(Word.Document curDoc, Word.Range scopeRange = null, int[] nArrOutlineLevels = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Boolean bForward = true;
            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null;
            int nLvl = 0;

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            int[] nArrLocalOutlineLevels = new int[10];

            if (nArrOutlineLevels == null)
            {
                int j = 0;
                for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
                {
                    nArrLocalOutlineLevels[j] = (int)lvl;
                    j++;
                }
            }
            else
            {
                for (int i = 0; i < nArrOutlineLevels.GetLength(0); i++)
                {
                    nArrLocalOutlineLevels[i] = nArrOutlineLevels[i];
                }
            }


            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            Word.Paragraph prevPara = null, nextPara = null;

            // for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            for (int i = 0; i < nArrLocalOutlineLevels.GetLength(0); i++)
            {
                nLvl = nArrLocalOutlineLevels[i];
                if (nLvl < (int)Word.WdOutlineLevel.wdOutlineLevel1 || nLvl > (int)Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    //return null;
                    continue;
                }

                if (bForward)
                {
                    sel.Start = scopeRngStartPos;
                }
                else
                {
                    sel.Start = scopeRngEndPos;
                }
                sel.End = sel.Start;
                sel.Range.GoTo();

                sel.Find.ClearFormatting();

                sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;
                //sel.Find.ParagraphFormat.WordWrap = -1;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";

                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                sel.Find.Forward = bForward;

                bool bRet = sel.Find.Execute();

                while (sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (para.OutlineLevel != (Word.WdOutlineLevel)nLvl)
                        {
                            break;
                        }

                        if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                        //if (para.Range.InRange(scopeRange))
                        {
                            // ignore empty para
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                arrParas.Add(para);
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nextPara = para.Next();
                            if (null == nextPara)
                            {
                                break;
                            }
                        }
                        else
                        {
                            prevPara = para.Previous();
                            if (null == prevPara)
                            {
                                break;
                            }
                        }

                    }

                    if (bForward)
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        //nextPara = para.Next();
                        if (nextPara != null)
                        {
                            sel.Start = nextPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                    else
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        //prevPara = para.Previous();
                        if (prevPara != null)
                        {
                            sel.Start = prevPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                    bRet = sel.Find.Execute();
                }// while

                if (!bAppIsWps)
                {
                    bForward = !bForward;
                }

            }// for


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            arrParas.Sort(paraComparer);

            return arrParas;
        }


        private ArrayList getSpecificHeadingParasInScope_v1(Word.Document curDoc, Word.Range scopeRange = null, int[] nArrOutlineLevels = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null;
            int nLvl = 0;

            //for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            int[] nArrLocalOutlineLevels = new int[10];

            if (nArrOutlineLevels == null)
            {
                int j = 0;
                for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
                {
                    nArrLocalOutlineLevels[j] = (int)lvl;
                    j++;
                }
            }
            else
            {
                for (int i = 0; i < nArrOutlineLevels.GetLength(0); i++)
                {
                    nArrLocalOutlineLevels[i] = nArrOutlineLevels[i];
                }
            }


            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            // for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            for (int i = 0; i < nArrLocalOutlineLevels.GetLength(0); i++)
            {
                nLvl = nArrLocalOutlineLevels[i];
                if (nLvl < (int)Word.WdOutlineLevel.wdOutlineLevel1 || nLvl > (int)Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    //return null;
                    continue;
                }


                if (scopeRange == null)
                {
                    sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
                }
                else
                {
                    //if (sel.Range.Paragraphs.Count > 0)
                    if (scopeRange.Paragraphs.Count > 0)
                    {
                        // para = sel.Range.Paragraphs[1];
                        para = scopeRange.Paragraphs[1];

                        sel.Start = para.Range.Start;
                        sel.End = sel.Start;
                    }
                    else
                    {
                        sel.Start = nStartPos;
                        sel.End = sel.Start;
                    }
                    // sel.End = nEndPos;
                }

                sel.Range.GoTo();

                sel.Find.ClearFormatting();

                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";
                sel.Find.Forward = true;
                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                bool bRet = sel.Find.Execute();

                while (sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (para.OutlineLevel != ((Word.WdOutlineLevel)nLvl))
                        {
                            break;
                        }

                        if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                        //if (para.Range.InRange(scopeRange))
                        {
                            // ignore empty para
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                arrParas.Add(para);
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (null == para.Next())
                        {
                            break;
                        }
                    }

                    /*
                    para = para.Next();
                    if (para != null)
                    {
                        //sel.Start = para.Range.Start;
                        //sel.End = sel.Start;
                        // para.Range.Select();
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    }
                    else
                    {
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                    */

                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    bRet = sel.Find.Execute();
                }// while

            }// for


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            arrParas.Sort(paraComparer);

            return arrParas;
        }

        private ArrayList getHeadingParasInScopeByNavNoChangeView_v1(Word.Application app, Word.Document doc, 
                                                                 Boolean bIgnoreToc, Boolean bIgnoreTable,
                                                                 Boolean bIgnorePages, uint nIgnorePages,
                                                                 Word.Range scopeRange = null)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            app.Browser.Target = Word.WdBrowseTarget.wdBrowseHeading;
            ArrayList arrParas = new ArrayList();
            Word.Paragraph fndPara = null, prevPara = null, para = null;
            Boolean bInToc = false, bInTables = false, bInIgnorePages = false;

            Word.Range oRng = sel.Range;

            //             int oStart = sel.Start;
            //             int oEnd = sel.End;

            if (scopeRange == null)
            {
                oRng = doc.Content;
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            }

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;


            sel.Start = nStartPos;
            sel.End = sel.Start;

            //sel.GoTo();

            if (itemPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                fndPara = itemPara;
            }
            else
            {
                app.Browser.Next();
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            // int nLine1 = 0, nLine2 = 0;
            while (fndPara != null && fndPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                if (prevPara != null)
                {
                    if (fndPara.Range.IsEqual(prevPara.Range))
                    {
                        break;
                    }
                }

                if (scopeRange == null || RangeOverlap(fndPara.Range, scopeRange))
                {
                    bInToc = false;
                    bInTables = false;
                    bInIgnorePages = false;

                    if (bIgnoreToc)
                    {
                        foreach (Word.TableOfContents toc in doc.TablesOfContents)
                        {
                            if (fndPara.Range.InRange(toc.Range))
                            {
                                bInToc = true;
                                break;
                            }
                        }
                       
                    }//


                    if (bIgnoreTable)
                    {
                        bInTables = fndPara.Range.get_Information(Word.WdInformation.wdWithInTable);
                    }


                    if (bIgnorePages)
                    {
                        int nPageSn = fndPara.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                        bInIgnorePages = (nPageSn <= nIgnorePages);
                    }

                    if (!((bIgnoreToc && bInToc) || (bIgnoreTable && bInTables) || (bIgnorePages && bInIgnorePages)))
                    {
                        if (!String.IsNullOrWhiteSpace(fndPara.Range.Text.Trim(m_trimChars)))
                        {
                            arrParas.Add(fndPara);
                        }
                    }
                }
                else
                {
                    break;
                }

                prevPara = fndPara;

                app.Browser.Next();
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            return arrParas;
        }



        private ArrayList getHeadingParasInScopeByNavNoChangeView(Word.Application app, Word.Document doc, Word.Range scopeRange = null)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            app.Browser.Target = Word.WdBrowseTarget.wdBrowseHeading;
            ArrayList arrParas = new ArrayList();
            Word.Paragraph fndPara = null, prevPara = null, para = null;

            Word.Range oRng = sel.Range;

//             int oStart = sel.Start;
//             int oEnd = sel.End;

            if (scopeRange == null)
            {
                oRng = doc.Content;
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            }

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

//             Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
//             if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
//             {
//                 doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
//             }
//             else
//             {
//                 doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
//             }


            sel.Start = nStartPos;
            sel.End = sel.Start;

            //sel.GoTo();

            if (itemPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                fndPara = itemPara;
            }
            else
            {
                app.Browser.Next();
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            // int nLine1 = 0, nLine2 = 0;
            while (fndPara != null && fndPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                if (prevPara != null)
                {
                    if (fndPara.Range.IsEqual(prevPara.Range))
                    {
                        break;
                    }
                }

                if (scopeRange == null || RangeOverlap(fndPara.Range, scopeRange))
                {
                    if (!String.IsNullOrWhiteSpace(fndPara.Range.Text.Trim(m_trimChars)))
                    {
                        arrParas.Add(fndPara);
                    }
                }
                else
                {
                    break;
                }

                prevPara = fndPara;

                app.Browser.Next();
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

//             doc.ActiveWindow.View.Type = oViewType;
// 
//             sel.Start = oStart;
//             sel.End = oStart;
//             // sel.Range.Select();
//             sel.Range.GoTo();
//             doc.ActiveWindow.ScrollIntoView(sel.Range);

            return arrParas;
        }


        private ArrayList getHeadingParasInScopeByNav(Word.Application app,Word.Document doc, Word.Range scopeRange = null)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;// 测试代码

            app.Browser.Target = Word.WdBrowseTarget.wdBrowseHeading;
            ArrayList arrParas = new ArrayList();
            Word.Paragraph fndPara = null, prevPara = null, para = null;

            Word.Range oRng = sel.Range;

            int oStart = sel.Start;
            int oEnd = sel.End;

            if (scopeRange == null)
            {
                oRng = doc.Content;
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            }

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }


            sel.Start = nStartPos;
            sel.End = sel.Start;

            //sel.GoTo();

            if (itemPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                fndPara = itemPara;
            }
            else
            {
                app.Browser.Next();
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            // int nLine1 = 0, nLine2 = 0;
            while (fndPara != null && fndPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                if (prevPara != null)
                {
                    if (fndPara.Range.IsEqual(prevPara.Range))
                    {
                        break;
                    }
                }

                if (scopeRange == null || RangeOverlap(fndPara.Range, scopeRange))
                {
                    if (!String.IsNullOrWhiteSpace(fndPara.Range.Text.Trim(m_trimChars)))
                    {
                        arrParas.Add(fndPara);
                    }
                }
                else
                {
                    break;
                }

                prevPara = fndPara;

                app.Browser.Next();
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            doc.ActiveWindow.View.Type = oViewType;

            sel.Start = oStart;
            sel.End = oStart;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return arrParas;
        }


        private ArrayList getHeadingParasInScopeByGoto(Word.Document doc, Word.Range scopeRange = null)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;

            ArrayList arrParas = new ArrayList();
            Word.Paragraph fndPara = null, prevPara = null, para = null;

            Word.Range oRng = sel.Range;

            int oStart = sel.Start;
            int oEnd = sel.End;

            if (scopeRange == null)
            {
                oRng = doc.Content;
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            }

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }


            sel.Start = nStartPos;
            sel.End = sel.Start;

            //sel.GoTo();

            if (itemPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                fndPara = itemPara;
            }
            else
            {
                sel.GoToNext(Word.WdGoToItem.wdGoToHeading);
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            // int nLine1 = 0, nLine2 = 0;
            while (fndPara != null && fndPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
            {
                if (prevPara != null)
                {
                    if (fndPara.Range.IsEqual(prevPara.Range))
                    {
                        break;
                    }
                }

                if (scopeRange == null || RangeOverlap(fndPara.Range, scopeRange))
                {
                    if (!String.IsNullOrWhiteSpace(fndPara.Range.Text.Trim(m_trimChars)))
                    {
                        arrParas.Add(fndPara);
                    }
                }
                else
                {
                    break;
                }

                prevPara = fndPara;

                sel.GoToNext(Word.WdGoToItem.wdGoToHeading);
                if (sel.Paragraphs.Count == 1)
                {
                    fndPara = sel.Paragraphs[1]; // 
                }
            }

            doc.ActiveWindow.View.Type = oViewType;

            sel.Start = oStart;
            sel.End = oStart;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return arrParas;
        }

        /*
        public ArrayList getSpecificPicsParasInScope(Word.Document curDoc, Word.Range scopeRange = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }


            if (scopeRange == null)
            {
                sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            }
            else
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    para = scopeRange.Paragraphs[1];

                    sel.Start = para.Range.Start;
                    sel.End = sel.Start;
                }
                else
                {
                    sel.Start = nStartPos;
                    sel.End = sel.Start;
                }
                // sel.End = nEndPos;
            }

            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Text = "^g";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = false; // true
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (sel.Range.Paragraphs.Count > 0)
                {
                    para = sel.Range.Paragraphs[1];

                    if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                    //if (para.Range.InRange(scopeRange))
                    {
                        arrParas.Add(para);
                    }
                    else
                    {
                        break;
                    }

                    if (null == para.Next())
                    {
                        break;
                    }
                }

                /*
                para = para.Next();
                if (para != null)
                {
                    //sel.Start = para.Range.Start;
                    //sel.End = sel.Start;
                    //sel.Range.GoTo();

                    para.Range.Select();
                    sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }
                else
                {
                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
                * /
                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                bRet = sel.Find.Execute();
            }// while


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }
        */

        public ArrayList getInlineShpsInScope(Word.Document doc,
                        ArrayList arrIsolatePicsNotInTbl, ArrayList arrNotIsolatePicsNotInTbl,
                        ArrayList arrIsolatePicsInTbl, ArrayList arrNotIsolatePicsInTbl,
                        Boolean bParagraph = true,
                        Word.Range scopeRange = null,
                        int nShowInterval = 20)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            ArrayList arrAllPics = new ArrayList();
            Word.Paragraph para = null;

            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            Word.InlineShapes inShps = null;

            if (scopeRange == null)
            {
                inShps = doc.InlineShapes;
            }
            else
            {
                inShps = scopeRange.InlineShapes;
            }

            HashSet<int> hsParas = new HashSet<int>();

            Word.Paragraph inShpPara = null;
            Boolean bInTbl = false, bIsIsolate = false;
            int nCnt = 0;

            foreach (Word.InlineShape inShp in inShps)
            {
                nCnt++;

                if (nCnt == nShowInterval)
                {
                    doc.ActiveWindow.ScrollIntoView(inShp.Range);
                    nCnt = 0;
                }

                inShpPara = inShp.Range.Paragraphs[1];
                if (bParagraph)
                {
                    if (hsParas.Contains(inShpPara.Range.Start + 1) && hsParas.Contains(inShpPara.Range.End - 1))
                    {
                        continue;
                    }
                    else
                    {
                        hsParas.Add(inShpPara.Range.Start + 1);
                        hsParas.Add(inShpPara.Range.End - 1);
                    }

                    arrAllPics.Add(inShpPara);
                }
                else
                {
                    arrAllPics.Add(inShp);
                }

                bInTbl = inShp.Range.get_Information(Word.WdInformation.wdWithInTable);
                bIsIsolate = isIsolatePic(inShpPara);

                if (!bInTbl)
                {
                    if (bIsIsolate)
                    {
                        if (bParagraph)
                        {
                            arrIsolatePicsNotInTbl.Add(inShpPara);
                        }
                        else
                        {
                            arrIsolatePicsNotInTbl.Add(inShp);
                        }
                    }
                    else
                    {
                        if (bParagraph)
                        {
                            arrNotIsolatePicsNotInTbl.Add(inShpPara);
                        }
                        else
                        {
                            arrNotIsolatePicsNotInTbl.Add(inShp);
                        }
                    }
                }
                else
                {
                    if (bIsIsolate)
                    {
                        if (bParagraph)
                        {
                            arrIsolatePicsInTbl.Add(inShpPara);
                        }
                        else
                        {
                            arrIsolatePicsInTbl.Add(inShp);
                        }
                    }
                    else
                    {
                        if (bParagraph)
                        {
                            arrNotIsolatePicsInTbl.Add(inShpPara);
                        }
                        else
                        {
                            arrNotIsolatePicsInTbl.Add(inShp);
                        }
                    }
                }
            }

            doc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);


            return arrAllPics;
        }


        // 
        private ArrayList getInlineShpsInScopeV1(Word.Document curDoc,
                                ArrayList arrIsolatePicsNotInTbl, ArrayList arrNotIsolatePicsNotInTbl,
                                ArrayList arrIsolatePicsInTbl, ArrayList arrNotIsolatePicsInTbl,
                                Boolean bParagraph = false,
                                Word.Range scopeRange = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            ArrayList arrAllPics = new ArrayList();
            Word.Paragraph para = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }


            if (scopeRange == null)
            {
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            }
            else
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    para = scopeRange.Paragraphs[1];

                    sel.Start = para.Range.Start;
                    sel.End = sel.Start;
                }
                else
                {
                    sel.Start = nStartPos;
                    sel.End = sel.Start;
                }
                // sel.End = nEndPos;

                sel.Range.GoTo();
            }

            sel.Find.ClearFormatting();

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Text = "^g";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = false; // true
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            bool bRet = sel.Find.Execute();

            Boolean bInTbl = false, bIsIsolate = false;

            while (sel.Find.Found)
            {
                para = sel.Range.Paragraphs[1];

                if (para != null && sel.Range.InlineShapes.Count > 0)
                {
                    Word.InlineShape inShp = sel.Range.InlineShapes[1];

                    if (scopeRange == null || RangeOverlap(inShp.Range, scopeRange))
                    {
                        if (bParagraph)
                        {
                            arrAllPics.Add(para);
                        }
                        else
                        {
                            arrAllPics.Add(inShp);
                        }

                        bInTbl = inShp.Range.get_Information(Word.WdInformation.wdWithInTable);
                        bIsIsolate = isIsolatePic(para); // inShp.Range.Paragraphs[1]

                        if (bInTbl)
                        {
                            if (bIsIsolate)
                            {
                                if (arrIsolatePicsInTbl != null)
                                {
                                    if (bParagraph)
                                    {
                                        arrIsolatePicsInTbl.Add(para);
                                    }
                                    else
                                    {
                                        arrIsolatePicsInTbl.Add(inShp);
                                    }
                                }
                            }
                            else
                            {
                                if (arrNotIsolatePicsInTbl != null)
                                {
                                    if (bParagraph)
                                    {
                                        arrNotIsolatePicsInTbl.Add(para);
                                    }
                                    else
                                    {
                                        arrNotIsolatePicsInTbl.Add(inShp);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (bIsIsolate)
                            {
                                if (arrIsolatePicsNotInTbl != null)
                                {
                                    if (bParagraph)
                                    {
                                        arrIsolatePicsNotInTbl.Add(para);
                                    }
                                    else
                                    {
                                        arrIsolatePicsNotInTbl.Add(inShp);
                                    }
                                }
                            }
                            else
                            {
                                if (arrNotIsolatePicsNotInTbl != null)
                                {
                                    if (bParagraph)
                                    {
                                        arrNotIsolatePicsNotInTbl.Add(para);
                                    }
                                    else
                                    {
                                        arrNotIsolatePicsNotInTbl.Add(inShp);
                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        break;
                    }

                }

                if (bParagraph)
                {
                    sel.Start = para.Range.End;
                    sel.End = para.Range.End;

                    sel.Range.GoTo();
                }

                //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                bRet = sel.Find.Execute();
            }// while

            //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            return arrAllPics;
        }


        private ArrayList getHighlightInScope(Word.Document curDoc, Word.Range scopeRange = null,Boolean bOnlyObj = true)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null,nextPara = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            //if (scopeRange == null)
            //{
            //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            //}
            //else
            //{
            //    if (scopeRange.Paragraphs.Count > 0)
            //    {
            //        para = scopeRange.Paragraphs[1];

            //        sel.Start = para.Range.Start;
            //        sel.End = sel.Start;
            //    }
            //    else
            //    {
            //        sel.Start = nStartPos;
            //        sel.End = sel.Start;
            //    }
            //    // sel.End = nEndPos;
            //}

            sel.Start = scopeRngStartPos;
            sel.End = sel.Start;
            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = true;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                if (!bOnlyObj)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (scopeRange == null || RangeOverlap(para.Range, scopeRange))
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)) )
                            {
                                arrParas.Add(para);
                            }
                        }
                        else
                        {
                            break;
                        }

                        nextPara = para.Next();
                        if (null == nextPara)
                        {
                            break;
                        }
                    }

                    if (nextPara != null)
                    {
                        sel.Start = nextPara.Range.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }
                                        
                }
                else
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if (scopeRange == null || RangeOverlap(sel.Range, scopeRange))
                        {
                            Word.Range chRng = null;
                            Boolean bNotEmpty = false;

                            for (int i = 1; i <= sel.Range.Characters.Count; i++)
                            {
                                chRng = sel.Range.Characters[i];
                                if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                {
                                    bNotEmpty = true;
                                    break;
                                }
                            }

                            if (bNotEmpty)
                            {
                                arrParas.Add(sel.Range);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                bRet = sel.Find.Execute();
            }// while


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrParas;
        }

        public Hashtable getSpecificHighlightInScopeSet(Word.Document curDoc, Word.Range scopeRange = null, int nType = 1, Boolean[] bDstArrColor = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            int nTmpStart = -1, nTmpEnd = -1;

            Word.WdColorIndex fndHighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            int nColorIndex = -1;
            int nSolColorIndex = 99;

            // ArrayList arrParas = new ArrayList();
            HashSet<SortedSet<int>> itemSet = null;
            SortedSet<int> rngSet = null;

            Hashtable hashColorIndexParas = new Hashtable();

            if (bDstArrColor == null)
            {
                // hashColorIndexParas.Add(nSolColorIndex, new ArrayList());
                hashColorIndexParas.Add(nSolColorIndex, new HashSet<SortedSet<int> >());
            }
            else
            {
                for (int nIndex = (int)Word.WdColorIndex.wdAuto; nIndex <= (int)Word.WdColorIndex.wdGray25; nIndex++)
                {
                    if (bDstArrColor[nIndex])
                    {
                        hashColorIndexParas[nIndex] = new HashSet<SortedSet<int> >();
                    }
                }
            }

            Word.Range nxChRng = null;
            Word.Paragraph para = null, nextPara = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            sel.Start = scopeRngStartPos;
            sel.End = sel.Start;
            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            // sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = true;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                fndHighlightColorIndex = sel.Range.HighlightColorIndex;

                if (nType == 2 || nType == 3) // para body, whole para
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if ((scopeRange == null || RangeOverlap(para.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                if (bDstArrColor == null)
                                {
                                    itemSet = (HashSet<SortedSet<int>>)hashColorIndexParas[nSolColorIndex];

                                    nTmpStart = para.Range.Start;
                                    nTmpEnd = para.Range.End;

                                    if (nType == 2)
                                    {
                                        if (nTmpEnd > nTmpStart)
                                        {
                                            nTmpEnd--;
                                        }
                                    }

                                    rngSet = new SortedSet<int>();
                                    for (int k = nTmpStart; k <= nTmpEnd; k++)
                                    {
                                        rngSet.Add(k);
                                    }
                                    // arrParas.Add(para.Range);
                                    itemSet.Add(rngSet);
                                }
                                else
                                {
                                    if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                itemSet = (HashSet<SortedSet<int>>)hashColorIndexParas[nColorIndex];

                                                nTmpStart = para.Range.Start;
                                                nTmpEnd = para.Range.End;

                                                if (nType == 2)
                                                {
                                                    if (nTmpEnd > nTmpStart)
                                                    {
                                                        nTmpEnd--;
                                                    }
                                                }

                                                rngSet = new SortedSet<int>();
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    rngSet.Add(k);
                                                }
                                                itemSet.Add(rngSet);
                                                // rngSet.Add(para.Range);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        nColorIndex = (int)fndHighlightColorIndex;

                                        if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                        {
                                            itemSet = (HashSet<SortedSet<int>>)hashColorIndexParas[nColorIndex];

                                            nTmpStart = para.Range.Start;
                                            nTmpEnd = para.Range.End;

                                            if (nType == 2)
                                            {
                                                if (nTmpEnd > nTmpStart)
                                                {
                                                    nTmpEnd--;
                                                }
                                            }

                                            rngSet = new SortedSet<int>();
                                            for (int k = nTmpStart; k <= nTmpEnd; k++)
                                            {
                                                rngSet.Add(k);
                                            }
                                            itemSet.Add(rngSet);
                                            // rngSet.Add(para.Range);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        nextPara = para.Next();
                        if (null == nextPara)
                        {
                            break;
                        }
                    }

                    if (nextPara != null)
                    {
                        sel.Start = nextPara.Range.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }

                }
                else if (nType == 1) // only obj
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if ((scopeRange == null || RangeOverlap(sel.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            Boolean bNotEmpty = false;

                            if (bDstArrColor == null)
                            {
                                foreach (Word.Range chRng in sel.Range.Characters)
                                {
                                    if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                    {
                                        bNotEmpty = true;
                                        break;
                                    }
                                }

                                if (bNotEmpty)
                                {
                                    itemSet = (HashSet<SortedSet<int>>)hashColorIndexParas[nSolColorIndex];

                                    nTmpStart = sel.Range.Start;
                                    nTmpEnd = sel.Range.End;

                                    rngSet = new SortedSet<int>();
                                    for (int k = nTmpStart; k <= nTmpEnd; k++)
                                    {
                                        rngSet.Add(k);
                                    }
                                    itemSet.Add(rngSet);
                                    // rngSet.Add(sel.Range);
                                }
                            }
                            else
                            {
                                if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                {
                                    foreach (Word.Range chRng in sel.Range.Characters)
                                    {
                                        if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                itemSet = (HashSet<SortedSet<int>>)hashColorIndexParas[nColorIndex];

                                                nTmpStart = chRng.Start;
                                                nTmpEnd = chRng.End;

                                                rngSet = new SortedSet<int>();
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    rngSet.Add(k);
                                                }
                                                itemSet.Add(rngSet);
                                                // rngSet.Add(chRng);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    nColorIndex = (int)fndHighlightColorIndex;

                                    if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                itemSet = (HashSet<SortedSet<int>>)hashColorIndexParas[nColorIndex];

                                                nTmpStart = sel.Range.Start;
                                                nTmpEnd = sel.Range.End;

                                                rngSet = new SortedSet<int>();
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    rngSet.Add(k);
                                                }
                                                itemSet.Add(rngSet);
                                                // rngSet.Add(sel.Range);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    nxChRng = sel.Range.Characters[sel.Range.Characters.Count].Next(Word.WdUnits.wdCharacter);

                    if (nxChRng != null)
                    {
                        sel.Start = nxChRng.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }
                }

                bRet = sel.Find.Execute();
            }// while


            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return hashColorIndexParas;
        }

        public Hashtable getSpecificHighlightInScope(Word.Document curDoc, Word.Range scopeRange = null, int nType = 1, Boolean[] bDstArrColor = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            int nTmpStart = -1, nTmpEnd = -1;

            Word.WdColorIndex fndHighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            int nColorIndex = -1;
            int nSolColorIndex = 99;

            // ArrayList arrParas = new ArrayList();
            HashSet<int> rngSet = null;

            Hashtable hashColorIndexParas = new Hashtable();

            if (bDstArrColor == null)
            {
                // hashColorIndexParas.Add(nSolColorIndex, new ArrayList());
                hashColorIndexParas.Add(nSolColorIndex, new HashSet<int>());
            }
            else
            {
                for (int nIndex = (int)Word.WdColorIndex.wdAuto; nIndex <= (int)Word.WdColorIndex.wdGray25; nIndex++)
                {
                    if (bDstArrColor[nIndex])
                    {
                        hashColorIndexParas[nIndex] = new HashSet<int>();
                    }
                }
            }

            Word.Range nxChRng = null;
            Word.Paragraph para = null,nextPara = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            sel.Start = scopeRngStartPos;
            sel.End = sel.Start;
            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            // sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = true;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                fndHighlightColorIndex = sel.Range.HighlightColorIndex;

                if (nType == 2 || nType == 3) // para body, whole para
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if ((scopeRange == null || RangeOverlap(para.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                if (bDstArrColor == null)
                                {
                                    rngSet = (HashSet<int>)hashColorIndexParas[nSolColorIndex];

                                    nTmpStart = para.Range.Start;
                                    nTmpEnd = para.Range.End;

                                    if(nType == 2)
                                    {
                                        if(nTmpEnd > nTmpStart)
                                        {
                                            nTmpEnd--;
                                        }
                                    }

                                    for (int k = nTmpStart; k <= nTmpEnd; k++)
                                    {
                                        rngSet.Add(k);
                                    }
                                    // arrParas.Add(para.Range);
                                }
                                else
                                {
                                    if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                rngSet = (HashSet<int>)hashColorIndexParas[nColorIndex];

                                                nTmpStart = para.Range.Start;
                                                nTmpEnd = para.Range.End;

                                                if(nType == 2)
                                                {
                                                    if(nTmpEnd > nTmpStart)
                                                    {
                                                        nTmpEnd--;
                                                    }
                                                }

                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    rngSet.Add(k);
                                                }

                                                // rngSet.Add(para.Range);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        nColorIndex = (int)fndHighlightColorIndex;

                                        if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                        {
                                            rngSet = (HashSet<int>)hashColorIndexParas[nColorIndex];

                                            nTmpStart = para.Range.Start;
                                            nTmpEnd = para.Range.End;

                                            if(nType == 2)
                                            {
                                                if(nTmpEnd > nTmpStart)
                                                {
                                                    nTmpEnd--;
                                                }
                                            }

                                            for (int k = nTmpStart; k <= nTmpEnd; k++)
                                            {
                                                rngSet.Add(k);
                                            }

                                            // rngSet.Add(para.Range);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        nextPara = para.Next();
                        if (null == nextPara)
                        {
                            break;
                        }
                    }

                    if (nextPara != null)
                    {
                        sel.Start = nextPara.Range.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }
                    
                }
                else if(nType == 1) // only obj
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if ((scopeRange == null || RangeOverlap(sel.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            Boolean bNotEmpty = false;

                            if (bDstArrColor == null)
                            {
                                foreach (Word.Range chRng in sel.Range.Characters)
                                {
                                    if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                    {
                                        bNotEmpty = true;
                                        break;
                                    }
                                }

                                if (bNotEmpty)
                                {
                                    rngSet = (HashSet<int>)hashColorIndexParas[nSolColorIndex];

                                    nTmpStart = sel.Range.Start;
                                    nTmpEnd = sel.Range.End;
                                    for (int k = nTmpStart; k <= nTmpEnd; k++)
                                    {
                                        rngSet.Add(k);
                                    }

                                    // rngSet.Add(sel.Range);
                                }
                            }
                            else
                            {
                                if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                {
                                    foreach (Word.Range chRng in sel.Range.Characters)
                                    {
                                        if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                rngSet = (HashSet<int>)hashColorIndexParas[nColorIndex];

                                                nTmpStart = chRng.Start;
                                                nTmpEnd = chRng.End;
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    rngSet.Add(k);
                                                }

                                                // rngSet.Add(chRng);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    nColorIndex = (int)fndHighlightColorIndex;

                                    if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                rngSet = (HashSet<int>)hashColorIndexParas[nColorIndex];

                                                nTmpStart = sel.Range.Start;
                                                nTmpEnd = sel.Range.End;
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    rngSet.Add(k);
                                                }

                                                // rngSet.Add(sel.Range);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    nxChRng = sel.Range.Characters[sel.Range.Characters.Count].Next(Word.WdUnits.wdCharacter);

                    if (nxChRng != null)
                    {
                        sel.Start = nxChRng.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }
                }

                bRet = sel.Find.Execute();
            }// while


            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return hashColorIndexParas;
        }

        private Hashtable getSpecificHighlightInScope2(Word.Document curDoc, Word.Range scopeRange = null, int nType = 1, Boolean[] bDstArrColor = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            int nTmpStart = -1, nTmpEnd = -1;

            Word.WdColorIndex fndHighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            int nColorIndex = -1;
            int nSolColorIndex = 99;

            // ArrayList arrParas = new ArrayList();
            // HashSet<int> rngSet = null;
            HashSet<HashSet<int> > rngSet = null;
            HashSet<int> tmpSet = null;

            Hashtable hashColorIndexParas = new Hashtable();

            if (bDstArrColor == null)
            {
                // hashColorIndexParas.Add(nSolColorIndex, new ArrayList());
                hashColorIndexParas.Add(nSolColorIndex, new HashSet<HashSet<int>>() );
            }
            else
            {
                for (int nIndex = (int)Word.WdColorIndex.wdAuto; nIndex <= (int)Word.WdColorIndex.wdGray25; nIndex++)
                {
                    if (bDstArrColor[nIndex])
                    {
                        // hashColorIndexParas[nIndex] = new HashSet<int>();
                        hashColorIndexParas[nIndex] = new HashSet<HashSet<int> >();
                    }
                }
            }

            Word.Range nxChRng = null;
            Word.Paragraph para = null, nextPara = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }

            sel.Start = scopeRngStartPos;
            sel.End = sel.Start;
            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            // sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = true;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                fndHighlightColorIndex = sel.Range.HighlightColorIndex;

                if (nType == 2 || nType == 3) // para body, whole para
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if ((scopeRange == null || RangeOverlap(para.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                if (bDstArrColor == null)
                                {
                                    rngSet = (HashSet<HashSet<int> >)hashColorIndexParas[nSolColorIndex];

                                    nTmpStart = para.Range.Start;
                                    nTmpEnd = para.Range.End;

                                    if (nType == 2)
                                    {
                                        if (nTmpEnd > nTmpStart)
                                        {
                                            nTmpEnd--;
                                        }
                                    }

                                    tmpSet = new HashSet<int>();
                                    for (int k = nTmpStart; k <= nTmpEnd; k++)
                                    {
                                        tmpSet.Add(k);
                                    }
                                    rngSet.Add(tmpSet);
                                    // arrParas.Add(para.Range);
                                }
                                else
                                {
                                    if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                rngSet = (HashSet<HashSet<int> >)hashColorIndexParas[nColorIndex];

                                                nTmpStart = para.Range.Start;
                                                nTmpEnd = para.Range.End;

                                                if (nType == 2)
                                                {
                                                    if (nTmpEnd > nTmpStart)
                                                    {
                                                        nTmpEnd--;
                                                    }
                                                }

                                                tmpSet = new HashSet<int>();
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    tmpSet.Add(k);
                                                }
                                                rngSet.Add(tmpSet);

                                                // rngSet.Add(para.Range);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        nColorIndex = (int)fndHighlightColorIndex;

                                        if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                        {
                                            rngSet = (HashSet<HashSet<int> >)hashColorIndexParas[nColorIndex];

                                            nTmpStart = para.Range.Start;
                                            nTmpEnd = para.Range.End;

                                            if (nType == 2)
                                            {
                                                if (nTmpEnd > nTmpStart)
                                                {
                                                    nTmpEnd--;
                                                }
                                            }

                                            tmpSet = new HashSet<int>();
                                            for (int k = nTmpStart; k <= nTmpEnd; k++)
                                            {
                                                tmpSet.Add(k);
                                            }
                                            rngSet.Add(tmpSet);

                                            // rngSet.Add(para.Range);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        nextPara = para.Next();
                        if (null == nextPara)
                        {
                            break;
                        }
                    }

                    if (nextPara != null)
                    {
                        sel.Start = nextPara.Range.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }

                }
                else if (nType == 1) // only obj
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if ((scopeRange == null || RangeOverlap(sel.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            Boolean bNotEmpty = false;

                            if (bDstArrColor == null)
                            {
                                foreach (Word.Range chRng in sel.Range.Characters)
                                {
                                    if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                    {
                                        bNotEmpty = true;
                                        break;
                                    }
                                }

                                if (bNotEmpty)
                                {
                                    rngSet = (HashSet<HashSet<int> >)hashColorIndexParas[nSolColorIndex];

                                    nTmpStart = sel.Range.Start;
                                    nTmpEnd = sel.Range.End;

                                    tmpSet = new HashSet<int>();
                                    for (int k = nTmpStart; k <= nTmpEnd; k++)
                                    {
                                        tmpSet.Add(k);
                                    }
                                    rngSet.Add(tmpSet);

                                    // rngSet.Add(sel.Range);
                                }
                            }
                            else
                            {
                                if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                {
                                    foreach (Word.Range chRng in sel.Range.Characters)
                                    {
                                        if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                rngSet = (HashSet<HashSet<int> >)hashColorIndexParas[nColorIndex];

                                                nTmpStart = chRng.Start;
                                                nTmpEnd = chRng.End;

                                                tmpSet = new HashSet<int>();
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    tmpSet.Add(k);
                                                }
                                                rngSet.Add(tmpSet);

                                                // rngSet.Add(chRng);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    nColorIndex = (int)fndHighlightColorIndex;

                                    if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                rngSet = (HashSet<HashSet<int> >)hashColorIndexParas[nColorIndex];

                                                nTmpStart = sel.Range.Start;
                                                nTmpEnd = sel.Range.End;

                                                tmpSet = new HashSet<int>();
                                                for (int k = nTmpStart; k <= nTmpEnd; k++)
                                                {
                                                    tmpSet.Add(k);
                                                }
                                                rngSet.Add(tmpSet);

                                                // rngSet.Add(sel.Range);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    nxChRng = sel.Range.Characters[sel.Range.Characters.Count].Next(Word.WdUnits.wdCharacter);

                    if (nxChRng != null)
                    {
                        sel.Start = nxChRng.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }
                }

                bRet = sel.Find.Execute();
            }// while


            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return hashColorIndexParas;
        }

        private Hashtable getSpecificHighlightInScope_v1(Word.Document curDoc, Word.Range scopeRange = null, Boolean bOnlyObj = true, Boolean[] bDstArrColor = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Word.WdColorIndex fndHighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            int nColorIndex = -1;
            int nSolColorIndex = 99;

            // ArrayList arrParas = new ArrayList();
            ArrayList arrParas = null;
            Hashtable hashColorIndexParas = new Hashtable();

            if (bDstArrColor == null)
            {
                hashColorIndexParas.Add(nSolColorIndex, new ArrayList());
            }
            else
            {
                for (int nIndex = (int)Word.WdColorIndex.wdAuto; nIndex <= (int)Word.WdColorIndex.wdGray25; nIndex++)
                {
                    if (bDstArrColor[nIndex])
                    {
                        hashColorIndexParas[nIndex] = new ArrayList();
                    }
                }
            }

            Word.Paragraph para = null;

            Word.WdViewType oViewType = curDoc.ActiveWindow.View.Type;
            if (curDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                curDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            }
            else
            {
                curDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            }


            if (scopeRange == null)
            {
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
            }
            else
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    para = scopeRange.Paragraphs[1];

                    sel.Start = para.Range.Start;
                    sel.End = sel.Start;
                }
                else
                {
                    sel.Start = nStartPos;
                    sel.End = sel.Start;
                }
                // sel.End = nEndPos;
            }

            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = true;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                fndHighlightColorIndex = sel.Range.HighlightColorIndex;

                if (!bOnlyObj)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if ((scopeRange == null || RangeOverlap(para.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                if (bDstArrColor == null)
                                {
                                    arrParas = (ArrayList)hashColorIndexParas[nSolColorIndex];
                                    arrParas.Add(para.Range);
                                }
                                else
                                {
                                    if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                arrParas = (ArrayList)hashColorIndexParas[nColorIndex];
                                                arrParas.Add(para.Range);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        nColorIndex = (int)fndHighlightColorIndex;

                                        if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                        {
                                            arrParas = (ArrayList)hashColorIndexParas[nColorIndex];
                                            arrParas.Add(para.Range);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (null == para.Next())
                        {
                            break;
                        }
                    }

                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    /*
                    para = para.Next();
                    if (para != null)
                    {
                        // sel.Start = para.Range.Start;
                        // sel.End = sel.Start;
                        // sel.Range.GoTo();

                        para.Range.Select();
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    }
                    else
                    {
                        sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                     * */
                }
                else
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if ((scopeRange == null || RangeOverlap(sel.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            Boolean bNotEmpty = false;

                            if (bDstArrColor == null)
                            {
                                foreach (Word.Range chRng in sel.Range.Characters)
                                {
                                    if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                    {
                                        bNotEmpty = true;
                                        break;
                                    }
                                }

                                if (bNotEmpty)
                                {
                                    arrParas = (ArrayList)hashColorIndexParas[nSolColorIndex];
                                    arrParas.Add(sel.Range);
                                }
                            }
                            else
                            {
                                if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                {
                                    foreach (Word.Range chRng in sel.Range.Characters)
                                    {
                                        if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                arrParas = (ArrayList)hashColorIndexParas[nColorIndex];
                                                arrParas.Add(chRng);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    nColorIndex = (int)fndHighlightColorIndex;

                                    if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                arrParas = (ArrayList)hashColorIndexParas[nColorIndex];
                                                arrParas.Add(sel.Range);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }

                bRet = sel.Find.Execute();
            }// while


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            curDoc.ActiveWindow.View.Type = oViewType;

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return hashColorIndexParas;
        }

        public ArrayList findNextSpecificHighlightInScope(Word.Document curDoc, Boolean bForward = true, Word.Range scopeRange = null, Boolean bOnlyObj = true, Boolean[] bDstArrColor = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Word.WdColorIndex fndHighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            int nColorIndex = -1;

            Word.Range fndRng = null, nxChRng = null;
            Word.Paragraph para = null,nextPara = null,prevPara = null;
            ArrayList arrFndRngs = new ArrayList();


            sel.Find.ClearFormatting();

            //if (bForward)
            //{
            //    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //}
            //else
            //{
            //    sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            //}

            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = bForward;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                fndHighlightColorIndex = sel.Range.HighlightColorIndex;

                if (!bOnlyObj)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if ((scopeRange == null || RangeOverlap(para.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                if (bDstArrColor == null)
                                {
                                    // arrParas.Add(para);
                                    //fndRng = para.Range;
                                    arrFndRngs.Add(para.Range);
                                }
                                else
                                {
                                    if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                // arrParas.Add(para);
                                                // fndRng = para.Range;
                                                arrFndRngs.Add(para.Range);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        nColorIndex = (int)fndHighlightColorIndex;

                                        if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                        {
                                            // arrParas.Add(para);
                                            // fndRng = para.Range;
                                            arrFndRngs.Add(para.Range);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nextPara = para.Next();
                            if (null == nextPara)
                            {
                                break;
                            }
                        }
                        else
                        {
                            prevPara = para.Previous();
                            if (null == prevPara)
                            {
                                break;
                            }
                        }
                    }

                    if (bForward)
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        //nextPara = para.Next();
                        if (nextPara != null)
                        {
                            sel.Start = nextPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                    else
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        //prevPara = para.Previous();
                        if (prevPara != null)
                        {
                            sel.Start = prevPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                }
                else
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if ((scopeRange == null || RangeOverlap(sel.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            Boolean bNotEmpty = false;

                            if (bDstArrColor == null)
                            {
                                foreach (Word.Range chRng in sel.Range.Characters)
                                {
                                    if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                    {
                                        bNotEmpty = true;
                                        break;
                                    }
                                }

                                if (bNotEmpty)
                                {
                                    // arrParas.Add(sel.Range);
                                    // fndRng = sel.Range;
                                    arrFndRngs.Add(sel.Range);
                                }
                            }
                            else
                            {
                                if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                {
                                    int nPreColorIndex = -1;
                                    Boolean bFound = false;

                                    if (bForward)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                nColorIndex = (int)chRng.HighlightColorIndex;

                                                if (bFound && nPreColorIndex != -1 && nColorIndex != nPreColorIndex)
                                                {
                                                    break;
                                                }

                                                if (nColorIndex != -1 && nColorIndex <= 16 &&
                                                    (nPreColorIndex == -1 || nColorIndex == nPreColorIndex) &&
                                                    bDstArrColor[nColorIndex])
                                                {
                                                    // arrParas.Add(chRng);
                                                    // fndRng = chRng;
                                                    arrFndRngs.Add(chRng);
                                                    bFound = true;

                                                    nPreColorIndex = nColorIndex;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Word.Range chRng = null;
                                        for (int i = sel.Range.Characters.Count; i > 0;i-- )
                                        {
                                            chRng = (Word.Range)sel.Range.Characters[i];

                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                nColorIndex = (int)chRng.HighlightColorIndex;

                                                if (bFound && nPreColorIndex != -1 && nColorIndex != nPreColorIndex)
                                                {
                                                    break;
                                                }

                                                if (nColorIndex != -1 && nColorIndex <= 16 &&
                                                   (nPreColorIndex == -1 || nColorIndex == nPreColorIndex) &&
                                                    bDstArrColor[nColorIndex])
                                                {
                                                    // arrParas.Add(chRng);
                                                    // fndRng = chRng;
                                                    arrFndRngs.Add(chRng);

                                                    bFound = true;

                                                    nPreColorIndex = nColorIndex;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    nColorIndex = (int)fndHighlightColorIndex;

                                    if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                // fndRng = sel.Range;
                                                // arrParas.Add(sel.Range);
                                                arrFndRngs.Add(sel.Range);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nxChRng = sel.Range.Characters[sel.Range.Characters.Count].Next(Word.WdUnits.wdCharacter);
                        }
                        else
                        {
                            nxChRng = sel.Range.Characters[1].Previous(Word.WdUnits.wdCharacter);
                        }

                        if (nxChRng != null)
                        {
                            if (bForward)
                            {
                                sel.Start = nxChRng.Start;
                            }
                            else
                            {
                                sel.Start = nxChRng.End;
                            }
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                }

                if (arrFndRngs.Count > 0)
                {
                    break;
                }

                bRet = sel.Find.Execute();
            }// while

            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            return arrFndRngs;
        }


        private ArrayList findSpecificHighlightInScope(Word.Document curDoc, Boolean bForward = true, Word.Range scopeRange = null, Boolean bOnlyObj = true, Boolean[] bDstArrColor = null)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.Paragraph itemPara = sel.Paragraphs[1];

            int nStartPos = itemPara.Range.Start;
            int nEndPos = nStartPos;

            Word.WdColorIndex fndHighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            int nColorIndex = -1;

            Word.Range fndRng = null, nxChRng = null;
            Word.Paragraph para = null,prevPara = null,nextPara = null;
            ArrayList arrFndRngs = new ArrayList();

            int scopeRngStartPos = curDoc.Content.Start, scopeRngEndPos = curDoc.Content.End;

            if (scopeRange != null)
            {
                if (scopeRange.Paragraphs.Count > 0)
                {
                    scopeRngStartPos = scopeRange.Paragraphs[1].Range.Start;
                    scopeRngEndPos = scopeRange.Paragraphs[scopeRange.Paragraphs.Count].Range.End;
                }
            }


            if (bForward)
            {
                sel.Start = scopeRngStartPos;
                sel.End = sel.Start;
            }
            else
            {
                sel.Start = scopeRngEndPos;
                sel.End = sel.Start;
            }


            sel.Range.GoTo();

            sel.Find.ClearFormatting();

            //if (bForward)
            //{
            //    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //}
            //else
            //{
            //    sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            //}

            // sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)nLvl;

            sel.Find.Highlight = 1;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = true;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Forward = bForward;

            bool bRet = sel.Find.Execute();

            while (sel.Find.Found)
            {
                fndHighlightColorIndex = sel.Range.HighlightColorIndex;

                if (!bOnlyObj)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if ((scopeRange == null || RangeOverlap(para.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                            {
                                if (bDstArrColor == null)
                                {
                                    // arrParas.Add(para);
                                    //fndRng = para.Range;
                                    arrFndRngs.Add(para.Range);
                                }
                                else
                                {
                                    if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                // arrParas.Add(para);
                                                // fndRng = para.Range;
                                                arrFndRngs.Add(para.Range);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        nColorIndex = (int)fndHighlightColorIndex;

                                        if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                        {
                                            // arrParas.Add(para);
                                            // fndRng = para.Range;
                                            arrFndRngs.Add(para.Range);
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nextPara = para.Next();
                            if (null == nextPara)
                            {
                                break;
                            }
                        }
                        else
                        {
                            prevPara = para.Previous();
                            if (null == prevPara)
                            {
                                break;
                            }
                        }
                    }

                    if (bForward)
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        //nextPara = para.Next();
                        if (nextPara != null)
                        {
                            sel.Start = nextPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }
                    else
                    {
                        //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        //prevPara = para.Previous();
                        if (prevPara != null)
                        {
                            sel.Start = prevPara.Range.Start;
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                }
                else
                {
                    if (sel.Range.Characters.Count > 0)
                    {
                        if ((scopeRange == null || RangeOverlap(sel.Range, scopeRange)) && fndHighlightColorIndex != Word.WdColorIndex.wdNoHighlight)
                        {
                            Boolean bNotEmpty = false;

                            if (bDstArrColor == null)
                            {
                                foreach (Word.Range chRng in sel.Range.Characters)
                                {
                                    if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                    {
                                        bNotEmpty = true;
                                        break;
                                    }
                                }

                                if (bNotEmpty)
                                {
                                    // arrParas.Add(sel.Range);
                                    // fndRng = sel.Range;
                                    arrFndRngs.Add(sel.Range);
                                }
                            }
                            else
                            {
                                if ((int)Word.WdConstants.wdUndefined == (int)fndHighlightColorIndex)
                                {
                                    foreach (Word.Range chRng in sel.Range.Characters)
                                    {
                                        if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                        {
                                            nColorIndex = (int)chRng.HighlightColorIndex;
                                            if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                            {
                                                // arrParas.Add(chRng);
                                                // fndRng = chRng;
                                                arrFndRngs.Add(chRng);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    nColorIndex = (int)fndHighlightColorIndex;

                                    if (nColorIndex != -1 && nColorIndex <= 16 && bDstArrColor[nColorIndex])
                                    {
                                        foreach (Word.Range chRng in sel.Range.Characters)
                                        {
                                            if (chRng.Text.IndexOfAny(m_trimChars) == -1)
                                            {
                                                // fndRng = sel.Range;
                                                // arrParas.Add(sel.Range);
                                                arrFndRngs.Add(sel.Range);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (bForward)
                        {
                            nxChRng = sel.Range.Characters[sel.Range.Characters.Count].Next(Word.WdUnits.wdCharacter);
                        }
                        else
                        {
                            nxChRng = sel.Range.Characters[1].Previous(Word.WdUnits.wdCharacter);
                        }

                        if (nxChRng != null)
                        {
                            if (bForward)
                            {
                                sel.Start = nxChRng.Start;
                            }
                            else
                            {
                                sel.Start = nxChRng.End;
                            }
                            sel.End = sel.Start;
                            sel.Range.GoTo();
                        }
                    }

                }

                if (arrFndRngs.Count > 0)
                {
                    break;
                }

                bRet = sel.Find.Execute();
            }// while


            // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ClearFormatting();

            //curDoc.ActiveWindow.View.Type = oViewType;

            //sel.Start = nOStart;
            //sel.End = nOEnd;
            //// sel.Range.Select();
            //sel.Range.GoTo();
            //curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            //ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            //arrParas.Sort(paraComparer);

            return arrFndRngs;
        }


        public TreeNode buildOutlineTree(Word.Document curDoc, Boolean bOnlyBody = false)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // search all NOT text body paragraph
            int nOStart = sel.Start;
            int nOEnd = sel.End;


            ArrayList arrParas = new ArrayList();
            Word.Paragraph para = null,nextPara = null;

            for (Word.WdOutlineLevel lvl = Word.WdOutlineLevel.wdOutlineLevel1; lvl <= Word.WdOutlineLevel.wdOutlineLevel9; lvl++)
            {
                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove); // sel.Start = nOStart;sel.End = nOStart;
                sel.Find.ClearFormatting();

                // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                sel.Find.ParagraphFormat.OutlineLevel = lvl;

                sel.Find.Text = "";
                sel.Find.Replacement.Text = "";
                sel.Find.Forward = true;
                sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                sel.Find.Format = true;
                sel.Find.MatchCase = false;
                sel.Find.MatchWholeWord = false;
                sel.Find.MatchByte = false;
                sel.Find.MatchWildcards = false;
                sel.Find.MatchSoundsLike = false;
                sel.Find.MatchAllWordForms = false;

                bool bRet = sel.Find.Execute();

                while (sel.Find.Found)
                {
                    if (sel.Range.Paragraphs.Count > 0)
                    {
                        para = sel.Range.Paragraphs[1];

                        if (!String.IsNullOrWhiteSpace(para.Range.Text.Trim(m_trimChars)))
                        {
                            arrParas.Add(para);
                        }

                        nextPara = para.Next();
                        if (null == nextPara)
                        {
                            break;
                        }
                    }

                    // sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    if (nextPara != null)
                    {
                        sel.Start = nextPara.Range.Start;
                        sel.End = sel.Start;
                        sel.Range.GoTo();
                    }

                    bRet = sel.Find.Execute();
                }// while

            }// for


            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            sel.Start = nOStart;
            sel.End = nOEnd;
            // sel.Range.Select();
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            // sort by range
            // 
            ClassParagraphComparer paraComparer = new ClassParagraphComparer();
            arrParas.Sort(paraComparer);

            // build outline
            TreeNode rootNode = buildOutLine(arrParas, bOnlyBody);

            return rootNode;

        }


        private TreeNode buildOutLine(ArrayList arrParas, Boolean bOnlyBody = false)
        {
            TreeNode rootNode = new TreeNode();
            TreeNode preNode = rootNode;

            String strItem = "", strPrefix = "";

            int i = 0;
            // 章节
            foreach (Word.Paragraph para in arrParas)
            {
                if (!bOnlyBody)
                {
                    if (!para.Range.ListFormat.ListString.Equals(""))
                    {
                        strPrefix = para.Range.ListFormat.ListString;
                    }
                    else
                    {
                        strPrefix = "";
                    }
                }

                strItem = strPrefix + para.Range.Text;
                strItem.Replace(Environment.NewLine, "");

                if (strItem.Trim().Equals(""))
                {
                    i++;
                    strItem = "无内容章节" + i;
                }

                TreeNode newNode = null;
                switch (para.OutlineLevel)
                {
                    case Word.WdOutlineLevel.wdOutlineLevel1:
                        newNode = new TreeNode(strItem);
                        newNode.Name = strItem;
                        newNode.ImageIndex = newNode.SelectedImageIndex = (int)para.OutlineLevel;
                        newNode.Tag = para;

                        rootNode.Nodes.Add(newNode);
                        preNode = newNode;
                        break;

                    case Word.WdOutlineLevel.wdOutlineLevel2:
                    case Word.WdOutlineLevel.wdOutlineLevel3:
                    case Word.WdOutlineLevel.wdOutlineLevel4:
                    case Word.WdOutlineLevel.wdOutlineLevel5:
                    case Word.WdOutlineLevel.wdOutlineLevel6:
                    case Word.WdOutlineLevel.wdOutlineLevel7:
                    case Word.WdOutlineLevel.wdOutlineLevel8:
                    case Word.WdOutlineLevel.wdOutlineLevel9:
                    case Word.WdOutlineLevel.wdOutlineLevelBodyText:
                        if (para.Range.Start == para.Range.End - 1)
                        {
                            // how to handle this kind of situation?
                            continue;
                        }

                        while (preNode != rootNode)
                        {
                            Word.Paragraph prePara = (Word.Paragraph)preNode.Tag;
                            if ((int)para.OutlineLevel - (int)prePara.OutlineLevel == 1)
                            {
                                break;
                            }
                            else if ((int)para.OutlineLevel - (int)prePara.OutlineLevel > 0)
                            {
                                break;
                            }
                            preNode = preNode.Parent;
                        }
                        newNode = new TreeNode(strItem);
                        newNode.Name = strItem;
                        newNode.ImageIndex = newNode.SelectedImageIndex = (int)para.OutlineLevel;
                        newNode.Tag = para;
                        preNode.Nodes.Add(newNode);
                        preNode = newNode;
                        break;

                }// switch

            }

            return rootNode;
        }


        private void buildOutline(ref TreeNode rootNode, String[] strOutlineNames, int[] nOutlineLevels)
        {
            String strItem = "";
            int nCurOutline = 0;
            TreeNode newNode = null;

            rootNode.Nodes.Clear();
            TreeNode preNode = rootNode;

            for (int i = 0; i < strOutlineNames.GetLength(0); i++)
            {
                newNode = null;
                nCurOutline = nOutlineLevels[i];

                strItem = strOutlineNames[i];

                switch ((Word.WdOutlineLevel)nCurOutline)
                {
                    case Word.WdOutlineLevel.wdOutlineLevel1:
                        newNode = new TreeNode(strItem);
                        newNode.Name = strItem;
                        newNode.ImageIndex = newNode.SelectedImageIndex = (int)nOutlineLevels[i];
                        newNode.Tag = (Word.WdOutlineLevel)nOutlineLevels[i];

                        rootNode.Nodes.Add(newNode);
                        preNode = newNode;
                        break;

                    case Word.WdOutlineLevel.wdOutlineLevel2:
                    case Word.WdOutlineLevel.wdOutlineLevel3:
                    case Word.WdOutlineLevel.wdOutlineLevel4:
                    case Word.WdOutlineLevel.wdOutlineLevel5:
                    case Word.WdOutlineLevel.wdOutlineLevel6:
                    case Word.WdOutlineLevel.wdOutlineLevel7:
                    case Word.WdOutlineLevel.wdOutlineLevel8:
                    case Word.WdOutlineLevel.wdOutlineLevel9:
                    case Word.WdOutlineLevel.wdOutlineLevelBodyText:

                        while (preNode != rootNode)
                        {
                            Word.WdOutlineLevel preParaOutline = (Word.WdOutlineLevel)preNode.Tag;
                            if ((int)nCurOutline - (int)preParaOutline == 1)
                            {
                                break;
                            }
                            else if ((int)nCurOutline - (int)preParaOutline > 0)
                            {
                                break;
                            }
                            preNode = preNode.Parent;
                        }
                        newNode = new TreeNode(strItem);
                        newNode.Name = strItem;
                        newNode.ImageIndex = newNode.SelectedImageIndex = (int)nOutlineLevels[i];
                        newNode.Tag = (Word.WdOutlineLevel)nOutlineLevels[i];
                        preNode.Nodes.Add(newNode);
                        preNode = newNode;
                        break;

                }// switch

            }

            return;
        }




        private void findHeadingsParas(TreeNode node, ref Word.Selection sel)
        {
            if (node.Tag == null)
                return;

            Word.Paragraph para = null;

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sel.Find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)node.Tag;

            sel.Find.Text = "";
            sel.Find.Replacement.Text = "";
            sel.Find.Forward = true;
            sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
            sel.Find.Format = true;
            sel.Find.MatchCase = false;
            sel.Find.MatchWholeWord = false;
            sel.Find.MatchByte = false;
            sel.Find.MatchWildcards = false;
            sel.Find.MatchSoundsLike = false;
            sel.Find.MatchAllWordForms = false;

            sel.Find.Execute();

            //             sel.Find.Execute2007(ref FindText, ref MatchCase, ref MatchWholeWord, ref MatchWildcards,
            //                                                 ref MatchSoundsLike, ref MatchAllWordForms, ref Forward,
            //                                                 ref Wrap, ref Format, ref ReplaceWith, ref Replace,
            //                                                 ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza,
            //                                                 ref MatchControl, ref MatchPrefix, ref MatchSuffix,
            //                                                 ref MatchPhrase, ref IgnoreSpace, ref IgnorePunct);

            if (sel.Find.Found && sel.Range.Paragraphs.Count > 0)
            {
                para = sel.Range.Paragraphs[1];
                String strParaCnt = para.Range.Text.Trim();
                String strOutline = "";

                if (!para.Range.ListFormat.ListString.Equals(""))
                {
                    strOutline = para.Range.ListFormat.ListString;

                    if (node.Text.Equals(strOutline + " " + strParaCnt))
                    {
                        node.Tag = para;
                    }
                }
                else
                {
                    if (node.Text.Equals(strParaCnt))
                    {
                        node.Tag = para;
                    }
                }
            }
            else
            {
                node.Tag = null;
            }

            foreach (TreeNode childNode in node.Nodes)
            {
                findHeadingsParas(childNode, ref sel);
            }

            return;
        }


        private TreeNode getDocOutlineTree(Word.Document curDoc)
        {
            Word.Selection sel = curDoc.ActiveWindow.Selection;

            // 此技术不能完全覆盖大纲级别的章节
            Object objType = Word.WdReferenceType.wdRefTypeHeading;
            Object dynHeadings = curDoc.GetCrossReferenceItems(objType);

            Array strHeadings = (Array)dynHeadings;

            int[] nLevelArr = null;
            String[] strOutlineNames = null;
            int nIndex = 0, i = 0;

            if (strHeadings != null)
            {
                char[] chArr = null;
                nLevelArr = new int[strHeadings.GetLength(0)];

                foreach (String strItem in strHeadings)
                {
                    // count blanks with starting
                    chArr = strItem.ToCharArray();

                    for (i = 0; i < chArr.GetLength(0); i++)
                    {
                        if (chArr[i] != ' ')
                        {
                            break;
                        }
                    }

                    nLevelArr[nIndex] = (i >> 1) + 1;
                    nIndex++;
                }

                strOutlineNames = new String[strHeadings.GetLength(0)];

                i = 0;
                foreach (String strItem in strHeadings)
                {
                    strOutlineNames[i] = strItem.Trim();
                    i++;
                }

            }

            TreeNode rootNode = new TreeNode();
            buildOutline(ref rootNode, strOutlineNames, nLevelArr);

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            sel.Find.ClearFormatting();
            // 
            foreach (TreeNode childNode in rootNode.Nodes)
            {
                findHeadingsParas(childNode, ref sel);
            }

            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            sel.Start = nOStart;
            sel.End = nOEnd;
            sel.Range.GoTo();
            curDoc.ActiveWindow.ScrollIntoView(sel.Range);

            return rootNode;
        }

        static public Hashtable Decode(String strInput)
        {
            ArrayList strArr = new ArrayList();

            strArr = parseProtocol(strInput);

            Hashtable hashFields = parseFields(strArr);

            return hashFields;
        }


        static public String Encode(Hashtable hashFields)
        {
            String strResult = "";

            String strName = "", strValue = "", strItem = "";

            foreach (DictionaryEntry ent in hashFields)
            {
                strName = (String)ent.Key;
                strValue = (String)ent.Value;

                strItem = "[" + strName + ":" + strValue + "]";

                strResult += strItem;

            }

            return strResult;
        }


        static private Hashtable parseFields(ArrayList strArr)
        {
            Hashtable hashFields = new Hashtable();

            int nIndex = -1;

            String strName = "", strValue = "";

            foreach (String strItem in strArr)
            {
                nIndex = strItem.IndexOf(':');

                if (nIndex != -1)
                {
                    strName = strItem.Substring(0, nIndex);

                    if ((nIndex + 1) < strItem.Length)
                    {
                        strValue = strItem.Substring(nIndex + 1);
                    }
                    else
                    {
                        strValue = "";
                    }

                    hashFields[strName] = strValue;
                }

            }

            return hashFields;
        }



        // parse (\[([^\[\]])*\])*
        static private ArrayList parseProtocol(String strInput)
        {
            ArrayList strArr = new ArrayList();

            char[] chArr = strInput.ToCharArray();
            int nStart = -1, nEnd = -1, i = 0;

            for (i = 0; i < chArr.GetLength(0); i++)
            {
                if (chArr[i] == '[')
                {
                    nStart = i + 1;
                }
                else if (chArr[i] == ']')
                {
                    nEnd = i - 1;
                    String strItem = "";
                    strItem = strInput.Substring(nStart, nEnd - nStart + 1);
                    strArr.Add(strItem);
                }
            }

            return strArr;
        }


        public Office.CommandBarControl getFontDialogControl(Word.Application app)
        {
            Office.CommandBar cb = app.CommandBars["Headings"];

            Office.CommandBarControl cbcFont = null;

            if(cb != null)
            {
                foreach (Office.CommandBarControl cbc in cb.Controls)
                {
                    if (cbc.accName.Equals("字体​​...") || cbc.Id == 253) // word 2007
                    {
                        cbcFont = cbc;
                        break;
                    }
                }
            }

            return cbcFont;

        }


        public Office.CommandBarControl getParagraphFormatControl(Word.Application app)
        {
            Office.CommandBar cb = app.CommandBars["Headings"];

            Office.CommandBarControl cbcParagraphFormat = null;

            if (cb != null)
            {
                foreach (Office.CommandBarControl cbc in cb.Controls)
                {
                    if (cbc.accName.Equals("段落​​...") || cbc.Id == 779)// word2007
                    {
                        cbcParagraphFormat = cbc;
                    }
                }
            }

            return cbcParagraphFormat;
        }


        public void copyFontStyle(Word.Font srcFont, Word.Font curFnt)
        {
            curFnt.AllCaps = srcFont.AllCaps;
            curFnt.Animation = srcFont.Animation;

            curFnt.Bold = srcFont.Bold;
            curFnt.BoldBi = srcFont.BoldBi;

            // copyBorders(srcFont.Borders, curFnt.Borders);

            curFnt.Color = srcFont.Color;
            curFnt.ColorIndex = srcFont.ColorIndex;
            curFnt.ColorIndexBi = srcFont.ColorIndexBi;

            curFnt.DiacriticColor = srcFont.DiacriticColor;
            curFnt.DisableCharacterSpaceGrid = srcFont.DisableCharacterSpaceGrid;
            curFnt.DoubleStrikeThrough = srcFont.DoubleStrikeThrough;

            curFnt.Emboss = srcFont.Emboss;
            curFnt.EmphasisMark = srcFont.EmphasisMark;
            curFnt.Engrave = srcFont.Engrave;
            curFnt.Hidden = srcFont.Hidden;
            curFnt.Italic = srcFont.Italic;
            curFnt.ItalicBi = srcFont.ItalicBi;
            curFnt.Kerning = srcFont.Kerning;
            curFnt.Name = srcFont.Name;
            curFnt.NameAscii = srcFont.NameAscii;
            curFnt.NameBi = srcFont.NameBi;
            curFnt.NameFarEast = srcFont.NameFarEast;
            curFnt.NameOther = srcFont.NameOther;
            curFnt.Outline = srcFont.Outline;
            curFnt.Position = srcFont.Position;
            curFnt.Scaling = srcFont.Scaling;

            curFnt.Shadow = srcFont.Shadow;
            curFnt.Size = srcFont.Size;
            curFnt.SizeBi = srcFont.SizeBi;
            curFnt.SmallCaps = srcFont.SmallCaps;
            curFnt.Spacing = srcFont.Spacing;
            curFnt.StrikeThrough = srcFont.StrikeThrough;
            curFnt.Subscript = srcFont.Subscript;
            curFnt.Superscript = srcFont.Superscript;
            curFnt.Underline = srcFont.Underline;
            curFnt.UnderlineColor = srcFont.UnderlineColor;

            return;
        }

        public void copyParagraphFormat(Word.ParagraphFormat srcParaFormat, Word.ParagraphFormat dstParaFormat)
        {
            dstParaFormat.AddSpaceBetweenFarEastAndAlpha = srcParaFormat.AddSpaceBetweenFarEastAndAlpha;

            dstParaFormat.AddSpaceBetweenFarEastAndDigit = srcParaFormat.AddSpaceBetweenFarEastAndDigit;
            dstParaFormat.Alignment = srcParaFormat.Alignment;

            dstParaFormat.AutoAdjustRightIndent = srcParaFormat.AutoAdjustRightIndent;
            dstParaFormat.BaseLineAlignment = srcParaFormat.BaseLineAlignment;

            dstParaFormat.CharacterUnitFirstLineIndent = srcParaFormat.CharacterUnitFirstLineIndent;
            dstParaFormat.CharacterUnitLeftIndent = srcParaFormat.CharacterUnitLeftIndent;
            dstParaFormat.CharacterUnitRightIndent = srcParaFormat.CharacterUnitRightIndent;
            dstParaFormat.DisableLineHeightGrid = srcParaFormat.DisableLineHeightGrid;
            dstParaFormat.FarEastLineBreakControl = srcParaFormat.FarEastLineBreakControl;
            dstParaFormat.FirstLineIndent = srcParaFormat.FirstLineIndent;
            dstParaFormat.HalfWidthPunctuationOnTopOfLine = srcParaFormat.HalfWidthPunctuationOnTopOfLine;
            dstParaFormat.HangingPunctuation = srcParaFormat.HangingPunctuation;
            dstParaFormat.Hyphenation = srcParaFormat.Hyphenation;
            dstParaFormat.KeepTogether = srcParaFormat.KeepTogether;
            dstParaFormat.KeepWithNext = srcParaFormat.KeepWithNext;
            dstParaFormat.LeftIndent = srcParaFormat.LeftIndent;
            dstParaFormat.LineSpacing = srcParaFormat.LineSpacing;
            dstParaFormat.LineSpacingRule = srcParaFormat.LineSpacingRule;
            dstParaFormat.LineUnitAfter = srcParaFormat.LineUnitAfter;
            dstParaFormat.LineUnitBefore = srcParaFormat.LineUnitBefore;
            dstParaFormat.MirrorIndents = srcParaFormat.MirrorIndents;
            dstParaFormat.OutlineLevel = srcParaFormat.OutlineLevel;
            dstParaFormat.NoLineNumber = srcParaFormat.NoLineNumber;
            dstParaFormat.PageBreakBefore = srcParaFormat.PageBreakBefore;
            dstParaFormat.ReadingOrder = srcParaFormat.ReadingOrder;
            dstParaFormat.RightIndent = srcParaFormat.RightIndent;

            dstParaFormat.SpaceAfter = srcParaFormat.SpaceAfter;
            dstParaFormat.SpaceAfterAuto = srcParaFormat.SpaceAfterAuto;
            dstParaFormat.SpaceBefore = srcParaFormat.SpaceBefore;
            dstParaFormat.SpaceBeforeAuto = srcParaFormat.SpaceBeforeAuto;
            dstParaFormat.TabStops = srcParaFormat.TabStops;

            if (bAppIsWps)
            {

            }
            else
            {
                dstParaFormat.TextboxTightWrap = srcParaFormat.TextboxTightWrap;
                dstParaFormat.WidowControl = srcParaFormat.WidowControl;
                dstParaFormat.WordWrap = srcParaFormat.WordWrap;
            }

            return;
        }

        //
        public String genTableTopoKey(Word.Table tbl)
        {
            String strKey = "";

            if (tbl != null)
            {
                Word.Cell cel = null;
                String strTopoInfo = "";

                for (int i = 1; i <= tbl.Rows.Count; i++)
                {
                    for (int j = 1; j <= tbl.Columns.Count; j++)
                    {
                        try
                        {
                            cel = tbl.Cell(i, j);
                            strTopoInfo += "[" + cel.RowIndex + "," + cel.ColumnIndex + "]";
                        }
                        catch (System.Exception ex)
                        {
                            continue;
                        }
                    }
                }

                strKey = ClassEncryptUtils.MD5Encrypt(strTopoInfo);
            }

            return strKey;
        }


        public TreeNode locateNode(TreeNode rootNode, String strFullPath, ref int nPathDepth)
        {
            String[] arrStrPath = strFullPath.Split('\\');

            if (arrStrPath.GetLength(0) == 0)
            {
                return null;
            }

            return locateNode(rootNode, arrStrPath, ref nPathDepth);
        }


        public TreeNode locateNode(TreeNode rootNode, String[] arrStrPath, ref int nPathDepth)
        {
            TreeNode fndNode = null;

            if (nPathDepth < arrStrPath.GetLength(0))
            {
                if (rootNode.Text.Equals(arrStrPath[nPathDepth]))
                {
                    nPathDepth++;
                    foreach (TreeNode childNd in rootNode.Nodes)
                    {
                        fndNode = locateNode(childNd, arrStrPath, ref nPathDepth);
                        if (fndNode != null)
                        {
                            return fndNode;
                        }
                    }

                    return rootNode;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }

        }

        public void RecordMultiSel(Word.Range rng)
        {
            Word.Editors editors = rng.Editors;
            Word.Editor edt = editors.Add(Word.WdEditorType.wdEditorEveryone); // 加入

            return;
        }

        public void ExecMultiSel(Word.Document doc)
        {
            doc.SelectAllEditableRanges(Word.WdEditorType.wdEditorEveryone); // 选择
            doc.DeleteAllEditableRanges(Word.WdEditorType.wdEditorEveryone);

            return;
        }

        public int getNavKeyWordBookmk( Word.Document doc,String strStartKeyWord,
                                    ref Word.Bookmark firstBkmk, ref Word.Bookmark lastBkmk,
                                    ref Word.Bookmark nearstPrevBkmk, ref Word.Bookmark nearstNextBkmk)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;

            int nMinStart = -1, nMaxStart = -1;
            int nPrevDis = -1, nNextDis = -1;
            int nCurDisPrev = 0, nCurDisNext = 0;

            foreach (Word.Bookmark bkmk in doc.Bookmarks)
            {
                if (!bkmk.Name.StartsWith(strStartKeyWord))
                {
                    continue;
                }

                if (nMinStart == -1 || bkmk.Range.Start < nMinStart)
                {
                    nMinStart = bkmk.Range.Start;
                    firstBkmk = bkmk;
                }

                if (nMaxStart == -1 || bkmk.Range.Start > nMaxStart)
                {
                    nMaxStart = bkmk.Range.Start;
                    lastBkmk = bkmk;
                }

                nCurDisPrev = sel.Start - bkmk.Range.Start;
                if (nCurDisPrev > 0)
                {
                    if (nPrevDis == -1 || nCurDisPrev < nPrevDis)
                    {
                        nPrevDis = nCurDisPrev;
                        nearstPrevBkmk = bkmk;
                    }
                }

                nCurDisNext = bkmk.Range.Start - sel.Start;
                if (nCurDisNext > 0)
                {
                    if (nNextDis == -1 || nCurDisNext < nNextDis)
                    {
                        nNextDis = nCurDisNext;
                        nearstNextBkmk = bkmk;
                    }
                }
            }

            return 0;
        }


        public int getNavBookmk(Word.Document doc,ref Word.Bookmark firstBkmk, ref Word.Bookmark lastBkmk,
                                 ref Word.Bookmark nearstPrevBkmk, ref Word.Bookmark nearstNextBkmk)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;

            int nMinStart = -1, nMaxStart = -1;
            int nPrevDis = -1, nNextDis = -1;
            int nCurDisPrev = 0, nCurDisNext = 0;

            foreach (Word.Bookmark bkmk in doc.Bookmarks)
            {
                if (nMinStart == -1 || bkmk.Range.Start < nMinStart)
                {
                    nMinStart = bkmk.Range.Start;
                    firstBkmk = bkmk;
                }

                if (nMaxStart == -1 || bkmk.Range.Start > nMaxStart)
                {
                    nMaxStart = bkmk.Range.Start;
                    lastBkmk = bkmk;
                }

                nCurDisPrev = sel.Start - bkmk.Range.Start;
                if (nCurDisPrev > 0)
                {
                    if (nPrevDis == -1 || nCurDisPrev < nPrevDis)
                    {
                        nPrevDis = nCurDisPrev;
                        nearstPrevBkmk = bkmk;
                    }
                }

                nCurDisNext = bkmk.Range.Start - sel.Start;
                if (nCurDisNext > 0)
                {
                    if (nNextDis == -1 || nCurDisNext < nNextDis)
                    {
                        nNextDis = nCurDisNext;
                        nearstNextBkmk = bkmk;
                    }
                }
            }

            return 0;
        }


        public int changeHeadingsParasStyle(Word.Application app, Word.Document dstDoc, ref String strRetMsg,
                                   ArrayList arrClassFont, ArrayList arrParaFmt,
                                   Boolean bIgnoreToc, Boolean bIgnoreTable,
                                   Boolean bIgnorePages, uint nIgnorePages,
                                   Boolean bKeepFont, Boolean bKeepParagraphFmt,
                                   Word.Range scopeRange = null,
                                   ClassListLevel[] cListLevels = null,
                                   Word.ListLevels wListLevels = null,
                                   ProgressBar progBar = null)
        {

            if (bKeepFont && bKeepParagraphFmt)
            {
                strRetMsg = "同时保留字体和段落样式，则无须改变";
                return -1;
            }

            if (bIgnorePages)
            {
                int nPages = dstDoc.Content.get_Information(Word.WdInformation.wdNumberOfPagesInDocument); // .ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (nIgnorePages >= nPages)
                {
                    strRetMsg = "忽略页数大于文档总页数";
                    return -1;
                }
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;
            ClassFont srcFnt = null;
            ClassParagraphFormat srcParaFmt = null;
            Word.Style headingStyle = null;
            String strHeading = "标题 ";
            dynamic styleDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatStyle];

            ArrayList arrParaArrs = null, arrParas = null;
            Boolean bWillChange = false;

            Word.Document doc = null;

            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("无活动文档，不能应用");
                return -1;
            }
            finally
            {
            }

            int oSaveInterval = app.Options.SaveInterval;
            app.Options.SaveInterval = 120;


            int[] nHeadings = new int[9];

            for (int i = 0; i < 9; i++)
            {
                nHeadings[i] = (i+1);
            }

            arrParaArrs = getSpecificHeadingParaArrsInScope(dstDoc, scopeRange, nHeadings, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages);

            if (progBar != null)
            {
                progBar.Maximum = 9;//arrHeadingParas.Count;

                if (cListLevels != null || wListLevels != null)
                {
                    progBar.Maximum += 9;
                }
            }

            for (int i = 1; i <= 9; i++)
            {
                arrParas = (ArrayList)arrParaArrs[i];
                if (arrParas.Count > 0)
                {
                    bWillChange = true;
                    break;
                }
            }

            if (!bWillChange)
            {
                strRetMsg = "没有符合条件的章节段落";
                return -1;
            }


            Word.Style st = null;

            Word.ListGallery listGallery = null;
            Word.ListTemplate lstTemplate = null;
            Word.ListLevels lstLvels = null;
            Object objIndex = 1;

            if (bAppIsWps)
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

            //Object objIndex = 1;
            //Word.ListGallery listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];
            //Word.ListTemplate lstTemplate = listGallery.ListTemplates[objIndex];
            //Word.ListLevels lstLvels = listGallery.ListTemplates[objIndex].ListLevels;

            Object objListLevel = 1;

            if (cListLevels != null)
            {
                if (progBar != null)
                {
                    progBar.Value += 9;
                }

                setTemplateList(app, lstLvels, cListLevels);
            }
            else if (wListLevels != null)
            {
                if (progBar != null)
                {
                    progBar.Value += 9;
                }

                setTemplateList(app, lstLvels, wListLevels);
            }
            else
            {
            }

            int nTotalParas = 0;
            String strStat = "";

            for (int i = 1; i <= 9; i++)
            {
                arrParas = (ArrayList)arrParaArrs[i];

                if (arrParas.Count > 0)
                {
                    nTotalParas += arrParas.Count;
                    strStat += i + "级：" + arrParas.Count + "\r\n";

                    //foreach (Word.Paragraph para in arrParas)
                    //{
                    //    RecordMultiSel(para.Range);
                    //}
                    //ExecMultiSel(dstDoc);

                    // copy form into styles
                    strHeading = "标题 " + i;

                    try
                    {
                        st = dstDoc.Styles[strHeading];
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("没有内置样式名：\"" + strHeading + "\"的样式\r\n" + ex.Message);
                    }
                    finally
                    {

                    }


                    if (st != null)
                    {
                        if ((i - 1) < arrClassFont.Count)
                        {
                            srcFnt = (ClassFont)arrClassFont[i - 1];
                            srcParaFmt = (ClassParagraphFormat)arrParaFmt[i - 1];
                        }
                        else
                        {
                            srcFnt = null;
                            srcParaFmt = null;
                        }

                        if (!bKeepFont && srcFnt != null)
                        {
                            srcFnt.copy2(st.Font);
                        }

                        if (!bKeepParagraphFmt && srcParaFmt != null)
                        {
                            srcParaFmt.copy2(st.ParagraphFormat);
                        }

                        if (cListLevels != null || wListLevels != null)
                        {
                            if (bAppIsWps)
                            {
                                copyFontStyle(st.Font, lstLvels[i].Font);
                            }

                            objListLevel = i;
                            if (st.ListTemplate != null)
                            {
                                st.LinkToListTemplate(null);
                            }

                            st.LinkToListTemplate(lstTemplate, objListLevel);

                            if (progBar != null)
                            {
                                progBar.Value++;
                            }
                        }

                        foreach (Word.Paragraph para in arrParas)
                        {
                            para.Range.set_Style(st);

                            if (bAppIsWps)
                            {
                                if (para.Range.ListFormat != null && para.Range.ListFormat.ListTemplate != null)
                                {
                                    // no effect
                                    copyFontStyle(st.Font, para.Range.ListFormat.ListTemplate.ListLevels[i].Font);
                                }
                            }
                        }

                    }

                    // sel.set_Style(st);
                }

                if (progBar != null)
                {
                    progBar.Value++;
                }
            }

            if (nTotalParas > 0)
            {
                strRetMsg += "章节总数：" + nTotalParas + "\r\n" + strStat;
            }

            return 0;
        }


        private int changeHeadingsParasStyle_v1(Word.Application app, Word.Document dstDoc, ref String strRetMsg,
                                   ArrayList arrClassFont, ArrayList arrParaFmt,
                                   Boolean bIgnoreToc, Boolean bIgnoreTable,
                                   Boolean bIgnorePages, uint nIgnorePages,
                                   Boolean bKeepFont, Boolean bKeepParagraphFmt,
                                   Word.Range scopeRange = null,
                                   ClassListLevel[] cListLevels = null,
                                   Word.ListLevels wListLevels = null,
                                   ProgressBar progBar = null)
        {

            if (bKeepFont && bKeepParagraphFmt)
            {
                strRetMsg = "同时保留字体和段落样式，则无须改变";
                return -1;
            }

            if (bIgnorePages)
            {
                int nPages = dstDoc.Content.get_Information(Word.WdInformation.wdNumberOfPagesInDocument); // .ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (nIgnorePages >= nPages)
                {
                    strRetMsg = "忽略页数大于文档总页数";
                    return -1;
                }
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;
            ClassFont srcFnt = null;
            ClassParagraphFormat srcParaFmt = null;
            Word.Style headingStyle = null;
            String strHeading = "标题 ";

            //             int nOStart = sel.Start;
            //             int nOEnd = sel.End;
            // 
            //             Word.WdViewType oViewType = dstDoc.ActiveWindow.View.Type;

            // get all headings
            // ArrayList targetArr = getHeadingParasInScopeByNavNoChangeView(app, dstDoc, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages, scopeRange);

            Boolean bInToc = false;
            Boolean bInTables = false;
            Boolean bInIgnorePages = false;

            ArrayList targetArr = new ArrayList();
            ArrayList tmpArr = getSpecificHeadingParasInScope(dstDoc,scopeRange);

            foreach(Word.Paragraph paraItem in tmpArr)
            {
                bInToc = false;
                bInTables = false;
                bInIgnorePages = false;

                if (bIgnoreToc)
                {
                    foreach (Word.TableOfContents toc in dstDoc.TablesOfContents)
                    {
                        if (paraItem.Range.InRange(toc.Range))
                        {
                            bInToc = true;
                            break;
                        }
                    }

                }//

                if (bIgnoreTable)
                {
                    bInTables = paraItem.Range.get_Information(Word.WdInformation.wdWithInTable);
                }


                if (bIgnorePages)
                {
                    int nPageSn = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                    bInIgnorePages = (nPageSn <= nIgnorePages);
                }

                if (!((bIgnoreToc && bInToc) || (bIgnoreTable && bInTables) || (bIgnorePages && bInIgnorePages)))
                {
                    if (!String.IsNullOrWhiteSpace(paraItem.Range.Text.Trim(m_trimChars)))
                    {
                        targetArr.Add(paraItem);
                    }
                }
            }


            // (app, dstDoc, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages, scopeRange);

            if (targetArr.Count == 0)
            {
                strRetMsg = "没有符合条件的章节段落";
                return -1;
            }

            if (progBar != null)
            {
                progBar.Maximum += targetArr.Count + 9;

                if (cListLevels != null || wListLevels != null)
                {
                    progBar.Maximum += 9;
                }
            }

            foreach (Word.Paragraph tPara in targetArr)
            {
                if (progBar != null)
                {
                    progBar.Value++;
                }

                if (tPara.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText)
                {
                    continue;
                }

                tPara.Range.GoTo();
                dstDoc.ActiveWindow.ScrollIntoView(tPara.Range);

                int nOutlineLevel = (int)tPara.OutlineLevel;

                try
                {
                    headingStyle = dstDoc.Styles[strHeading + nOutlineLevel]; // 取章节
                }
                catch (System.Exception ex)
                {
                    //MessageBox.Show("关联的内置样式：\'" + strHeading + nLvl + "\'异常，请检查此文档的此名称内置样式是否存在！");
                    if (strRetMsg.IndexOf(strHeading + nOutlineLevel) == -1)
                    {
                        strRetMsg += "关联的内置样式：\'" + strHeading + nOutlineLevel + "\'异常，请检查此文档的此名称内置样式是否存在！\r\n";
                    }

                    continue;
                }
                finally
                {

                }

                if (headingStyle != null)
                {
                    tPara.set_Style(headingStyle); // 设置样式
                }

            }

            // change heading styles in style
            Word.Style st = null;
            Object objIndex = 1;

            Word.ListGallery listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];

            Word.ListTemplate lstTemplate = listGallery.ListTemplates[objIndex];
            Word.ListLevels lstLvels = listGallery.ListTemplates[objIndex].ListLevels;
            Object objListLevel = 1;

            Object objContinue = false;
            Object objApplyTo = Word.WdListApplyTo.wdListApplyToSelection;// wdListApplyToWholeList;
            Object objDefaultBehav = Word.WdDefaultListBehavior.wdWord10ListBehavior;

            if (cListLevels != null)
            {
                if (progBar != null)
                {
                    progBar.Value += 9;
                }

                setTemplateList(app, lstLvels,cListLevels);
            }
            else if (wListLevels != null)
            {
                if (progBar != null)
                {
                    progBar.Value += 9;
                }

                setTemplateList(app, lstLvels,wListLevels);
            }
            else
            {
            }

            for (int nLvl = (int)Word.WdOutlineLevel.wdOutlineLevel1; nLvl < (int)Word.WdOutlineLevel.wdOutlineLevelBodyText; nLvl++)
            {
                if (progBar != null)
                {
                    progBar.Value++;
                }

                if ((nLvl - 1) < arrClassFont.Count)
                {
                    srcFnt = (ClassFont)arrClassFont[nLvl - 1];
                    srcParaFmt = (ClassParagraphFormat)arrParaFmt[nLvl - 1];
                }
                else
                {
                    srcFnt = null;
                    srcParaFmt = null;
                }

                // copy form into styles
                strHeading = "标题 " + nLvl;

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

                if (st != null)
                {
                    if (!bKeepFont && srcFnt != null)
                    {
                        srcFnt.copy2(st.Font);
                    }

                    if (!bKeepParagraphFmt && srcParaFmt != null)
                    {
                        srcParaFmt.copy2(st.ParagraphFormat);
                    }

                    if (cListLevels != null || wListLevels != null)
                    {
                        objListLevel = nLvl;
                        if (st.ListTemplate != null)
                        {
                            st.LinkToListTemplate(null);
                        }
                        st.LinkToListTemplate(lstTemplate, objListLevel);
                    }

                }

            } // for

            //             // 恢复特定view
            //             dstDoc.ActiveWindow.View.Type = oViewType;
            // 
            //             // restore original position
            //             sel.Start = nOStart;
            //             sel.End = nOEnd;
            //             // sel.Range.Select();
            //             sel.Range.GoTo();
            //             dstDoc.ActiveWindow.ScrollIntoView(sel.Range); // 视角恢复

            return 0;
        }


        public int changeTextBodyParasStyle(Word.Application app, Word.Document dstDoc, ref String strRetMsg,
                                           ArrayList arrClassFont, ArrayList arrParaFmt,
                                           Boolean bIgnoreToc, Boolean bIgnoreTable,
                                           Boolean bIgnorePages, uint nIgnorePages,
                                           Boolean bKeepFont, Boolean bKeepParagraphFmt,
                                           Word.Range scopeRange = null,
                                           ClassListLevel[] cListLevels = null,
                                           Word.ListLevels wListLevels = null,
                                           ProgressBar progBar = null)
        {

            if (bKeepFont && bKeepParagraphFmt)
            {
                strRetMsg = "同时保留字体和段落样式，则无须改变";
                return -1;
            }


            if (bIgnorePages)
            {
                int nPages = dstDoc.Content.get_Information(Word.WdInformation.wdNumberOfPagesInDocument); // .ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (nIgnorePages >= nPages)
                {
                    strRetMsg = "忽略页数大于文档总页数";
                    return -1;
                }
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;
            ClassFont srcFnt = null;
            ClassParagraphFormat srcParaFmt = null;

            int nLvl = (int)Word.WdOutlineLevel.wdOutlineLevelBodyText;

            if ((nLvl - 1) < arrClassFont.Count)
            {
                srcFnt = (ClassFont)arrClassFont[nLvl - 1];
                srcParaFmt = (ClassParagraphFormat)arrParaFmt[nLvl - 1];
            }
            else
            {
                srcFnt = null;
                srcParaFmt = null;
            }

            if (srcFnt == null && srcParaFmt == null)
            {
                strRetMsg = "没有源正文样式";
                return -1;
            }

            //
            ArrayList arrTextBodyParas = getSpecificTextBodyParasInScopeNoChangeView(dstDoc, scopeRange, bIgnoreTable, bIgnoreToc, bIgnorePages, nIgnorePages);

            ArrayList targetTextBodyParas = new ArrayList();

            if (arrTextBodyParas.Count == 0)
            {
                strRetMsg = "没有目标段落";
                return -1;
            }

            Boolean bInIgnorePages = false;

            if (bIgnorePages)
            {
                if (progBar != null)
                {
                    progBar.Maximum += arrTextBodyParas.Count;
                }

                foreach (Word.Paragraph para in arrTextBodyParas)
                {
                    if (progBar != null)
                    {
                        progBar.Value++;
                    }

                    if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        continue;
                    }

                    bInIgnorePages = false;

                    int nPageSn = para.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                    bInIgnorePages = (nPageSn <= nIgnorePages);

                    if (bInIgnorePages)
                    {
                        continue;
                    }

                    targetTextBodyParas.Add(para);
                }
            }
            else
            {
                targetTextBodyParas = arrTextBodyParas;
            }

            if (targetTextBodyParas.Count == 0)
            {
                strRetMsg = "没有符合条件的段落";
                return -1;
            }


            if (progBar != null)
            {
                progBar.Maximum += targetTextBodyParas.Count;
            }

            // 新建或取 样式：“目标正文”
            // 
            // 设置上述样式的FONT/PARAGRAPHFORMAT
            // 
            // 每100个多选，设置样式名：“目标正文”
            // 

            const int nMaxBatchNum = 100;

            String strStyleName = "目标正文";
            Object objType = Word.WdStyleType.wdStyleTypeParagraph;
            Word.Style myTxtStyle = null;
            Word.Paragraph tPara = null;

            tPara = (Word.Paragraph)targetTextBodyParas[0];
            sel.Start = tPara.Range.Start;
            sel.End = tPara.Range.End;

            sel.Range.GoTo();

            try
            {
                myTxtStyle = dstDoc.Styles[strStyleName];
            }
            catch (System.Exception ex)
            {
                myTxtStyle = dstDoc.Styles.Add(strStyleName, objType);
            }
            finally
            {
            }

            if(myTxtStyle != null)
            {
                myTxtStyle.LinkToListTemplate(null);

                if (!bKeepFont && srcFnt != null)
                {
                    srcFnt.copy2(myTxtStyle.Font);
                }

                if (!bKeepParagraphFmt && srcParaFmt != null)
                {
                    srcParaFmt.copy2(myTxtStyle.ParagraphFormat);
                }
            }

            
            int nUpper = 0;
            dynamic styleDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatStyle];

            strRetMsg = "正文总数：" + targetTextBodyParas.Count;

            while(targetTextBodyParas.Count > 0)
            {
                nUpper = Math.Min(nMaxBatchNum, targetTextBodyParas.Count);

                for (int i = 0; i < nUpper; i++)
                {
                    tPara = (Word.Paragraph)targetTextBodyParas[i];
                    RecordMultiSel(tPara.Range);
                }

                ExecMultiSel(dstDoc);

                dstDoc.ActiveWindow.ScrollIntoView(tPara.Range);

                try
                {
                    //styleDialog.Name = strStyleName;
                    //styleDialog.Apply();
                    sel.set_Style(myTxtStyle);
                }
                catch (System.Exception ex)
                {
                    // MessageBox.Show("没有内置样式名：'" + "标题 " + i + "'的样式\r\n" + ex.Message);
                }
                finally
                {
                }

                targetTextBodyParas.RemoveRange(0, nUpper);
                if (progBar != null)
                {
                    progBar.Value += nUpper;
                }
            }

            return 0;
        }


        private int changeTextBodyParasStyle_v1(Word.Application app, Word.Document dstDoc, ref String strRetMsg,
                                           ArrayList arrClassFont, ArrayList arrParaFmt,
                                           Boolean bIgnoreToc, Boolean bIgnoreTable,
                                           Boolean bIgnorePages, uint nIgnorePages,
                                           Boolean bKeepFont, Boolean bKeepParagraphFmt,
                                           Word.Range scopeRange = null,
                                           ClassListLevel[] cListLevels = null,
                                           Word.ListLevels wListLevels = null,
                                           ProgressBar progBar = null)
        {

            if (bKeepFont && bKeepParagraphFmt)
            {
                strRetMsg = "同时保留字体和段落样式，则无须改变";
                return -1;
            }


            if (bIgnorePages)
            {
                int nPages = dstDoc.Content.get_Information(Word.WdInformation.wdNumberOfPagesInDocument); // .ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (nIgnorePages >= nPages)
                {
                    strRetMsg = "忽略页数大于文档总页数";
                    return -1;
                }
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;
            ClassFont srcFnt = null;
            ClassParagraphFormat srcParaFmt = null;

            int nLvl = (int)Word.WdOutlineLevel.wdOutlineLevelBodyText;

            if ((nLvl - 1) < arrClassFont.Count)
            {
                srcFnt = (ClassFont)arrClassFont[nLvl - 1];
                srcParaFmt = (ClassParagraphFormat)arrParaFmt[nLvl - 1];
            }
            else
            {
                srcFnt = null;
                srcParaFmt = null;
            }

            if (srcFnt == null && srcParaFmt == null)
            {
                strRetMsg = "没有源正文样式";
                return -1;
            }

            //             int nOStart = sel.Start;
            //             int nOEnd = sel.End;
            // 
            //             Word.WdViewType oViewType = dstDoc.ActiveWindow.View.Type;

            //
            ArrayList arrTextBodyParas = getSpecificTextBodyParasInScopeNoChangeView(dstDoc, scopeRange, bIgnoreTable, bIgnoreToc, bIgnorePages, nIgnorePages);

            ArrayList targetTextBodyParas = new ArrayList();

            if (arrTextBodyParas.Count == 0)
            {
                strRetMsg = "没有目标段落";
                return -1;
            }

            Boolean bInIgnorePages = false;

            if (bIgnorePages)
            {
                if (progBar != null)
                {
                    progBar.Maximum += arrTextBodyParas.Count;
                }

                foreach (Word.Paragraph para in arrTextBodyParas)
                {
                    if (progBar != null)
                    {
                        progBar.Value++;
                    }

                    if (para.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                    {
                        continue;
                    }

                    bInIgnorePages = false;

                    int nPageSn = para.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                    bInIgnorePages = (nPageSn <= nIgnorePages);

                    if (bInIgnorePages)
                    {
                        continue;
                    }

                    targetTextBodyParas.Add(para);
                }
            }
            else
            {
                targetTextBodyParas = arrTextBodyParas;
            }

            if (targetTextBodyParas.Count == 0)
            {
                strRetMsg = "没有符合条件的段落";
                return -1;
            }


            //             // 切换到normal view
            //             if (dstDoc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            //             {
            //                 dstDoc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdNormalView;
            //             }
            //             else
            //             {
            //                 dstDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            //             }

            if (progBar != null)
            {
                progBar.Maximum += targetTextBodyParas.Count;
            }

            foreach (Word.Paragraph tPara in targetTextBodyParas)
            {
                if (progBar != null)
                {
                    progBar.Value++;
                }

                tPara.Range.GoTo();
                dstDoc.ActiveWindow.ScrollIntoView(tPara.Range);

                if (!bKeepFont && srcFnt != null)
                {
                    srcFnt.copy2(tPara.Range.Font);
                }

                if (!bKeepParagraphFmt && srcParaFmt != null)
                {
                    srcParaFmt.copy2(tPara.Range.ParagraphFormat);
                }
            }

            //             // 恢复特定view
            //             dstDoc.ActiveWindow.View.Type = oViewType;
            // 
            //             // restore original position
            //             sel.Start = nOStart;
            //             sel.End = nOEnd;
            //             // sel.Range.Select();
            //             sel.Range.GoTo();
            //             dstDoc.ActiveWindow.ScrollIntoView(sel.Range); // 视角恢复

            return 0;
        }

        // change heading styles in highest performance
        public int changeTargetParasStyle(Word.Application app, Word.Document dstDoc, ref String strRetMsg,
                                           ArrayList arrClassFont, ArrayList arrParaFmt,
                                           Boolean bIgnoreToc, Boolean bIgnoreTable,
                                           Boolean bIgnorePages, uint nIgnorePages,
                                           Boolean bIgnoreHeading, Boolean bIgnoreTextBody,
                                           Boolean bKeepFont, Boolean bKeepParagraphFmt,
                                           Word.Range scopeRange = null,
                                           ClassListLevel[] cListLevels = null,
                                           Word.ListLevels wListLevels = null,
                                           ProgressBar progBar = null)
        {
            if (bIgnoreHeading && bIgnoreTextBody)
            {
                strRetMsg = "同时忽略章节和正文，则无操作对象";
                return -1;
            }

            if (bKeepFont && bKeepParagraphFmt)
            {
                strRetMsg = "同时保留字体和段落样式，则无须改变";
                return -1;
            }


            if (bIgnorePages)
            {
                int nPages = dstDoc.Content.get_Information(Word.WdInformation.wdNumberOfPagesInDocument); // .ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (nIgnorePages >= nPages)
                {
                    strRetMsg = "忽略页数大于文档总页数";
                    return -1;
                }
            }

            Word.Selection sel = dstDoc.ActiveWindow.Selection;
            ClassFont srcFnt = null;
            ClassParagraphFormat srcParaFmt = null;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = dstDoc.ActiveWindow.View.Type;

            dstDoc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
            app.Options.Pagination = false;

            if (bIgnoreTextBody) // only headings
            {
                changeHeadingsParasStyle(app, dstDoc, ref strRetMsg, arrClassFont, arrParaFmt, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages, bKeepFont, bKeepParagraphFmt, scopeRange, cListLevels, wListLevels);
            }
            else if (bIgnoreHeading) // only text body
            {
                //changeTextBodyParasStyle(app, dstDoc, ref strRetMsg, arrClassFont, arrParaFmt, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages, bKeepFont, bKeepParagraphFmt, scopeRange, cListLevels, wListLevels);
            }
            else // headings and text body
            {
                String strRet1 = "", strRet2 = "";

                changeHeadingsParasStyle(app, dstDoc, ref strRet1, arrClassFont, arrParaFmt, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages, bKeepFont, bKeepParagraphFmt, scopeRange, cListLevels, wListLevels);
                //changeTextBodyParasStyle(app, dstDoc, ref strRet2, arrClassFont, arrParaFmt, bIgnoreToc, bIgnoreTable, bIgnorePages, nIgnorePages, bKeepFont, bKeepParagraphFmt, scopeRange, cListLevels, wListLevels);

                strRetMsg = strRet1 + "\r\n" + strRet2;
            }

            //app.Options.Pagination = true;

            // 恢复特定view
            dstDoc.ActiveWindow.View.Type = oViewType;

            // restore original position
            sel.Start = nOStart;
            sel.End = nOEnd;
            //// sel.Range.Select();
            sel.Range.GoTo();
            dstDoc.ActiveWindow.ScrollIntoView(sel.Range); // 视角恢复

            return 0;
        }

        public void setTemplateList(Word.Application app, Word.ListLevels lstLvels, Word.ListLevels oListLvels)
        {
            if (oListLvels != null)
            {
                for (int i = 1; i <= lstLvels.Count; i++) // 遍历
                {
                    lstLvels[i].NumberFormat = oListLvels[i].NumberFormat;  // 赋值
                    lstLvels[i].TrailingCharacter = oListLvels[i].TrailingCharacter; // 赋值
                    lstLvels[i].NumberStyle = oListLvels[i].NumberStyle; // 赋值
                    lstLvels[i].NumberPosition = oListLvels[i].NumberPosition; // 赋值
                    lstLvels[i].Alignment = oListLvels[i].Alignment; // 赋值
                    lstLvels[i].TextPosition = oListLvels[i].TextPosition; // 赋值
                    lstLvels[i].TabPosition = oListLvels[i].TabPosition; // 赋值
                    lstLvels[i].ResetOnHigher = oListLvels[i].ResetOnHigher; // 赋值
                    lstLvels[i].StartAt = oListLvels[i].StartAt; // 赋值
                    lstLvels[i].LinkedStyle = oListLvels[i].LinkedStyle; // 赋值

                    copyFontStyle(oListLvels[i].Font, lstLvels[i].Font);  // 赋font值
                }
            }
            else
            {

                // Word.Document doc = app.ActiveDocument;
                //Word.Selection sel = doc.ActiveWindow.Selection;

                // 缺省值，见word的帮助多级列表，下同
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
            }

            return;
        }

        public void setTemplateList(Word.Application app, Word.ListLevels lstLvels, ClassListLevel[] oListLvels)
        {
            if (oListLvels != null)
            {
                for (int i = 1; i <= lstLvels.Count; i++) // 遍历
                {   // 设置
                    lstLvels[i].NumberFormat = oListLvels[i - 1].NumberFormat;
                    lstLvels[i].TrailingCharacter = oListLvels[i - 1].TrailingCharacter;
                    lstLvels[i].NumberStyle = oListLvels[i - 1].NumberStyle;
                    lstLvels[i].NumberPosition = oListLvels[i - 1].NumberPosition;
                    lstLvels[i].Alignment = oListLvels[i - 1].Alignment;
                    lstLvels[i].TextPosition = oListLvels[i - 1].TextPosition;
                    lstLvels[i].TabPosition = oListLvels[i - 1].TabPosition;
                    lstLvels[i].ResetOnHigher = oListLvels[i - 1].ResetOnHigher;
                    lstLvels[i].StartAt = oListLvels[i - 1].StartAt;
                    lstLvels[i].LinkedStyle = oListLvels[i - 1].LinkedStyle;

                    if (lstLvels[i].Font != null)
                    {
                        oListLvels[i - 1].Font.copy2(lstLvels[i].Font); // 复制字体
                    }
                }
            }
            else
            {
                // Word.Document doc = app.ActiveDocument;
                // Word.Selection sel = doc.ActiveWindow.Selection;
                // 参照word的设置进行配置，下同
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
            }

            return;
        }


        // 设置多级序号
        private void setTemplateList_v1(Word.Application app, ClassListLevel[] oListLvels)
        {
            // Word.Application app = m_addin.Application;
            // 自动编号 
            Word.ListGallery listGallery = null;
            Word.ListLevels lstLvels = null;
            Object objIndex = 1;

            if (bAppIsWps)
            {
                listGallery = app.ListGalleries[(Word.WdListGalleryType)4];

                if (listGallery.ListTemplates.Count == 0)
                {
                    Object objOutlineNumbered = true;
                    Word.ListTemplate lstTemplate = listGallery.ListTemplates.Add(objOutlineNumbered);
                    lstLvels = listGallery.ListTemplates[objIndex].ListLevels;
                }
                else
                {
                    lstLvels = listGallery.ListTemplates[objIndex].ListLevels;
                }

            }
            else
            {
                listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];
                lstLvels = listGallery.ListTemplates[objIndex].ListLevels;
            }


            if (oListLvels != null)
            {
                for (int i = 1; i <= lstLvels.Count; i++) // 遍历
                {   // 设置
                    lstLvels[i].NumberFormat = oListLvels[i - 1].NumberFormat;
                    lstLvels[i].TrailingCharacter = oListLvels[i - 1].TrailingCharacter;
                    lstLvels[i].NumberStyle = oListLvels[i - 1].NumberStyle;
                    lstLvels[i].NumberPosition = oListLvels[i - 1].NumberPosition;
                    lstLvels[i].Alignment = oListLvels[i - 1].Alignment;
                    lstLvels[i].TextPosition = oListLvels[i - 1].TextPosition;
                    lstLvels[i].TabPosition = oListLvels[i - 1].TabPosition;
                    lstLvels[i].ResetOnHigher = oListLvels[i - 1].ResetOnHigher;
                    lstLvels[i].StartAt = oListLvels[i - 1].StartAt;
                    lstLvels[i].LinkedStyle = oListLvels[i - 1].LinkedStyle;

                    if (lstLvels[i].Font != null)
                    {
                        oListLvels[i - 1].Font.copy2(lstLvels[i].Font); // 复制字体
                    }
                }
            }
            else
            {
                // Word.Document doc = app.ActiveDocument;
                // Word.Selection sel = doc.ActiveWindow.Selection;
                // 参照word的设置进行配置，下同
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
            }

            // listGallery.ListTemplates[objIndex].Name = "myList";
            return;
        }


        /// <summary>
        /// 设置当前模板List
        /// </summary>
        /// <param name="oListLvels"></param>
        private void setTemplateList_v1(Word.Application app, Word.ListLevels oListLvels)
        {
            // Word.Application app = m_addin.Application;
            // 自动编号 
            Word.ListGallery listGallery = app.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery];

            Object objIndex = 1;
            Word.ListLevels lstLvels = listGallery.ListTemplates[objIndex].ListLevels;


            if (oListLvels != null)
            {
                for (int i = 1; i <= lstLvels.Count; i++) // 遍历
                {
                    lstLvels[i].NumberFormat = oListLvels[i].NumberFormat;  // 赋值
                    lstLvels[i].TrailingCharacter = oListLvels[i].TrailingCharacter; // 赋值
                    lstLvels[i].NumberStyle = oListLvels[i].NumberStyle; // 赋值
                    lstLvels[i].NumberPosition = oListLvels[i].NumberPosition; // 赋值
                    lstLvels[i].Alignment = oListLvels[i].Alignment; // 赋值
                    lstLvels[i].TextPosition = oListLvels[i].TextPosition; // 赋值
                    lstLvels[i].TabPosition = oListLvels[i].TabPosition; // 赋值
                    lstLvels[i].ResetOnHigher = oListLvels[i].ResetOnHigher; // 赋值
                    lstLvels[i].StartAt = oListLvels[i].StartAt; // 赋值
                    lstLvels[i].LinkedStyle = oListLvels[i].LinkedStyle; // 赋值

                    copyFontStyle(oListLvels[i].Font, lstLvels[i].Font);  // 赋font值
                }
            }
            else
            {

                // Word.Document doc = app.ActiveDocument;
                //Word.Selection sel = doc.ActiveWindow.Selection;

                // 缺省值，见word的帮助多级列表，下同
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
            }

            // listGallery.ListTemplates[objIndex].Name = "myList";

        }


        public double DateDiff(DateTime DateTime1, DateTime DateTime2, int nType = 0)
        {
            double dbRet = 0;

            TimeSpan ts1 = new TimeSpan(DateTime1.Ticks);
            TimeSpan ts2 = new TimeSpan(DateTime2.Ticks);

            TimeSpan ts = ts1.Subtract(ts2);

            switch (nType)
            {
                case 1:
                    dbRet = ts.TotalHours;
                    break;

                case 2:
                    dbRet = ts.TotalMinutes;
                    break;

                case 3:
                    dbRet = ts.TotalSeconds;
                    break;

                default:
                case 0:
                    dbRet = ts.TotalDays;
                    break;
            }

            return dbRet;
        }


        public String CopyTmpDocFile(String strDoc)
        {
            String strTmpPath = Path.GetTempPath(); // 临时目录
            String strTmpFile = "", strPostx = "";

            if (!File.Exists(strDoc))
            {
                return null;
            }

            // 拼装临时文件名
            strPostx = DateTime.Now.ToString("yyyyMMdd_hhmmssffff");
            strTmpFile = strTmpPath + Path.GetFileNameWithoutExtension(strDoc) + "_" + strPostx + Path.GetExtension(strDoc);

            if (File.Exists(strTmpFile)) // 判断存在
            {
                try
                {
                    File.Delete(strTmpFile);
                }
                catch (System.Exception ex)
                {
                    return null;
                }
                finally
                {
                }
            }

            File.Copy(strDoc, strTmpFile);

            return strTmpFile;
        }



        public HashSet<int> excludeSpecificRng(HashSet<int> exSet, HashSet<int> selSet)
        {
            HashSet<int> tmpExSet = new HashSet<int>();
            HashSet<int> tmpSelSet = new HashSet<int>(selSet);

            int nPreIdx = -1;
            int nRngStart = -1;
            int nCnt = 0;

            foreach (int nIdx in exSet)
            {
                nCnt++;

                if (nRngStart == -1)
                {
                    nRngStart = nIdx;
                }

                if (nIdx - nPreIdx > 1) 
                {
                    if (nPreIdx > nRngStart)
                    {
                        for (int k = nRngStart + 1; k < nPreIdx; k++)
                        {
                            tmpExSet.Add(k);
                        }
                    }

                    nRngStart = nIdx;
                }
                else if(nCnt == exSet.Count)
                {
                    nPreIdx = nIdx;
                    if (nPreIdx > nRngStart)
                    {
                        for (int k = nRngStart + 1; k < nPreIdx; k++)
                        {
                            tmpExSet.Add(k);
                        }
                    }
                }

                nPreIdx = nIdx;
            }

            tmpSelSet.ExceptWith(tmpExSet);

            return tmpSelSet;
        }


        private HashSet<int> excludeSpecificRng_v2(ArrayList arrExclude, ArrayList arrSeledRngs, Word.WdColorIndex wdExcludeColor)
        {
            HashSet<int> tmpExSet = new HashSet<int>();
            HashSet<int> exSet = new HashSet<int>();
            HashSet<int> selSet = new HashSet<int>();

            int nStart = -1, nEnd = -1;

            foreach(Word.Range rng in arrExclude)
            {
                nStart = rng.Start;
                nEnd = rng.End;

                for (int k = nStart; k <= nEnd; k++)
                {
                    tmpExSet.Add(k);
                }
            }

            foreach(Word.Range rng in arrSeledRngs)
            {
                nStart = rng.Start;
                nEnd = rng.End;

                for (int k = nStart; k <= nEnd; k++)
                {
                    selSet.Add(k);
                }
            }


            int nPreIdx = -1;
            int nRngStart = -1;
            int nCnt = 0;

            foreach (int nIdx in tmpExSet)
            {
                nCnt++;

                if (nRngStart == -1)
                {
                    nRngStart = nIdx;
                }

                if (nIdx - nPreIdx > 1 || nCnt == tmpExSet.Count)
                {
                    if (nPreIdx > nRngStart)
                    {
                        for (int k = nRngStart + 1; k < nPreIdx; k++)
                        {
                            exSet.Add(k);
                        }
                    }

                    nRngStart = nIdx;
                }

                nPreIdx = nIdx;
            }

            selSet.ExceptWith(exSet);

            return selSet;
        }



        private ArrayList excludeSpecificRng_v1(ArrayList arrExclude, ArrayList arrSeledRngs,Word.WdColorIndex wdExcludeColor)
        {
            ArrayList arrExcludedRngs = new ArrayList();
            Boolean bOverlap = false;

            foreach (Word.Range selRng in arrSeledRngs)
            {
                bOverlap = false;

                foreach (Word.Range exRng in arrExclude)
                {
                    if (exRng.Start <= selRng.Start && exRng.End >= selRng.End)
                    {
                        // bRet = true;
                        // completely exclude current Para
                        // 
                        bOverlap = true;
                    }
                    else if (selRng.Start <= exRng.Start && selRng.End >= exRng.End)
                    {
                        // bRet = true;
                        // 
                        foreach (Word.Range chRng in selRng.Characters)
                        {
                            if (chRng.Start <= exRng.Start && chRng.End <= exRng.Start)
                            {
                                if (chRng.HighlightColorIndex != wdExcludeColor)
                                {
                                    arrExcludedRngs.Add(chRng);
                                    bOverlap = true;
                                }
                            }
                            else if (chRng.Start >= exRng.End && chRng.End >= exRng.End)
                            {
                                if (chRng.HighlightColorIndex != wdExcludeColor)
                                {
                                    arrExcludedRngs.Add(chRng);
                                    bOverlap = true;
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                    else if (exRng.Start <= selRng.Start && exRng.End > selRng.Start && exRng.End <= selRng.End)
                    {
                        // bRet = true;
                        foreach (Word.Range chRng in selRng.Characters)
                        {
                            if (chRng.Start >= exRng.End && chRng.End > exRng.End)
                            {
                                if (chRng.HighlightColorIndex != wdExcludeColor)
                                {
                                    arrExcludedRngs.Add(chRng);
                                    bOverlap = true;
                                }
                            }
                        }
                    }
                    else if (selRng.Start <= exRng.Start && selRng.End > exRng.Start && selRng.End <= exRng.End)
                    {
                        // bRet = true;
                        foreach (Word.Range chRng in selRng.Characters)
                        {
                            if (chRng.Start <= exRng.Start && chRng.End <= exRng.Start)
                            {
                                if (chRng.HighlightColorIndex != wdExcludeColor)
                                {
                                    arrExcludedRngs.Add(chRng);
                                    bOverlap = true;
                                }
                            }
                        }
                    }
                }

                if (!bOverlap)
                {
                    arrExcludedRngs.Add(selRng);
                }
            }

            return arrExcludedRngs;
        }



        private int getNavListParagraphs(Word.Document doc, ref Word.Paragraph firstListPara, ref Word.Paragraph lastListPara,
                         ref Word.Paragraph nearstPrevListPara, ref Word.Paragraph nearstNextListPara)
        {
            Word.Selection sel = doc.ActiveWindow.Selection;

            int nMinStart = -1, nMaxStart = -1;
            int nPrevDis = -1, nNextDis = -1;
            int nCurDisPrev = 0, nCurDisNext = 0;

            ArrayList arrParas = new ArrayList();

            // doc.ListParagraphs.AsQueryable
            foreach (Word.Paragraph para in doc.ListParagraphs)
            {
                arrParas.Add(para);
            }

            if (arrParas.Count > 1)
            {
                ClassParagraphComparer cmp = new ClassParagraphComparer();

                arrParas.Sort(cmp);

                firstListPara = (Word.Paragraph)arrParas[0];
                lastListPara = (Word.Paragraph)arrParas[arrParas.Count - 1];
            }

            


            //foreach(Word.Paragraph para in doc.ListParagraphs)
            //{
            //    if (nMinStart == -1 || para.Range.Start < nMinStart)
            //    {
            //        nMinStart = para.Range.Start;
            //        firstListPara = para;
            //    }

            //    if (nMaxStart == -1 || para.Range.Start > nMaxStart)
            //    {
            //        nMaxStart = para.Range.Start;
            //        lastListPara = para;
            //    }

            //    nCurDisPrev = sel.Start - para.Range.Start;
            //    if (nCurDisPrev > 0)
            //    {
            //        if (nPrevDis == -1 || nCurDisPrev < nPrevDis)
            //        {
            //            nPrevDis = nCurDisPrev;
            //            nearstPrevListPara = para;
            //        }
            //    }

            //    nCurDisNext = para.Range.Start - sel.Start;
            //    if (nCurDisNext > 0)
            //    {
            //        if (nNextDis == -1 || nCurDisNext < nNextDis)
            //        {
            //            nNextDis = nCurDisNext;
            //            nearstNextListPara = para;
            //        }
            //    }
            //}

            return 0;
        }


        public Boolean isPageCenterHideTbl(Word.Table tbl)
        {
            if (tbl == null)
            {
                return false;
            }

            if (tbl.Rows.Count != 1)
            {
                return false;
            }

            if (tbl.Columns.Count != 1)
            {
                return false;
            }


            if (tbl.Borders.InsideLineStyle != Word.WdLineStyle.wdLineStyleNone)
            {
                return false;
            }

            if ((int)tbl.Borders.InsideLineWidth != 0)
            {
                return false;
            }


            if (tbl.Borders.OutsideLineStyle != Word.WdLineStyle.wdLineStyleNone)
            {
                return false;
            }

            if ((int)tbl.Borders.OutsideLineWidth != 0)
            {
                return false;
            }
            
            if (tbl.Rows.HeightRule != Word.WdRowHeightRule.wdRowHeightAtLeast)
            {
                return false;
            }

            if(tbl.Rows.Height == (float)(Word.WdConstants.wdUndefined))
            {
                return false;
            }

            int nPages = tbl.Range.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            if (nPages != 2)
            {
                return false;
            }

            Word.PageSetup pgsetup = tbl.Range.PageSetup;
            float fHeight = (pgsetup.PageHeight - pgsetup.TopMargin - pgsetup.BottomMargin);// in points, / 28.34f;

            float fRatio = (float)(tbl.Rows.Height / fHeight);

            if( fRatio > 0.95f && fRatio <= 1.00f)
            {
                return false;
            }

            //Word.Cell cel = tbl.Cell(1, 1);

            //if(cel.VerticalAlignment != Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter)
            //{
            //    return false;
            //}
            
            //if(cel.Range.ParagraphFormat.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter)
            //{
            //    return false;
            //}

            return true;

            /*
            int nPages = tbl.Range.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

            Word.PageSetup pgsetup = tbl.Range.PageSetup;

            float fHeight = (pgsetup.PageHeight - pgsetup.TopMargin - pgsetup.BottomMargin) ;// in points, / 28.34f;

            Boolean b1Cell = (tbl.Rows.Count == 1 && tbl.Columns.Count == 1);

            Boolean b1Page = (nPages == 1);

            Boolean bInsideColor = (tbl.Borders.InsideLineStyle == Word.WdLineStyle.wdLineStyleNone &&
                                    (int)tbl.Borders.InsideLineWidth == 0);
            Boolean bOutsideColor = (tbl.Borders.OutsideLineStyle == Word.WdLineStyle.wdLineStyleNone &&
                                    (int)tbl.Borders.OutsideLineWidth == 0);

            Boolean bSize = (tbl.Rows.HeightRule != Word.WdRowHeightRule.wdRowHeightAuto && (tbl.Rows.Height / fHeight) > 0.95);
            // ?? 
            Boolean bCellAlignment = false;

            if (b1Cell)
            {
                Word.Cell cel = tbl.Cell(1, 1);

                bCellAlignment = (cel.VerticalAlignment == Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter &&
                                    cel.Range.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter);
            }
            // ??

            return true;
             *
             */
        }


        public Boolean isHideTbl(Word.Table tbl)
        {
            if (tbl == null)
            {
                return false;
            }

            if (tbl.Borders.InsideLineStyle != Word.WdLineStyle.wdLineStyleNone)
            {
                return false;
            }

            if ((int)tbl.Borders.InsideLineWidth != 0)
            {
                return false;
            }

            if (tbl.Borders.OutsideLineStyle != Word.WdLineStyle.wdLineStyleNone)
            {
                return false;
            }

            if ((int)tbl.Borders.OutsideLineWidth != 0)
            {
                return false;
            }

            return true;
        }


        public ArrayList getTiZhuParasInScope(Word.Document doc,Word.Range scopeRange = null)
        {
            ArrayList arrs = new ArrayList();

            Word.Selection sel = doc.ActiveWindow.Selection;

            int nPosStart = sel.Start;
            int nPosEnd = sel.End;


            if (scopeRange != null)
            {
                if (scopeRange.Fields.Count == 0)
                {
                    return arrs;
                }

                sel.Start = scopeRange.Start;
                sel.End = sel.Start;
                sel.Range.GoTo();
                doc.ActiveWindow.ScrollIntoView(sel.Range);
            }
            else
            {
                if (doc.Fields.Count == 0)
                {
                    return arrs;
                }

                sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            }

            Word.Range curRng = null, prevRng = null;
            Word.Field fld = null;
            Boolean bInToc = false;

            while (true)
            {
                curRng = sel.GoTo(Word.WdGoToItem.wdGoToField, Word.WdGoToDirection.wdGoToNext, 1, "");

                if(scopeRange != null && !RangeOverlap(curRng.Paragraphs[1].Range,scopeRange))
                {
                    break;
                }

                if (!bInToc)
                {
                    foreach (Word.TableOfContents cnts in doc.TablesOfContents)
                    {
                        if (curRng.InRange(cnts.Range))
                        {
                            sel.Start = cnts.Range.End;
                            sel.End = sel.Start;
                            sel.Range.GoTo();

                            bInToc = true;

                            // curRng = sel.GoTo(Word.WdGoToItem.wdGoToField, Word.WdGoToDirection.wdGoToNext, 1, "");

                            break;
                        }
                    }

                    if (bInToc)
                    {
                        continue;
                    }
                }


                fld = null;

                if (curRng != null && curRng.Paragraphs[1].Range.Fields.Count > 0)
                {
                    if (prevRng != null && curRng != null && (prevRng.Start == curRng.Start && prevRng.End == curRng.End))
                    {
                        break; // no found any more
                    }
                    else
                    {
                        // next
                        prevRng = curRng;
                    }

                    foreach (Word.Field tmpfld in curRng.Paragraphs[1].Range.Fields)
                    {
                        if ((curRng.Start + 1) == tmpfld.Code.Start)
                        {
                            fld = tmpfld;
                        }
                    }

                    if (fld != null && fld.Type == Word.WdFieldType.wdFieldSequence) // fld.Code.Text.StartsWith("SEQ")
                    {
                        arrs.Add(fld.Result.Paragraphs[1]); // found
                    }
                }
                else
                {
                    if (prevRng == null && curRng != null)
                    {
                        break;
                    }
                }
            }


            sel.Start = nPosStart;
            sel.End = nPosEnd;
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return arrs;
        }


        public ArrayList mergeRange(ArrayList rngArr)
        {
            int nPreIdx = -1;
            int nRngStart = -1;
            int nCnt = 0;

            ArrayList edRngArr = new ArrayList();

            foreach (int nIdx in rngArr)
            {
                nCnt++;

                if (nRngStart == -1)
                {
                    nRngStart = nIdx;
                }

                if (nIdx - nPreIdx > 1)
                {
                    if (nPreIdx > nRngStart)
                    {
                        edRngArr.Add(nRngStart);
                        edRngArr.Add(nPreIdx);
                    }

                    nRngStart = nIdx;
                }
                else if (nCnt == rngArr.Count)
                {
                    nPreIdx = nIdx;
                    if (nPreIdx > nRngStart)
                    {
                        edRngArr.Add(nRngStart);
                        edRngArr.Add(nPreIdx);
                    }
                }

                nPreIdx = nIdx;

            } // foreach

            return edRngArr;
        }


        public int clearTiZhu(Word.Document doc, Word.Range rng = null, Boolean bRemoveOnlyBody = true)
        {
            Word.Selection sel = doc.Application.ActiveWindow.Selection;

            int nOStart = sel.Start;
            int nOEnd = sel.End;

            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;

            doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;

            if ((rng == null && doc.Fields.Count == 0) || (rng != null && rng.Fields.Count == 0))
            {
                return -1;
            }

            ArrayList arrTiZhus = getTiZhuParasInScope(doc, rng);

            int nCnt = 0, nSelCnt = 0, nTotalCnt = arrTiZhus.Count;
            int nRngStart = -1, nRngEnd = -1;

            foreach (Word.Paragraph para in arrTiZhus)
            {
                nCnt++;

                if (bRemoveOnlyBody)
                {
                    nRngStart = para.Range.Start;
                    nRngEnd = para.Range.End;

                    if (nRngEnd > nRngStart)
                    {
                        nRngEnd--;
                    }

                    sel.SetRange(nRngStart, nRngEnd);
                    RecordMultiSel(sel.Range);
                }
                else
                {
                    RecordMultiSel(para.Range);
                }

                if (nSelCnt == 50 || nCnt == nTotalCnt)
                {
                    ExecMultiSel(doc);
                    sel.Delete();

                    nSelCnt = 0;
                }
                else
                {
                    nSelCnt++;
                }
            }
              
            doc.ActiveWindow.View.Type = oViewType;

            // restore original position
            sel.Start = nOStart;
            sel.End = nOEnd;
            //// sel.Range.Select();
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return nTotalCnt;
        }


        public Boolean isTiZhu(Word.Range rng)
        {
            foreach(Word.Field fld in rng.Fields)
            {
                if (fld.Type == Word.WdFieldType.wdFieldSequence)
                {
                    return true;
                }
            }

            return false;
        }


        public Hashtable getHeadingScope(Word.Document doc)
        {
            // Word.Selection sel = doc.ActiveWindow.Selection;

            ArrayList arrHeadings = getHeadingParas(doc);

            // HashSet<int> setHeading = new HashSet<int>();
            Hashtable hashScopeHeading = new Hashtable();

            int nDocRngStart = doc.Content.Start;
            int nDocRngEnd = doc.Content.End;

            String strHeading = "";

            Word.WdOutlineLevel wdPrevParaLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
            int nRngStart = -1, nRngEnd = -1, nPreParaStart = -1,nPreParaEnd = -1;
            int nCount = arrHeadings.Count;

            if (nCount > 0)
            {
                TiZhuHeadingItem headingItem = null;

                Word.Paragraph firstPara = (Word.Paragraph)arrHeadings[0];
                Word.Paragraph lastPara = (Word.Paragraph)arrHeadings[nCount - 1];
                Word.Paragraph para = null;

                nPreParaStart = firstPara.Range.Start;
                nPreParaEnd = firstPara.Range.End;
                strHeading = firstPara.Range.Text.Trim();
                wdPrevParaLevel = firstPara.OutlineLevel;

                for (int i = 1; i < nCount; i++)
                {
                    para = (Word.Paragraph)arrHeadings[i];

                    nRngStart = nPreParaEnd;
                    nRngEnd = para.Range.Start;

                    //if (nRngStart < nDocRngEnd)
                    //{
                    //    nRngStart++;
                    //}

                    //if (nRngEnd > 1)
                    //{
                    //    nRngEnd--;
                    //}
                    headingItem = new TiZhuHeadingItem();

                    headingItem.strHeadingName = strHeading;
                    headingItem.nRngStart = nPreParaStart;
                    headingItem.nRngEnd = nPreParaEnd;
                    headingItem.nCoverRngStart = nRngStart;
                    headingItem.nCoverRngEnd = nRngEnd;
                    headingItem.wdLevel = wdPrevParaLevel;
                    headingItem.nTblCnt = 0;
                    headingItem.nInShpCnt = 0;

                    for (int k = nRngStart; k <= nRngEnd; k++)
                    {
                        // setHeading.Add(k);
                        hashScopeHeading.Add(k, headingItem);
                    }

                    nPreParaStart = para.Range.Start;
                    nPreParaEnd = para.Range.End;
                    strHeading = para.Range.Text.Trim();
                    wdPrevParaLevel = para.OutlineLevel;
                }

                // last heading scope
                para = lastPara;

                nRngStart = nPreParaEnd;
                nRngEnd = nDocRngEnd;

                nPreParaStart = para.Range.Start;
                nPreParaEnd = para.Range.End;
                strHeading = para.Range.Text.Trim();
                wdPrevParaLevel = para.OutlineLevel;

                //if (nRngStart < nDocRngEnd)
                //{
                //    nRngStart++;
                //}

                //if (nRngEnd > 1)
                //{
                //    nRngEnd--;
                //}

                headingItem = new TiZhuHeadingItem();

                headingItem.strHeadingName = strHeading;
                headingItem.nRngStart = para.Range.Start;
                headingItem.nRngEnd = para.Range.End;
                headingItem.nCoverRngStart = nRngStart;
                headingItem.nCoverRngEnd = nRngEnd;
                headingItem.wdLevel = para.OutlineLevel;
                headingItem.nTblCnt = 0;
                headingItem.nInShpCnt = 0;

                for (int k = nRngStart; k <= nRngEnd; k++)
                {
                    // setHeading.Add(k);
                    hashScopeHeading.Add(k, headingItem);
                }

            }

            return hashScopeHeading;
        }


        public void bulkAddTiZhus(Word.Application app,Word.Document doc,
                                  ArrayList arrsTbls, ArrayList arrsInShpIsolateNotInTbl,
                                  Word.CaptionLabel TblCapLbl,
                                  Word.WdCaptionPosition wdTblPos,
                                  int nTblAlignment,
                                  String strTxtTblCapLblPreFix,
                                  Boolean bTblCaplblGetFromHeading,
                                  String strTxtTblCapLblPostFix,
                                  Word.CaptionLabel InShpCapLbl, 
                                  Word.WdCaptionPosition wdInShpPos,
                                  int nInShpAlignment,
                                  String strTxtInShpCapLblPreFix,
                                  Boolean bInShpCaplblGetFromHeading,
                                  String strTxtInShpCapLblPostFix,
                                  Hashtable hashHeadings, Hashtable hashSimpNum,
                                  Boolean bTblNeedSn, Boolean bInShpNeedSn,
                                  ClassFont tblFnt = null, ClassParagraphFormat tblParaFmt = null,
                                  ClassFont inShpFnt = null, ClassParagraphFormat inShpParaFmt = null
                                  )
        {
            Word.Selection sel = doc.ActiveWindow.Selection;

            Boolean bPagination = app.Options.Pagination;
            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
            int nOStart = sel.Start;
            int nOEnd = sel.End;


            app.Options.Pagination = false;
            doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;


            // insert a cap label
            // copy field
            // insert a FAKE cap label in every tables and pictures
            // find specail FAKE cap label and paste
            // could it be multi selection?

            Word.Style tizhuStyle = null;

            try
            {
                tizhuStyle = doc.Styles["题注"];
            }
            catch (System.Exception ex)
            {
                tizhuStyle = null;
            }

            Word.Paragraph para = null;
            int nSelCnt = 0;

            int nTblCnt = 0, nInShpCnt = 0;
            Word.Field tblStyleRefField = null, tblSeqField = null;
            Word.Range fldRng = null;
            String strTblHyphen = "", strHeadingTxt = "";

            const String strTblStyleRefPlaceHolder = "#XTBLSTYREFX#";
            const String strTblSeqRefPlaceHolder = "#XTBLREPX#";
            String strTblFieldPlaceHolder = "";

            TiZhuHeadingItem curHeadingItem = null;
            TiZhuHeadingItem PrevHeadingItem = null;
            String strHeadingSn = "";
            int nHeadingCnt = 0;


            if (TblCapLbl != null)
            {
                PrevHeadingItem = null;

                foreach (Word.Table tblItem in arrsTbls)
                {
                    nTblCnt++;

                    if (bTblCaplblGetFromHeading && hashHeadings.Count > 0)
                    {
                        if (hashHeadings.Contains(("表" + nTblCnt)))
                        {
                            curHeadingItem = (TiZhuHeadingItem)hashHeadings[("表" + nTblCnt)];
                            if (curHeadingItem != null)
                            {
                                strHeadingTxt = curHeadingItem.strHeadingName;
                            }
                            else
                            {
                                strHeadingTxt = "";
                            }

                            strHeadingSn = "";

                            if (bTblNeedSn && curHeadingItem != null)
                            {
                                if (curHeadingItem.nTblCnt < 2)
                                {
                                    PrevHeadingItem = curHeadingItem;
                                    nHeadingCnt = 0;
                                }
                                else
                                {
                                    if (PrevHeadingItem != null)
                                    {
                                        if (curHeadingItem.nRngStart == PrevHeadingItem.nRngStart &&
                                            curHeadingItem.nRngEnd == PrevHeadingItem.nRngEnd)
                                        {
                                            nHeadingCnt++;
                                        }
                                        else
                                        {
                                            PrevHeadingItem = curHeadingItem;
                                            nHeadingCnt = 1;
                                        }
                                    }
                                    else
                                    {
                                        PrevHeadingItem = curHeadingItem;
                                        nHeadingCnt = 1;
                                    }

                                    strHeadingSn = "";
                                    if (hashSimpNum.Contains(nHeadingCnt))
                                    {
                                        strHeadingSn = (String)hashSimpNum[nHeadingCnt];
                                    }
                                }
                            }
                        }
                    }

                    doc.ActiveWindow.ScrollIntoView(tblItem.Range);
                    tblItem.Range.GoTo();

                    if (nTblCnt == 1)
                    {
                        tblItem.Range.Select();

                        if (bAppIsWps)
                        {
                            sel.InsertCaption(TblCapLbl.Name, "", "", wdTblPos, false);
                        }
                        else
                        {
                            sel.InsertCaption(TblCapLbl, "", "", wdTblPos, false);
                        }

                        //foreach (Word.Field fld in sel.Paragraphs[1].Range.Fields)
                        //{
                        //    if (fld.Type == Word.WdFieldType.wdFieldStyleRef)
                        //    {
                        //        if (tblStyleRefField == null)
                        //        {
                        //            tblStyleRefField = fld;
                        //        }
                        //    }

                        //    if (fld.Type == Word.WdFieldType.wdFieldSequence)
                        //    {
                        //        if (tblSeqField == null)
                        //        {
                        //            tblSeqField = fld;
                        //        }
                        //    }
                        //}

                        if (strTxtTblCapLblPreFix.Length > 0)
                        {
                            sel.TypeText(strTxtTblCapLblPreFix);
                        }

                        if (bTblCaplblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                        {
                            sel.TypeText(strHeadingTxt);
                        }

                        if (strTxtTblCapLblPostFix.Length > 0)
                        {
                            sel.TypeText(strTxtTblCapLblPostFix);
                        }

                        if (bTblCaplblGetFromHeading && bTblNeedSn &&
                            !String.IsNullOrWhiteSpace(strHeadingSn))
                        {
                            sel.TypeText(strHeadingSn);
                        }

                        foreach (Word.Field fld in sel.Paragraphs[1].Range.Fields)
                        {
                            if (fld.Type == Word.WdFieldType.wdFieldStyleRef)
                            {
                                if (tblStyleRefField == null)
                                {
                                    tblStyleRefField = fld;
                                }
                            }

                            if (fld.Type == Word.WdFieldType.wdFieldSequence)
                            {
                                if (tblSeqField == null)
                                {
                                    tblSeqField = fld;
                                }
                            }
                        }

                        strTblFieldPlaceHolder = TblCapLbl.Name + " " + strTblSeqRefPlaceHolder;

                        if (tblStyleRefField != null && tblSeqField != null)
                        {
                            fldRng = doc.Range(tblStyleRefField.Result.End, tblSeqField.Result.Start);

                            strTblHyphen = fldRng.Text;

                            strTblFieldPlaceHolder = TblCapLbl.Name + " " + strTblStyleRefPlaceHolder + strTblHyphen + strTblSeqRefPlaceHolder;
                        }

                    }
                    else
                    {
                        if (wdTblPos == Word.WdCaptionPosition.wdCaptionPositionBelow)
                        {
                            Word.Paragraph nextPara = null;

                            nextPara = tblItem.Range.Paragraphs.Last.Next();
                            
                            if (nextPara == null || !String.IsNullOrWhiteSpace(nextPara.Range.Text.Trim(m_trimChars)) ||
                                nextPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                int nPos = tblItem.Range.Paragraphs.Last.Range.End;
                                sel.SetRange(nPos, nPos);

                                sel.InsertParagraphAfter();
                                sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                            }
                            else
                            {
                                sel.SetRange(nextPara.Range.Start, nextPara.Range.End -1);
                            }

                            sel.TypeText(strTblFieldPlaceHolder);

                            if (strTxtTblCapLblPreFix.Length > 0)
                            {
                                sel.TypeText(strTxtTblCapLblPreFix);
                            }

                            if (bTblCaplblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                            {
                                sel.TypeText(strHeadingTxt);
                            }

                            if (strTxtTblCapLblPostFix.Length > 0)
                            {
                                sel.TypeText(strTxtTblCapLblPostFix);
                            }

                            if (bTblCaplblGetFromHeading && bTblNeedSn && !String.IsNullOrWhiteSpace(strHeadingSn))
                            {
                                sel.TypeText(strHeadingSn);
                            }

                            if (tizhuStyle != null)
                            {
                                sel.Paragraphs[1].set_Style(tizhuStyle);
                            }
                        }
                        else // above
                        {
                            Word.Paragraph PrevPara = tblItem.Range.Paragraphs.First.Previous();

                            // MessageBox.Show(PrevPara.Range.Text);

                            if (/*PrevPara == null || */!String.IsNullOrWhiteSpace(PrevPara.Range.Text.Trim(m_trimChars)) ||
                                PrevPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                // int nPos = tblItem.Range.Paragraphs.First.Range.Start;
                                int nPos = PrevPara.Range.End - 1;
                                sel.SetRange(nPos, nPos);

                                sel.InsertParagraphBefore();
                                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            }
                            else
                            {
                                sel.SetRange(PrevPara.Range.Start, PrevPara.Range.End -1);
                            }

                            sel.TypeText(strTblFieldPlaceHolder);

                            if (strTxtTblCapLblPreFix.Length > 0)
                            {
                                sel.TypeText(strTxtTblCapLblPreFix);
                            }

                            if (bTblCaplblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                            {
                                sel.TypeText(strHeadingTxt);
                            }

                            if (strTxtTblCapLblPostFix.Length > 0)
                            {
                                sel.TypeText(strTxtTblCapLblPostFix);
                            }

                            if (bTblCaplblGetFromHeading && bTblNeedSn && !String.IsNullOrWhiteSpace(strHeadingSn))
                            {
                                sel.TypeText(strHeadingSn);
                            }

                            // sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                            if (tizhuStyle != null)
                            {
                                sel.Paragraphs[1].set_Style(tizhuStyle);
                            }

                            sel.Paragraphs[1].Format.KeepWithNext = -1; //true
                        }
                    }

                    sel.Paragraphs[1].Alignment = (Word.WdParagraphAlignment)nTblAlignment;
                    
                    if (tblFnt != null)
                    {
                        tblFnt.SelCopy2(sel.Paragraphs[1].Range.Font);
                    }

                    if (tblParaFmt != null)
                    {
                        tblParaFmt.SelCopy2(sel.Paragraphs[1].Range.ParagraphFormat);
                    }

                    //switch (nTblAlignment)
                    //{
                    //    case 1:
                    //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    //        break;

                    //    case 2:
                    //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //        break;

                    //    default:
                    //    case 0:
                    //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //        break;
                    //}
                }

                //if (tblSeqField != null)
                //{
                //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);

                //    sel.Find.ClearFormatting();

                //    sel.Find.Text = strTblSeqRefPlaceHolder;
                //    sel.Find.Replacement.Text = "";
                //    sel.Find.Forward = true;
                //    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                //    sel.Find.Format = false;
                //    sel.Find.MatchCase = false;
                //    sel.Find.MatchWholeWord = false;
                //    sel.Find.MatchByte = false;
                //    sel.Find.MatchWildcards = false;
                //    sel.Find.MatchSoundsLike = false;
                //    sel.Find.MatchAllWordForms = false;

                //    sel.Find.Execute();

                //    tblSeqField.Copy();

                //    while (sel.Find.Found)
                //    {
                //        para = sel.Paragraphs[1];

                //        if (bAppIsWps)
                //        {
                //            sel.Paste();
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                //        }
                //        else
                //        {
                //            RecordMultiSel(sel.Range);
                //            nSelCnt++;
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //            if (nSelCnt == 50)
                //            {
                //                ExecMultiSel(doc);
                //                sel.Paste();
                //                nSelCnt = 0;
                //            }
                //        }

                //        if (para.Next() == null)
                //        {
                //            if (bAppIsWps)
                //            {

                //            }
                //            else
                //            {
                //                if (nSelCnt > 0)
                //                {
                //                    ExecMultiSel(doc);
                //                    sel.Paste();
                //                    nSelCnt = 0;
                //                }
                //            }

                //            break;
                //        }

                //        sel.Find.Execute();
                //    }

                //    if (bAppIsWps)
                //    {

                //    }
                //    else
                //    {
                //        if (nSelCnt > 0)
                //        {
                //            ExecMultiSel(doc);
                //            sel.Paste();
                //        }
                //    }
                //}

                //if (tblStyleRefField != null)
                //{
                //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);

                //    sel.Find.ClearFormatting();

                //    sel.Find.Text = strTblStyleRefPlaceHolder;
                //    sel.Find.Replacement.Text = "";
                //    sel.Find.Forward = true;
                //    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                //    sel.Find.Format = false;
                //    sel.Find.MatchCase = false;
                //    sel.Find.MatchWholeWord = false;
                //    sel.Find.MatchByte = false;
                //    sel.Find.MatchWildcards = false;
                //    sel.Find.MatchSoundsLike = false;
                //    sel.Find.MatchAllWordForms = false;

                //    sel.Find.Execute();


                //    tblStyleRefField.Copy();

                //    nSelCnt = 0;

                //    while (sel.Find.Found)
                //    {
                //        para = sel.Paragraphs[1];

                //        if (bAppIsWps)
                //        {
                //            sel.Paste();
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                //        }
                //        else
                //        {
                //            RecordMultiSel(sel.Range);
                //            nSelCnt++;
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //            if (nSelCnt == 50)
                //            {
                //                ExecMultiSel(doc);
                //                sel.Paste();
                //                nSelCnt = 0;
                //            }
                //        }


                //        if (para.Next() == null)
                //        {
                //            if (bAppIsWps)
                //            {

                //            }
                //            else
                //            {
                //                if (nSelCnt > 0)
                //                {
                //                    ExecMultiSel(doc);
                //                    sel.Paste();
                //                    nSelCnt = 0;
                //                }
                //            }

                //            break;
                //        }

                //        sel.Find.Execute();
                //    }

                //    if (bAppIsWps)
                //    {

                //    }
                //    else
                //    {
                //        if (nSelCnt > 0)
                //        {
                //            ExecMultiSel(doc);
                //            sel.Paste();
                //        }
                //    }
                //}
            }


            ///////////////////// inline shape

            Word.Field inshpStyleRefField = null, inshpSeqField = null;
            Boolean bIgnore = false;
            String strInShpHyphen = "";

            const String strInShpStyleRefPlaceHolder = "#XTUSTYREFX#";
            const String strInShpSeqRefPlaceHolder = "#XTUREPX#";
            String strInShpFieldPlaceHolder = "";

            if (InShpCapLbl != null)
            {
                PrevHeadingItem = null;

                nInShpCnt = 0;
                int nIdx = 0;

                foreach (Word.Paragraph inShpPara in arrsInShpIsolateNotInTbl)
                {
                    doc.ActiveWindow.ScrollIntoView(inShpPara.Range);
                    inShpPara.Range.GoTo();

                    nIdx++;

                    bIgnore = true;
                    foreach (Word.InlineShape inShp in inShpPara.Range.InlineShapes)
                    {
                        if (!(inShp.Type == Word.WdInlineShapeType.wdInlineShapePictureBullet ||
                            (inShp.OLEFormat != null && inShp.OLEFormat.DisplayAsIcon)))
                        {
                            bIgnore = false;
                            break;
                        }
                    }

                    if (bIgnore)
                    {
                        continue;
                    }

                    nInShpCnt++;

                    if (bInShpCaplblGetFromHeading && hashHeadings.Count > 0)
                    {
                        if (hashHeadings.Contains(("图" + nIdx)))
                        {
                            curHeadingItem = (TiZhuHeadingItem)hashHeadings[("图" + nIdx)];

                            if (curHeadingItem != null)
                            {
                                strHeadingTxt = curHeadingItem.strHeadingName;
                            }
                            else
                            {
                                strHeadingTxt = "";
                            }

                            strHeadingSn = "";

                            if (bInShpNeedSn && curHeadingItem != null)
                            {
                                if (curHeadingItem.nInShpCnt < 2)
                                {
                                    PrevHeadingItem = curHeadingItem;
                                    nHeadingCnt = 0;
                                }
                                else
                                {
                                    if (PrevHeadingItem != null)
                                    {
                                        if (curHeadingItem.nRngStart == PrevHeadingItem.nRngStart &&
                                            curHeadingItem.nRngEnd == PrevHeadingItem.nRngEnd)
                                        {
                                            nHeadingCnt++;
                                        }
                                        else
                                        {
                                            PrevHeadingItem = curHeadingItem;
                                            nHeadingCnt = 1;
                                        }
                                    }
                                    else
                                    {
                                        PrevHeadingItem = curHeadingItem;
                                        nHeadingCnt = 1;
                                    }

                                    strHeadingSn = "";
                                    if (hashSimpNum.Contains(nHeadingCnt))
                                    {
                                        strHeadingSn = (String)hashSimpNum[nHeadingCnt];
                                    }
                                }
                            }
                        }
                    }


                    if (nInShpCnt == 1)
                    {
                        inShpPara.Range.Select();

                        if (bAppIsWps)
                        {
                            sel.InsertCaption(InShpCapLbl.Name, "", "", wdInShpPos, false);
                        }
                        else
                        {
                            sel.InsertCaption(InShpCapLbl, "", "", wdInShpPos, false);
                        }

                        //foreach (Word.Field fld in sel.Paragraphs[1].Range.Fields)
                        //{
                        //    if (fld.Type == Word.WdFieldType.wdFieldStyleRef)
                        //    {
                        //        if (inshpStyleRefField == null)
                        //        {
                        //            inshpStyleRefField = fld;
                        //        }
                        //    }

                        //    if (fld.Type == Word.WdFieldType.wdFieldSequence)
                        //    {
                        //        if (inshpSeqField == null)
                        //        {
                        //            inshpSeqField = fld;
                        //        }
                        //    }
                        //}

                        if (strTxtInShpCapLblPreFix.Length > 0)
                        {
                            sel.TypeText(strTxtInShpCapLblPreFix);
                        }

                        if (bInShpCaplblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                        {
                            sel.TypeText(strHeadingTxt);
                        }

                        if (strTxtInShpCapLblPostFix.Length > 0)
                        {
                            sel.TypeText(strTxtInShpCapLblPostFix);
                        }

                        if (bInShpCaplblGetFromHeading && bInShpNeedSn && !String.IsNullOrWhiteSpace(strHeadingSn))
                        {
                            sel.TypeText(strHeadingSn);
                        }

                        sel.Paragraphs[1].Alignment = (Word.WdParagraphAlignment)nInShpAlignment;

                        if (inShpFnt != null)
                        {
                            inShpFnt.SelCopy2(sel.Paragraphs[1].Range.Font);
                        }

                        if (inShpParaFmt != null)
                        {
                            inShpParaFmt.SelCopy2(sel.Paragraphs[1].Range.ParagraphFormat);
                        }

                        //switch (nInShpAlignment)
                        //{
                        //    case 1:
                        //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        //        break;

                        //    case 2:
                        //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        //        break;

                        //    default:
                        //    case 0:
                        //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        //        break;
                        //}

                        foreach (Word.Field fld in sel.Paragraphs[1].Range.Fields)
                        {
                            if (fld.Type == Word.WdFieldType.wdFieldStyleRef)
                            {
                                if (inshpStyleRefField == null)
                                {
                                    inshpStyleRefField = fld;
                                }
                            }

                            if (fld.Type == Word.WdFieldType.wdFieldSequence)
                            {
                                if (inshpSeqField == null)
                                {
                                    inshpSeqField = fld;
                                }
                            }
                        }

                        strInShpFieldPlaceHolder = InShpCapLbl.Name + " " + strInShpSeqRefPlaceHolder;

                        if (inshpStyleRefField != null && inshpSeqField != null)
                        {
                            fldRng = doc.Range(inshpStyleRefField.Result.End, inshpSeqField.Result.Start);

                            strInShpHyphen = fldRng.Text;

                            strInShpFieldPlaceHolder = InShpCapLbl.Name + " " + strInShpStyleRefPlaceHolder + strInShpHyphen + strInShpSeqRefPlaceHolder;
                        }

                    }
                    else
                    {
                        if (wdInShpPos == Word.WdCaptionPosition.wdCaptionPositionBelow)
                        {
                            Word.Paragraph nextPara = inShpPara.Next();

                            if (nextPara == null || !String.IsNullOrWhiteSpace(nextPara.Range.Text.Trim(m_trimChars)) ||
                                nextPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                sel.SetRange(inShpPara.Range.End -1, inShpPara.Range.End -1);

                                sel.InsertParagraphAfter();
                                sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            }
                            else
                            {
                                sel.SetRange(nextPara.Range.Start, nextPara.Range.End-1);
                            }

                            sel.TypeText(strInShpFieldPlaceHolder);

                            if (strTxtInShpCapLblPreFix.Length > 0)
                            {
                                sel.TypeText(strTxtInShpCapLblPreFix);
                            }

                            if (bInShpCaplblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                            {
                                sel.TypeText(strHeadingTxt);
                            }

                            if (strTxtInShpCapLblPostFix.Length > 0)
                            {
                                sel.TypeText(strTxtInShpCapLblPostFix);
                            }

                            if (bInShpCaplblGetFromHeading && bInShpNeedSn && !String.IsNullOrWhiteSpace(strHeadingSn))
                            {
                                sel.TypeText(strHeadingSn);
                            }

                            //sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                            if (tizhuStyle != null)
                            {
                                sel.Paragraphs[1].set_Style(tizhuStyle);
                            }

                            sel.Paragraphs[1].Previous().Format.KeepWithNext = -1;
                        }
                        else // above
                        {
                            Word.Paragraph prevPara = inShpPara.Previous();

                            if (prevPara == null || !String.IsNullOrWhiteSpace(prevPara.Range.Text.Trim(m_trimChars)) ||
                                prevPara.OutlineLevel != Word.WdOutlineLevel.wdOutlineLevelBodyText)
                            {
                                sel.SetRange(inShpPara.Range.Start, inShpPara.Range.Start);
                                sel.InsertParagraphBefore();
                                sel.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                            }
                            else
                            {
                                sel.SetRange(prevPara.Range.Start, prevPara.Range.End - 1);
                            }

                            sel.TypeText(strInShpFieldPlaceHolder);

                            if (strTxtInShpCapLblPreFix.Length > 0)
                            {
                                sel.TypeText(strTxtInShpCapLblPreFix);
                            }

                            if (bInShpCaplblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                            {
                                sel.TypeText(strHeadingTxt);
                            }

                            if (strTxtInShpCapLblPostFix.Length > 0)
                            {
                                sel.TypeText(strTxtInShpCapLblPostFix);
                            }

                            if (bInShpCaplblGetFromHeading && bInShpNeedSn && !String.IsNullOrWhiteSpace(strHeadingSn))
                            {
                                sel.TypeText(strHeadingSn);
                            }

                            if (tizhuStyle != null)
                            {
                                sel.Paragraphs[1].set_Style(tizhuStyle);
                            }

                            sel.Paragraphs[1].Format.KeepWithNext = -1;
                        }

                        sel.Paragraphs[1].Alignment = (Word.WdParagraphAlignment)nInShpAlignment;

                        if (inShpFnt != null)
                        {
                            inShpFnt.SelCopy2(sel.Paragraphs[1].Range.Font);
                        }

                        if (inShpParaFmt != null)
                        {
                            inShpParaFmt.SelCopy2(sel.Paragraphs[1].Range.ParagraphFormat);
                        }

                        //switch (nInShpAlignment)
                        //{
                        //    case 1:
                        //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        //        break;

                        //    case 2:
                        //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                        //        break;

                        //    default:
                        //    case 0:
                        //        sel.Paragraphs[1].Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        //        break;
                        //}
                    }
                }

                //if (inshpSeqField != null)
                //{
                //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);

                //    sel.Find.ClearFormatting();

                //    sel.Find.Text = strInShpSeqRefPlaceHolder;
                //    sel.Find.Replacement.Text = "";
                //    sel.Find.Forward = true;
                //    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                //    sel.Find.Format = false;
                //    sel.Find.MatchCase = false;
                //    sel.Find.MatchWholeWord = false;
                //    sel.Find.MatchByte = false;
                //    sel.Find.MatchWildcards = false;
                //    sel.Find.MatchSoundsLike = false;
                //    sel.Find.MatchAllWordForms = false;

                //    sel.Find.Execute();

                //    inshpSeqField.Copy();

                //    nSelCnt = 0;

                //    while (sel.Find.Found)
                //    {
                //        para = sel.Paragraphs[1];

                //        if (bAppIsWps)
                //        {
                //            sel.Paste();
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                //        }
                //        else
                //        {
                //            RecordMultiSel(sel.Range);
                //            nSelCnt++;
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //            if (nSelCnt == 50)
                //            {
                //                ExecMultiSel(doc);
                //                sel.Paste();
                //                nSelCnt = 0;
                //            }
                //        }


                //        if (para.Next() == null)
                //        {
                //            if (bAppIsWps)
                //            {

                //            }
                //            else
                //            {
                //                if (nSelCnt > 0)
                //                {
                //                    ExecMultiSel(doc);
                //                    sel.Paste();
                //                    nSelCnt = 0;
                //                }
                //            }

                //            break;
                //        }

                //        sel.Find.Execute();
                //    }

                //    if (bAppIsWps)
                //    {

                //    }
                //    else
                //    {
                //        if (nSelCnt > 0)
                //        {
                //            ExecMultiSel(doc);
                //            sel.Paste();
                //        }
                //    }
                //}

                //if (inshpStyleRefField != null)
                //{
                //    sel.HomeKey(Word.WdUnits.wdStory,Word.WdMovementType.wdMove);

                //    sel.Find.ClearFormatting();

                //    sel.Find.Text = strInShpStyleRefPlaceHolder;
                //    sel.Find.Replacement.Text = "";
                //    sel.Find.Forward = true;
                //    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                //    sel.Find.Format = false;
                //    sel.Find.MatchCase = false;
                //    sel.Find.MatchWholeWord = false;
                //    sel.Find.MatchByte = false;
                //    sel.Find.MatchWildcards = false;
                //    sel.Find.MatchSoundsLike = false;
                //    sel.Find.MatchAllWordForms = false;

                //    sel.Find.Execute();

                //    inshpStyleRefField.Copy();

                //    nSelCnt = 0;

                //    while (sel.Find.Found)
                //    {
                //        para = sel.Paragraphs[1];

                //        if (bAppIsWps)
                //        {
                //            sel.Paste();
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                //        }
                //        else
                //        {
                //            RecordMultiSel(sel.Range);
                //            nSelCnt++;
                //            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                //            if (nSelCnt == 50)
                //            {
                //                ExecMultiSel(doc);
                //                sel.Paste();
                //                nSelCnt = 0;
                //            }
                //        }


                //        if (para.Next() == null)
                //        {
                //            if (bAppIsWps)
                //            {

                //            }
                //            else
                //            {
                //                if (nSelCnt > 0)
                //                {
                //                    ExecMultiSel(doc);
                //                    sel.Paste();
                //                    nSelCnt = 0;
                //                }
                //            }

                //            break;
                //        }

                //        sel.Find.Execute();
                //    }

                //    if (bAppIsWps)
                //    {

                //    }
                //    else
                //    {
                //        if (nSelCnt > 0)
                //        {
                //            ExecMultiSel(doc);
                //            sel.Paste();
                //        }
                //    }
                //}
            }

            if (TblCapLbl != null)
            {
                if (tblSeqField != null)
                {
                    sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

                    sel.Find.ClearFormatting();

                    sel.Find.Text = strTblSeqRefPlaceHolder;
                    sel.Find.Replacement.Text = "";
                    sel.Find.Forward = true;
                    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    sel.Find.Format = false;
                    sel.Find.MatchCase = false;
                    sel.Find.MatchWholeWord = false;
                    sel.Find.MatchByte = false;
                    sel.Find.MatchWildcards = false;
                    sel.Find.MatchSoundsLike = false;
                    sel.Find.MatchAllWordForms = false;

                    sel.Find.Execute();

                    tblSeqField.Copy();

                    while (sel.Find.Found)
                    {
                        para = sel.Paragraphs[1];

                        if (bAppIsWps)
                        {
                            sel.Paste();
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        else
                        {
                            RecordMultiSel(sel.Range);
                            nSelCnt++;
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            if (nSelCnt == 50)
                            {
                                ExecMultiSel(doc);
                                sel.Paste();
                                nSelCnt = 0;
                            }
                        }

                        if (para.Next() == null)
                        {
                            if (bAppIsWps)
                            {

                            }
                            else
                            {
                                if (nSelCnt > 0)
                                {
                                    ExecMultiSel(doc);
                                    sel.Paste();
                                    nSelCnt = 0;
                                }
                            }

                            break;
                        }

                        sel.Find.Execute();
                    }

                    if (bAppIsWps)
                    {

                    }
                    else
                    {
                        if (nSelCnt > 0)
                        {
                            ExecMultiSel(doc);
                            sel.Paste();
                        }
                    }
                }

                if (tblStyleRefField != null)
                {
                    sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

                    sel.Find.ClearFormatting();

                    sel.Find.Text = strTblStyleRefPlaceHolder;
                    sel.Find.Replacement.Text = "";
                    sel.Find.Forward = true;
                    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    sel.Find.Format = false;
                    sel.Find.MatchCase = false;
                    sel.Find.MatchWholeWord = false;
                    sel.Find.MatchByte = false;
                    sel.Find.MatchWildcards = false;
                    sel.Find.MatchSoundsLike = false;
                    sel.Find.MatchAllWordForms = false;

                    sel.Find.Execute();


                    tblStyleRefField.Copy();

                    nSelCnt = 0;

                    while (sel.Find.Found)
                    {
                        para = sel.Paragraphs[1];

                        if (bAppIsWps)
                        {
                            sel.Paste();
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        else
                        {
                            RecordMultiSel(sel.Range);
                            nSelCnt++;
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            if (nSelCnt == 50)
                            {
                                ExecMultiSel(doc);
                                sel.Paste();
                                nSelCnt = 0;
                            }
                        }


                        if (para.Next() == null)
                        {
                            if (bAppIsWps)
                            {

                            }
                            else
                            {
                                if (nSelCnt > 0)
                                {
                                    ExecMultiSel(doc);
                                    sel.Paste();
                                    nSelCnt = 0;
                                }
                            }

                            break;
                        }

                        sel.Find.Execute();
                    }

                    if (bAppIsWps)
                    {

                    }
                    else
                    {
                        if (nSelCnt > 0)
                        {
                            ExecMultiSel(doc);
                            sel.Paste();
                        }
                    }
                }
            }

            if (InShpCapLbl != null)
            {
                if (inshpSeqField != null)
                {
                    sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

                    sel.Find.ClearFormatting();

                    sel.Find.Text = strInShpSeqRefPlaceHolder;
                    sel.Find.Replacement.Text = "";
                    sel.Find.Forward = true;
                    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    sel.Find.Format = false;
                    sel.Find.MatchCase = false;
                    sel.Find.MatchWholeWord = false;
                    sel.Find.MatchByte = false;
                    sel.Find.MatchWildcards = false;
                    sel.Find.MatchSoundsLike = false;
                    sel.Find.MatchAllWordForms = false;

                    sel.Find.Execute();

                    inshpSeqField.Copy();

                    nSelCnt = 0;

                    while (sel.Find.Found)
                    {
                        para = sel.Paragraphs[1];

                        if (bAppIsWps)
                        {
                            sel.Paste();
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        else
                        {
                            RecordMultiSel(sel.Range);
                            nSelCnt++;
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            if (nSelCnt == 50)
                            {
                                ExecMultiSel(doc);
                                sel.Paste();
                                nSelCnt = 0;
                            }
                        }


                        if (para.Next() == null)
                        {
                            if (bAppIsWps)
                            {

                            }
                            else
                            {
                                if (nSelCnt > 0)
                                {
                                    ExecMultiSel(doc);
                                    sel.Paste();
                                    nSelCnt = 0;
                                }
                            }

                            break;
                        }

                        sel.Find.Execute();
                    }

                    if (bAppIsWps)
                    {

                    }
                    else
                    {
                        if (nSelCnt > 0)
                        {
                            ExecMultiSel(doc);
                            sel.Paste();
                        }
                    }
                }

                if (inshpStyleRefField != null)
                {
                    sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

                    sel.Find.ClearFormatting();

                    sel.Find.Text = strInShpStyleRefPlaceHolder;
                    sel.Find.Replacement.Text = "";
                    sel.Find.Forward = true;
                    sel.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    sel.Find.Format = false;
                    sel.Find.MatchCase = false;
                    sel.Find.MatchWholeWord = false;
                    sel.Find.MatchByte = false;
                    sel.Find.MatchWildcards = false;
                    sel.Find.MatchSoundsLike = false;
                    sel.Find.MatchAllWordForms = false;

                    sel.Find.Execute();

                    inshpStyleRefField.Copy();

                    nSelCnt = 0;

                    while (sel.Find.Found)
                    {
                        para = sel.Paragraphs[1];

                        if (bAppIsWps)
                        {
                            sel.Paste();
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        else
                        {
                            RecordMultiSel(sel.Range);
                            nSelCnt++;
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                            if (nSelCnt == 50)
                            {
                                ExecMultiSel(doc);
                                sel.Paste();
                                nSelCnt = 0;
                            }
                        }


                        if (para.Next() == null)
                        {
                            if (bAppIsWps)
                            {

                            }
                            else
                            {
                                if (nSelCnt > 0)
                                {
                                    ExecMultiSel(doc);
                                    sel.Paste();
                                    nSelCnt = 0;
                                }
                            }

                            break;
                        }

                        sel.Find.Execute();
                    }

                    if (bAppIsWps)
                    {

                    }
                    else
                    {
                        if (nSelCnt > 0)
                        {
                            ExecMultiSel(doc);
                            sel.Paste();
                        }
                    }
                }
            }

            sel.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);
            doc.Fields.Update();

            app.Options.Pagination = bPagination;
            doc.ActiveWindow.View.Type = oViewType;

            sel.SetRange(nOStart, nOEnd);
            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            return;
        }



        public String AddTiZhu(Word.Application app,Word.Document doc,
                              Word.CaptionLabel tblCapLbl,
                              Word.WdCaptionPosition tblCapPos,
                              int nTblAlign,
                              String strTblPrefix,
                              String strTblPostfix,
                              Boolean bTblGetFromHeading,
                              Boolean bTblNeedSn,
                              int nTblApplyScopeType,  // 范围：章节目录之后
                              ClassFont tblFnt,
                              ClassParagraphFormat tblParaFmt,

                              Word.CaptionLabel inShpCapLbl,
                              Word.WdCaptionPosition inShpCapPos,
                              int nInShpAlign,
                              String strInShpPrefix,
                              String strInShpPostfix,
                              Boolean bInShpGetFromHeading,
                              Boolean bInShpNeedSn,
                              int nInShpApplyScopeType,  // 范围：章节目录之后
                              ClassFont inShpFnt,
                              ClassParagraphFormat inShpParaFmt,

                              Boolean bRemoveTizhuFirst = true,
                              Boolean bRemoveTizhuBody = true
                              )
        {
            Word.Tables tbls = null;
            String strHeadingTxt = "";

            ArrayList arrIsolatePicsNotInTbl = new ArrayList();
            ArrayList arrNotIsolatePicsNotInTbl = new ArrayList();
            ArrayList arrIsolatePicsInTbl = new ArrayList();
            ArrayList arrNotIsolatePicsInTbl = new ArrayList();
            ArrayList arrInShps = null;
            ArrayList arrTbls = new ArrayList();

            Boolean bSelInShpCapLbl = false;
            Boolean bSelTblCapLbl = false;

            bSelTblCapLbl = ((tblCapLbl != null) &&!String.IsNullOrWhiteSpace(tblCapLbl.Name));
            bSelInShpCapLbl = ((inShpCapLbl != null) && !String.IsNullOrWhiteSpace(tblCapLbl.Name));

            if(!bSelTblCapLbl && !bSelInShpCapLbl)
            {
                return "错误：无题注设置";
            }
            
            Word.Selection sel = doc.ActiveWindow.Selection;

            Boolean bPagination = app.Options.Pagination;
            Word.WdViewType oViewType = doc.ActiveWindow.View.Type;
            int nOStart = sel.Start;
            int nOEnd = sel.End;


            app.Options.Pagination = false;
            doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;

            
            Boolean bTblAllDoc = false;
            Boolean bInShpAllDoc = false;

            int nPosLast = -1; // doc.Content.End;

            switch (nTblApplyScopeType)
            {
                case (int)TizhuScope.tizhuScopeAllDoc:
                    bTblAllDoc = true;
                    break;

                case (int)TizhuScope.tizhuScopeAfterToc:
                    foreach(Word.TableOfContents toc in doc.TablesOfContents)
                    {
                        if (toc.Range.End > nPosLast)
                        {
                            nPosLast = toc.Range.End;
                        }
                    }

                    sel.SetRange(nPosLast, doc.Content.End);

                    bTblAllDoc = false;

                    if (sel.End - sel.Start <= 1)
                    {
                        return @"错误：未做任何选择";
                    }

                    break;

                default:
                    bTblAllDoc = true;
                    break;
            }

            if (bTblAllDoc)
            {
                tbls = doc.Tables;
            }
            else
            {
                tbls = sel.Tables;
            }

            foreach (Word.Table wTbl in tbls)
            {
                if (isHideTbl(wTbl))
                {
                    continue;
                }

                arrTbls.Add(wTbl);
            }


            nPosLast = doc.Content.End;
            switch (nInShpApplyScopeType)
            {
                case (int)TizhuScope.tizhuScopeAllDoc:
                    bInShpAllDoc = true;
                    break;

                case (int)TizhuScope.tizhuScopeAfterToc:
                    foreach (Word.TableOfContents toc in doc.TablesOfContents)
                    {
                        if (toc.Range.End > nPosLast)
                        {
                            nPosLast = toc.Range.End;
                        }
                    }

                    sel.SetRange(nPosLast, doc.Content.End);

                    bInShpAllDoc = false;

                    if (sel.End - sel.Start <= 1)
                    {
                        return @"错误：未做任何选择";
                    }

                    break;

                default:
                    bInShpAllDoc = true;
                    break;
            }

            if (bInShpAllDoc)
            {
                arrInShps = getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                                                 arrIsolatePicsInTbl, arrNotIsolatePicsInTbl, true);
            }
            else
            {
                arrInShps = getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                                                 arrIsolatePicsInTbl, arrNotIsolatePicsInTbl, true, sel.Range);
            }

            if (arrTbls.Count + arrIsolatePicsNotInTbl.Count == 0)
            {
                return "错误：无题注对象";
            }


            int nClearTiZhuCnt = 0;
            Boolean bRemoveOnlyBody = true;
            String strRemoveFmt = "";
            String strRemoveBody = "删除段落体（不包括最后回车符），不会改变页面段落相对结构（推荐）";
            String strRemovePara = "删除整个段落（包括最后回车符），会引起下一段落向上移动一行，改变页面段落相对结构";

            if (bRemoveTizhuFirst)
            {
                bRemoveOnlyBody = bRemoveTizhuBody;

                if (bRemoveOnlyBody)
                {
                    strRemoveFmt = strRemoveBody;
                }
                else
                {
                    strRemoveFmt = strRemovePara;
                }

                if (bTblAllDoc || bInShpAllDoc)
                {
                    nClearTiZhuCnt = clearTiZhu(doc, null, bRemoveOnlyBody);
                }
                else
                {
                    nClearTiZhuCnt = clearTiZhu(doc, sel.Range, bRemoveOnlyBody);
                }

                doc.Fields.Update(); //?

                sel.Start = nOStart; // ?
                sel.End = nOEnd;
                sel.Range.GoTo();
                doc.ActiveWindow.ScrollIntoView(sel.Range);
            }

            Hashtable hashScopeHeading = null;
            Hashtable hashHeadings = null;
            Hashtable hashSimpNum = null;
            ClassOfficeCommon.TiZhuHeadingItem PrevHeadingItem = null;
            ClassOfficeCommon.TiZhuHeadingItem curHeadingItem = null;
            String strHeadingSn = "";
            int nHeadingCnt = 0;

            Boolean bTmpTblNeedSn = false, bTmpInShpNeedSn = false;

            if (bTblGetFromHeading)
            {
                bTmpTblNeedSn = bTblNeedSn;
                bTmpInShpNeedSn = bInShpNeedSn;

                hashScopeHeading = getHeadingScope(doc);
                hashHeadings = new Hashtable();

                hashSimpNum = new Hashtable();

                String[] strSimpNum = { "〇", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三",
                                      "十四","十五","十六","十七","十八","十九","二十","二十一","二十二","二十三","二十四",
                                      "二十五","二十六","二十七","二十八","二十九","三十","三十一","三十二","三十三","三十四",
                                      "三十五","三十六","三十七","三十八","三十九","四十","四十一","四十二","四十三","四十四",
                                      "四十五","四十六","四十七","四十八","四十九","五十","五十一","五十二","五十三","五十四",
                                      "五十五","五十六","五十七","五十八","五十九","六十","六十一","六十二","六十三","六十四",
                                      "六十五","六十六","六十七","六十八","六十九","七十","七十一","七十二","七十三","七十四",
                                      "七十五","七十六","七十七","七十八","七十九","八十","八十一", "八十二","八十三","八十四",
                                      "八十五","八十六","八十七","八十八","八十九","九十","九十一","九十二","九十三","九十四",
                                      "九十五","九十六","九十七","九十八","九十九","一百","一百〇一","一百〇二","一百〇三",
                                      "一百〇四","一百〇五","一百〇六","一百〇七","一百〇八","一百〇九","一百一十","一百一十一",
                                      "一百一十一","一百一十二","一百一十三","一百一十四","一百一十四","一百一十五","一百一十六",
                                      "一百一十七","一百一十八","一百一十九","一百二十"};

                int nLen = strSimpNum.GetLength(0);

                for (int i = 0; i < nLen; i++)
                {
                    hashSimpNum.Add(i, strSimpNum[i]);
                }

            }

            //Boolean bTblPosAbove = (tblCapPos == Word.WdCaptionPosition.wdCaptionPositionAbove);
            Word.CaptionLabel TblCapLbl = null;
            //Boolean bInShpPosAbove = (tblCapPos == Word.WdCaptionPosition.wdCaptionPositionAbove);
            Word.CaptionLabel InShpCapLbl = null;

            Word.WdCaptionPosition wdInShpPos = tblCapPos;
            Word.WdCaptionPosition wdTblPos = tblCapPos;

            Boolean bIgnore = true;


            //Word.Table tblItem = null;
            int nTblCnt = 0, nInShpCnt = 0;

            if (bSelTblCapLbl)
            {
                try
                {
                    TblCapLbl = tblCapLbl;// app.CaptionLabels[txtSelectedTblCapLbl.Text];
                }
                catch (System.Exception ex)
                {
                    TblCapLbl = null;
                }

                //if (bAllDoc)
                //{
                //    tbls = doc.Tables;
                //}
                //else
                //{
                //    tbls = sel.Tables;
                //}

                //foreach (Word.Table wTbl in tbls)
                //{
                //    if (isHideTbl(wTbl))
                //    {
                //        continue;
                //    }

                //    arrTbls.Add(wTbl);
                //}

                if (hashHeadings != null)
                {
                    foreach (Word.Table tblItem in arrTbls)
                    {
                        nTblCnt++;

                        TiZhuHeadingItem headingItem = null;

                        if (hashScopeHeading.Contains(tblItem.Range.Start))
                        {
                            headingItem = (TiZhuHeadingItem)hashScopeHeading[tblItem.Range.Start];
                            headingItem.nTblCnt++;

                            hashHeadings[("表" + nTblCnt)] = headingItem;
                        }
                    }
                }
                else
                {
                    nTblCnt = arrTbls.Count;
                }
            }

            // 
            if (bSelInShpCapLbl)
            {
                try
                {
                    InShpCapLbl = tblCapLbl;// app.CaptionLabels[txtSelectedInShpCapLbl.Text];
                }
                catch (System.Exception ex)
                {
                    InShpCapLbl = null;
                }


                //if (bAllDoc)
                //{
                //    arrInShps = getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                //                                     arrIsolatePicsInTbl, arrNotIsolatePicsInTbl, true);
                //}
                //else
                //{
                //    arrInShps = getInlineShpsInScope(doc, arrIsolatePicsNotInTbl, arrNotIsolatePicsNotInTbl,
                //                                     arrIsolatePicsInTbl, arrNotIsolatePicsInTbl, true, sel.Range);
                //}

                if (InShpCapLbl == null)
                {
                    // MessageBox.Show("'" + txtSelectedInShpCapLbl.Text + "'" + "题注失效，请重新设置");
                    return "'" + tblCapLbl.Name + "'" + "题注失效，请重新设置";
                }
                else
                {
                    //Word.Paragraph para = null;

                    if (hashHeadings != null)
                    {
                        //for (int nIdx = arrIsolatePicsNotInTbl.Count - 1; nIdx >= 0; nIdx--)
                        foreach (Word.Paragraph para in arrIsolatePicsNotInTbl)
                        {
                            nInShpCnt++;

                            bIgnore = true;
                            foreach (Word.InlineShape inShp in para.Range.InlineShapes)
                            {
                                if (!(inShp.Type == Word.WdInlineShapeType.wdInlineShapePictureBullet ||
                                    (inShp.OLEFormat != null && inShp.OLEFormat.DisplayAsIcon)))
                                {
                                    bIgnore = false;
                                    break;
                                }
                            }

                            if (bIgnore)
                            {
                                continue;
                            }

                            TiZhuHeadingItem headingItem = null;

                            if (hashScopeHeading.Contains(para.Range.Start))
                            {
                                headingItem = (TiZhuHeadingItem)hashScopeHeading[para.Range.Start];
                                headingItem.nInShpCnt++;

                                hashHeadings[("图" + nInShpCnt)] = headingItem;
                            }
                        }
                    }
                    else
                    {
                        nInShpCnt = arrIsolatePicsNotInTbl.Count;
                    }
                }
            }


            nHeadingCnt = 0;

            int nIconInShpCnt = 0;

            if ((nTblCnt + nInShpCnt) > 50)
            {
                //bulkAddTiZhus(app, doc, arrTbls, arrIsolatePicsNotInTbl,
                //            TblCapLbl, wdTblPos, cmbTblCapLblAlign.SelectedIndex,
                //            txtTblCapLblPreFix.Text, chkTblCaplblGetFromHeading.Checked,
                //            txtTblCapLblPostFix.Text, InShpCapLbl,
                //            wdInShpPos, cmbInShpCapLblAlign.SelectedIndex,
                //            txtInShpCapLblPreFix.Text, chkInShpCaplblGetFromHeading.Checked,
                //            txtInShpCapLblPostFix.Text, hashHeadings, hashSimpNum,
                //            bTmpTblNeedSn, bTmpInShpNeedSn);

                bulkAddTiZhus(app, doc, arrTbls, arrIsolatePicsNotInTbl,
                            TblCapLbl, wdTblPos, nTblAlign,
                            strTblPrefix, bTblGetFromHeading,
                            strTblPostfix, InShpCapLbl,
                            wdInShpPos, nInShpAlign,
                            strInShpPrefix, bInShpGetFromHeading,
                            strInShpPostfix, hashHeadings, hashSimpNum,
                            bTmpTblNeedSn, bTmpInShpNeedSn,
                            tblFnt,tblParaFmt,inShpFnt,inShpParaFmt);
            }
            else
            {
                if (bSelTblCapLbl)
                {
                    PrevHeadingItem = null;

                    if (TblCapLbl == null)
                    {
                        // MessageBox.Show("'" + txtSelectedTblCapLbl.Text + "'" + "题注失效，请重新设置");
                        return "'" + tblCapLbl.Name + "'" + "题注失效，请重新设置";
                    }
                    else
                    {
                        int nIdx = 0;
                        foreach (Word.Table tblItem in arrTbls)
                        {
                            nIdx++;

                            if (tblItem.Rows.WrapAroundText != 0)
                            {
                                Word.WdRowAlignment talign = tblItem.Rows.Alignment;
                                tblItem.Rows.WrapAroundText = 0;

                                tblItem.Rows.Alignment = talign;
                            }

                            doc.ActiveWindow.ScrollIntoView(tblItem.Range);
                            tblItem.Range.GoTo();
                            tblItem.Range.Select();

                            strHeadingTxt = "";
                            if (bTblGetFromHeading && hashHeadings.Count > 0)
                            {
                                if (hashHeadings.Contains(("表" + nIdx)))
                                {
                                    curHeadingItem = (TiZhuHeadingItem)hashHeadings[("表" + nIdx)];
                                    if (curHeadingItem == null)
                                    {
                                        strHeadingTxt = "";
                                    }
                                    else
                                    {
                                        strHeadingTxt = curHeadingItem.strHeadingName;
                                    }

                                    if (bTmpTblNeedSn && curHeadingItem != null)
                                    {
                                        strHeadingSn = "";

                                        if (curHeadingItem.nTblCnt < 2)
                                        {
                                            PrevHeadingItem = curHeadingItem;
                                            nHeadingCnt = 0;
                                        }
                                        else
                                        {
                                            if (PrevHeadingItem != null)
                                            {
                                                if (curHeadingItem.nRngStart == PrevHeadingItem.nRngStart &&
                                                    curHeadingItem.nRngEnd == PrevHeadingItem.nRngEnd)
                                                {
                                                    nHeadingCnt++;
                                                }
                                                else
                                                {
                                                    PrevHeadingItem = curHeadingItem;
                                                    nHeadingCnt = 1;
                                                }
                                            }
                                            else
                                            {
                                                PrevHeadingItem = curHeadingItem;
                                                nHeadingCnt = 1;
                                            }

                                            strHeadingSn = "";
                                            if (hashSimpNum.Contains(nHeadingCnt))
                                            {
                                                strHeadingSn = (String)hashSimpNum[nHeadingCnt];
                                            }
                                        }
                                    }
                                }
                            }

                            if (bAppIsWps)
                            {
                                sel.InsertCaption(TblCapLbl.Name, "", "", wdTblPos, false);
                            }
                            else
                            {
                                sel.InsertCaption(TblCapLbl, "", "", wdTblPos, false);
                            }

                            if (strTblPrefix.Length > 0)
                            {
                                sel.TypeText(strTblPrefix);
                            }

                            if (bTblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                            {
                                sel.TypeText(strHeadingTxt);
                            }

                            if (strTblPostfix.Length > 0)
                            {
                                sel.TypeText(strTblPostfix);
                            }

                            if (bTblGetFromHeading && bTmpTblNeedSn &&
                                !String.IsNullOrWhiteSpace(strHeadingSn))
                            {
                                sel.TypeText(strHeadingSn);
                            }

                            sel.Paragraphs[1].Alignment = (Word.WdParagraphAlignment)nTblAlign;

                            if (tblFnt != null)
                            {
                                tblFnt.SelCopy2(sel.Paragraphs[1].Range.Font);
                            }

                            if (tblParaFmt != null)
                            {
                                tblParaFmt.SelCopy2(sel.Paragraphs[1].Range.ParagraphFormat);
                            }

                        }// foreach
                    }
                }

                if (bSelInShpCapLbl)
                {
                    PrevHeadingItem = null;

                    int nCnt = 0, nIdx = 0;
                    foreach (Word.Paragraph para in arrIsolatePicsNotInTbl)
                    {
                        nIdx++;

                        foreach (Word.InlineShape inShp in para.Range.InlineShapes)
                        {
                            if (!(inShp.Type == Word.WdInlineShapeType.wdInlineShapePictureBullet ||
                                (inShp.OLEFormat != null && inShp.OLEFormat.DisplayAsIcon)))
                            {
                                nIconInShpCnt++;

                                bIgnore = false;
                                break;
                            }
                        }

                        if (bIgnore)
                        {
                            continue;
                        }

                        nCnt++;

                        doc.ActiveWindow.ScrollIntoView(para.Range);
                        para.Range.GoTo();
                        para.Range.Select();

                        strHeadingTxt = "";
                        if (bTblGetFromHeading && hashHeadings.Count > 0)
                        {
                            if (hashHeadings.Contains(("图" + nIdx)))
                            {
                                curHeadingItem = (ClassOfficeCommon.TiZhuHeadingItem)hashHeadings[("图" + nIdx)];
                                if (curHeadingItem != null)
                                {
                                    strHeadingTxt = curHeadingItem.strHeadingName;
                                }
                                else
                                {
                                    strHeadingTxt = "";
                                }

                                strHeadingSn = "";

                                if (bTmpInShpNeedSn && curHeadingItem != null)
                                {
                                    if (curHeadingItem.nInShpCnt < 2)
                                    {
                                        PrevHeadingItem = curHeadingItem;
                                        nHeadingCnt = 0;
                                    }
                                    else
                                    {
                                        if (PrevHeadingItem != null)
                                        {
                                            if (curHeadingItem.nRngStart == PrevHeadingItem.nRngStart &&
                                                curHeadingItem.nRngEnd == PrevHeadingItem.nRngEnd)
                                            {
                                                nHeadingCnt++;
                                            }
                                            else
                                            {
                                                PrevHeadingItem = curHeadingItem;
                                                nHeadingCnt = 1;
                                            }
                                        }
                                        else
                                        {
                                            PrevHeadingItem = curHeadingItem;
                                            nHeadingCnt = 1;
                                        }

                                        strHeadingSn = "";
                                        if (hashSimpNum.Contains(nHeadingCnt))
                                        {
                                            strHeadingSn = (String)hashSimpNum[nHeadingCnt];
                                        }
                                    }
                                }
                            }
                        }

                        if (bAppIsWps)
                        {
                            sel.InsertCaption(InShpCapLbl.Name, "", "", wdInShpPos, false);
                        }
                        else
                        {
                            sel.InsertCaption(InShpCapLbl, "", "", wdInShpPos, false);
                        }

                        if (strTblPrefix.Length > 0)
                        {
                            sel.TypeText(strTblPrefix);
                        }

                        if (bTblGetFromHeading && !String.IsNullOrWhiteSpace(strHeadingTxt))
                        {
                            sel.TypeText(strHeadingTxt);
                        }

                        if (strTblPostfix.Length > 0)
                        {
                            sel.TypeText(strTblPostfix);
                        }

                        if (bTblGetFromHeading && bTmpInShpNeedSn &&
                            !String.IsNullOrWhiteSpace(strHeadingSn))
                        {
                            sel.TypeText(strHeadingSn);
                        }

                        sel.Paragraphs[1].Alignment = (Word.WdParagraphAlignment)nTblAlign;

                        if (inShpFnt != null)
                        {
                            inShpFnt.SelCopy2(sel.Paragraphs[1].Range.Font);
                        }

                        if (inShpParaFmt != null)
                        {
                            inShpParaFmt.SelCopy2(sel.Paragraphs[1].Range.ParagraphFormat);
                        }
                    }
                }//if
            }//


            app.Options.Pagination = bPagination;
            // 恢复特定view
            if (doc.ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)
            {
                doc.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView;
            }
            else
            {
                doc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            }

            sel.Start = nOStart;
            sel.End = nOEnd;

            sel.Range.GoTo();
            doc.ActiveWindow.ScrollIntoView(sel.Range);

            // compile response info
            nTblCnt = (arrTbls != null) ? arrTbls.Count : 0;

            int nTotalPicParasCnt = (arrInShps != null) ? arrInShps.Count : 0;
            int nIsolatePicParasNotInTblCnt = arrIsolatePicsNotInTbl.Count;
            int nNotIsolatePicParasNotInTblCnt = arrNotIsolatePicsNotInTbl.Count;
            int nIsolatePicParasInTblCnt = arrIsolatePicsInTbl.Count;
            int nNotIsolatePicParasInTblCnt = arrNotIsolatePicsInTbl.Count;

            // int nShpCnt = (arrInShps != null)? arrInShps.Count : 0;
            // int nNonIsolatePicCnt = arrNoneIsolatePicSn.Count;

            String strTblResInfo = "", strShpResInfo = "";

            if (bSelTblCapLbl)
            {
                if (nTblCnt > 0)
                {
                    strTblResInfo = "\r\n成功添加表格题注：" + nTblCnt + "个" + "\r\n    总表格数：" + nTblCnt + "个" + "\r\n    失败数：" + 0 + "个";
                }
                else
                {
                    strTblResInfo = "\r\n没有表格，不能添加题注";
                }
            }

            if (bSelInShpCapLbl)
            {
                if (nTotalPicParasCnt > 0) // succ
                {
                    String strPosInfoNotIso = "图片位置：";
                    String strPosInfoInTbl = "图片位置：";

                    int nAbsPage = 0, nPage = 0, nLineNum = 0;

                    foreach (Word.Paragraph paraItem in arrNotIsolatePicsNotInTbl)
                    {
                        nAbsPage = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        nPage = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
                        nLineNum = paraItem.Range.get_Information(Word.WdInformation.wdFirstCharacterLineNumber);

                        strPosInfoNotIso += "总页码：" + nAbsPage + " 页码:" + nPage + " 行：" + nLineNum + "/";
                    }

                    foreach (Word.Paragraph paraItem in arrIsolatePicsInTbl)
                    {
                        nAbsPage = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        nPage = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
                        nLineNum = paraItem.Range.get_Information(Word.WdInformation.wdFirstCharacterLineNumber);

                        strPosInfoInTbl += "总页码：" + nAbsPage + " 页码:" + nPage + " 行：" + nLineNum + "/";
                    }

                    foreach (Word.Paragraph paraItem in arrNotIsolatePicsInTbl)
                    {
                        nAbsPage = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndPageNumber);
                        nPage = paraItem.Range.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber);
                        nLineNum = paraItem.Range.get_Information(Word.WdInformation.wdFirstCharacterLineNumber);

                        strPosInfoInTbl += "总页码：" + nAbsPage + " 页码:" + nPage + " 行：" + nLineNum + "/";
                    }

                    // 
                    strShpResInfo = "\r\n总图片数：" + nTotalPicParasCnt + "\r\n    完成图片段落题注数：" + nIsolatePicParasNotInTblCnt + "个" +
                        "\r\n    忽略图标图片段落数：" + nIconInShpCnt + "个" +
                        "\r\n    忽略非单独成行图片段落数：" + nNotIsolatePicParasNotInTblCnt + "个" +
                        ((nNotIsolatePicParasNotInTblCnt > 0) ? ("(" + strPosInfoNotIso + ")") : "") +
                        "\r\n    忽略表格内图片段落数：" + (nIsolatePicParasInTblCnt + nNotIsolatePicParasInTblCnt) + "个" +
                        (((nIsolatePicParasInTblCnt + nNotIsolatePicParasInTblCnt) > 0) ? ("(" + strPosInfoInTbl + ")") : "");

                }
                else
                {
                    strShpResInfo = "\r\n没有图片段落，不能添加题注";
                }
            }

            String strClearedTiZhu = "";
            if (bRemoveTizhuFirst)
            {
                if (nClearTiZhuCnt < 0)
                {
                    strClearedTiZhu = "\r\n清除原题注：0个\r\n";
                }
                else
                {
                    strClearedTiZhu = "\r\n清除原题注：" + nClearTiZhuCnt + "个\r\n";
                }
            }

            return "完成\r\n" + strClearedTiZhu + strTblResInfo + "\r\n" + strShpResInfo;
        }


        // 
        public String paraAlignment2Name(Word.WdParagraphAlignment wAlign)
        {
            String strAlignment = "";
            switch (wAlign)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphLeft: // 居左
                    strAlignment = "居左";
                    break;

                case Word.WdParagraphAlignment.wdAlignParagraphRight: // 居右
                    strAlignment = "居右";
                    break;

                case Word.WdParagraphAlignment.wdAlignParagraphCenter: // 居中
                    strAlignment = "居中";
                    break;

                case Word.WdParagraphAlignment.wdAlignParagraphDistribute: // 两端对齐
                    strAlignment = "两端对齐";
                    break;

                case Word.WdParagraphAlignment.wdAlignParagraphJustify: // 分散对齐
                    strAlignment = "分散对齐";
                    break;

                default:
                    break;
            }

            return strAlignment;
        }


        public Word.WdParagraphAlignment paraName2Alignment(String strAlignment)
        {

            if (strAlignment.Equals("居左"))
            {
                return Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
            else if (strAlignment.Equals("居中"))
            {
                return Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            else if (strAlignment.Equals("居右"))
            {
                return Word.WdParagraphAlignment.wdAlignParagraphRight;
            }
            else if (strAlignment.Equals("两端对齐"))
            {
                return Word.WdParagraphAlignment.wdAlignParagraphDistribute;
            }
            else if (strAlignment.Equals("分散对齐"))
            {
                return Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }

            return (Word.WdParagraphAlignment)(-1);
        }


        public float str2float(String strItem)
        {
            float fValue = float.NaN;

            char[] chs = strItem.ToCharArray();


            String strNum = "";
            int nCnt = 0, nPrePos = 0;

            foreach (char ch in chs)
            {
                nCnt++;

                if (ch == '-' || ch == '.' || (ch >= '0' && ch <= '9') )
                {
                    if (nPrePos > 1 && nCnt - nPrePos > 1)
                    {
                        break;
                    }

                    strNum += ch;
                    nPrePos = nCnt;
                }
            }

            if (float.TryParse(strNum, out fValue))
            {
                return fValue;
            }

            return float.NaN;

        }//


        public String formateTimeDiff(DateTime dtStart, DateTime dtEnd)
        {
            String strFmt = "";

            TimeSpan ts1 = new TimeSpan(dtStart.Ticks);
            TimeSpan ts2 = new TimeSpan(dtEnd.Ticks);
            TimeSpan ts = ts2.Subtract(ts1).Duration();

            int days = ts.Days;
            int hours = ts.Hours;
            int mins = ts.Minutes;
            int secs = ts.Seconds;
            int millsecs = ts.Milliseconds;

            if (days > 0)
            {
                strFmt += days + "天";
            }

            if (hours > 0)
            {
                strFmt += hours + "小时";
            }

            if (mins > 0)
            {
                strFmt += mins + "分钟";
            }

            if (millsecs > 0)
            {
                float fVal = (float)millsecs / 1000.0f;

                fVal += secs;
                strFmt += fVal.ToString("0.##") + "秒";
            }
            else
            {
                if (secs > 0)
                {
                    strFmt += secs + "秒";
                }
            }

            float fSecs = (float)ts.TotalMilliseconds / 1000.0f;

            strFmt += "(总:" + fSecs.ToString("0.##") + "秒)";

            return strFmt;
        }




    }
}