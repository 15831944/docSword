using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Collections;


namespace OfficeTools.Common
{
    public class ClassNumTrans
    {


        // 数值转换，数额转换
        // 


        public static void digitTranslate(double dbNum, out String strArabicNum,
                                          out String strSimpChNum, out String strBigSimpChNum)
        {
            String strNum = dbNum.ToString();

            digitTranslate(strNum, out strArabicNum, out strSimpChNum, out strBigSimpChNum);
            return;
        }


        public static void digitTranslate(String strOrigText, out String strArabicNum,
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

            }

            strArabicNum = new String(chArrArabicNum);
            strSimpChNum = new String(chArrSimpChNum);
            strBigSimpChNum = new String(chArrBigSimpChNum);

            return;
        }
    }
}
