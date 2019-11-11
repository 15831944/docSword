using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

//using Microsoft.Office.Tools.Word;
//using Microsoft.Office.Tools;

using System.Collections;
using System.Collections.Specialized;
using Microsoft.Office.Core;

namespace OfficeTools.Common
{
    public class ClassDocDbMgr
    {

        private Word.Document m_doc = null;
        private Hashtable m_hashCategoryAllocator = new Hashtable();
        private Hashtable m_hashCategoryMaxIndex = new Hashtable();
        public readonly int SLOT_MAX_CONTENT_SIZE = 255;

        public void bindDoc(Word.Document doc)
        {
            m_doc = doc;
            return;
        }


        // ~ClassDocCustomItemsMgr()
        public void releaseAll2Doc()
        {
            String strCategory = "";
            ArrayList arrAllocator = null;

            foreach (DictionaryEntry ent in m_hashCategoryAllocator)
            {
                strCategory = (String)ent.Key;
                arrAllocator = (ArrayList)ent.Value;

                if (String.IsNullOrWhiteSpace(strCategory) || arrAllocator == null || arrAllocator.Count == 0)
                {
                    continue;
                }

                foreach (int nItem in arrAllocator)
                {
                    releaseToDoc(strCategory + nItem);
                }

            }

            return ;
        }


        // 
        //public int rebuild(Word.Document odoc = null)
        //{
        //    Word.Document doc = null;

        //    if (odoc == null)
        //    {
        //        doc = m_doc;
        //    }
        //    else
        //    {
        //        doc = odoc;
        //    }

        //    if (doc == null)
        //    {
        //        return -1;
        //    }

        //    m_hashCategoryAllocator.Clear();
        //    m_hashCategoryMaxIndex.Clear();

        //    DocumentProperties customProps = doc.CustomDocumentProperties;
        //    // DocumentProperty prop = null;

        //    foreach (DocumentProperty propItem in customProps)
        //    {
        //        // propItem.Name
        //    }

        //    return 0;
        //}

        public int calcSlotsNum(String strSavedInfo)
        {
            int nCnt = calcSlotsNum(strSavedInfo.Length);

            return nCnt;
        }

        public int calcSlotsNum(int nLen)
        {
            // int nCnt = (int)(nLen >> 8);
            float fVal = (float)nLen / (float)SLOT_MAX_CONTENT_SIZE;
            int nVal = nLen / SLOT_MAX_CONTENT_SIZE;
            int nCnt = nVal;

            if (fVal > (float)nVal)
            {
                nCnt += 1;
            }

            return nCnt;
        }


        public int saveIntoDoc(String strCategory,String strSavedInfo, ArrayList arrRes)
        {
            int nCnt = 0, nIndex = 0, nRet = 0;
            int nSpareSlotSn = 0;
            int nLeftLength = strSavedInfo.Length;
            int nItemLen = SLOT_MAX_CONTENT_SIZE;
            String strItem = "";

            nCnt = calcSlotsNum(strSavedInfo.Length);

            if (arrRes.Count != nCnt)
            {
                return -1;
            }


            nCnt = 0;
            while (nLeftLength > 0)
            {
                nItemLen = Math.Min(SLOT_MAX_CONTENT_SIZE, nLeftLength);
                strItem = strSavedInfo.Substring(nIndex, nItemLen);

                nSpareSlotSn = (int)arrRes[nCnt];
                nRet = saveIntoDoc(strCategory, nSpareSlotSn, strItem);

                if (nRet != 0)
                {
                    // MessageBox.Show("不能保存数据到文档存储空间中");
                    return -2;
                }

                nLeftLength -= nItemLen;
                nIndex += nItemLen;
                nCnt++;
            }

            return 0;
        }

        public int saveIntoDoc(String strCategory,int nSlotSn,String strValue)
        {
            DocumentProperties customProps = m_doc.CustomDocumentProperties;
            DocumentProperty prop = null;

            try
            {
                prop = (DocumentProperty)customProps[strCategory + nSlotSn];
                if (prop != null)
                {
                    prop.Delete();
                }
                customProps.Add(strCategory + nSlotSn, false, MsoDocProperties.msoPropertyTypeString, strValue);
            }
            catch (System.Exception ex)
            {
                // customProps.Add(strCategory + nSlotSn, false, MsoDocProperties.msoPropertyTypeString, strValue);
                return -1;
            }
            finally
            {
            }

            return 0;
        }


        public int getFromDoc(String strCategory, ArrayList arrSlots, ref String strValue)
        {
            DocumentProperties customProps = m_doc.CustomDocumentProperties;
            DocumentProperty prop = null;

            int nCnt = 0;
            foreach (int nSlotSn in arrSlots)
            {
                try
                {
                    prop = (DocumentProperty)customProps[strCategory + nSlotSn];
                    if (prop != null && prop.Type == MsoDocProperties.msoPropertyTypeString)
                    {
                        strValue += (String)prop.Value;
                        nCnt++;
                    }

                    // customProps.Add(strCategory + nSlotSn, false, MsoDocProperties.msoPropertyTypeString, strValue);
                }
                catch (System.Exception ex)
                {
                    // customProps.Add(strCategory + nSlotSn, false, MsoDocProperties.msoPropertyTypeString, strValue);
                    continue;
                }
                finally
                {
                }
            }

            return nCnt;
        }


        public int getFromDoc(String strCategory, int nSlotSn, ref String strValue)
        {
            DocumentProperties customProps = m_doc.CustomDocumentProperties;
            DocumentProperty prop = null;

            try
            {
                prop = (DocumentProperty)customProps[strCategory + nSlotSn];
                if (prop != null && prop.Type == MsoDocProperties.msoPropertyTypeString)
                {
                    strValue = (String)prop.Value;
                }
                else
                {
                    return -1;
                }

                // customProps.Add(strCategory + nSlotSn, false, MsoDocProperties.msoPropertyTypeString, strValue);
            }
            catch (System.Exception ex)
            {
                // customProps.Add(strCategory + nSlotSn, false, MsoDocProperties.msoPropertyTypeString, strValue);
                return -2;
            }
            finally
            {
            }

            return 0;
        }


        // to user
        // alloc 分配to outside
        public int alloc(String strCategory, int nNum, ref ArrayList arrSlots)
        {
            int nMax = 0;

            ArrayList arrAllocator = (ArrayList)m_hashCategoryAllocator[strCategory];

            if (arrAllocator == null)
            {
                arrAllocator = new ArrayList();
                m_hashCategoryAllocator[strCategory] = arrAllocator;
            }

            if (arrAllocator.Count <= nNum)
            {
                if (m_hashCategoryMaxIndex.Contains(strCategory))
                {
                    nMax = (int)m_hashCategoryMaxIndex[strCategory];
                }
                else
                {
                    nMax = 0;
                }

                int nRet = allocInDoc(strCategory, nMax, 50, ref arrAllocator);

                if (nRet != 0)
                {
                    return -1;
                }

                nMax += 50;
                m_hashCategoryMaxIndex[strCategory] = nMax;

            }

            
            int nSlotIndex = 0;

            for (int i = 0; i < nNum; i++)
            {
                nSlotIndex = (int)arrAllocator[i];
                arrSlots.Add(nSlotIndex);
            }

            arrAllocator.RemoveRange(0, (int)nNum);

            return 0;
        }



        // release 释放
        public int release(String strCategory, ArrayList arrSlots)
        {
            ArrayList arrAllocator = (ArrayList)m_hashCategoryAllocator[strCategory];

            if (arrAllocator == null)
            {
                return -1; // never happen
            }

            int nSlotIndex = 0;

            for (int i = 0; i < arrSlots.Count; i++)
            {
                nSlotIndex = (int)arrSlots[i];
                arrAllocator.Add(nSlotIndex);
            }

            return 0;
        }


        


        // to doc
        // alloc
        // release

        private int allocInDoc(String strCategory, int nMaxIndex, int nNum, ref ArrayList retArr)
        {
            if (m_doc == null)
            {
                return -1;
            }

            DocumentProperties customProps = m_doc.CustomDocumentProperties;
            DocumentProperty prop = null;

            for (int i = nMaxIndex; i < (nMaxIndex + nNum); i++)
            {
                try
                {
                    prop = (DocumentProperty)customProps[strCategory + i];
                    if (prop != null)
                    {
                        prop.Delete();
                    }
                    customProps.Add(strCategory + i, false, MsoDocProperties.msoPropertyTypeString, strCategory + i);
                }
                catch (System.Exception ex)
                {
                    customProps.Add(strCategory + i, false, MsoDocProperties.msoPropertyTypeString, strCategory + i);
                    // return -2;
                }
                finally
                {
                }

                retArr.Add(i);
            }

            return 0;
        }

        // 
        private int releaseToDoc(String strName)
        {
            if (m_doc == null)
            {
                return -1;
            }

            DocumentProperties customProps = m_doc.CustomDocumentProperties;
            DocumentProperty prop = null;

            try
            {
                prop = (DocumentProperty)customProps[strName];
                if (prop != null)
                {
                    prop.Delete();
                }
            }
            catch (System.Exception ex)
            {
                // customProps.Add("Name" + (i + 1), false, MsoDocProperties.msoPropertyTypeString, "NameValue" + (i + 1));
                return -2;
            }
            finally
            {
            }

            return 0;
        }


        private int splitNameNumber(String strNameNumber, ref String strName, ref String strNumber)
        {
            Char[] chs = strNameNumber.ToCharArray();
            Char ch = (Char)0;
            int nLen = chs.GetLength(0);
            int nIndex = -1;

            for (int i = nLen - 1; i > 0;i--)
            {
                ch = chs[i];
                if (ch >= '0' && ch <= '9')
                {
                    nIndex = i;
                }
                else
                {
                    break;
                }
            }

            if (nIndex == -1)
            {
                strName = strNameNumber;
                strNumber = "";
            }
            else
            {
                strName = strNameNumber.Substring(0, nIndex);
                strNumber = strNameNumber.Substring(nIndex);
            }

            return 0;
        }

        // 
        public int rebuild(Word.Document doc)
        {
            // 
            if (m_doc == null)
            {
                m_doc = doc;
            }

            m_hashCategoryAllocator.Clear();
            m_hashCategoryMaxIndex.Clear();

            DocumentProperties customProps = m_doc.CustomDocumentProperties;
            String strName = "", strNumber = "";
            int nRet = 0, nMax = 0, nVal = 0;

            foreach(DocumentProperty prop in customProps)
            {
                // split number and name 2 parts

                nRet = splitNameNumber(prop.Name, ref strName, ref strNumber);

                if (String.IsNullOrWhiteSpace(strName))
                {
                    continue;
                }
                else
                {
                    if (!m_hashCategoryAllocator.Contains(strName))
                    {
                        ArrayList arr = new ArrayList();

                        m_hashCategoryAllocator[strName] = arr;
                        m_hashCategoryMaxIndex[strName] = 0;
                    }
                }

                if (!String.IsNullOrWhiteSpace(strNumber) && int.TryParse(strNumber, out nVal))
                {
                    nMax = (int)m_hashCategoryMaxIndex[strName];
                    nMax = (nVal > nMax)? nVal : nMax;
                    m_hashCategoryMaxIndex[strName] = nMax + 1;
                }
            }

            return 0;
        }

    }
}
