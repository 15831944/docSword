using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Collections;

namespace OfficeAssistCommon
{
    public class ClassProtocol
    {
        public Hashtable Decode(String strInput)
        {
            ArrayList strArr = new ArrayList();

            strArr = parseProtocol(strInput);

            Hashtable hashFields = parseFields(strArr);

            return hashFields;
        }


        public String Encode(Hashtable hashFields)
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


        private Hashtable parseFields(ArrayList strArr)
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
        private ArrayList parseProtocol(String strInput)
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

    }
}
