using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Windows.Forms;
using System.Collections;

namespace OfficeAssist.docPub
{
    public class docPubDataLevelMgr
    {

        private Hashtable m_hshDocPubScheme = new Hashtable();

        public docPubScheme GetDocPubScheme(String strType, String strSchemeName)
        {
            String strKey = strType + @"/" + strSchemeName;

            docPubScheme docPubSchemeObj = (docPubScheme)m_hshDocPubScheme[strKey];

            return docPubSchemeObj;
        }

        public Boolean IsSchemeLoaded(String strType,String strSchemeName)
        {
            String strKey = strType + @"/" + strSchemeName;

            Boolean bRet = m_hshDocPubScheme.Contains(strKey);

            return bRet;
        }


        // 
        public int LoadRemoteSchemeNames()
        {
            return 0;
        }


        public int CreateBuiltInTypes()
        {
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";

            String[] strTypes = {"服务端(连线)","服务端(离线)","预置","自定义" };
            String strTypeDir = strDocPubDir;

            foreach(String strType in strTypes)
            {
                strTypeDir = strDocPubDir + strType;

                if (!Directory.Exists(strTypeDir))
                {
                    Directory.CreateDirectory(strTypeDir);
                }
            }

            return 0;
        }


        // 
        public int LoadNames(ref TreeNode trn)
        {
            CreateBuiltInTypes();

            // get dir
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";

            if (!Directory.Exists(strDocPubDir))
            {
                return -1;
            }
            

            IEnumerable<String> SubDirs = Directory.EnumerateDirectories(strDocPubDir);
            // IEnumerator<String> enSubDirs = SubDirs.GetEnumerator();

            TreeNode subTr = null, schemeTr = null, tmpTr = null;
            String strItem = "", strSchemeName = "";
            int nPos = -1, nLen = -1;
            ArrayList arrXmlFiles = new ArrayList();

            trn.Nodes.Clear();
            foreach(String strSubDir in SubDirs)
            {
                nPos = strSubDir.LastIndexOf('\\');
                nLen = strSubDir.Length;

                if (nPos != -1 && nLen > nPos)
                {
                    //strItem = strSubDir.Substring(nPos+1,nLen - nPos);
                    strItem = strSubDir.Substring(nPos + 1);
                    subTr = new TreeNode(strItem);
                    // 
                    IEnumerable<String> xmlFiles = Directory.EnumerateFiles(strSubDir, "*.xml");

                    arrXmlFiles.Clear();
                    foreach(String strXmlFile in xmlFiles)
                    {
                        arrXmlFiles.Add(strXmlFile);
                    }
                    
                    arrXmlFiles.Sort();

                    foreach (String strXmlFile in arrXmlFiles)
                    {
                        strSchemeName = Path.GetFileNameWithoutExtension(strXmlFile);
                        schemeTr = new TreeNode(strSchemeName);

                        tmpTr = new TreeNode();
                        schemeTr.Nodes.Add(tmpTr);

                        subTr.Nodes.Add(schemeTr);
                    }
                    trn.Nodes.Add(subTr);
                }
                
            }

            return 0;
        }


        public docPubScheme LoadObject(String strType, String strSchemeName)
        {
            // 
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";

            String strXmlFile = strDocPubDir + strType + @"\" + strSchemeName + ".xml";

            docPubScheme docPubSchemeObj = null;

            docPubSchemeObj = LoadObject(strXmlFile);

            if (docPubSchemeObj != null)
            {
                String strKey = strType + @"/" + strSchemeName;
                m_hshDocPubScheme[strKey] = docPubSchemeObj;
            }

            return docPubSchemeObj;
        }



        // 指定Load, Lazy Load
        public docPubScheme LoadObject(String strXmlFile)
        {
            if (!File.Exists(strXmlFile))
            {
                return null;
            }

            StreamReader sr = new StreamReader(strXmlFile);
            String strXmlContent = sr.ReadToEnd();
            sr.Close();

            if (String.IsNullOrWhiteSpace(strXmlContent))
            {
                return null;
            }

            docPubScheme docPubSchemeObj = null;

            try
            {
                docPubSchemeObj = docPub.XmlUtility.DeserializeToObject<docPubScheme>(strXmlContent);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
            }

            return docPubSchemeObj;
        }


        public int Save2Xml(docPubScheme docPubSchemeObj,String strType, String strSchemeName)
        {
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";
            String strTypeDir = strDocPubDir + @"" + strType + @"\";
            String strXmlFile = strTypeDir + strSchemeName + @".xml";

            if (!Directory.Exists(strTypeDir))
            {
                Directory.CreateDirectory(strTypeDir);
            }

            int nInvalid = CheckFileNameValid(strType, strSchemeName);

            if (nInvalid < 0)
            {
                return nInvalid;
            }

            String strXml = "";
            
            try
            {
            	strXml = docPub.XmlUtility.SerializeToXml<docPubScheme>(docPubSchemeObj);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -1;
            }
            finally
            {
            }

            StreamWriter sw = null;
            
            try
            {
	            sw = new StreamWriter(strXmlFile);
                sw.Write(strXml);
                sw.Flush();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -2;
            }
            finally
            {
                sw.Close();
            }

            return 0;
        }


        public Boolean IsSchemeNameValid(String strSchemeName)
        {
            if (String.IsNullOrWhiteSpace(strSchemeName))
            {
                return false;
            }

            if (strSchemeName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                //含有非法字符 \ / : * ? " < > | 等
                // Show("模板名称含有非法字符，请重新输入", "错误", Error, OK);
                return false;
            }

            return true;
        }


        // 
        public int CheckFileNameValid(String strType, String strSchemeName)
        {
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";
            String strTypeDir = strDocPubDir + @"\" + strType + @"\";
            String strXmlFile = strTypeDir + strSchemeName + @".xml";

            if (!Directory.Exists(strDocPubDir))
            {
                return -1;
            }

            if (!Directory.Exists(strTypeDir))
            {
                return -2;
            }

            if (strXmlFile.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                //含有非法字符 \ / : * ? " < > | 等
                // Show("模板名称含有非法字符，请重新输入", "错误", Error, OK);
                return -3;
            }

            if (File.Exists(strXmlFile))
            {
                return -4;
            }

            return 0;
        }



        // 验证某个XML是否合法
        public int Verify(docPubScheme docPubSchemeObj)
        {
            String strXml = "";

            try
            {
                strXml = docPub.XmlUtility.SerializeToXml<docPub.docPubScheme>(docPubSchemeObj);
            }
            catch (System.Exception ex)
            {
                // MessageBox.Show(ex.Message);
                return -1;
            }
            finally
            {
            }

            return 0;
        }


        // 
        public int Add(String strType, String strSchemeName, docPubScheme docPubSchemeObj)
        {
            int nRet = Save2Xml(docPubSchemeObj,strType, strSchemeName);
            return nRet;
        }

        // 
        public int Remove(String strType, String strSchemeName)
        {
            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";
            String strTypeDir = strDocPubDir + @"\" + strType + @"\";
            String strXmlFile = strTypeDir + strSchemeName + @".xml";

            if (!File.Exists(strXmlFile))
            {
                return -1;
            }

            try
            {
            	File.Delete(strXmlFile);
            }
            catch (System.Exception ex)
            {
                return -2;
            }
            finally
            {
            }

            return 0;
        }


        // 
        public int Export(docPubScheme docPubSchemeObj, String strXmlFile)
        {
            String strXml = "";

            try
            {
                strXml = docPub.XmlUtility.SerializeToXml<docPubScheme>(docPubSchemeObj);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -1;
            }
            finally
            {
            }

            StreamWriter sw = null;

            try
            {
                sw = new StreamWriter(strXmlFile);
                sw.Write(strXml);
                sw.Flush();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -2;
            }
            finally
            {
                sw.Close();
            }


            return 0;
        }


        // 
        public int Import(String strInXmlFile, String strType, String strExpSchemeName = null)
        {
            if (!File.Exists(strInXmlFile))
            {
                return -1;
            }
            // 
            String strSchemeName = strExpSchemeName;

            if (String.IsNullOrWhiteSpace(strExpSchemeName))
            {
                strSchemeName = Path.GetFileNameWithoutExtension(strInXmlFile);
            }

            int nRet = CheckFileNameValid(strType,strSchemeName);

            if (nRet < 0)
            {
                return nRet;
            }

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strDocPubDir = strBaseDir + @"docPub\";
            String strTypeDir = strDocPubDir + @"\" + strType + @"\";
            String strXmlFile = strTypeDir + strSchemeName + @".xml";

            if (!Directory.Exists(strTypeDir))
            {
                Directory.CreateDirectory(strTypeDir);
            }

            try
            {
            	File.Copy(strInXmlFile, strXmlFile);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -2;
            }
            finally
            {
            }

            return 0;
        }


    }
}
