using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Web;
using System.Collections.Specialized;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Runtime.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeTools.Common;

namespace OfficeAssist
{
    /// <summary>
    /// HTTP处理类
    /// </summary>
    public class ShareContributorOper
    {
        private String m_cfgLoginUrl = "";
        private String m_cfgUploadFileUrl = "";
        public String m_cfgDocRepositoryUrl = "";
        public String m_cfgTempFileLoc = @"d:\temp\";

        public ShareContributorOper()
        {
            return;
        }

        // config
        public void setConfig(String cfgLoginUrl,String cfgUploadFileUrl,
                              String cfgDocRepositoryUrl,String cfgTempFileLoc)
        {
            m_cfgLoginUrl = cfgLoginUrl;
            m_cfgUploadFileUrl = cfgUploadFileUrl;
            m_cfgDocRepositoryUrl = cfgDocRepositoryUrl;
            m_cfgTempFileLoc = cfgTempFileLoc + "\\";

            return;
        }

        private ClassOfficeCommon m_commTools = null;

        public void setCommonTools(ClassOfficeCommon commTools)
        {
            m_commTools = commTools;
            return;
        }


        // 
        [DataContract()]
        public class TypeNode
        {
            // {"status":"success","value":"0","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String id;
            [DataMember(Order = 1)]
            public String name;
            [DataMember(Order = 2)]
            public String parentId;
            [DataMember(Order = 3)]
            public String orderId;
            [DataMember(Order = 4)]
            public String type; // 1:板块, 2:folder, 其它:file
            [DataMember(Order = 6)]
            public Int32 isLeaf; // 1,true; 0,false

        }

        [DataContract()]
        public class FileNode
        {
            // {"status":"success","value":"0","message":"登录成功"} 
          
            [DataMember(Order = 0)]
            public String nodeId; // parentid, typeid
            [DataMember(Order = 1)]
            public String templateFileId;
            [DataMember(Order = 2)]
            public String templateFileName;
            [DataMember(Order = 3)]
            public String exampleFileId;
            [DataMember(Order = 4)]
            public String exampleFileName;

        }

        [DataContract()]
        public class fileResult
        {
            // [{\"name\":\"样板样式.docx\",\"file_id\":\"a3d69ea6-37d0-4d9c-b887-f20351e8475b\",\"parent_id\":\"F961D297-4719-6DE6-CE02-57F550314182\"}]}\r\n"

            [DataMember(Order = 0)]
            public String name;
            //public String sheetId;

            [DataMember(Order = 1)]
            public String file_id;
            //public String dataStatusId;

            [DataMember(Order = 2)]
            public String parent_id;
            //public String createDate;

/*
            [DataMember(Order = 3)]
            public String createUserId;
            [DataMember(Order = 4)]
            public String createUser;
            [DataMember(Order = 5)]
            public String patternId;
            [DataMember(Order = 6)]
            public String node_Id;
            [DataMember(Order = 7)]
            public String template_File_Id;
            [DataMember(Order = 8)]
            public String template_File_Name;
            [DataMember(Order = 9)]
            public String create_date;
            [DataMember(Order = 10)]
            public String create_responser;
            [DataMember(Order = 11)]
            public String path;
 */
        }



        [DataContract()]
        public class fileResult_v1
        {
            [DataMember(Order = 0)]
            public String sheetId;
            [DataMember(Order = 1)]
            public String dataStatusId;
            [DataMember(Order = 2)]
            public String createDate;
            [DataMember(Order = 3)]
            public String createUserId;
            [DataMember(Order = 4)]
            public String createUser;
            [DataMember(Order = 5)]
            public String patternId;
            [DataMember(Order = 6)]
            public String node_Id;
            [DataMember(Order = 7)]
            public String template_File_Id;
            [DataMember(Order = 8)]
            public String template_File_Name;
            [DataMember(Order = 9)]
            public String create_date;
            [DataMember(Order = 10)]
            public String create_responser;
            [DataMember(Order = 11)]
            public String path;
        }

        [DataContract()]
        public class loginReturnResult
        {
            // {"status":"success","value":"0","message":"登录成功!","result":"70e37a71-184b-4f04-885d-e32cada84fa5"} 
            [DataMember(Order = 0)]
            public String status { get; set; }
            [DataMember(Order = 1)]
            public int value { get; set; }
            [DataMember(Order = 2)]
            public String message { get; set; }
            [DataMember(Order = 3)]
            public String result { get; set; }
        }

        [DataContract()]
        public class fileReturnResult
        {
            // "{\"message\":\"文件上传成功！\",\"status\":\"1\",
            //  \"data\":[{\"name\":\"样板样式.docx\",\"file_id\":\"a3d69ea6-37d0-4d9c-b887-f20351e8475b\",\"parent_id\":\"F961D297-4719-6DE6-CE02-57F550314182\"}]}\r\n"
            //
            //xx {"status":"success","value":"0","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String message;

            [DataMember(Order = 1)]
            public int status;

            [DataMember(Order = 2)]
            public fileResult[] data;

            /*
            [DataMember(Order = 0)]
            public String status;
            [DataMember(Order = 1)]
            public int value;
            [DataMember(Order = 2)]
            public String message;
            [DataMember(Order = 3)]
            public fileResult[] result;
             */ 
        }

        [DataContract()]
        public class typeReturnResult
        {
            // {"status":"success","value":"0","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String status;
            [DataMember(Order = 1)]
            public int value;
            [DataMember(Order = 2)]
            public String message;
            [DataMember(Order = 3)]
            public TypeNode[] result;
        }


        [DataContract()]
        public class typeFilePermissionReturnResult
        {
            // {"status":"success","value":"0","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String status;
            [DataMember(Order = 1)]
            public int value;
            [DataMember(Order = 2)]
            public String message;
            [DataMember(Order = 3)]
            public typeFilePermission[] result;
        }


        [DataContract()]
        public class typeFilePermission
        {
            // [{"ID":"c583a058-48a6-44bb-a022-89d74097325e","FOLDERID":"E78162FC-EF25-97E1-2EAF-2616D5881D11","VALUE":"3"},{"ID":"1740651f-f917-425e-98c5-0cae1dd10b41","FOLDERID":"D3CA67A3-491E-4BA8-9B08-445DB165E5E8","VALUE":"1"}]

            [DataMember(Order = 0)]
            public String ID;
            [DataMember(Order = 1)]
            public String FOLDERID;
            [DataMember(Order = 2)]
            public int VALUE;
        }


        [DataContract()]
        public class dirReturnResultData
        {
            // {\"dir_id\":\"b041847f-b4a5-4b2a-ba9f-f6093795b263\",\"parent_dir_id\":\"BDCF2818-3740-49BB-8A3A-B648549671CB\",\"dir_name\":\"dirdir2\"}}\r\n"

            [DataMember(Order = 0)]
            public String dir_id;
            [DataMember(Order = 1)]
            public String parent_dir_id;
            [DataMember(Order = 2)]
            public String dir_name;
        }


        [DataContract()]
        public class dirReturnResult
        {
            // "{\"message\":\"�����ļ��гɹ���\",
            // \"result\":[],
            // \"status\":\"1\",
            // \"data\":{\"dir_id\":\"b041847f-b4a5-4b2a-ba9f-f6093795b263\",\"parent_dir_id\":\"BDCF2818-3740-49BB-8A3A-B648549671CB\",\"dir_name\":\"dirdir2\"}}\r\n"

            [DataMember(Order = 0)]
            public String message;

            [DataMember(Order = 1)]
            public typeFilePermission[] result;

            [DataMember(Order = 2)]
            public int status;

            [DataMember(Order = 3)]
            public dirReturnResultData data;


            /*
            [DataMember(Order = 3)]
            public String dir_id;

            [DataMember(Order = 4)]
            public String dir_name;

            [DataMember(Order = 5)]
            public String parent_dir_id;
            */
            

            /*
            [DataMember(Order = 0)]
            public String status;
            [DataMember(Order = 1)]
            public int value;
            [DataMember(Order = 2)]
            public String message;
            [DataMember(Order = 3)]
            public fileResult[] result;
             */
        }


        public int login_v2(String username, String pwd, ref String strRetMsg)
        {
            NameValueCollection form = new NameValueCollection();

            form.Add("username", username);
            form.Add("password", pwd);


            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;

            //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");

            int nRet = -1;
            strRetMsg = "异常";

            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String responseText = Encoding.UTF8.GetString(responseArray);

                // responseText = "{\"status\":\"success\",\"value\":\"OK\",\"message\":\"ok\",\"result\":\"39a68e4a-88e4-4788-bfcb-c0363aea2a8b\"}";

                loginReturnResult res = (loginReturnResult)JsonConvert.DeserializeObject(responseText, typeof(loginReturnResult));
                nRet = res.value;
                strRetMsg = res.message;

            }
            catch (WebException exp)
            {
                MessageBox.Show(exp.Message, "Error");
            }
            finally
            {
                wc.Dispose();
            }

            return nRet;
        }


        // login
        public int login(String username, String pwd, ref String strRetMsg)
        {
            NameValueCollection form = new NameValueCollection();

            form.Add("username", username);
            form.Add("password", pwd);


            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;

            //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");

            int nRet = -1;
            strRetMsg = "异常";

            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgLoginUrl, "POST", form);
                String responseText = Encoding.UTF8.GetString(responseArray);

                // responseText = "{\"status\":\"success\",\"value\":\"OK\",\"message\":\"ok\",\"result\":\"39a68e4a-88e4-4788-bfcb-c0363aea2a8b\"}";

                loginReturnResult res = (loginReturnResult)JsonConvert.DeserializeObject(responseText, typeof(loginReturnResult));

                if (res != null)
                {
                    nRet = res.value;
                    strRetMsg = res.message;
                }

            }
            catch (WebException exp)
            {
                MessageBox.Show(exp.Message, "Error");
            }
            finally
            {
                wc.Dispose();
            }

            return nRet;
        }

        //
        private int iconIndex(String strFileName)
        {
            int nIndex = 17;

            String strExt = Path.GetExtension(strFileName);

            if (strExt != null && !strExt.Equals(""))
            {
                strExt = strExt.ToUpper();

                if (strExt.Equals(".DOC") || strExt.Equals(".DOCX")) // word
                {
                    nIndex = 0;
                }
                else if (strExt.Equals(".XLS") || strExt.Equals(".XLSX")) // excel
                {
                    nIndex = 14;
                }
                else if (strExt.Equals(".PPT") || strExt.Equals(".PPTX")) // ppt
                {
                    nIndex = 15;
                }
                else if (strExt.Equals(".PDF") ) // pdf
                {
                    nIndex = 16;
                }

            }

            return nIndex;
        }


//         public TreeNode getNodes_v1(String typeParam, String fileParam, ref TreeNode privLib)
//         {
//             TreeNode rootNd = new TreeNode("root");
// 
//             Hashtable hashCommonLibNodes = new Hashtable();
//             Hashtable hashPrivLibNodes = new Hashtable();
// 
// 
//             NameValueCollection form = new NameValueCollection();
// 
//             // form.Add("projecttree", typeParam);
//             form.Add("username", typeParam);
// 
//             WebClient wc = new WebClient();
// 
//             wc.Encoding = Encoding.UTF8;
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgGetTypeTreeUrl, "POST", form);
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 JArray jsonArr = JArray.Parse(jsonText);
//                 int nLib = 0;
// 
//                 foreach (JObject jobj in jsonArr)
//                 {
//                     TypeNode pt = new TypeNode();
// 
//                     pt.id = (String)jobj["ID"];
//                     pt.name = (String)jobj["NAME"];
// 
//                     pt.parentId = (String)jobj["PARENTID"];
//                     pt.orderId = (String)jobj["ORDERID"];
//                     pt.type = (String)jobj["TYPE"];
// 
//                     nLib = (int)jobj["LIB"];
// 
//                     TreeNode typeNd = new TreeNode(pt.name);
//                     typeNd.Name = pt.name;
//                     typeNd.Tag = pt;
//                     typeNd.ImageIndex = typeNd.SelectedImageIndex = 13;
// 
//                     if (nLib == 1)
//                     {
//                         hashCommonLibNodes.Add(pt.id, typeNd);
//                     }
//                     else // 1
//                     {
//                         hashPrivLibNodes.Add(pt.id, typeNd);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 wc.Dispose();
//                 return null;
//             }
//             finally
//             {
//                 //wc.Dispose();
//             }
// 
//             // 
//             String strParentId = "";
//             TreeNode nd = null, fndNd = null;
//             TypeNode tmpTypeNode = null;
// 
//             IDictionaryEnumerator enumRator = hashCommonLibNodes.GetEnumerator();
//             DictionaryEntry entry;
//             while (enumRator.MoveNext())
//             {
//                 entry = (DictionaryEntry)enumRator.Current;
//                 nd = (TreeNode)(entry.Value);
// 
//                 tmpTypeNode = (nd.Tag as TypeNode);
// 
//                 strParentId = tmpTypeNode.parentId;
// 
//                 if (tmpTypeNode.type.Equals("0"))
//                 {
//                     nd.Tag = "#" + tmpTypeNode.id; // type node
//                     nd.ImageIndex = nd.SelectedImageIndex = 13;
//                 }
//                 else
//                 {
//                     nd.Tag = "$" + tmpTypeNode.id; // type node
//                     nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
//                 }
// 
//                 if (strParentId != null)
//                 {
//                     fndNd = (TreeNode)hashCommonLibNodes[strParentId];
//                     if (fndNd == null)
//                     {
//                         rootNd.Nodes.Add(nd);
//                     }
//                     else
//                     {
//                         fndNd.Nodes.Add(nd);
//                     }
//                 }
//                 else
//                 {
//                     rootNd.Nodes.Add(nd);                    
//                 }
//             }
// 
// 
//             enumRator = hashPrivLibNodes.GetEnumerator();
//             while (enumRator.MoveNext())
//             {
//                 entry = (DictionaryEntry)enumRator.Current;
//                 nd = (TreeNode)(entry.Value);
// 
//                 tmpTypeNode = (nd.Tag as TypeNode);
// 
//                 strParentId = tmpTypeNode.parentId;
// 
//                 if (tmpTypeNode.type.Equals("0"))
//                 {
//                     nd.Tag = "#" + tmpTypeNode.id; // type node
//                     nd.ImageIndex = nd.SelectedImageIndex = 13;
//                 }
//                 else
//                 {
//                     nd.Tag = "$" + tmpTypeNode.id; // type node
//                     nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
//                 }
// 
//                 if (strParentId != null)
//                 {
//                     fndNd = (TreeNode)hashPrivLibNodes[strParentId];
//                     if (fndNd == null)
//                     {
//                         privLib.Nodes.Add(nd);
//                     }
//                     else
//                     {
//                         fndNd.Nodes.Add(nd);
//                     }
//                 }
//                 else
//                 {
//                     privLib.Nodes.Add(nd);
//                 }
//             }
// 
// 
// /*
// 
//             // get file list 
//             form.Clear();
//             form.Add("param", fileParam);
// 
//             wc.Encoding = Encoding.UTF8;
//             // wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgGetFilesUrl, "POST", form);
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 JArray jsonArr = JArray.Parse(jsonText);
// 
//                 foreach (JObject jobj in jsonArr)
//                 {
//                     FileNode fileNd = new FileNode();
// 
//                     fileNd.nodeId = (String)jobj["NODE_ID"];
//                     fileNd.templateFileId = (String)jobj["TEMPLATE_FILE_ID"];
//                     fileNd.templateFileName = (String)jobj["TEMPLATE_FILE_NAME"];
// 
//                     fileNd.exampleFileId = (String)jobj["EXAMPLE_FILE_ID"];
//                     fileNd.exampleFileName = (String)jobj["EXAMPLE_FILE_NAME"];
// 
//                     fndNd = (TreeNode)hashNodes[fileNd.nodeId];
// 
//                     if (fndNd == null)
//                         fndNd = rootNd;
// 
//                     if (fileNd.templateFileName != null / *&& Globals.ThisAddIn.IsSupportFileFormat(fileNd.templateFileName)* /)
//                     {
//                         TreeNode templateFileNd = new TreeNode(fileNd.templateFileName);
//                         templateFileNd.Name = fileNd.templateFileName;
//                         templateFileNd.Tag = "$" + fileNd.templateFileId;
//                         templateFileNd.ImageIndex = templateFileNd.SelectedImageIndex = 0;
//                         fndNd.Nodes.Add(templateFileNd);
//                     }
// 
//                     if (fileNd.exampleFileName != null / *&& Globals.ThisAddIn.IsSupportFileFormat(fileNd.exampleFileName)* /)
//                     {
//                         TreeNode exampleFileNd = new TreeNode(fileNd.exampleFileName);
// 
//                         exampleFileNd.Name = fileNd.exampleFileName;
//                         exampleFileNd.Tag = "$" + fileNd.exampleFileId;
//                         exampleFileNd.ImageIndex = exampleFileNd.SelectedImageIndex = 0;
//                         fndNd.Nodes.Add(exampleFileNd);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 return null;
//             }
//             finally
//             {
//                 wc.Dispose();
//             }*/
// 
//             return rootNd;
// 
//         }


        // 2017/02/16 15:15已测新文库，可行
        // download file
        public int downloadFile(String userName,String fileId, ref String strLocFileUrl)
        {
            NameValueCollection form = new NameValueCollection();

            form.Add("username", userName);
            form.Add("fileId", fileId);
            form.Add("serverId", "client_fms_002");


            WebClient wc = new WebClient();
            //wc.Headers.
            wc.Encoding = Encoding.UTF8;
            //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");

            int nRet = 0;
            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);

                String resStr = wc.ResponseHeaders.Get("Content-Disposition");

                if (resStr != null)
                {
                    String strTransed = HttpUtility.UrlDecode(resStr);

                    int nStart = strTransed.IndexOf('\"');
                    int nEnd = strTransed.LastIndexOf('\"');
                    String strFileName = strTransed.Substring(nStart + 1, nEnd - nStart - 1);

                    try
                    {
                        if (strLocFileUrl.Equals(""))
                        {
                            strLocFileUrl = m_cfgTempFileLoc + strFileName;
                        }
                        else
                        {
                            strLocFileUrl = strLocFileUrl + "\\" + strFileName;
                        }

                        FileStream fs = new FileStream(strLocFileUrl, FileMode.Create);

                        fs.Write(responseArray, 0, responseArray.Length);
                        fs.Flush();
                        fs.Close();
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error");
                        nRet = -1;
                    }
                    finally
                    {
                    }
                }
                else
                {
                    MessageBox.Show("文件不存在");
                    nRet = -1;
                }
            }
            catch (WebException exp)
            {
                MessageBox.Show(exp.Message, "Error");
                nRet = -1;
            }
            finally
            {
                wc.Dispose();
            }


            
//             try
//             {
// 	            byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
// 	            String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 downloadFileResult res = (downloadFileResult)JsonConvert.DeserializeObject(jsonText, typeof(downloadFileResult));
// 	
// 	            if (res == null || res.status == 0) // failure
// 	            {
// 	                return -1;
// 	            }
//             }
//             catch (System.Exception ex)
//             {
//             	
//             }
//             finally
//             {
//             }


            return nRet;
        }

//         public int downloadFile_v1(String fileId, ref String strLocFileUrl)
//         {
//             NameValueCollection form = new NameValueCollection();
// 
//             form.Add("fileId", fileId);
// 
//             WebClient wc = new WebClient();
//             //wc.Headers.
//             wc.Encoding = Encoding.UTF8;
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             int nRet = 0;
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgDownloadFileUrl, "POST", form);
// 
//                 String resStr = wc.ResponseHeaders.Get("Content-Disposition");
// 
//                 if (resStr != null)
//                 {
//                     String strTransed = HttpUtility.UrlDecode(resStr);
// 
//                     int nStart = strTransed.IndexOf('\"');
//                     int nEnd = strTransed.LastIndexOf('\"');
//                     String strFileName = strTransed.Substring(nStart + 1, nEnd - nStart - 1);
// 
//                     try
//                     {
//                         if (strLocFileUrl.Equals(""))
//                         {
//                             strLocFileUrl = m_cfgTempFileLoc + strFileName;
//                         }
//                         else
//                         {
//                             strLocFileUrl = strLocFileUrl + "\\" + strFileName;
//                         }
// 
//                         FileStream fs = new FileStream(strLocFileUrl, FileMode.Create);
// 
//                         fs.Write(responseArray, 0, responseArray.Length);
//                         fs.Flush();
//                         fs.Close();
//                     }
//                     catch (System.Exception ex)
//                     {
//                         MessageBox.Show(ex.Message, "Error");
//                         nRet = -1;
//                     }
//                     finally
//                     {
//                     }
//                 }
//                 else
//                 {
//                     MessageBox.Show("文件不存在");
//                     nRet = -1;
//                 }
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 nRet = -1;
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return nRet;
//         }


//         public TreeNode updateFile_v1(String dirId, String nodeId, String username, String fileLoc, ref String retMessage)
//         {
//             TreeNode newFileNd = null;
// 
//             string responseText = "";
// 
//             /*
//             String strFileName = Path.GetFileNameWithoutExtension(fileLoc);
//             String strExtName = Path.GetExtension(fileLoc);
//             String strPath = Path.GetFullPath(fileLoc);
//             String strRnd = DateTime.Now.Ticks.ToString("x");
// 
//             String strNewFile = strPath + strFileName + strRnd + strExtName;
//             */
// 
// 
//             String strNewFile = Path.GetTempPath() + Path.GetFileName(fileLoc);
// 
//             try
//             {
//                 File.Copy(fileLoc, strNewFile, true);
//             }
//             catch (System.Exception ex)
//             {
//                 //File.Delete(strNewFile);
//                 return null;
//             }
//             finally
//             {
// 
//             }
// 
// 
//             byte[] fileBytes = null;
//             try
//             {
//                 FileStream fs = new FileStream(strNewFile, System.IO.FileMode.Open, System.IO.FileAccess.Read);
//                 fileBytes = new byte[fs.Length];
//                 fs.Read(fileBytes, 0, fileBytes.Length);
//                 fs.Close();
//                 fs.Dispose();
//             }
//             catch (System.Exception ex)
//             {
//                 return null;
//             }
//             finally
//             {
//             }
// 
// 
//             HttpRequestClient httpRequestClient = new HttpRequestClient();
//             httpRequestClient.SetFieldValue("type", "2");
//             httpRequestClient.SetFieldValue("dir_id", dirId);
//             httpRequestClient.SetFieldValue("file_id", nodeId);
//             httpRequestClient.SetFieldValue("user_code", username);
//             httpRequestClient.SetFieldValue("uploadfile", strNewFile, "application/octet-stream", fileBytes);
// 
//             bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);
// 
//             //File.Delete(strNewFile);
// 
//             if (!bRet)
//             {
//                 retMessage = responseText;
//                 return newFileNd;
//             }
// 
//             try
//             {
// 
//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
// 
//                 if (res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return newFileNd;
//                 }
// 
//                 if (res.data.GetLength(0) > 0)
//                 {
//                     fileResult result = res.data[0];
// 
//                     if (result.name != null)
//                     {
//                         //newFileNd = new TreeNode(result.template_File_Name);
//                         //newFileNd.Name = result.template_File_Name;
//                         //newFileNd.Tag = "$" + result.template_File_Id;
// 
//                         newFileNd = new TreeNode(result.name);
//                         newFileNd.Name = result.name;
//                         newFileNd.Tag = "$" + result.file_id;
// 
//                         newFileNd.ImageIndex = newFileNd.SelectedImageIndex = iconIndex(result.name);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 retMessage = exp.Message;
//                 MessageBox.Show(exp.Message, "Error");
//             }
//             finally
//             {
//                 // wc.Dispose();
//             }
// 
//             return newFileNd;
//         }

        // String dirId, String nodeId, String username, String fileLoc, ref String retMessage
        public TreeNode updateFile(String nodeId, String username, String fileLoc, ref String retMessage, ref Hashtable hashNodePermission)
        {
            TreeNode newFileNd = null;

            string responseText = "";

            /*
            String strFileName = Path.GetFileNameWithoutExtension(fileLoc);
            String strExtName = Path.GetExtension(fileLoc);
            String strPath = Path.GetFullPath(fileLoc);
            String strRnd = DateTime.Now.Ticks.ToString("x");

            String strNewFile = strPath + strFileName + strRnd + strExtName;
            */


            String strNewFile = Path.GetTempPath() + Path.GetFileName(fileLoc);

            if (File.Exists(strNewFile))
            {
                try
                {
                    File.Delete(strNewFile);
                }
                catch (System.Exception ex)
                {
                    return null;
                }
            }

            try
            {
                File.Copy(fileLoc, strNewFile, true);
            }
            catch (System.Exception ex)
            {
                //File.Delete(strNewFile);
                return null;
            }
            finally
            {

            }


            byte[] fileBytes = null;
            try
            {
                FileStream fs = new FileStream(strNewFile, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                fileBytes = new byte[fs.Length];
                fs.Read(fileBytes, 0, fileBytes.Length);
                fs.Close();
                fs.Dispose();
            }
            catch (System.Exception ex)
            {
                return null;
            }
            finally
            {
            }


            HttpRequestClient httpRequestClient = new HttpRequestClient();

            httpRequestClient.SetFieldValue("serverId", "client_fms_005");
            // httpRequestClient.SetFieldValue("dir_id", dirId);
            httpRequestClient.SetFieldValue("fileId", nodeId);
            httpRequestClient.SetFieldValue("username", username);

            //             httpRequestClient.SetFieldValue("type", "1");
            //             httpRequestClient.SetFieldValue("dir_id", dirId);
            //             httpRequestClient.SetFieldValue("file_id", nodeId);
            //             httpRequestClient.SetFieldValue("user_code", username);

            httpRequestClient.SetFieldValue("uploadfile", strNewFile, "application/octet-stream", fileBytes);

            bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);

            //File.Delete(strNewFile);

            if (!bRet)
            {
                retMessage = responseText;
                return newFileNd;
            }

            try
            {
                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(responseText, typeof(nodesResult));

                if (res == null || res.status == 0) // failure
                {
                    return null;
                }

                //                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
                // 
                //                 if (res.status == 0) // failure
                //                 {
                //                     retMessage = res.message;
                //                     return newFileNd;
                //                 }

                if (res.data != null && res.data.GetLength(0) > 0)
                {
                    nodeInfo result = res.data[0];
                    // fileResult result = res.data[0];

                    if (result.name != null)
                    {
                        //newFileNd = new TreeNode(result.template_File_Name);
                        //newFileNd.Name = result.template_File_Name;
                        //newFileNd.Tag = "$" + result.template_File_Id;

                        TypeNode pt = new TypeNode();

                        pt.id = result.id;
                        pt.name = result.name;
                        pt.parentId = result.parentid;
                        pt.orderId = result.orderid;
                        pt.type = "3";// result.type;
                        pt.isLeaf = 0;

                        newFileNd = new TreeNode(result.name);
                        newFileNd.Name = result.name;
                        newFileNd.Tag = pt;//"$" + result.id;// result.file_id;

                        newFileNd.ImageIndex = newFileNd.SelectedImageIndex = iconIndex(result.name);


                        int[] arrPermVal = new int[16];
                        if (!String.IsNullOrWhiteSpace(result.pmisItem))
                        {
                            String[] arrPmis = result.pmisItem.Split(',');
                            int nVal = 0;

                            foreach (String strItem in arrPmis)
                            {
                                if (int.TryParse(strItem, out nVal))
                                {
                                    if (nVal == 99)
                                        nVal = 0;

                                    if (nVal >= 0 && nVal < arrPermVal.GetLength(0))
                                    {
                                        arrPermVal[nVal] = 1;
                                    }
                                }
                            }
                        }

                        hashNodePermission[result.id] = arrPermVal;

                    }
                }
                else
                {
                    retMessage = "接口返回值异常，请联系系统管理员";
                }

            }
            catch (WebException exp)
            {
                retMessage = exp.Message;
                MessageBox.Show(exp.Message, "Error");
            }
            finally
            {
                // wc.Dispose();
            }

            return newFileNd;
        }



        public TreeNode uploadFile(String dirId, String nodeId, String username, String fileLoc, ref String retMessage, ref Hashtable hashNodePermission)
        {
            TreeNode newFileNd = null;

            string responseText = "";

            /*
            String strFileName = Path.GetFileNameWithoutExtension(fileLoc);
            String strExtName = Path.GetExtension(fileLoc);
            String strPath = Path.GetFullPath(fileLoc);
            String strRnd = DateTime.Now.Ticks.ToString("x");

            String strNewFile = strPath + strFileName + strRnd + strExtName;
            */

            // = fileLoc;
            
            
            String strNewFile = m_commTools.CopyTmpDocFile(fileLoc);

            if (String.IsNullOrWhiteSpace(strNewFile))
            {
                MessageBox.Show("创建临时文件失败，请确保磁盘空间");
                return null;
            }


            /*
            String strNewFile = Path.GetTempPath() + Path.GetFileName(fileLoc);

            if (File.Exists(strNewFile))
            {
                try
                {
                    File.Delete(strNewFile);
                }
                catch (System.Exception ex)
                {
                    return null;
                }
            }

            try
            {
                File.Copy(fileLoc, strNewFile, true);
            }
            catch (System.Exception ex)
            {
                //File.Delete(strNewFile);
                return null;
            }
            finally
            {

            }
            */

            


            byte[] fileBytes = null;
            try
            {
                FileStream fs = new FileStream(strNewFile, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                fileBytes = new byte[fs.Length];
                fs.Read(fileBytes, 0, fileBytes.Length);
                fs.Close();
                fs.Dispose();
            }
            catch (System.Exception ex)
            {
                return null;
            }
            finally
            {
            }


            HttpRequestClient httpRequestClient = new HttpRequestClient();

            httpRequestClient.SetFieldValue("serverId", "client_fms_003");
            httpRequestClient.SetFieldValue("dir_id", dirId);
            httpRequestClient.SetFieldValue("username", username);

//             httpRequestClient.SetFieldValue("type", "1");
//             httpRequestClient.SetFieldValue("dir_id", dirId);
//             httpRequestClient.SetFieldValue("file_id", nodeId);
//             httpRequestClient.SetFieldValue("user_code", username);

            httpRequestClient.SetFieldValue("uploadfile", fileLoc, "application/octet-stream", fileBytes);

            bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);

            //File.Delete(strNewFile);

            if (!bRet)
            {
                retMessage = responseText;
                return newFileNd;
            }

            try
            {
                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(responseText, typeof(nodesResult));

                if (res == null || res.status == 0) // failure
                {
                    return null;
                }

//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
// 
//                 if (res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return newFileNd;
//                 }

                // nodeInfo result = res.data[0];
                // fileResult result = res.data[0];

                //newFileNd = new TreeNode(result.template_File_Name);
                //newFileNd.Name = result.template_File_Name;
                //newFileNd.Tag = "$" + result.template_File_Id;
                if(res.data != null && res.data.GetLength(0) > 0)
                {
                    nodeInfo ndInfo = res.data[0];

                    newFileNd = new TreeNode(ndInfo.name);
                    newFileNd.Name = ndInfo.name;

                    TypeNode pt = new TypeNode();

                    pt.id = ndInfo.id;
                    pt.name = ndInfo.name;
                    pt.parentId = ndInfo.parentid;
                    pt.orderId = ndInfo.orderid;
                    pt.type = "3";// ndInfo.type;
                    pt.isLeaf = 0;

                    newFileNd.Tag = pt;

                    newFileNd.ImageIndex = newFileNd.SelectedImageIndex = iconIndex(ndInfo.name);


                    int[] arrPermVal = new int[16];
                    if (!String.IsNullOrWhiteSpace(ndInfo.pmisItem))
                    {
                        String[] arrPmis = ndInfo.pmisItem.Split(',');
                        int nVal = 0;

                        foreach (String strItem in arrPmis)
                        {
                            if (int.TryParse(strItem, out nVal))
                            {
                                if (nVal == 99)
                                    nVal = 0;

                                if (nVal >= 0 && nVal < arrPermVal.GetLength(0))
                                {
                                    arrPermVal[nVal] = 1;
                                }
                            }
                        }
                    }

                    hashNodePermission[ndInfo.id] = arrPermVal;
                }
                else
                {
                    retMessage = "返回数据异常，请联系系统管理员";
                }

            }
            catch (WebException exp)
            {
                retMessage = exp.Message;
                MessageBox.Show(exp.Message, "错误");
            }
            finally
            {
                // wc.Dispose();
            }

            return newFileNd;
        }


        public TreeNode uploadFile_v1(String dirId, String nodeId, String username, String fileLoc, ref String retMessage)
        {
            TreeNode newFileNd = null;

            string responseText = "";

            /*
            String strFileName = Path.GetFileNameWithoutExtension(fileLoc);
            String strExtName = Path.GetExtension(fileLoc);
            String strPath = Path.GetFullPath(fileLoc);
            String strRnd = DateTime.Now.Ticks.ToString("x");

            String strNewFile = strPath + strFileName + strRnd + strExtName;
            */


            String strNewFile = Path.GetTempPath() + Path.GetFileName(fileLoc);

            if (File.Exists(strNewFile))
            {
                try
                {
                    File.Delete(strNewFile);
                }
                catch (System.Exception ex)
                {
                    return null;
                }
            }

            try
            {
                File.Copy(fileLoc, strNewFile,true);
            }
            catch (System.Exception ex)
            {
                //File.Delete(strNewFile);
                return null;
            }
            finally
            {

            }


            byte[] fileBytes = null;
            try
            {
                FileStream fs = new FileStream(strNewFile, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                fileBytes = new byte[fs.Length];
                fs.Read(fileBytes, 0, fileBytes.Length);
                fs.Close();
                fs.Dispose();
            }
            catch (System.Exception ex)
            {
                return null;
            }
            finally
            {
            }


            HttpRequestClient httpRequestClient = new HttpRequestClient();
            httpRequestClient.SetFieldValue("type", "1");
            httpRequestClient.SetFieldValue("dir_id", dirId);
            httpRequestClient.SetFieldValue("file_id", nodeId);
            httpRequestClient.SetFieldValue("user_code", username);
            httpRequestClient.SetFieldValue("uploadfile", strNewFile, "application/octet-stream", fileBytes);

            bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);

            //File.Delete(strNewFile);

            if (!bRet)
            {
                retMessage = responseText;
                return newFileNd;
            }

            try
            {

                fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));

                if (res == null || res.status == 0) // failure
                {
                    if (res != null)
                    {
                        retMessage = res.message;
                    }

                    return newFileNd;
                }

                if (res.data != null && res.data.GetLength(0) > 0)
                {
                    fileResult result = res.data[0];

                    if (result.name != null)
                    {
                        //newFileNd = new TreeNode(result.template_File_Name);
                        //newFileNd.Name = result.template_File_Name;
                        //newFileNd.Tag = "$" + result.template_File_Id;

                        newFileNd = new TreeNode(result.name);
                        newFileNd.Name = result.name;
                        newFileNd.Tag = "$" + result.file_id;

                        newFileNd.ImageIndex = newFileNd.SelectedImageIndex = iconIndex(result.name);
                    }
                }

            }
            catch (WebException exp)
            {
                retMessage = exp.Message;
                MessageBox.Show(exp.Message, "Error");
            }
            finally
            {
                // wc.Dispose();
            }

            return newFileNd;
        }





/*
        public TreeNode uploadFile_v1(String nodeId, String username,String fileLoc, ref String retMessage)
        {
            TreeNode newFileNd = null;

            string responseText = "";
            FileStream fs = new FileStream(fileLoc, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            byte[] fileBytes = new byte[fs.Length];
            fs.Read(fileBytes, 0, fileBytes.Length);
            fs.Close(); 
            fs.Dispose();

            HttpRequestClient httpRequestClient = new HttpRequestClient();
            httpRequestClient.SetFieldValue("nodeid", nodeId);
            httpRequestClient.SetFieldValue("createuser", username);
            httpRequestClient.SetFieldValue("uploadfile", fileLoc, "application/octet-stream", fileBytes);
            
            bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);

            if (!bRet)
            {
                retMessage = responseText;
                return newFileNd;
            }

            try
            {
                
                fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));

                if (res.value != 0)
                {
                    retMessage = res.message;
                    return newFileNd;
                }

                if (res.result.GetLength(0) > 0)
                {
                    fileResult result = res.result[0];

                    if (result.template_File_Name != null)
                    {
                        newFileNd = new TreeNode(result.template_File_Name);
                        newFileNd.Name = result.template_File_Name;
                        newFileNd.Tag = "$" + result.template_File_Id;
                        newFileNd.ImageIndex = newFileNd.SelectedImageIndex = 0;
                    }
                }

            }
            catch (WebException exp)
            {
                retMessage = exp.Message;
                MessageBox.Show(exp.Message, "Error");
            }
            finally
            {
                // wc.Dispose();
            }

            return newFileNd;
        }*/


        public int removeFile_v2(String userName, String fileId, ref String retMessage)
        {
            NameValueCollection form = new NameValueCollection();
            int nLibType = -1;

            form.Add("username", userName);
            form.Add("serverId", "client_fms_004");
            form.Add("fileId", fileId);


            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
	            byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
	            String jsonText = Encoding.UTF8.GetString(responseArray);
	
	            nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));
	
	            if (res == null || res.status == 0) // failure
	            {
	                return -1;
	            }
            }
            catch (System.Exception ex)
            {
            	
            }
            finally
            {
            }




            //TreeNode newFileNd = null;
            int nRet = 0;
            string responseText = "";

            byte[] fileBytes = new Byte[] {1 };


            HttpRequestClient httpRequestClient = new HttpRequestClient();
//             httpRequestClient.SetFieldValue("type", "0");
//             httpRequestClient.SetFieldValue("dir_id", dirId);
//             httpRequestClient.SetFieldValue("file_id", nodeId);
//             httpRequestClient.SetFieldValue("user_code", username);
//             httpRequestClient.SetFieldValue("uploadfile", "xxx", "application/octet-stream", fileBytes);

            bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);

            //File.Delete(strNewFile);

            if (!bRet)
            {
                nRet = -1;
                retMessage = responseText;
                return nRet;
            }

            try
            {
                fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));

                if (res.status == 0) // failure
                {
                    nRet = -1;
                    retMessage = res.message;
                    return nRet;
                }
            }
            catch (WebException exp)
            {
                retMessage = exp.Message;
                MessageBox.Show(exp.Message, "Error");
            }
            finally
            {
                // wc.Dispose();
            }

            return nRet;
        }


        public int removeFile(String nodeId, String username, ref String retMessage)
        {
            NameValueCollection form = new NameValueCollection();

            form.Add("username", username);
            form.Add("serverId", "client_fms_004");
            form.Add("fileId", nodeId);

            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));

                if (res == null || res.status == 0) // failure
                {
                    if (res != null)
                    {
                        retMessage = res.message;
                    }
                    return -1;
                }

            }
            catch (WebException exp)
            {
                // MessageBox.Show(exp.Message, "Error");
                retMessage = exp.Message;
                wc.Dispose();
                return -1;
            }
            finally
            {
                //wc.Dispose();
            }

            return 0;
        }


//         public int removeFile_v1p2(String dirId, String nodeId, String username, ref String retMessage)
//         {
//             //TreeNode newFileNd = null;
//             int nRet = 0;
//             string responseText = "";
// 
//             byte[] fileBytes = new Byte[] { 1 };
// 
// 
//             HttpRequestClient httpRequestClient = new HttpRequestClient();
//             httpRequestClient.SetFieldValue("type", "0");
//             httpRequestClient.SetFieldValue("dir_id", dirId);
//             httpRequestClient.SetFieldValue("file_id", nodeId);
//             httpRequestClient.SetFieldValue("user_code", username);
//             httpRequestClient.SetFieldValue("uploadfile", "xxx", "application/octet-stream", fileBytes);
// 
//             bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);
// 
//             //File.Delete(strNewFile);
// 
//             if (!bRet)
//             {
//                 nRet = -1;
//                 retMessage = responseText;
//                 return nRet;
//             }
// 
//             try
//             {
//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
// 
//                 if (res.status == 0) // failure
//                 {
//                     nRet = -1;
//                     retMessage = res.message;
//                     return nRet;
//                 }
//             }
//             catch (WebException exp)
//             {
//                 retMessage = exp.Message;
//                 MessageBox.Show(exp.Message, "Error");
//             }
//             finally
//             {
//                 // wc.Dispose();
//             }
// 
//             return nRet;
//         }

        // removeFile
//         public int removeFile_v1(String username, String fileId, ref String strRetMessage)
//         {
//             NameValueCollection form = new NameValueCollection();
// 
//             form.Add("username", username);
//             form.Add("fileId", fileId);
// 
// 
//             WebClient wc = new WebClient();
// 
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgRemoveFileUrl, "POST", form);
//                 String responseText = Encoding.UTF8.GetString(responseArray);
// 
//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
//                 nRet = (res.status == 1? 0:-1);
//                 strRetMessage = res.message;
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 nRet = -1;
//                 strRetMessage = exp.Message;
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return nRet;
//         }

        // removeType
//         public int removeType_v1(String typeId, ref String strRetMessage)
//         {
//             NameValueCollection form = new NameValueCollection();
// 
//             form.Add("id", typeId);
// 
// 
//             WebClient wc = new WebClient();
// 
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgRemoveTypeUrl, "POST", form);
//                 String responseText = Encoding.UTF8.GetString(responseArray);
// 
//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
//                 nRet = (res.status == 1? 0: -1);
//                 strRetMessage = res.message;
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 nRet = -1;
//                 strRetMessage = exp.Message;
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return nRet;
//         }


        public Hashtable getVstoPermission(String strUserName)
        {
            Hashtable hashVstoPermission = new Hashtable();

            NameValueCollection form = new NameValueCollection();

            form.Add("username", strUserName);
            form.Add("serverId", "client_fms_009");

            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                vstoNodesResult res = (vstoNodesResult)JsonConvert.DeserializeObject(jsonText, typeof(vstoNodesResult));

                if (res != null || res.status != 0) // failure
                {
                    if (res.data != null)
                    {
                        foreach (vstoNodeInfo vstoNdInfo in res.data)
                        {
                            if (!hashVstoPermission.Contains(vstoNdInfo.controlname))
                            {
                                hashVstoPermission.Add(vstoNdInfo.controlname, vstoNdInfo.value);
                            }
                        }
                    }
                }

            }
            catch (WebException exp)
            {
                
            }
            finally
            {
                //wc.Dispose();
            }

            return hashVstoPermission;

        }


//         public Hashtable getVstoPermission_v1(String strUserName)
//         {
//             Hashtable hashVstoPermission = new Hashtable();
// 
//             NameValueCollection form = new NameValueCollection();
//             form.Add("username", strUserName);
// 
//             WebClient wc = new WebClient();
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             String strRetMessage = "";
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgVstoPermissionUrl, "POST", form);
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 /* 
//                 [{\"SHEETID\":\"241CB27E-66DC-45F7-B178-D385599C0ECB\",\"SHOWNAME\":\"连续格式刷\",\"CONTROLNAME\":\"btnPasteFormat\",\"CONTROLTYPE\":\"RibbonButton\",\"VALUE\":\"0\",\"PARENTID\":\"902B2C92-630D-4B72-B4D8-41BC3DCC0400\"}, ... ]
//                 */
// 
//                 JArray jsonArr = JArray.Parse(jsonText);
//                 String strName = "";
//                 int nVal = 0;
//                 
//                 foreach (JObject jobj in jsonArr)
//                 {
//                     strName = (String)jobj["CONTROLNAME"];
//                     nVal = (int)jobj["VALUE"];
//                     
//                     if (!hashVstoPermission.Contains(strName))
//                     {
//                         hashVstoPermission.Add(strName, nVal);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 // wc.Dispose();
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return hashVstoPermission;
//         }


//         // 不可用, 2016-06-07
//         // 
//         public Hashtable geteFileLibPermission(String strUserName)
//         {
//             Hashtable hashFileLibPermission = new Hashtable();
// 
//             NameValueCollection form = new NameValueCollection();
//             form.Add("username", strUserName);
// 
//             WebClient wc = new WebClient();
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             String strRetMessage = "";
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgFileLibPermissionUrl, "POST", form);
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 /* 
//                 [{"RESOURCE_ID":"D218230D-4BE6-9B17-F5E0-39662BC14632","OPER_ID":"0,100"}]
// 
//                 [{"RESOURCE_ID":"D218230D-4BE6-9B17-F5E0-39662BC14632","OPER_ID":"0,100"}]
//                 resource_id:文件夹ID（projecttree中的ID，filerecord中的NODE_ID）
//                 oper_id:
//                 100	访问
//                 4	删除
//                 1	只读
//                 2	新增
//                 3	修改
// 
//                 * */
// 
//                 JArray jsonArr = JArray.Parse(jsonText);
//                 String strResourceId = "", strOperPermission = "";
//                 int nVal = 0;
// 
//                 foreach (JObject jobj in jsonArr)
//                 {
//                     strResourceId = (String)jobj["RESOURCE_ID"];
//                     strOperPermission = (String)jobj["OPER_ID"];
// 
//                     if (!hashFileLibPermission.Contains(strResourceId))
//                     {
//                         hashFileLibPermission.Add(strResourceId, strOperPermission);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 // wc.Dispose();
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return hashFileLibPermission;
// 
//         }



//         public void Test()
//         {
//             NameValueCollection form = new NameValueCollection();
// 
//             form.Add("username", "lidong");
// 
// 
//             WebClient wc = new WebClient();
// 
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             String strRetMessage = "";
// 
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgVstoPermissionUrl, "POST", form);
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 /* 
//                 [{\"SHEETID\":\"241CB27E-66DC-45F7-B178-D385599C0ECB\",\"SHOWNAME\":\"连续格式刷\",\"CONTROLNAME\":\"btnPasteFormat\",\"CONTROLTYPE\":\"RibbonButton\",\"VALUE\":\"0\",\"PARENTID\":\"902B2C92-630D-4B72-B4D8-41BC3DCC0400\"}, ... ]
//                 */
//                 
//                 JArray jsonArr = JArray.Parse(jsonText);
// 
//                 foreach (JObject jobj in jsonArr)
//                 {
//                     TypeNode pt = new TypeNode();
// 
//                     pt.id = (String)jobj["ID"];
//                     pt.name = (String)jobj["NAME"];
//                     pt.parentId = (String)jobj["PARENTID"];
//                     pt.orderId = (String)jobj["ORDERID"];
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 // wc.Dispose();
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return ;
//         }


        public TreeNode updateType(String dirId, String username, String newName, ref String retMessage)
        {
            TreeNode newFolderNd = null;

            NameValueCollection form = new NameValueCollection();

            form.Add("username", username);
            form.Add("serverId", "client_fms_008");
            form.Add("dir_new_name", newName);
            form.Add("dir_id", dirId);


            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;

            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));

                if (res == null || res.status == 0) // failure
                {
                    if (res != null)
                    {
                        retMessage = res.message;
                    }
                    return newFolderNd;
                }

                // update permission
                newFolderNd = new TreeNode(newName);
                newFolderNd.Name = newName;

//                 nodeInfo ndInfo = res.data[0];
// 
//                 TypeNode pt = new TypeNode();
//                 pt.id = ndInfo.id;
//                 pt.name = ndInfo.name;
//                 pt.parentId = ndInfo.parentid;
//                 pt.orderId = ndInfo.orderid;
//                 pt.type = ndInfo.type;
//                 pt.isLeaf = 0;
// 
//                 newFolderNd.Tag = pt;// "#" + dirId;

                newFolderNd.ImageIndex = newFolderNd.SelectedImageIndex = 13;
            }
            catch (WebException exp)
            {
                // MessageBox.Show(exp.Message, "Error");
                retMessage = exp.Message;
                wc.Dispose();
                return newFolderNd;
            }
            finally
            {
                //wc.Dispose();
            }

            return newFolderNd;
        }


//         public TreeNode updateType_v1(String dirId, String nodeId, String username, String newName, ref String retMessage)
//         {
//             TreeNode newFolderNd = null;
// 
//             //string responseText = "";
// 
//             NameValueCollection form = new NameValueCollection();
//             form.Add("type", "2");
//             form.Add("user_code", username);
//             form.Add("dir_id", nodeId);
//             form.Add("parent_dir_id", dirId);
//             form.Add("dir_new_name", newName);
// 
// 
//             WebClient wc = new WebClient();
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             String strRetMessage = "";
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgDirOperateUrl, "POST", form);
// 
//                 if (responseArray == null)
//                 {
//                     return null;
//                 }
// 
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 dirReturnResult res = (dirReturnResult)JsonConvert.DeserializeObject(jsonText, typeof(dirReturnResult));
// 
//                 if (res == null || res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return newFolderNd;
//                 }
// 
//                 if (res.data != null)
//                 {
//                     // update permission
//                     newFolderNd = new TreeNode(res.data.dir_name);
//                     newFolderNd.Name = res.data.dir_name;
//                     newFolderNd.Tag = "#" + res.data.dir_id;
// 
//                     newFolderNd.ImageIndex = newFolderNd.SelectedImageIndex = 13;
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 // wc.Dispose();
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return newFolderNd;
//         }


        public int removeType(String parentDirID, String nodeId, String username,ref String retMessage)
        {
            NameValueCollection form = new NameValueCollection();

            form.Add("username", username);
            form.Add("serverId", "client_fms_007");
            form.Add("dir_id", nodeId);


            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));

                if (res == null || res.status == 0) // failure
                {
                    if (res != null)
                    {
                        retMessage = res.message;
                    }
                    return -1;
                }

            }
            catch (WebException exp)
            {
                // MessageBox.Show(exp.Message, "Error");
                retMessage = exp.Message;
                wc.Dispose();
                return -1;
            }
            finally
            {
                //wc.Dispose();
            }

            return 0;
        }


//         public int removeType_v1(String parentDirID, String nodeId, String username, String newName, ref String retMessage)
//         {
//             // int nRet = -1;
// 
//             //string responseText = "";
// 
//             NameValueCollection form = new NameValueCollection();
//             form.Add("type", "0");
//             form.Add("user_code", username);
//             form.Add("dir_id", nodeId);
//             form.Add("parent_dir_id", parentDirID);
//             form.Add("dir_new_name", newName);
// 
// 
//             WebClient wc = new WebClient();
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             String strRetMessage = "";
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgDirOperateUrl, "POST", form);
// 
//                 if (responseArray == null)
//                 {
//                     return -1;
//                 }
// 
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 dirReturnResult res = (dirReturnResult)JsonConvert.DeserializeObject(jsonText, typeof(dirReturnResult));
// 
//                 if (res == null || res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return -1;
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 // wc.Dispose();
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
// 
//             return 0;
//         }


        public TreeNode AddType(String dirId, String username, String newName, String parentType,
                                String libType,ref String retMessage, ref Hashtable hashNodePermission)
        {
            TreeNode newFolderNd = null;

            NameValueCollection form = new NameValueCollection();

            form.Add("username", username);
            form.Add("serverId", "client_fms_006");
            form.Add("nodename", newName);
            form.Add("parentid", dirId);
            form.Add("parent_type", parentType); // 父目录类型 （1板块2文件夹）
            form.Add("lib_type", libType); // 库类别（0个人库 1公共库）


            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodeResult res = (nodeResult)JsonConvert.DeserializeObject(jsonText, typeof(nodeResult));

                if (res == null || res.status == 0) // failure
                {
                    if (res != null)
                    {
                        retMessage = res.message;
                    }

                    return newFolderNd;
                }

                TypeNode pt = new TypeNode();

                pt.id = res.id;
                pt.name = res.name;
                pt.parentId = res.parentid;
                pt.orderId = res.orderid;
                pt.type = "2";// res.type;
                pt.isLeaf = 0;

                // update permission
                newFolderNd = new TreeNode(res.name);
                newFolderNd.Name = res.name;
                newFolderNd.Tag = pt; // "#" + res.id;

                newFolderNd.ImageIndex = newFolderNd.SelectedImageIndex = 13;

                int[] arrPermVal = new int[16];
                if (!String.IsNullOrWhiteSpace(res.pmisItem))
                {
                    String[] arrPmis = res.pmisItem.Split(',');
                    int nVal = 0;

                    foreach (String strItem in arrPmis)
                    {
                        if (int.TryParse(strItem, out nVal))
                        {
                            if (nVal == 99)
                                nVal = 0;

                            if (nVal >= 0 && nVal < arrPermVal.GetLength(0))
                            {
                                arrPermVal[nVal] = 1;
                            }
                        }
                    }
                }

                hashNodePermission[res.id] = arrPermVal;

            }
            catch (WebException exp)
            {
                // MessageBox.Show(exp.Message, "Error");
                retMessage = exp.Message;
                wc.Dispose();
                return newFolderNd;
            }
            finally
            {
                //wc.Dispose();
            }

            return newFolderNd;
        }


//         public TreeNode AddType_v1p2(String dirId, String nodeId, String username,String newName, ref String retMessage)
//         {
//             TreeNode newFolderNd = null;
// 
//             //string responseText = "";
// 
//             NameValueCollection form = new NameValueCollection();
//             form.Add("type", "1");
//             form.Add("user_code", username);
//             form.Add("dir_id", "");
//             form.Add("parent_dir_id", dirId);
//             form.Add("dir_new_name", newName);
// 
// 
//             WebClient wc = new WebClient();
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             String strRetMessage = "";
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgDirOperateUrl, "POST", form);
// 
//                 if (responseArray == null )
//                 {
//                     return null;
//                 }
// 
//                 String jsonText = Encoding.UTF8.GetString(responseArray);
// 
//                 dirReturnResult res = (dirReturnResult)JsonConvert.DeserializeObject(jsonText, typeof(dirReturnResult));
// 
//                 if (res == null || res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return newFolderNd;
//                 }
// 
// 
//                 if (res.data != null )
//                 {
//                     // update permission
//                     newFolderNd = new TreeNode(res.data.dir_name);
//                     newFolderNd.Name = res.data.dir_name;
//                     newFolderNd.Tag = "#" + res.data.dir_id;
// 
//                     newFolderNd.ImageIndex = newFolderNd.SelectedImageIndex = 13;
//                 }
// 
// /*
// 
//                 JArray jsonArr = JArray.Parse(jsonText);
//                 String strResourceId = "", strOperPermission = "";
//                 int nVal = 0;
// 
//                 foreach (JObject jobj in jsonArr)
//                 {
//                     strResourceId = (String)jobj["RESOURCE_ID"];
//                     strOperPermission = (String)jobj["OPER_ID"];
// 
//                     if (!hashFileLibPermission.Contains(strResourceId))
//                     {
//                         hashFileLibPermission.Add(strResourceId, strOperPermission);
//                     }
//                 }*/
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 // wc.Dispose();
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             //File.Delete(strNewFile);
// 
//             /*if (!bRet)
//             {
//                 retMessage = responseText;
//                 return newFolderNd;
//             }
// 
//             try
//             {
// 
//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
// 
//                 if (res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return newFolderNd;
//                 }
// 
//                 if (res.data.GetLength(0) > 0)
//                 {
//                     fileResult result = res.data[0];
// 
//                     if (result.name != null)
//                     {
//                         //newFileNd = new TreeNode(result.template_File_Name);
//                         //newFileNd.Name = result.template_File_Name;
//                         //newFileNd.Tag = "$" + result.template_File_Id;
// 
//                         newFolderNd = new TreeNode(result.name);
//                         newFolderNd.Name = result.name;
//                         newFolderNd.Tag = "$" + result.file_id;
// 
//                         newFolderNd.ImageIndex = newFolderNd.SelectedImageIndex = iconIndex(result.name);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 retMessage = exp.Message;
//                 MessageBox.Show(exp.Message, "Error");
//             }
//             finally
//             {
//                 // wc.Dispose();
//             }
// */
// 
//             return newFolderNd;
//         }



//         public TreeNode AddType_v1(String dirId, String nodeId, String username, String fileLoc, ref String retMessage)
//         {
//             TreeNode newFileNd = null;
// 
//             string responseText = "";
// 
//             /*
//             String strFileName = Path.GetFileNameWithoutExtension(fileLoc);
//             String strExtName = Path.GetExtension(fileLoc);
//             String strPath = Path.GetFullPath(fileLoc);
//             String strRnd = DateTime.Now.Ticks.ToString("x");
// 
//             String strNewFile = strPath + strFileName + strRnd + strExtName;
//             */
// 
// 
//             String strNewFile = Path.GetTempPath() + Path.GetFileName(fileLoc);
// 
//             try
//             {
//                 File.Copy(fileLoc, strNewFile, true);
//             }
//             catch (System.Exception ex)
//             {
//                 //File.Delete(strNewFile);
//                 return null;
//             }
//             finally
//             {
// 
//             }
// 
// 
//             byte[] fileBytes = null;
//             try
//             {
//                 FileStream fs = new FileStream(strNewFile, System.IO.FileMode.Open, System.IO.FileAccess.Read);
//                 fileBytes = new byte[fs.Length];
//                 fs.Read(fileBytes, 0, fileBytes.Length);
//                 fs.Close();
//                 fs.Dispose();
//             }
//             catch (System.Exception ex)
//             {
//                 return null;
//             }
//             finally
//             {
//             }
// 
// 
//             HttpRequestClient httpRequestClient = new HttpRequestClient();
//             httpRequestClient.SetFieldValue("type", "1");
//             httpRequestClient.SetFieldValue("dir_id", dirId);
//             httpRequestClient.SetFieldValue("file_id", nodeId);
//             httpRequestClient.SetFieldValue("user_code", username);
//             httpRequestClient.SetFieldValue("uploadfile", strNewFile, "application/octet-stream", fileBytes);
// 
//             bool bRet = httpRequestClient.Upload(m_cfgUploadFileUrl, out responseText);
// 
//             //File.Delete(strNewFile);
// 
//             if (!bRet)
//             {
//                 retMessage = responseText;
//                 return newFileNd;
//             }
// 
//             try
//             {
// 
//                 fileReturnResult res = (fileReturnResult)JsonConvert.DeserializeObject(responseText, typeof(fileReturnResult));
// 
//                 if (res.status == 0) // failure
//                 {
//                     retMessage = res.message;
//                     return newFileNd;
//                 }
// 
//                 if (res.data.GetLength(0) > 0)
//                 {
//                     fileResult result = res.data[0];
// 
//                     if (result.name != null)
//                     {
//                         //newFileNd = new TreeNode(result.template_File_Name);
//                         //newFileNd.Name = result.template_File_Name;
//                         //newFileNd.Tag = "$" + result.template_File_Id;
// 
//                         newFileNd = new TreeNode(result.name);
//                         newFileNd.Name = result.name;
//                         newFileNd.Tag = "$" + result.file_id;
// 
//                         newFileNd.ImageIndex = newFileNd.SelectedImageIndex = iconIndex(result.name);
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 retMessage = exp.Message;
//                 MessageBox.Show(exp.Message, "Error");
//             }
//             finally
//             {
//                 // wc.Dispose();
//             }
// 
//             return newFileNd;
//         }



        // removeType
//         public TreeNode AddType_v1(String orderid, String nodename, String parentid, ref String strRetMessage)
//         {
//             TreeNode newNd = null;
// 
//             NameValueCollection form = new NameValueCollection();
// 
//             form.Add("orderid", orderid);
//             form.Add("nodename", nodename);
//             form.Add("parentid", parentid);
// 
//             WebClient wc = new WebClient();
// 
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgAddNewTypeUrl, "POST", form);
//                 String responseText = Encoding.UTF8.GetString(responseArray);
// 
//                 typeReturnResult res = (typeReturnResult)JsonConvert.DeserializeObject(responseText, typeof(typeReturnResult));
//                 nRet = res.value;
//                 strRetMessage = res.message;
// 
//                 if (nRet == 0)
//                 {
//                     newNd = new TreeNode(nodename);
//                     newNd.Name = nodename;
//                     newNd.Tag = "#" + res.result[0].id;
//                     newNd.ImageIndex = newNd.SelectedImageIndex = 13;
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 nRet = -1;
//                 strRetMessage = exp.Message;
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return newNd;
//         }


//         public Hashtable getFilePermissions(String strUserLoginName, ref String strRetMessage)
//         {
//             Hashtable hashFilePermission = new Hashtable();
// 
//             NameValueCollection form = new NameValueCollection();
// 
//             form.Add("username", strUserLoginName);
// 
//             WebClient wc = new WebClient();
// 
//             wc.Encoding = Encoding.UTF8;
// 
//             //wc.Headers.Add("Cookie", "USERID=21A723860BBB0887790C9140FC17AAE9");
// 
//             int nRet = 0;
// 
//             try
//             {
//                 byte[] responseArray = wc.UploadValues(m_cfgFilePermissionUrl, "POST", form);
//                 String responseText = Encoding.UTF8.GetString(responseArray);
// 
//                 typeFilePermissionReturnResult res = (typeFilePermissionReturnResult)JsonConvert.DeserializeObject(responseText, typeof(typeFilePermissionReturnResult));
//                 nRet = res.value;
//                 strRetMessage = res.message;
// 
//                 int[] values = null;
//                 String strID = "";
// 
//                 if (nRet > 0)
//                 {
//                     if (res.result != null && res.result.GetLength(0) > 0)
//                     {
//                         foreach(typeFilePermission filePermission in res.result)
//                         {
//                             strID = filePermission.FOLDERID;
// 
//                             if (hashFilePermission.Contains(strID))
//                             {
//                                 values = (int[])hashFilePermission[strID];
//                                 if(filePermission.VALUE >= 0 && filePermission.VALUE < values.GetLength(0))
//                                 {
//                                     values[filePermission.VALUE] = 1;
//                                 }
//                             }
//                             else
//                             {
//                                 values = new int[16];
//                                 
//                                 if (filePermission.VALUE >= 0 && filePermission.VALUE < values.GetLength(0))
//                                 {
//                                     values[filePermission.VALUE] = 1;
//                                 }
// 
//                                 hashFilePermission[strID] = values;
// 
//                             }
//                         }
//                     }
//                 }
// 
//             }
//             catch (WebException exp)
//             {
//                 MessageBox.Show(exp.Message, "Error");
//                 nRet = -1;
//                 strRetMessage = exp.Message;
//             }
//             finally
//             {
//                 wc.Dispose();
//             }
// 
//             return hashFilePermission;
//         }


        [DataContract()]
        public class nodeResult
        {
            // {"status":"success","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String message;
            [DataMember(Order = 1)]
            public int status;
            [DataMember(Order = 2)]
            public String id;
            [DataMember(Order = 3)]
            public String name;
            [DataMember(Order = 4)]
            public String parentid;
            [DataMember(Order = 5)]
            public String pmisItem;
            [DataMember(Order = 6)]
            public String orderid;
            [DataMember(Order = 7)]
            public String type;
        }



        [DataContract()]
        public class nodesResult
        {
            // {"status":"success","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String message; 
            [DataMember(Order = 1)]
            public int status;
            [DataMember(Order = 2)]
            public nodeInfo[] data;
        }


        [DataContract()]
        public class nodeInfo
        {
            // {\"id\":\"7239e4be-4b57-4b95-b398-c2ce0cfcf1d1\",\"name\":\"权限测试\",\"parentid\":null,\"pmisItem\":\"01,02,03,04,05,06,07,08,09,99\",\"orderid\":\"151\",\"type\":\"0\"},
            [DataMember(Order = 0)]
            public String id;
            [DataMember(Order = 1)]
            public String name;
            [DataMember(Order = 2)]
            public String parentid;
            [DataMember(Order = 3)]
            public String pmisItem;
            [DataMember(Order = 4)]
            public String orderid;
            [DataMember(Order = 5)]
            public String type;
            [DataMember(Order = 6)]
            public String lib_type;
            [DataMember(Order = 7)]
            public Int32 isLeaf;
        }



        [DataContract()]
        public class vstoNodesResult
        {
            // {"status":"success","message":"登录成功"} 

            [DataMember(Order = 0)]
            public String message;
            [DataMember(Order = 1)]
            public int status;
            [DataMember(Order = 2)]
            public vstoNodeInfo[] data;
        }



        [DataContract()]
        public class vstoNodeInfo
        {
            [DataMember(Order = 0)]
            public String sheetid;
            [DataMember(Order = 1)]
            public String showname;
            [DataMember(Order = 2)]
            public String controlname;
            [DataMember(Order = 3)]
            public String controltype;
            [DataMember(Order = 4)]
            public int value;
            [DataMember(Order = 5)]
            public String parentid;
        }


        public TreeNode getSubNodes(String userName, String fileParam, String strObjId, String strObjType, ref Hashtable hashNodePermission)
        {
            TreeNode rootNd = new TreeNode("root");

            Hashtable hashCommonLibNodes = new Hashtable();
            Hashtable hashPrivLibNodes = new Hashtable();

            Hashtable hashFileNodes = new Hashtable();


            NameValueCollection form = new NameValueCollection();
            // int nLibType = -1;
            String strLibType = "";

            form.Add("username", userName);
            form.Add("serverId", "client_fms_001");
            form.Add("objectId", strObjId);
            form.Add("objectType", strObjType);

            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));

                if (res == null || res.status == 0 || res.data == null) // failure
                {
                    return null;
                }

                foreach (nodeInfo ndInfo in res.data)
                {
                    TypeNode pt = new TypeNode();

                    pt.id = ndInfo.id;
                    pt.name = ndInfo.name;
                    pt.parentId = ndInfo.parentid;
                    pt.orderId = ndInfo.orderid;
                    pt.type = ndInfo.type;
                    pt.isLeaf = ndInfo.isLeaf;

                    strLibType = ndInfo.lib_type;

                    TreeNode typeNd = new TreeNode(pt.name);
                    typeNd.Name = pt.name;
                    typeNd.Tag = pt;

                    typeNd.ImageIndex = typeNd.SelectedImageIndex = 13;


                    if (strLibType == null)
                    {
                        hashFileNodes.Add(pt.id, typeNd);
                    }
                    else
                    {
                        if (strLibType.Equals("1"))
                        {
                            hashCommonLibNodes.Add(pt.id, typeNd);
                        }
                        else // 1
                        {
                            hashPrivLibNodes.Add(pt.id, typeNd);
                        }
                    }

                    // parse permission and insert into hash
                    // 
                    int[] arrPermVal = new int[16];
                    if (!String.IsNullOrWhiteSpace(ndInfo.pmisItem))
                    {
                        String[] arrPmis = ndInfo.pmisItem.Split(',');
                        int nVal = 0;

                        foreach (String strItem in arrPmis)
                        {
                            if (int.TryParse(strItem, out nVal))
                            {
                                if (nVal == 99)
                                    nVal = 0;

                                if (nVal >= 0 && nVal < arrPermVal.GetLength(0))
                                {
                                    arrPermVal[nVal] = 1;
                                }
                            }
                        }
                    }

                    hashNodePermission[ndInfo.id] = arrPermVal;
                }

            }
            catch (WebException exp)
            {
                MessageBox.Show(exp.Message, "Error");
                wc.Dispose();
                return null;
            }
            finally
            {
                //wc.Dispose();
            }


            // 
            String strParentId = "";
            TreeNode nd = null, fndNd = null;
            TypeNode tmpTypeNode = null;

            IDictionaryEnumerator enumRator = hashCommonLibNodes.GetEnumerator();
            DictionaryEntry entry;
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 22;
                }
                else if (tmpTypeNode.type.Equals("2"))
                {
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // file type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;


                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashCommonLibNodes[strParentId];
                    if (fndNd == null)
                    {
                        rootNd.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    rootNd.Nodes.Add(nd);
                }
            }

            enumRator = hashPrivLibNodes.GetEnumerator();
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 22;
                }
                else if (tmpTypeNode.type.Equals("2"))
                {
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;

                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashPrivLibNodes[strParentId];
                    if (fndNd == null)
                    {
                        rootNd.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    rootNd.Nodes.Add(nd);
                }
            }


            enumRator = hashFileNodes.GetEnumerator();
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 22;
                }
                else if (tmpTypeNode.type.Equals("2"))
                {
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;

                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashCommonLibNodes[strParentId];

                    if (fndNd == null)
                    {
                        fndNd = (TreeNode)hashPrivLibNodes[strParentId];

                        if (fndNd == null)
                        {
                            // ignore
                            rootNd.Nodes.Add(nd);
                        }
                        else
                        {
                            fndNd.Nodes.Add(nd);
                        }
                        // privLib.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    rootNd.Nodes.Add(nd);

                    // privLib.Nodes.Add(nd);
                    // never happen
                }
            }

            return rootNd;
        }


        public TreeNode getNodes(String userName,String fileParam, String strObjId, String strObjType, ref TreeNode privLib, ref Hashtable hashNodePermission)
        {
            TreeNode rootNd = new TreeNode("root");

            Hashtable hashCommonLibNodes = new Hashtable();
            Hashtable hashPrivLibNodes = new Hashtable();

            Hashtable hashFileNodes = new Hashtable();


            NameValueCollection form = new NameValueCollection();
            // int nLibType = -1;
            String strLibType = "";

            form.Add("username", userName);
            form.Add("serverId", "client_fms_001");
            form.Add("objectId", strObjId);
            form.Add("objectType", strObjType);

            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));

                if (res == null || res.status == 0 || res.data == null) // failure
                {
                    return null;
                }
                
                foreach (nodeInfo ndInfo in res.data)
                {
                    TypeNode pt = new TypeNode();

                    pt.id = ndInfo.id;
                    pt.name = ndInfo.name;
                    pt.parentId = ndInfo.parentid;
                    pt.orderId = ndInfo.orderid;
                    pt.type = ndInfo.type;
                    pt.isLeaf = ndInfo.isLeaf;

                    strLibType = ndInfo.lib_type;

                    TreeNode typeNd = new TreeNode(pt.name);
                    typeNd.Name = pt.name;
                    typeNd.Tag = pt;

                    typeNd.ImageIndex = typeNd.SelectedImageIndex = 13;
                  

                    if (strLibType == null)
                    {
                        hashFileNodes.Add(pt.id, typeNd);
                    }
                    else
                    {
                        if (strLibType.Equals("1"))
                        {
                            hashCommonLibNodes.Add(pt.id, typeNd);
                        }
                        else // 1
                        {
                            hashPrivLibNodes.Add(pt.id, typeNd);
                        }
                    }

                    // parse permission and insert into hash
                    // 
                    int[] arrPermVal = new int[16];
                    if (!String.IsNullOrWhiteSpace(ndInfo.pmisItem))
                    {
                        String[] arrPmis = ndInfo.pmisItem.Split(',');
                        int nVal = 0;
                        
                        foreach (String strItem in arrPmis)
                        {
                            if (int.TryParse(strItem, out nVal))
                            {
                                if (nVal == 99)
                                    nVal = 0;

                                if (nVal >= 0 && nVal < arrPermVal.GetLength(0))
                                {
                                    arrPermVal[nVal] = 1;
                                }
                            }
                        }
                    }

                    hashNodePermission[ndInfo.id] = arrPermVal;
                }

            }
            catch (WebException exp)
            {
                MessageBox.Show(exp.Message, "Error");
                wc.Dispose();
                return null;
            }
            finally
            {
                //wc.Dispose();
            }


            // 
            String strParentId = "";
            TreeNode nd = null, fndNd = null;
            TypeNode tmpTypeNode = null;

            IDictionaryEnumerator enumRator = hashCommonLibNodes.GetEnumerator();
            DictionaryEntry entry;
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 22;
                }
                else if(tmpTypeNode.type.Equals("2"))
                {
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // file type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;


                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashCommonLibNodes[strParentId];
                    if (fndNd == null)
                    {
                        rootNd.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    rootNd.Nodes.Add(nd);
                }
            }

            enumRator = hashPrivLibNodes.GetEnumerator();
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 22;
                }
                else if (tmpTypeNode.type.Equals("2"))
                {
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;

                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashPrivLibNodes[strParentId];
                    if (fndNd == null)
                    {
                        privLib.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    privLib.Nodes.Add(nd);
                }
            }


            enumRator = hashFileNodes.GetEnumerator();
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 22;
                }
                else if (tmpTypeNode.type.Equals("2"))
                {
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;

                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashCommonLibNodes[strParentId];

                    if (fndNd == null)
                    {
                        fndNd = (TreeNode)hashPrivLibNodes[strParentId];

                        if (fndNd == null)
                        {
                            // ignore
                            rootNd.Nodes.Add(nd);
                        }
                        else
                        {
                           fndNd.Nodes.Add(nd);
                        }
                        // privLib.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    rootNd.Nodes.Add(nd);

                    // privLib.Nodes.Add(nd);
                    // never happen
                }
            }


            return rootNd;
        }

        /*
        public TreeNode getChildNodes(String userName, String fileParam, String strObjId, String strObjType,ref Hashtable hashNodePermission)
        {
            TreeNode rootNd = new TreeNode("root");

            Hashtable hashCommonLibNodes = new Hashtable();
            Hashtable hashPrivLibNodes = new Hashtable();

            NameValueCollection form = new NameValueCollection();
            int nLibType = -1;

            form.Add("username", userName);
            form.Add("serverId", "client_fms_001");
            form.Add("objectId", strObjId);
            form.Add("objectType", strObjType);

            WebClient wc = new WebClient();

            wc.Encoding = Encoding.UTF8;


            try
            {
                byte[] responseArray = wc.UploadValues(m_cfgDocRepositoryUrl, "POST", form);
                String jsonText = Encoding.UTF8.GetString(responseArray);

                nodesResult res = (nodesResult)JsonConvert.DeserializeObject(jsonText, typeof(nodesResult));

                if (res == null || res.status == 0 || res.data == null) // failure
                {
                    return null;
                }

                foreach (nodeInfo ndInfo in res.data)
                {
                    TypeNode pt = new TypeNode();

                    pt.id = ndInfo.id;
                    pt.name = ndInfo.name;
                    pt.parentId = ndInfo.parentid;
                    pt.orderId = ndInfo.orderid;
                    pt.type = ndInfo.type;

                    nLibType = ndInfo.lib_type;

                    TreeNode typeNd = new TreeNode(pt.name);
                    typeNd.Name = pt.name;
                    typeNd.Tag = pt;
                    typeNd.ImageIndex = typeNd.SelectedImageIndex = 13;

                    if (nLibType == 1)
                    {
                        hashCommonLibNodes.Add(pt.id, typeNd);
                    }
                    else // 1
                    {
                        hashPrivLibNodes.Add(pt.id, typeNd);
                    }

                    // parse permission and insert into hash
                    // 
                    int[] arrPermVal = new int[16];
                    if (!String.IsNullOrWhiteSpace(ndInfo.pmisItem))
                    {
                        String[] arrPmis = ndInfo.pmisItem.Split(',');
                        int nVal = 0;

                        foreach (String strItem in arrPmis)
                        {
                            if (int.TryParse(strItem, out nVal))
                            {
                                if (nVal == 99)
                                    nVal = 0;

                                if (nVal >= 0 && nVal < arrPermVal.GetLength(0))
                                {
                                    arrPermVal[nVal] = 1;
                                }
                            }
                        }
                    }

                    hashNodePermission[ndInfo.id] = arrPermVal;
                }

            }
            catch (WebException exp)
            {
                MessageBox.Show(exp.Message, "Error");
                wc.Dispose();
                return null;
            }
            finally
            {
                //wc.Dispose();
            }


            // 
            String strParentId = "";
            TreeNode nd = null, fndNd = null;
            TypeNode tmpTypeNode = null;

            IDictionaryEnumerator enumRator = hashCommonLibNodes.GetEnumerator();
            DictionaryEntry entry;
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1") || tmpTypeNode.type.Equals("2"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // folder type node
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // file type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;


                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashCommonLibNodes[strParentId];
                    if (fndNd == null)
                    {
                        rootNd.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    rootNd.Nodes.Add(nd);
                }
            }

            enumRator = hashPrivLibNodes.GetEnumerator();
            while (enumRator.MoveNext())
            {
                entry = (DictionaryEntry)enumRator.Current;
                nd = (TreeNode)(entry.Value);

                tmpTypeNode = (nd.Tag as TypeNode);

                strParentId = tmpTypeNode.parentId;

                if (tmpTypeNode.type.Equals("1") || tmpTypeNode.type.Equals("2"))
                {
                    // nd.Tag = "#" + tmpTypeNode.id; // type node
                    nd.ImageIndex = nd.SelectedImageIndex = 13;
                }
                else
                {
                    // nd.Tag = "$" + tmpTypeNode.id; // type node
                    nd.ImageIndex = nd.SelectedImageIndex = iconIndex(tmpTypeNode.name);
                }

                nd.Tag = tmpTypeNode;

                if (strParentId != null)
                {
                    fndNd = (TreeNode)hashPrivLibNodes[strParentId];
                    if (fndNd == null)
                    {
                        privLib.Nodes.Add(nd);
                    }
                    else
                    {
                        fndNd.Nodes.Add(nd);
                    }
                }
                else
                {
                    privLib.Nodes.Add(nd);
                }
            }

            return rootNd;
        }
        */

    }
}
