using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Xml;


namespace OfficeAssist
{
    class TreeExXMLCls
    {
        private TreeView thetreeview;
        private string xmlfilepath;
        XmlTextWriter textWriter;
        XmlNode Xmlroot;
        XmlDocument textdoc;

        String[,] m_arrStrs = {
{"~","_英文波浪号_"},
{"`","_英文顿号_"}, 
{"!","_英文感叹号_"}, 
{"@","_英文AT号_"}, 
{"#","_英文井号_"}, 
{"$","_英文美刀号_"}, 
{"%","_英文百分号_"}, 
{"^","_英文分词号_"}, 
{"&","_英文AND号_"}, 
{"*","_英文星号_"}, 
{"(","_英文左圆括号_"}, 
{")","_英文右圆括号_"}, 
{"-","_英文中划号_"}, 
{"+","_英文加号_"}, 
{"=","_英文等号_"}, 
{"|","_英文连接号_"}, 
{"\\","_英文左斜号_"}, 
{"{","_英文左大括号_"}, 
{"}","_英文右大括号_"}, 
{"[","_英文左方括号_"}, 
{"]","_英文右方括号_"}, 
{":","_英文冒号_"}, 
{";","_英文分号_"}, 
{"\"","_英文双引号_"}, 
{"'","_英文单引号_"}, 
{"<","_英文小于号_"}, 
{">","_英文大于号_"}, 
{",","_英文逗号_"}, 
{".","_英文句号_"}, 
{"?","_英文问号_"}, 
{"/","_英文右斜号_"}, 
{"\r","_英文回车号_"}, 
{"\n","_英文换行号_"},
{"\a","_英文段落号_"},

{"~","_中文波浪号_"},
{"·","_中文连词号_"},
{"！","_中文感叹号_"},
{"@","_中文AT号_"},
{"#","_中文井号_"},
{"￥","_中文货币号_"},
{"%","_中文百分号_"},
{"……","_中文省略号_"},
{"&","_中文连接号_"},
{"*","_中文星号_"},
{"（","_中文左圆括号_"},
{"）","_中文右圆括号_"},
{"-","_中文中划线号_"},
{"——","_中文下划线号_"},
{"+","_中文加号_"},
{"=","_中文等号_"},
{"｜","_中文竖号_"},
{"、","_中文顿号_"},
{"｛","_中文左大括号_"},
{"｝","_中文右大括号_"},
{"【","_中文左书名号_"},
{"】","_中文右书名号_"},
{"：","_中文冒号_"},
{"；","_中文分号_"},
{"“","_中文左双引号_"},
{"”","_中文右双引号_"}, 
{"‘","_中文左单引号_"},
{"’","_中文右单引号_"}, 
{"《","_中文左双书名号_"},
{"》","_中文右双书名号_"},
{"，","_中文逗号_"},
{"。","_中文句号_"},
{"？","_中文问号_"},
{"、","_中文顿号_"},

{"0","_英文0数字_"},
{"1","_英文1数字_"},
{"2","_英文2数字_"},
{"3","_英文3数字_"},
{"4","_英文4数字_"},
{"5","_英文5数字_"},
{"6","_英文6数字_"},
{"7","_英文7数字_"},
{"8","_英文8数字_"},
{"9","_英文9数字_"}

                   };


        public TreeExXMLCls()
        {
            //----构造函数  
            textdoc = new XmlDocument();

        }

        ~TreeExXMLCls()
        {
            //----析构函数  
        }

        #region 遍历treeview并实现向XML的转化
        /// <summary>     
        /// 遍历treeview并实现向XML的转化  
        /// </summary>     
        /// <param name="TheTreeView">树控件对象</param>     
        /// <param name="XMLFilePath">XML输出路径</param>     
        /// <returns>0表示函数顺利执行</returns>     

        public int TreeToXML(TreeView TheTreeView, string XMLFilePath)
        {
            //-------初始化转换环境变量  
            thetreeview = TheTreeView;
            xmlfilepath = XMLFilePath;
            textWriter = new XmlTextWriter(xmlfilepath, null);

            //-------创建XML写操作对象  
            textWriter.Formatting = Formatting.Indented;

            //-------开始写过程，调用WriteStartDocument方法  
            textWriter.WriteStartDocument();
            
            //-------写入说明  
            // textWriter.WriteComment("this XML is created from a tree");
            // textWriter.WriteComment("By 思月行云");

            //-------添加第一个根节点  
            textWriter.WriteStartElement("FillGatherSchemes");
            textWriter.WriteEndElement();

            //------ 写文档结束，调用WriteEndDocument方法  
            textWriter.WriteEndDocument();

            //-----关闭输入流  
            textWriter.Close();

            //-------创建XMLDocument对象  
            textdoc.Load(xmlfilepath);

            textdoc.CreateXmlDeclaration("1.0", "utf-8", "yes");

            //------选中根节点
            if (thetreeview.Nodes.Count > 0)
            {
                XmlElement Xmlnode = textdoc.CreateElement(thetreeview.Nodes[0].Text);
                Xmlroot = textdoc.SelectSingleNode("FillGatherSchemes");

                //------遍历原treeview控件，并生成相应的XML  
                TransTreeSav(thetreeview.Nodes, (XmlElement)Xmlroot);
            }

            return 0;
        }


        private String transString2InXml(String strIn)
        {
            String strInXml = strIn;

            int nRowCnt = m_arrStrs.GetLength(0);
            for (int i = 0; i < nRowCnt; i++)
            {
                strInXml = strInXml.Replace(m_arrStrs[i,0], m_arrStrs[i,1]);
            }

            return strInXml;
        }


        private String restoreInXml2String(String strInXml)
        {
            String strTxt = strInXml;

            int nRowCnt = m_arrStrs.GetLength(0);
            for (int i = 0; i < nRowCnt; i++)
            {
                strTxt = strTxt.Replace(m_arrStrs[i, 1], m_arrStrs[i, 0]);
            }

            return strTxt;
        }


        private int TransTreeSav(TreeNodeCollection nodes, XmlElement ParXmlnode)
        {

            //-------遍历树的各个故障节点，同时添加节点至XML  
            XmlElement xmlnode;
            Xmlroot = textdoc.SelectSingleNode("FillGatherSchemes");

            String strInXmlTxt = "";

            foreach (TreeNode node in nodes)
            {
                strInXmlTxt = transString2InXml(node.Text);

                xmlnode = textdoc.CreateElement(strInXmlTxt);
                if (node.Tag != null)
                {
                    strInXmlTxt = transString2InXml((String)node.Tag);

                    xmlnode.SetAttribute("tag", strInXmlTxt);
                }

                ParXmlnode.AppendChild(xmlnode);

                if (node.Nodes.Count > 0)
                {
                    TransTreeSav(node.Nodes, xmlnode);
                }
            }
            textdoc.Save(xmlfilepath);
            return 0;
        }

        #endregion

        #region 遍历XML并实现向tree的转化
        /// <summary>     
        /// 遍历treeview并实现向XML的转化  
        /// </summary>     
        /// <param name="XMLFilePath">XML输出路径</param>     
        /// <param name="TheTreeView">树控件对象</param>     
        /// <returns>0表示函数顺利执行</returns>     

        public int XMLToTree(string XMLFilePath, TreeView TheTreeView)
        {
            //-------重新初始化转换环境变量  
            thetreeview = TheTreeView;
            xmlfilepath = XMLFilePath;

            //-------重新对XMLDocument对象赋值  
            textdoc.Load(xmlfilepath);

            XmlNode root = textdoc.SelectSingleNode("FillGatherSchemes");

            String strTxt = "";

            foreach (XmlNode subXmlnod in root.ChildNodes)
            {
                TreeNode trerotnod = new TreeNode();

                strTxt = restoreInXml2String(subXmlnod.Name);

                trerotnod.Text = strTxt;
                if (subXmlnod.Attributes["tag"] != null)
                {
                    strTxt = restoreInXml2String(subXmlnod.Attributes["tag"].Value);
                    trerotnod.Tag = strTxt;
                }

                thetreeview.Nodes.Add(trerotnod);
                TransXML(subXmlnod.ChildNodes, trerotnod);

            }
            return 0;
        }

        private int TransXML(XmlNodeList Xmlnodes, TreeNode partrenod)
        {
            String strTxt = "";

            //------遍历XML中的所有节点，仿照treeview节点遍历函数  
            foreach (XmlNode xmlnod in Xmlnodes)
            {
                TreeNode subtrnod = new TreeNode();

                strTxt = restoreInXml2String(xmlnod.Name);

                subtrnod.Text = strTxt;

                if (xmlnod.Attributes["tag"] != null)
                {
                    strTxt = restoreInXml2String(xmlnod.Attributes["tag"].Value);
                    subtrnod.Tag = strTxt;
                }

                partrenod.Nodes.Add(subtrnod);

                if (xmlnod.ChildNodes.Count > 0)
                {
                    TransXML(xmlnod.ChildNodes, subtrnod);
                }
            }
            return 0;
        }
        #endregion
    }
}
