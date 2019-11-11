using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;

using System.Collections;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;


namespace Antlr4.Parser
{
    public class ClassVarsNetWork
    {
        public Boolean m_bInUndoRedo = false;

        // 全部变量的查找定位HASH表
        private Hashtable m_hashVarsNetWork = new Hashtable();

        private Hashtable m_hashAllVars = null;

        public void SetAllVarsHash(Hashtable oHashAllVars)
        {
            m_hashAllVars = oHashAllVars;
        }

        public class NetWorkNode
        {
            public String strName = "";
            public String strValue = "";
            public String strOpRules = ""; // 计算公式

            public Hashtable hashDependentVars = new Hashtable();   // link up
            public Hashtable hashImpactVars = new Hashtable();      // link down

            public void Clear()
            {
                hashDependentVars.Clear();
                hashImpactVars.Clear();
            }

        }

        public NetWorkNode findNode(String strVarName)
        {
            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            return node;
        }

        public void Clear()
        {
            m_hashVarsNetWork.Clear();

            return;
        }


        // 
        public int AddVar(String strVarName, String strVarValue, String strVarOpRules, ref String strRetMsg)
        {
            int nRet = -1;
            String strRet = "", strMsg = "";

            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            if (node == null)
            {
                node = new NetWorkNode();
            }
            else
            {
                // nRet = UpdateVar(node.strName, strVarName, strVarValue, strVarOpRules,ref strRetMsg);
                nRet = SafeRemoveVar(node.strName, ref strRetMsg, false);
            }


            node.strName = strVarName;
            node.strOpRules = strVarOpRules;
            node.strValue = strVarValue;

            Hashtable hashDependentVars = new Hashtable();

            if (!node.strOpRules.Equals(""))
            {
                strRet = Computing(node.strOpRules, m_hashAllVars, ref hashDependentVars, ref strMsg);

                if (strRet != null)
                {
                    node.strValue = strRet;
                }
                else
                {
                    node.strValue = "#INVALID:" + strMsg;
                }
            }

            if (hashDependentVars.Count > 0)
            {
                foreach (DictionaryEntry entry in hashDependentVars)
                {
                    String strItem = (String)entry.Key;

                    NetWorkNode dependentNode = (NetWorkNode)m_hashVarsNetWork[strItem];
                    if (dependentNode == null)
                    {
                        dependentNode = new NetWorkNode();
                        // fill 
                        dependentNode.strName = strItem;
                        dependentNode.strValue = "";
                        dependentNode.strOpRules = "";

                        m_hashVarsNetWork[strItem] = dependentNode;
                    }

                    // record dependent vars in current node
                    node.hashDependentVars[strItem] = dependentNode;
                    // update its impact node link
                    dependentNode.hashImpactVars[strVarName] = node;

                }

            }

            m_hashVarsNetWork[strVarName] = node;


            //重新计算,变化传导
            nRet = ChangeConduct(node, ref strRetMsg);

            return nRet;
        }


        public int SafeRemoveVar(String strVarName, ref String strRetMsg, Boolean bConduct = true)
        {
            Boolean bRet = false;
            int nRet = -1;
            String strMsg = "";

            // recalc
            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            if (node != null)
            {
                // Hashtable hashPastedVars = new Hashtable();
                if (bConduct)
                {
                    ChangeConduct(node, ref strMsg);
                }

                foreach (DictionaryEntry entry in node.hashDependentVars)
                {
                    NetWorkNode dependentNode = (NetWorkNode)entry.Value;
                    dependentNode.hashImpactVars.Remove(strVarName);
                }
            }

            bRet = IsAlone(strVarName);

            if (bRet)
            {
                nRet = RemoveVar(strVarName, ref strRetMsg);
            }
            else
            {
                nRet = -1;
                strRetMsg = "非孤立变量";
            }

            return nRet;

        }

        // 
        private int RemoveVar(String strVarName, ref String strRetMsg)
        {
            int nRet = -1;

            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            if (node == null)
            {
                nRet = -1;
                strRetMsg = "无此变量";

                return nRet;
            }

            foreach (DictionaryEntry entry in node.hashDependentVars)
            {
                NetWorkNode dependentNode = (NetWorkNode)entry.Value;
                dependentNode.hashImpactVars.Remove(strVarName);
            }

            foreach (DictionaryEntry entry in node.hashImpactVars)
            {
                NetWorkNode impactedNode = (NetWorkNode)entry.Value;
                impactedNode.hashDependentVars.Remove(strVarName);
            }

            m_hashVarsNetWork.Remove(strVarName);


            nRet = 0;
            strRetMsg = "删除成功";
            return nRet;
        }

        // 
        public Boolean IsAlone(String strVarName)
        {
            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            if (node == null)
            {
                return true;
            }

            Boolean bRet = (node.hashDependentVars.Count == 0 && node.hashImpactVars.Count == 0);

            return bRet;

        }


        public Boolean IsHasRef(String strVarName)
        {
            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            if (node == null)
            {
                return false;
            }

            Boolean bRet = (node.hashImpactVars.Count > 0);

            return bRet;
        }


        // 
        public int UpdateVar(String strOldVarName, String strVarName, String strVarValue, String strVarOpRules, ref String strRetMsg)
        {
            int nRet = -1;

            // nRet = SafeRemoveVar(strOldVarName, ref strRetMsg,false);

            nRet = AddVar(strVarName, strVarValue, strVarOpRules, ref strRetMsg);

            return nRet;
        }


        public int ChangeConduct(String strVarName, String strNewValue, String strNewOpRules, ref String strRetMsg)
        {
            int nRet = -1;

            NetWorkNode node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            if (node == null)
            {
                nRet = -1;
                strRetMsg = "无此变量";

                return nRet;
            }

            //@TODO, diff to decide whether update



            nRet = UpdateVar(strVarName, strVarName, strNewValue, strNewOpRules, ref strRetMsg);
            node = (NetWorkNode)m_hashVarsNetWork[strVarName];

            nRet = ChangeConduct(node, ref strRetMsg);

            nRet = 0;
            strRetMsg = "变化传导成功";

            return nRet;
        }


        public Boolean IsCycleRef(NetWorkNode node,ref Hashtable hashPastVars)
        {

            if (hashPastVars.Contains(node.strName))
            {
                return true;
            }

            hashPastVars.Add(node.strName, node);

            foreach (DictionaryEntry entry in node.hashImpactVars)
            {
                NetWorkNode impactedNode = (NetWorkNode)entry.Value;

                if (IsCycleRef(impactedNode, ref hashPastVars))
                {
                    return true;
                }
            }

            hashPastVars.Remove(node.strName);

            return false;
        }


        public int ChangeConduct(NetWorkNode node, ref String strRetMsg)
        {
            int nRet = -1;
            Hashtable hashDependentVars = new Hashtable();
            Hashtable hashPastVars = new Hashtable();

            if (IsCycleRef(node,ref hashPastVars))
            {
                node.strValue = "#INVALID:存在循环引用:" + node.strName;
                strRetMsg = node.strValue;

                Word.ContentControl cnt = (Word.ContentControl)m_hashAllVars[node.strName];
                if (cnt != null)
                {
                    Boolean bLock1 = cnt.LockContentControl, bLock2 = cnt.LockContents;

                    cnt.LockContentControl = false;
                    cnt.LockContents = false;

                    try
                    {
                        if (!m_bInUndoRedo)
                        {
                            cnt.Range.Text = node.strValue;
                        }
                    }
                    catch (System.Exception ex)
                    {

                    }
                    finally
                    {
                    }

                    cnt.LockContentControl = bLock1;
                    cnt.LockContents = bLock2;
                }

                return nRet;
            }

            hashPastVars.Clear();

            nRet = ChangeConduct(node, ref hashDependentVars, ref hashPastVars, ref strRetMsg);

            return nRet;
        }


        // 变化传导
        public int ChangeConduct(NetWorkNode node, ref Hashtable hashDependentVars, ref Hashtable hashPastVars, ref String strRetMsg)
        {
            int nRet = -1;
            String strRet = "", strMsg = "";
            Word.ContentControl cnt = null;
            Boolean bIsCycle = false;

            if (hashPastVars.Contains(node.strName))
            {
                node.strValue = "#INVALID:存在循环引用:" + node.strName;
                bIsCycle = true;
            }
            else
            {
                if (!node.strOpRules.Equals(""))
                {
                    // re calc
                    hashDependentVars.Clear();

                    strRet = Computing(node.strOpRules, m_hashAllVars, ref hashDependentVars, ref strMsg);

                    if (strRet != null)
                    {
                        node.strValue = strRet;
                    }
                    else
                    {
                        node.strValue = "#INVALID:" + strMsg;
                    }
                }

                hashPastVars.Add(node.strName, node);
            }

            cnt = (Word.ContentControl)m_hashAllVars[node.strName];
            if (cnt != null)
            {
                Boolean bLock1 = cnt.LockContentControl, bLock2 = cnt.LockContents;

                cnt.LockContentControl = false;
                cnt.LockContents = false;

                try
                {
                    if (!m_bInUndoRedo)
                    {
                        cnt.Range.Text = node.strValue;
                    }
                }
                catch (System.Exception ex)
                {

                }
                finally
                {
                }

                cnt.LockContentControl = bLock1;
                cnt.LockContents = bLock2;
            }

            if (bIsCycle)
            {
                nRet = -1;
                strRetMsg = node.strValue;
                hashPastVars.Remove(node.strName);
                return nRet;
            }

            foreach (DictionaryEntry entry in node.hashImpactVars)
            {
                NetWorkNode impactedNode = (NetWorkNode)entry.Value;

                nRet = ChangeConduct(impactedNode, ref hashDependentVars, ref hashPastVars, ref strRetMsg);
            }

            hashPastVars.Remove(node.strName);

            return nRet;
        }


        public Boolean IsValid(String strExp, ref String strRetMsg)
        {
            Boolean bRet = true;
            calcExprLexer lex = new calcExprLexer(new AntlrInputStream("=" + strExp));
            CommonTokenStream tokens = new CommonTokenStream(lex);
            calcExprParser parser = new calcExprParser(tokens);

            IParseTree tree = null;

            try
            {
                tree = parser.parse();

                if (parser.NumberOfSyntaxErrors > 0)
                {
                    bRet = false;
                    strRetMsg = "公式:" + "'" + strExp + "'有语法错误，请查看帮助中公式的语法要求";
                }

            }
            catch (RecognitionException ex)
            {
                bRet = false;
                strRetMsg = "位置：" + parser.GetErrorHeader(ex) + ":公式:" + "'" + strExp + "'有语法错误，请查看帮助中公式的语法要求";
            }

            return bRet;
        }


        private String Computing(String strExp, Hashtable hashAllVars, ref Hashtable refVarsHash, ref String strMsg)
        {
            calcExprLexer lex = new calcExprLexer(new AntlrInputStream("=" + strExp));
            CommonTokenStream tokens = new CommonTokenStream(lex);
            calcExprParser parser = new calcExprParser(tokens);

            //             ParseTreeWalker walker = new ParseTreeWalker();
            //             calcExprBaseListener listener = new calcExprBaseListener();


            EvalExprVisitor visitor = new EvalExprVisitor();
            visitor.SetRefVarsHash(refVarsHash);
            visitor.SetVarHash(hashAllVars);//m_varsHash);

            IParseTree tree = null;
            EvalResult ret = null;

            try
            {
                tree = parser.parse();
                if (tree != null)
                {
                    ret = visitor.Visit(tree);
                }
            }
            catch (RecognitionException ex)
            {
                strMsg = ex.Message;
                System.Console.Error.WriteLine(ex.StackTrace);
                return null;
            }

            if (ret == null)
                return null;

            if (ret.bInvalid)
                return ret.strExceptionMsg;

            return ret.Value1.ToString();
        }
    }

}
