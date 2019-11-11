using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace OfficeAssist.docPub.Model
{
    [Serializable]
    [XmlType]
    public class tocStyle
    {
        // 章节目录
        [XmlAttribute]
        public Boolean bHeadingTocEnable = true;
        [XmlElement]
        public nhFont hdTocFnt { get; set; }
        [XmlElement]
        public nhParaFmt hdTocParaFmt { get; set; }

        [XmlArrayAttribute]
        public List<nhFont> headingTocFnts { get; set; }//= new nhFont();
        [XmlArrayAttribute]
        public List<nhParaFmt> headingTocParaFmts { get; set; }//= new nhParaFmt();

        // 图目录
        [XmlAttribute]
        public Boolean bInShpTocEnable = true;
        [XmlElement]
        public nhFont inShpTocFnt { get; set; }//= new nhFont();
        [XmlElement]
        public nhParaFmt inShpTocParaFmt { get; set; }//= new nhParaFmt();

        // 表目录
        [XmlAttribute]
        public Boolean bTblTocEnable = true;
        [XmlElement]
        public nhFont tblTocFnt { get; set; }//= new nhFont();
        [XmlElement]
        public nhParaFmt tblTocParaFmt { get; set; }// = new nhParaFmt();


        public tocStyle()
        {
            bHeadingTocEnable = true;
            hdTocFnt = new nhFont();
            hdTocParaFmt = new nhParaFmt();

            headingTocFnts = new List<nhFont>();
            for (int i = 0; i < 9; i++)
            {
                headingTocFnts.Add(new nhFont());
            }

            headingTocParaFmts = new List<nhParaFmt>();
            for (int i = 0; i < 9; i++)
            {
                headingTocParaFmts.Add(new nhParaFmt());
            }

            inShpTocFnt = new nhFont();
            inShpTocParaFmt = new nhParaFmt();

            tblTocFnt = new nhFont();
            tblTocParaFmt = new nhParaFmt();

            return;
        }


        public void clone(tocStyle oth)
        {
            bHeadingTocEnable = oth.bHeadingTocEnable;

            hdTocFnt.clone(oth.hdTocFnt);
            hdTocParaFmt.clone(oth.hdTocParaFmt);

            nhFont inFnt = null, oFnt = null;
            nhParaFmt inParaFmt = null, oParaFmt = null;

            int nCnt = headingTocFnts.Count;
            for (int i = 0; i < nCnt; i++ )
            {
                inFnt = (nhFont)headingTocFnts[i];
                oFnt = (nhFont)oth.headingTocFnts[i];

                inFnt.clone(oFnt);
            }

            nCnt = headingTocParaFmts.Count;
            for (int i = 0; i < nCnt; i++)
            {
                inParaFmt = (nhParaFmt)headingTocParaFmts[i];
                oParaFmt = (nhParaFmt)oth.headingTocParaFmts[i];

                inParaFmt.clone(oParaFmt);
            }

            bInShpTocEnable = oth.bInShpTocEnable;
            inShpTocFnt.clone(oth.inShpTocFnt);
            inShpTocParaFmt.clone(oth.inShpTocParaFmt);

            bTblTocEnable = oth.bTblTocEnable;
            tblTocFnt.clone(oth.tblTocFnt);
            tblTocParaFmt.clone(oth.tblTocParaFmt);

            return;
        }


        public void copy2(tocStyle oth)
        {
            oth.bHeadingTocEnable = bHeadingTocEnable;
            oth.hdTocFnt.clone(hdTocFnt);
            oth.hdTocParaFmt.clone(hdTocParaFmt);

            nhFont inFnt = null, oFnt = null;
            nhParaFmt inParaFmt = null, oParaFmt = null;

            int nCnt = headingTocFnts.Count;
            for (int i = 0; i < nCnt; i++)
            {
                inFnt = (nhFont)headingTocFnts[i];
                oFnt = (nhFont)oth.headingTocFnts[i];

                oFnt.clone(inFnt);
            }

            nCnt = headingTocParaFmts.Count;
            for (int i = 0; i < nCnt; i++)
            {
                inParaFmt = (nhParaFmt)headingTocParaFmts[i];
                oParaFmt = (nhParaFmt)oth.headingTocParaFmts[i];

                oParaFmt.clone(inParaFmt);
            }

            oth.bInShpTocEnable = bInShpTocEnable;
            oth.inShpTocFnt.clone(inShpTocFnt);
            oth.inShpTocParaFmt.clone(inShpTocParaFmt);

            oth.bTblTocEnable = bTblTocEnable;
            oth.tblTocFnt.clone(tblTocFnt);
            oth.tblTocParaFmt.clone(tblTocParaFmt);

            return;
        }

        public int formatString(RichTextBox rchTxt, String strPreBlanks)
        {
            String strRet = "";
            String strCurBlanks = strPreBlanks;
            String strNextLevelBlanks = strCurBlanks + "    ";

            if (bHeadingTocEnable)
            {
                strRet = strCurBlanks + "章节目录:启用\r\n";

                rchTxt.AppendText(strRet);

                hdTocFnt.formatString(rchTxt, strNextLevelBlanks);
                hdTocParaFmt.formatString(rchTxt, strNextLevelBlanks);

                for (int i = 0; i < 9; i++ )
                {
                    strRet = strNextLevelBlanks + (i + 1) + "级：\r\n";
                    rchTxt.AppendText(strRet);
                    headingTocFnts[i].formatString(rchTxt, strNextLevelBlanks);
                    headingTocParaFmts[i].formatString(rchTxt, strNextLevelBlanks);
                }

            }
            else
            {
                strRet = strCurBlanks + "章节目录:停用\r\n";
                rchTxt.AppendText(strRet);
            }

            if (bInShpTocEnable)
            {
                strRet = strCurBlanks + "图目录:启用\r\n";
                // 是否 取题注的 name？
                rchTxt.AppendText(strRet);

                inShpTocFnt.formatString(rchTxt, strNextLevelBlanks);
                inShpTocParaFmt.formatString(rchTxt, strNextLevelBlanks);
            }
            else
            {
                strRet = strCurBlanks + "图目录:停用\r\n";
                rchTxt.AppendText(strRet);
            }


            if (bTblTocEnable)
            {
                strRet = strCurBlanks + "表目录:启用\r\n";
                rchTxt.AppendText(strRet);

                tblTocFnt.formatString(rchTxt, strNextLevelBlanks);
                tblTocParaFmt.formatString(rchTxt, strNextLevelBlanks);
            }
            else
            {
                strRet = strCurBlanks + "表目录:停用\r\n";
                rchTxt.AppendText(strRet);
            }

            return 0;
        }


    }// class

}// namespace
