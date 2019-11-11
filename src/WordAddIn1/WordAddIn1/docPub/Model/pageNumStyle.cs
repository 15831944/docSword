using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace OfficeAssist.docPub.Model
{
    [Serializable]
    [XmlType]
    public class pageNumStyle
    {
        [XmlAttribute]
        public Boolean bEnable = true;

        [XmlAttribute]
        public int nSecNum = 0;
        [XmlElement]
        public nhFont Fnt { get; set; } // = new nhFont();
        [XmlElement]
        public nhParaFmt ParaFmt { get; set; } // = new nhParaFmt();
        [XmlAttribute]
        public int nPgNumSnStyle = 0; // 编号格式
        [XmlAttribute]
        public Boolean bIncludeHeadingSn = false; // 包含章节号
        [XmlAttribute]
        public int nPgNumHeadingStartStyle = -1; // 标题样式序号
        [XmlAttribute]
        public int nPgNumHeadingSplittor = -1; // 使用分隔符序号

        [XmlAttribute]
        public Boolean bPgNumFollowPrevSec = true; // 续前节
        [XmlAttribute]
        public int nPgNumStartPageNum = 1; // 起始页码


        private readonly String[] m_arrPageNumSplittors = { "- （连字符）", ". （句点）", ": （冒号）", "—（长划线）", "–（短划线）" };

        public pageNumStyle()
        {
            // 缺省值

            Fnt = new nhFont();
            ParaFmt = new nhParaFmt();

            return;
        }

        public void clone(pageNumStyle oth)
        {
            bEnable = oth.bEnable;
            nSecNum = oth.nSecNum;

            Fnt.clone(oth.Fnt);
            ParaFmt.clone(oth.ParaFmt);

            nPgNumSnStyle = oth.nPgNumSnStyle;
            bIncludeHeadingSn = oth.bIncludeHeadingSn;
            nPgNumHeadingStartStyle = oth.nPgNumHeadingStartStyle;
            nPgNumHeadingSplittor = oth.nPgNumHeadingSplittor;

            bPgNumFollowPrevSec = oth.bPgNumFollowPrevSec;
            nPgNumStartPageNum = oth.nPgNumStartPageNum;

            return;
        }

        public void copy2(pageNumStyle oth)
        {
            oth.bEnable = bEnable;
            oth.nSecNum = nSecNum;

            oth.Fnt.clone(Fnt);
            oth.ParaFmt.clone(ParaFmt);

            oth.nPgNumSnStyle = nPgNumSnStyle;
            oth.bIncludeHeadingSn = bIncludeHeadingSn;
            oth.nPgNumHeadingStartStyle = nPgNumHeadingStartStyle;
            oth.nPgNumHeadingSplittor = nPgNumHeadingSplittor;

            oth.bPgNumFollowPrevSec = bPgNumFollowPrevSec;
            oth.nPgNumStartPageNum = nPgNumStartPageNum;

            return;
        }

        public int formatString(RichTextBox rchTxt, String strPreBlanks)
        {
            if (!bEnable)
            {
                rchTxt.AppendText("停用\r\n");
                return 0;
            }

            String strCurLevelBlanks = strPreBlanks + "    ";
            String strNextLevelBlanks = strCurLevelBlanks + "    ";

            rchTxt.AppendText(strCurLevelBlanks + "启用\r\n");

            Fnt.formatString(rchTxt, strNextLevelBlanks);
            ParaFmt.formatString(rchTxt, strNextLevelBlanks);

            // ?
            String strRet = @"1，2，3，…";
            switch (nPgNumSnStyle)
            {
                case (int)Word.WdPageNumberStyle.wdPageNumberStyleArabic:
                    strRet = @"1，2，3，…";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleNumberInDash:
                    strRet = @"- 1 -，- 2 -，- 3 -，…";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleArabicFullWidth:
                    strRet = @"全角 …";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleLowercaseLetter:
                    strRet = @"a，b，c，…";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleUppercaseLetter:
                    strRet = @"A，B，C，…";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleLowercaseRoman:
                    strRet = @"i，ii，iii，…";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleUppercaseRoman:
                    strRet = @"I，II，III，…";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleSimpChinNum1:
                    strRet = @"一，二，三（简） …";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleSimpChinNum2:
                    strRet = @"壹，贰，叁 …";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleTradChinNum1:
                    strRet = @"甲，乙，丙 …";
                    break;

                case (int)Word.WdPageNumberStyle.wdPageNumberStyleTradChinNum2:
                    strRet = @"子，丑，寅 …";
                    break;

                //
            }

            rchTxt.AppendText(strNextLevelBlanks + "编号格式：\'" +  strRet + "\'\r\n");


            if (bIncludeHeadingSn)
            {
                rchTxt.AppendText(strNextLevelBlanks + "[包含章节号]");

                if (nPgNumStartPageNum != -1)
                {
                    rchTxt.AppendText("[章节号自：样式\'标题" + (nPgNumStartPageNum + 1) + "\']");
                }

                if (nPgNumHeadingSplittor != -1)
                {
                    rchTxt.AppendText("[分隔符：\'" + m_arrPageNumSplittors[nPgNumHeadingSplittor] + "\']"); // ???
                }

                rchTxt.AppendText("\r\n");
            }


            if (bPgNumFollowPrevSec)
            {
                rchTxt.AppendText(strNextLevelBlanks + "起始页码：[续前节]\r\n");
            }
            else
            {
                rchTxt.AppendText(strNextLevelBlanks + "起始页码：[" + nPgNumStartPageNum + "]\r\n");
            }

            return 0;
        }


    }
}
