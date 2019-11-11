using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using OfficeTools.Common;

using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace OfficeAssist.docPub.Model
{
    [Serializable]
    [XmlType]
    public class tizhuStyle
    {
        [XmlAttribute]
        public Boolean bEnable = true;

        [XmlAttribute]
        public String strCapLabelName;

        // 位置
        [XmlAttribute]
        public Word.WdCaptionPosition capPos;

        // 居中，左右
        [XmlAttribute]
        public Word.WdParagraphAlignment align;

        // 前缀
        [XmlAttribute]
        public String strPrefix;

        // 后缀
        [XmlAttribute]
        public String strPostfix;

        // 取就近标题
        [XmlAttribute]
        public Boolean bGetHeadingText;

        // 同标题则序号区分
        [XmlAttribute]
        public Boolean bSnWhileSameHeadingText;

        // 范围：目录之后
        [XmlAttribute]
        public Boolean bScopeAfterToc;

        // 题注字体
        [XmlElement]
        public nhFont Fnt { get; set; } // = new nhFont();
        [XmlElement]
        public nhParaFmt ParaFmt { get; set; } // = new nhFont();

        public tizhuStyle()
        {
            strCapLabelName = "";
            capPos = WdCaptionPosition.wdCaptionPositionBelow;
            align = WdParagraphAlignment.wdAlignParagraphCenter;
            strPrefix = "";
            strPostfix = "";
            bGetHeadingText = false;
            bSnWhileSameHeadingText = false;
            bScopeAfterToc = false;

            Fnt = new nhFont();
            ParaFmt = new nhParaFmt();

            return;
        }

        public void clone(tizhuStyle oth)
        {
            bEnable = oth.bEnable;
            strCapLabelName = oth.strCapLabelName;
            capPos = oth.capPos;
            align = oth.align;
            strPrefix = oth.strPrefix;
            strPostfix = oth.strPostfix;
            bGetHeadingText = oth.bGetHeadingText;
            bSnWhileSameHeadingText = oth.bSnWhileSameHeadingText;
            bScopeAfterToc = oth.bScopeAfterToc;

            Fnt.clone(oth.Fnt);
            ParaFmt.clone(oth.ParaFmt);

            return;
        }

        public void copy2(tizhuStyle oth)
        {
            oth.bEnable = bEnable;
            oth.strCapLabelName = strCapLabelName;
            oth.capPos = capPos;
            oth.align = align;
            oth.strPrefix = strPrefix;
            oth.strPostfix = strPostfix;
            oth.bGetHeadingText = bGetHeadingText;
            oth.bSnWhileSameHeadingText = bSnWhileSameHeadingText;
            oth.bScopeAfterToc = bScopeAfterToc;

            oth.Fnt.clone(Fnt);
            oth.ParaFmt.clone(ParaFmt);

            return;
        }

        public int formatString(RichTextBox rchTxt, String strPreBlanks)
        {
            if (bEnable)
            {
                rchTxt.AppendText("启用\r\n");

                if (!String.IsNullOrWhiteSpace(strCapLabelName))
                {
                    rchTxt.AppendText("题注名：" + strCapLabelName + "\r\n");

                    if (capPos == WdCaptionPosition.wdCaptionPositionBelow)
                    {
                        rchTxt.AppendText("居下\r\n");
                    }
                    else
                    {
                        rchTxt.AppendText("居上\r\n");
                    }

                    switch (align)
                    {
                        case WdParagraphAlignment.wdAlignParagraphLeft:
                            rchTxt.AppendText("左对齐\r\n");
                            break;

                        case WdParagraphAlignment.wdAlignParagraphRight:
                            rchTxt.AppendText("右对齐\r\n");
                            break;

                        case WdParagraphAlignment.wdAlignParagraphCenter:
                            rchTxt.AppendText("居中\r\n");
                            break;

                        default:
                            break;

                    }

                    if (!String.IsNullOrWhiteSpace(strPrefix))
                    {
                        rchTxt.AppendText("前缀文字：" + strPrefix + "\r\n");
                    }

                    if (bGetHeadingText)
                    {
                        rchTxt.AppendText("[取就近标题内容]\r\n");
                    }

                    if (bSnWhileSameHeadingText)
                    {
                        rchTxt.AppendText("[同标题则序号区分]\r\n");
                    }

                    if (!String.IsNullOrWhiteSpace(strPostfix))
                    {
                        rchTxt.AppendText("后缀文字：" + strPostfix + "\r\n");
                    }


                    Fnt.formatString(rchTxt, strPreBlanks);
                    ParaFmt.formatString(rchTxt, strPreBlanks);
                }

            }
            else
            {
                rchTxt.AppendText("停用\r\n");
            }

            return 0;
        }



    }
}
