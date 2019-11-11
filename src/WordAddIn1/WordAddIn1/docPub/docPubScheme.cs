using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

using OfficeAssist.docPub.Model;
using OfficeTools.Common;
using System.Xml.Serialization;
using System.IO;

using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Collections;

namespace OfficeAssist.docPub
{
    [Serializable]
    [XmlRoot("docPubScheme")]
    public class docPubScheme
    {
        //-------------------
        [XmlAttribute("SchemeName")]
        public String strSchemeName { get; set; } //  = "";

        [XmlAttribute]
        public int nTocEnable = 2;   // checkbox.CheckState, 0-unchecked, 1-checked, 2-Indeterminate
        [XmlElement]
        public nhFont tocFnt { get; set; }
        [XmlElement]
        public nhParaFmt tocParaFmt { get; set; }
        // 目录
        [XmlElement]
        public tocStyle tocStyles { get; set; } // = new tocStyle();


        [XmlAttribute]
        public Boolean bHdSnsEnable = true;
        [XmlElement]
        public nhFont hdSnFnt { get; set; }
        // 章节序号
        [XmlArrayAttribute("hdSns")]
        public List<headingSn> hdSns {get; set;} // = new headingSn[9];

        [XmlAttribute]
        public Boolean bHdStylesEnable = true;
        [XmlElement]
        public nhFont hdStyleFnt { get; set; }
        [XmlElement]
        public nhParaFmt hdStyleParaFmt { get; set; }
        // 章节样式
        [XmlArrayAttribute("hdStyles")]
        public List<headingStyle> hdStyles { get; set; } //  = new headingStyle[9];


        // 题注表
        [XmlAttribute]
        public int nTizhuEnable = 2;  // checkbox.CheckState, 0-unchecked, 1-checked, 2-Indeterminate
        [XmlElement]
        public nhFont tizhuFnt { get; set; }
        [XmlElement]
        public nhParaFmt tizhuParaFmt { get; set; }
        [XmlElement]
        public tizhuStyle tblTizhuStyle { get; set; }//= new tizhuStyle();  // 表题注
        [XmlElement]
        public tizhuStyle inShpTizhuStyle { get; set; }//= new tizhuStyle();// 图题注

        // 图（嵌入且独立成段落）、整表
        [XmlAttribute]
        public Boolean bTblEnable = true;
        // 整表字体，表内段落
        [XmlElement]
        public nhFont tblFont { get; set; } // = new nhFont();
        //[XmlElement]
        //public nhParaFmt tblEveryParaFmt { get; set; } // = new nhParaFmt();
        [XmlElement]
        public nhParaFmt tblParaFmt { get; set; } // = new nhParaFmt();
        [XmlAttribute]
        public Boolean bInShpEnable = true;
        [XmlElement]
        public nhParaFmt inShpParaFmt { get; set; } // = new nhParaFmt();
        

        // 序号段落
        [XmlAttribute]
        public int nListParaEnable = 2;  // checkbox.CheckState, 0-unchecked, 1-checked, 2-Indeterminate
        [XmlElement]
        public nhFont listParaFnt { get; set; }
        [XmlElement]
        public nhParaFmt listParaParaFmt { get; set; }

        [XmlElement]
        public listParaStyle outTblListParaStyle { get; set; } // = new listParaStyle(); // 表内
        [XmlElement]
        public listParaStyle inTblListParaStyle { get; set; } // = new listParaStyle();  // 表外

        // 正文
        [XmlAttribute]
        public int nTextBodyEnable = 2; // checkbox.CheckState, 0-unchecked, 1-checked, 2-Indeterminate
        [XmlElement]
        public nhFont textBodyFnt { get; set; }
        [XmlElement]
        public nhParaFmt textBodyParaFmt { get; set; }

        [XmlElement]
        public textBodyStyle outTblTextBodyStyle { get; set; } // = new textBodyStyle();// 表内
        [XmlElement]
        public textBodyStyle inTblTextBodyStyle { get; set; } // = new textBodyStyle(); // 表外

        // 页码格式 （9个节）
        [XmlAttribute]
        public int nPgNumEnable = 2; // checkbox.CheckState, 0-unchecked, 1-checked, 2-Indeterminate

        [XmlElement]
        public pageNumStyle pgNumStyle { get; set; }

        [XmlArrayAttribute("pgNumStyles")]
        public List<pageNumStyle> pgNumStyles { get; set; } // = new pageNumStyle[9];


        private String[] strsHdSnNumberFormat = {"%1","%1.%2","%1.%2.%3","%1.%2.%3.%4",
                                                "%1.%2.%3.%4.%5","%1.%2.%3.%4.%5.%6","%1.%2.%3.%4.%5.%6.%7",
                                                "%1.%2.%3.%4.%5.%6.%7.%8","%1.%2.%3.%4.%5.%6.%7.%8.%9"};
        private float[] fArrsNumberPosition = {0.0f,0.0f,0.0f,0.0f,0.0f,0.0f,0.0f,0.0f,0.0f };

        private float[] fArrsTextPosition = { 0.76f, 1.02f, 1.27f, 1.52f, 1.78f, 2.03f, 2.29f, 2.54f, 2.79f };

        private Hashtable m_hashSnStyle2Name = new Hashtable();

        public void fillHeadingSnDefault(headingSn hsn, int nIndex)
        {
            Word.Application app = Globals.ThisAddIn.Application;

            int i = nIndex;

            hsn.Index = i;
            hsn.NumberFormat = strsHdSnNumberFormat[i];
            hsn.NumberPosition = app.CentimetersToPoints(fArrsNumberPosition[i]);

            hsn.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            hsn.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;

            hsn.Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            hsn.TextPosition = app.CentimetersToPoints(fArrsTextPosition[i]);
            hsn.TabPosition = 0.0f;
            hsn.ResetOnHigher = i;
            hsn.StartAt = 1;
            hsn.LinkedStyle = "标题 " + (i + 1);

            return;
        }


        public void fillHeadingStyleDefault(headingStyle hs, int nIndex, String strFntName)
        {
            String[] strArrFntNameBi = new String[] {
                    "Times New Roman",
                    "+标题 CS",
                    "+正文 CS 字体",
                    "+标题 CS",
                    "+正文 CS 字体",
                    "+标题 CS",
                    "+正文 CS 字体",
                    "+标题 CS",
                    "+标题 CS",
                    "+正文 CS 字体"
            };

            /*
            float[] fArrFntSize = new float[] {
                36,26,24,22,18,16,15,14,12,10.0f
            };*/

            float[] fArrFntSize = new float[] {
                22.0f,16.0f,16.0f,14.0f,14.0f,12.0f,12.0f,12.0f,10f,10f};
            /*
            标题1：中文正文，宋体，2号，加粗
            标题2：宋体，3号，加粗
            标题3：宋体，3号，加粗
            标题4：宋体，4号，加粗
            标题5：宋体，4号，加粗
            标题6：宋体，小四号，加粗
            标题7：宋体，小四号，加粗
            标题8：宋体，小四
            标题9：宋体，5号
            正文：宋体，5号
            */

            float[] fArrFntKerning = new float[] {
                16,1,1,1,1,1,1,1,1,1
            };


            // float[] fArrParaFmtLineSpacing = new float[]{
            //    12,20.8f,20.8f,18.8f,18.8f,16,16,16,16,12
            // };

            float[] fArrParaFmtLineSpacing = new float[]{
                12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f,12.0f
            };

            float[] fArrParaFmtSpaceBefore = new float[]{
                5,13,13,14,14,12,12,12,12,0
            };

            float[] fArrParaFmtSpaceAfter = new float[]{
                5,13,13,14.5f,14.5f,3.2f,3.2f,3.2f,3.2f,0
            };

            int i = nIndex;

            hs.Fnt.Color = (int)Word.WdColor.wdColorAutomatic;// -16777216;
            hs.Fnt.UnderlineColor = (int)Word.WdColor.wdColorAutomatic;// -16777216;
            hs.Fnt.DiacriticColor = (int)(Word.WdColor)Word.WdConstants.wdUndefined;

            hs.Fnt.Name = strFntName;
            hs.Fnt.NameFarEast = strFntName;
            hs.Fnt.NameAscii = strFntName;//"+西文正文";
            //hs[i].m_fnt.NameOther = strFntName;//"+西文正文";

            hs.Fnt.NameBi = strArrFntNameBi[i];

            hs.Fnt.Size = fArrFntSize[i];
            hs.Fnt.SizeBi = fArrFntSize[i];


            hs.Fnt.Kerning = fArrFntKerning[i];
            hs.Fnt.Scaling = 100;


            // paragraph format

            hs.ParaFmt.OutlineLevel = (int)((Word.WdOutlineLevel)(i + 1));
            hs.ParaFmt.AddSpaceBetweenFarEastAndAlpha = -1;
            hs.ParaFmt.AddSpaceBetweenFarEastAndDigit = -1;
            hs.ParaFmt.Alignment = (int)Word.WdParagraphAlignment.wdAlignParagraphLeft;
            hs.ParaFmt.AutoAdjustRightIndent = -1;
            hs.ParaFmt.BaseLineAlignment = (int)Word.WdBaselineAlignment.wdBaselineAlignAuto; // 4
            hs.ParaFmt.CharacterUnitFirstLineIndent = 0;
            hs.ParaFmt.CharacterUnitLeftIndent = 0;
            hs.ParaFmt.CharacterUnitRightIndent = 0;
            hs.ParaFmt.DisableLineHeightGrid = 0;
            hs.ParaFmt.FarEastLineBreakControl = -1;
            hs.ParaFmt.FirstLineIndent = 0;
            hs.ParaFmt.HalfWidthPunctuationOnTopOfLine = 0;
            hs.ParaFmt.HangingPunctuation = -1;
            hs.ParaFmt.Hyphenation = -1;

            if (i < 9)
            {
                hs.ParaFmt.KeepTogether = -1;
                hs.ParaFmt.KeepWithNext = -1;
            }
            else
            {
                hs.ParaFmt.KeepTogether = 0;
                hs.ParaFmt.KeepWithNext = 0;
            }


            hs.ParaFmt.LeftIndent = 0;
            hs.ParaFmt.LineSpacing = fArrParaFmtLineSpacing[i];
            hs.ParaFmt.LineSpacingRule = (int)Word.WdLineSpacing.wdLineSpaceSingle;//Word.WdLineSpacing.wdLineSpaceMultiple; // 5
            hs.ParaFmt.LineUnitAfter = 0;
            hs.ParaFmt.LineUnitBefore = 0;
            hs.ParaFmt.MirrorIndents = 0;
            hs.ParaFmt.NoLineNumber = 0;
            hs.ParaFmt.PageBreakBefore = 0;
            hs.ParaFmt.ReadingOrder = (int)Word.WdReadingOrder.wdReadingOrderLtr; // 1
            hs.ParaFmt.RightIndent = 0;
            hs.ParaFmt.SpaceAfter = fArrParaFmtSpaceAfter[i];
            hs.ParaFmt.SpaceAfterAuto = 0; // [0] = -1, others = 0
            hs.ParaFmt.SpaceBefore = fArrParaFmtSpaceBefore[i];
            hs.ParaFmt.SpaceBeforeAuto = 0; // [0] = -1, others = 0
            hs.ParaFmt.TextboxTightWrap = 0;
            hs.ParaFmt.WidowControl = 0; // [0] = -1, others = 0
            hs.ParaFmt.WordWrap = -1;

            return;
        }


        public docPubScheme()
        {
            Word.Application app = Globals.ThisAddIn.Application;

            strSchemeName = "";

            nTocEnable = 2;
            tocFnt = new nhFont();
            tocParaFmt = new nhParaFmt();
            tocStyles = new tocStyle();

            bHdSnsEnable = true;
            hdSnFnt = new nhFont();
            hdSns = new List<headingSn>();
            for (int i = 0; i < 9; i++)
            {
                headingSn hsn = new headingSn();

                fillHeadingSnDefault(hsn, i);

                hdSns.Add(hsn);
            }

            bHdStylesEnable = true;
            hdStyleFnt = new nhFont();
            hdStyleParaFmt = new nhParaFmt();

            hdStyles = new List<headingStyle>();
            for (int i = 0; i < 9; i++)
            {
                headingStyle hs = new headingStyle();
                fillHeadingStyleDefault(hs, i, @"宋体");
                hdStyles.Add(hs);
            }

            nTizhuEnable = 2;
            tizhuFnt = new nhFont();
            tizhuParaFmt = new nhParaFmt();

            tblTizhuStyle = new tizhuStyle();
            inShpTizhuStyle = new tizhuStyle();


            bTblEnable = true;
            tblFont = new nhFont();
            //tblEveryParaFmt = new nhParaFmt();
            tblParaFmt = new nhParaFmt();

            bInShpEnable = true;
            inShpParaFmt = new nhParaFmt();


            nListParaEnable = 2;
            listParaFnt = new nhFont();
            listParaParaFmt = new nhParaFmt();

            outTblListParaStyle = new listParaStyle();
            inTblListParaStyle = new listParaStyle();

            // 
            nTextBodyEnable = 2;
            textBodyFnt = new nhFont();
            textBodyParaFmt = new nhParaFmt();

            outTblTextBodyStyle = new textBodyStyle();
            inTblTextBodyStyle = new textBodyStyle();


            nPgNumEnable = 2;
            pgNumStyle = new pageNumStyle();

            pgNumStyles = new List<pageNumStyle>();
            for (int i = 0; i < 9; i++)
            {
                pageNumStyle pg = new pageNumStyle();
                pgNumStyles.Add(pg);
            }


            m_hashSnStyle2Name.Clear();
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleNone, "");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleArabic, "1");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleUppercaseRoman, "I");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleLowercaseRoman, "i");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleUppercaseLetter, "A");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter, "a");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleSimpChinNum3, "一");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleSimpChinNum2, "壹");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleZodiac1, "甲");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleZodiac2, "子");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleOrdinal, "1st");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleCardinalText, "One");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleOrdinalText, "First");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleArabicLZ, "01");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleLegalLZ, "01");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleLegal, "1");
            m_hashSnStyle2Name.Add(Word.WdListNumberStyle.wdListNumberStyleNumberInCircle, "①");

            return;
        }

        public void clone(docPubScheme oth)
        {
            strSchemeName = oth.strSchemeName;

            nTocEnable = oth.nTocEnable;
            tocFnt.clone(oth.tocFnt);
            tocParaFmt.clone(oth.tocParaFmt);
            tocStyles.clone(oth.tocStyles);

            bHdSnsEnable = oth.bHdSnsEnable;
            hdSnFnt.clone(oth.hdSnFnt);
            headingSn hdSn = null, ohdSn = null;
            int nCnt = hdSns.Count;
            for (int i = 0; i < nCnt; i++ )
            {
                hdSn = (headingSn)hdSns[i];
                ohdSn = (headingSn)oth.hdSns[i];

                hdSn.clone(ohdSn);
            }

            bHdStylesEnable = oth.bHdStylesEnable;
            hdStyleFnt.clone(oth.hdStyleFnt);
            hdStyleParaFmt.clone(oth.hdStyleParaFmt);

            headingStyle hdStyle = null, ohdStyle = null;
            nCnt = hdStyles.Count;
            for (int i = 0; i < nCnt; i++ )
            {
                hdStyle = (headingStyle)hdStyles[i];
                ohdStyle = (headingStyle)oth.hdStyles[i];

                hdStyle.clone(ohdStyle);
            }

            nTizhuEnable = oth.nTizhuEnable;
            tizhuFnt.clone(oth.tizhuFnt);
            tizhuParaFmt.clone(oth.tizhuParaFmt);

            tblTizhuStyle.clone(oth.tblTizhuStyle);
            inShpTizhuStyle.clone(oth.inShpTizhuStyle);

            // 
            bTblEnable = oth.bTblEnable;
            tblFont.clone(oth.tblFont);
            //tblEveryParaFmt.clone(oth.tblEveryParaFmt);
            tblParaFmt.clone(oth.tblParaFmt);

            bInShpEnable = oth.bInShpEnable;
            inShpParaFmt.clone(oth.inShpParaFmt);


            nListParaEnable = oth.nListParaEnable;
            listParaFnt.clone(oth.listParaFnt);
            listParaParaFmt.clone(oth.listParaParaFmt);
            outTblListParaStyle.clone(oth.outTblListParaStyle);
            inTblListParaStyle.clone(oth.inTblListParaStyle);

            nTextBodyEnable = oth.nTextBodyEnable;
            textBodyFnt.clone(oth.textBodyFnt);
            textBodyParaFmt.clone(oth.textBodyParaFmt);

            outTblTextBodyStyle.clone(oth.outTblTextBodyStyle);
            inTblTextBodyStyle.clone(oth.inTblTextBodyStyle);


            nPgNumEnable = oth.nPgNumEnable;
            pgNumStyle.clone(oth.pgNumStyle);

            pageNumStyle pgNum = null, opgNum = null;
            nCnt = pgNumStyles.Count;

            for (int i = 0; i < nCnt; i++)
            {
                pgNum = (pageNumStyle)pgNumStyles[i];
                opgNum = (pageNumStyle)oth.pgNumStyles[i];

                // 
                pgNum.clone(opgNum);
            }

            return;
        }

        public void copy2(docPubScheme oth)
        {
            oth.strSchemeName = strSchemeName;
            oth.nTocEnable = nTocEnable;
            oth.tocFnt.clone(tocFnt);
            oth.tocParaFmt.clone(tocParaFmt);
            oth.tocStyles.clone(tocStyles);

            oth.bHdSnsEnable = bHdSnsEnable;
            oth.hdSnFnt.clone(hdSnFnt);

            headingSn hdSn = null, ohdSn = null;
            int nCnt = oth.hdSns.Count;
            for (int i = 0; i < nCnt; i++)
            {
                hdSn = (headingSn)hdSns[i];
                ohdSn = (headingSn)oth.hdSns[i];

                ohdSn.clone(hdSn);
            }

            oth.bHdStylesEnable = bHdStylesEnable;
            oth.hdStyleFnt.clone(hdStyleFnt);
            oth.hdStyleParaFmt.clone(hdStyleParaFmt);

            headingStyle hdStyle = null, ohdStyle = null;
            nCnt = oth.hdStyles.Count;
            for (int i = 0; i < nCnt; i++)
            {
                hdStyle = (headingStyle)hdStyles[i];
                ohdStyle = (headingStyle)oth.hdStyles[i];

                ohdStyle.clone(hdStyle);
            }

            oth.nTizhuEnable = nTizhuEnable;
            oth.tizhuFnt.clone(tizhuFnt);
            oth.tizhuParaFmt.clone(tizhuParaFmt);

            oth.tblTizhuStyle.clone(tblTizhuStyle);
            oth.inShpTizhuStyle.clone(inShpTizhuStyle);


            oth.bTblEnable = bTblEnable;
            oth.tblFont.clone(tblFont);
            //oth.tblEveryParaFmt.clone(tblEveryParaFmt);
            oth.tblParaFmt.clone(tblParaFmt);

            oth.bInShpEnable = bInShpEnable;
            oth.inShpParaFmt.clone(inShpParaFmt);


            oth.nListParaEnable = nListParaEnable;
            oth.listParaFnt.clone(listParaFnt);
            oth.listParaParaFmt.clone(listParaParaFmt);
            oth.outTblListParaStyle.clone(outTblListParaStyle);
            oth.inTblListParaStyle.clone(inTblListParaStyle);

            oth.nTextBodyEnable = nTextBodyEnable;
            oth.textBodyFnt.clone(textBodyFnt);
            oth.textBodyParaFmt.clone(textBodyParaFmt);

            oth.outTblTextBodyStyle.clone(outTblTextBodyStyle);
            oth.inTblTextBodyStyle.clone(inTblTextBodyStyle);


            oth.nPgNumEnable = nPgNumEnable;
            oth.pgNumStyle.clone(pgNumStyle);

            pageNumStyle pgNum = null, opgNum = null;
            nCnt = oth.pgNumStyles.Count;

            for (int i = 0; i < nCnt; i++)
            {
                pgNum = (pageNumStyle)pgNumStyles[i];
                opgNum = (pageNumStyle)oth.pgNumStyles[i];

                // 
                opgNum.clone(pgNum);
            }

            return; 
        }


        public int formatString(RichTextBox rchTxt, String strPreBlanks)
        {
            int nCnt = 0;

            String strCurLevelBlanks = strPreBlanks;
            String strNextLevelBlanks = strCurLevelBlanks + "    ";

            // 目录
            switch (nTocEnable)
            {
                case (int)CheckState.Checked:
                    rchTxt.AppendText("目录:全部启用[章节目录|图目录|表目录]\r\n");

                    nCnt = tocFnt.getSetCount();
                    if (nCnt > 0)
                    {
                        rchTxt.AppendText(strNextLevelBlanks + "统一字体：\r\n");
                        tocFnt.formatString(rchTxt, strNextLevelBlanks);
                    }

                    nCnt = tocParaFmt.getSetCount();
                    if (nCnt > 0)
                    {
                        rchTxt.AppendText(strNextLevelBlanks + "统一段落：\r\n");
                        tocParaFmt.formatString(rchTxt, strNextLevelBlanks);
                    }

                    break;

                case (int)CheckState.Indeterminate:
                    rchTxt.AppendText("目录:\r\n");

                    tocStyles.formatString(rchTxt, strNextLevelBlanks);

                    break;

                case (int)CheckState.Unchecked:
                    rchTxt.AppendText("目录:停用\r\n");
                    break;

                default:
                    break;
            }

            rchTxt.AppendText("\r\n");

            // 章节序号
            formatHdSnString(rchTxt, strPreBlanks);

            rchTxt.AppendText("\r\n");

            // 章节样式
            if (bHdStylesEnable)
            {
                rchTxt.AppendText("章节样式:启用\r\n");

                hdStyleFnt.formatString(rchTxt, strNextLevelBlanks);
                hdStyleParaFmt.formatString(rchTxt, strNextLevelBlanks);

                for (int i = 0; i < 9; i++ )
                {
                    rchTxt.AppendText(strNextLevelBlanks + (i + 1) + "级:");
                    hdStyles[i].formatString(rchTxt, strNextLevelBlanks);
                }

            }
            else
            {
                rchTxt.AppendText("章节样式:停用\r\n");
            }

            rchTxt.AppendText("\r\n");

            // 图
            if (bInShpEnable)
            {
                rchTxt.AppendText("图:启用\r\n");

                inShpParaFmt.formatString(rchTxt, strNextLevelBlanks);
            }
            else
            {
                rchTxt.AppendText("图:停用\r\n");
            }

            rchTxt.AppendText("\r\n");

            // 表
            if (bTblEnable)
            {
                rchTxt.AppendText("表:启用\r\n");
                tblFont.formatString(rchTxt, strNextLevelBlanks);
                tblParaFmt.formatString(rchTxt, strNextLevelBlanks);
            }
            else
            {
                rchTxt.AppendText("表:停用\r\n");
            }

            rchTxt.AppendText("\r\n");

            // 题注
            switch (nTizhuEnable)
            {
                case (int)CheckState.Checked:
                    rchTxt.AppendText("题注:全部启用[图|表]\r\n");

                    rchTxt.AppendText(strNextLevelBlanks + "题注:图\r\n");
                    inShpTizhuStyle.formatString(rchTxt, strNextLevelBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "题注:表\r\n");
                    tblTizhuStyle.formatString(rchTxt, strNextLevelBlanks);

                    tizhuFnt.formatString(rchTxt, strNextLevelBlanks);
                    tizhuParaFmt.formatString(rchTxt, strNextLevelBlanks);

                    break;

                case (int)CheckState.Indeterminate:
                    rchTxt.AppendText("题注:\r\n");
                    rchTxt.AppendText(strNextLevelBlanks + "图\r\n");
                    inShpTizhuStyle.formatString(rchTxt, strNextLevelBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "表\r\n");
                    tblTizhuStyle.formatString(rchTxt, strNextLevelBlanks);

                    tizhuFnt.formatString(rchTxt, strNextLevelBlanks);
                    tizhuParaFmt.formatString(rchTxt, strNextLevelBlanks);

                    break;

                case (int)CheckState.Unchecked:
                    rchTxt.AppendText("题注:全部停用[图|表]\r\n");
                    break;

                default:
                    break;
            }

            rchTxt.AppendText("\r\n");

            // 序号段落
            switch (nListParaEnable)
            {
                case (int)CheckState.Checked:
                    rchTxt.AppendText("序号段落:全部启用[表内|表外]\r\n");

                    rchTxt.AppendText(strNextLevelBlanks + "统一字体：\r\n");
                    listParaFnt.formatString(rchTxt, strNextLevelBlanks);

                    rchTxt.AppendText(strNextLevelBlanks + "统一段落：\r\n");
                    listParaParaFmt.formatString(rchTxt, strNextLevelBlanks);

                    rchTxt.AppendText("\r\n");
                    rchTxt.AppendText(strNextLevelBlanks + "序号段落:表内\r\n");
                    inTblListParaStyle.formatString(rchTxt, strPreBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "序号段落:表外\r\n");
                    outTblListParaStyle.formatString(rchTxt, strPreBlanks);

                    break;

                case (int)CheckState.Indeterminate:
                    rchTxt.AppendText("序号段落:\r\n");

                    rchTxt.AppendText(strNextLevelBlanks + "表内\r\n");
                    inTblListParaStyle.formatString(rchTxt, strNextLevelBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "表外\r\n");
                    outTblListParaStyle.formatString(rchTxt, strNextLevelBlanks);

                    break;

                case (int)CheckState.Unchecked:
                    rchTxt.AppendText("序号段落:停用\r\n");
                    break;

                default:
                    break;
            }

            rchTxt.AppendText("\r\n");
            // 正文
            switch (nTextBodyEnable)
            {
                case (int)CheckState.Checked:
                    rchTxt.AppendText("正文:全部启用[表内|表外]\r\n");

                    rchTxt.AppendText(strNextLevelBlanks + "统一字体：\r\n");
                    textBodyFnt.formatString(rchTxt, strNextLevelBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "统一段落：\r\n");
                    textBodyParaFmt.formatString(rchTxt, strNextLevelBlanks);

                    rchTxt.AppendText("\r\n");
                    rchTxt.AppendText(strNextLevelBlanks + "表内\r\n");
                    inTblTextBodyStyle.formatString(rchTxt, strNextLevelBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "表外\r\n");
                    outTblTextBodyStyle.formatString(rchTxt, strNextLevelBlanks);

                    break;

                case (int)CheckState.Indeterminate:
                    rchTxt.AppendText("正文:单项启用\r\n");

                    rchTxt.AppendText(strNextLevelBlanks + "表内\r\n");
                    inTblTextBodyStyle.formatString(rchTxt, strNextLevelBlanks);
                    rchTxt.AppendText(strNextLevelBlanks + "表外\r\n");
                    outTblTextBodyStyle.formatString(rchTxt, strNextLevelBlanks);

                    break;

                case (int)CheckState.Unchecked:
                    rchTxt.AppendText("序号段落:停用\r\n");
                    break;

                default:
                    break;
            }

            rchTxt.AppendText("\r\n");

            // 页码
            switch (nPgNumEnable)
            {
                case (int)CheckState.Checked:
                    rchTxt.AppendText("页码:全部启用[1-9节]\r\n");

                    pgNumStyle.formatString(rchTxt, strNextLevelBlanks);

                    break;  

                case (int)CheckState.Indeterminate:
                    rchTxt.AppendText("页码:\r\n");

                    for (int i = 0; i < 9; i++ )
                    {
                        rchTxt.AppendText(strNextLevelBlanks + "第" + (i + 1) + "节:");
                        pgNumStyles[i].formatString(rchTxt, strNextLevelBlanks);
                    }

                    break;

                case (int)CheckState.Unchecked:
                    rchTxt.AppendText("页码:停用\r\n");
                    break;

                default:
                    break;
            }

            return 0;
        }


        public int formatHdSnString(RichTextBox rchTxt, String strPreBlanks)
        {
            String strRet = "";
            String strNextLevelBlanks = strPreBlanks + "    ";

            if (bHdSnsEnable)
            {
                strRet += "章节序号：启用\r\n";
            }
            else
            {
                strRet += "章节序号：停用\r\n";
                rchTxt.AppendText(strRet);
                return 0;
            }

            strRet += strNextLevelBlanks + "格式：\r\n";
            rchTxt.AppendText(strRet);

            String strListLevels = "", strTmp = "", strTmp2 = "", strInfo = "";
            int nLen1 = 0, nLen2 = 0;

            String strNumFormat = "", strPreview = "";
            for (int i = 0; i < 9; i++)
            {
                strNumFormat = hdSns[i].NumberFormat;
                strPreview = buildHeadingSnPreview(hdSns, i);

                strTmp2 = strNextLevelBlanks + (i + 1) + "级：" + strPreview;

                nLen1 = Encoding.GetEncoding("GB2312").GetBytes(strTmp2).Length;
                nLen2 = strTmp2.Length;

                strInfo = strTmp2.PadRight(32 - (nLen1 - nLen2));

                strListLevels = strInfo + "(" + strNumFormat + ")\r\n";

                rchTxt.AppendText(strListLevels);
            }

            rchTxt.AppendText(strRet);

            hdSnFnt.formatString(rchTxt,strPreBlanks);

            return 0;
        }


        private String buildHeadingSnPreview(List<headingSn> curListLevels, int nCurIndex)
        {
            headingSn curListLvl = null;
            String strDefInput = curListLevels[nCurIndex].NumberFormat;
            String strItem = "";

            // for (int i = nCurIndex; i > -1; i--) // 遍历
            // for (int i = 0; i <= nCurIndex; i++) // 遍历
            curListLvl = curListLevels[nCurIndex];

            if (curListLvl.NumberStyle == Word.WdListNumberStyle.wdListNumberStyleLegalLZ) // 特定格式
            {
                for (int j = 0; j <= nCurIndex; j++)
                {   // 
                    if (curListLevels[j].NumberStyleSel != Word.WdListNumberStyle.wdListNumberStyleArabicLZ)  // 特定格式
                    {
                        strItem = "1"; // 名称
                    }
                    else
                    {
                        strItem = "01"; // 名称
                    }

                    strDefInput = strDefInput.Replace("%" + (j + 1), strItem); // 转换
                }
                //break;
            }
            else if (curListLvl.NumberStyle == Word.WdListNumberStyle.wdListNumberStyleLegal) // 特定格式
            {
                for (int j = 0; j <= nCurIndex; j++) // 遍历
                {
                    if (curListLevels[j].NumberStyleSel != Word.WdListNumberStyle.wdListNumberStyleArabicLZ) // 特定格式
                    {
                        strItem = "1"; // 名称
                    }
                    else
                    {
                        strItem = "01"; // 名称
                    }

                    strDefInput = strDefInput.Replace("%" + (j + 1), strItem); // 转换
                }
                //break;
            }
            else
            {
                for (int i = 0; i <= nCurIndex; i++) // 遍历
                {
                    curListLvl = curListLevels[i];
                    if (m_hashSnStyle2Name.Contains(curListLvl.NumberStyle)) // 判断
                    {
                        strItem = (String)m_hashSnStyle2Name[curListLvl.NumberStyle];

                        if (strItem != null)
                        {
                            strDefInput = strDefInput.Replace("%" + (i + 1), strItem); // 转换
                        }
                    }
                }
            }

            return strDefInput;
        }

    }// class


}// namespace
