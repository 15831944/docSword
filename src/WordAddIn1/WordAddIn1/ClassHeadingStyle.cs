using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using OfficeTools.Common;


namespace OfficeAssist
{
    // 自定义的章节序号类
    public class ClassHeadingStyle
    {
        public Boolean                  m_bFntAssigned; // 字体是否赋值的标志
        public Boolean                  m_bParaFmtAssigned;// 段落样式是否赋值的标志
        public ClassFont                m_fnt;// 字体信息类
        public ClassParagraphFormat     m_paraFmt;// 段落信息类

        public ClassHeadingStyle()
        {
            m_fnt = new ClassFont();
            m_paraFmt = new ClassParagraphFormat();

            int nInit = 0;

            m_fnt.AllCaps = nInit; // 保存字体同名参数信息
            m_fnt.Animation = (Word.WdAnimation)nInit;//.wdAnimationNone;// 保存字体同名参数信息
            m_fnt.Bold = nInit;// 保存字体同名参数信息
            m_fnt.BoldBi = nInit;// 保存字体同名参数信息

            m_fnt.Color = Word.WdColor.wdColorAutomatic;// 保存字体同名参数信息
            m_fnt.ColorIndex = (Word.WdColorIndex)nInit;//.wdAuto;// 保存字体同名参数信息
            m_fnt.ColorIndexBi = (Word.WdColorIndex)nInit;//.wdAuto;// 保存字体同名参数信息
            m_fnt.DiacriticColor = Word.WdColor.wdColorAutomatic;// 保存字体同名参数信息
            m_fnt.DisableCharacterSpaceGrid = false;// 保存字体同名参数信息
            m_fnt.DoubleStrikeThrough = nInit;// 保存字体同名参数信息
            m_fnt.Emboss = nInit;// 保存字体同名参数信息
            m_fnt.EmphasisMark = Word.WdEmphasisMark.wdEmphasisMarkNone;// 保存字体同名参数信息

            m_fnt.Engrave = nInit;// 保存字体同名参数信息
            m_fnt.Hidden = nInit;// 保存字体同名参数信息
            m_fnt.Italic = nInit;// 保存字体同名参数信息
            m_fnt.ItalicBi = nInit;// 保存字体同名参数信息

            m_fnt.NameBi = "";// 保存字体同名参数信息
            m_fnt.Outline = nInit;// 保存字体同名参数信息

            m_fnt.Position = nInit;// 保存字体同名参数信息

            m_fnt.Shadow = nInit;// 保存字体同名参数信息

            m_fnt.SmallCaps = nInit;// 保存字体同名参数信息
            m_fnt.Spacing = (float)nInit;// 保存字体同名参数信息
            m_fnt.StrikeThrough = nInit;// 保存字体同名参数信息
            m_fnt.Subscript = nInit;// 保存字体同名参数信息
            m_fnt.Superscript = nInit;// 保存字体同名参数信息
            m_fnt.Underline = Word.WdUnderline.wdUnderlineNone;// 保存字体同名参数信息
            m_fnt.UnderlineColor = Word.WdColor.wdColorAutomatic;// 保存字体同名参数信息
            m_fnt.DisableCharacterSpaceGrid = false;// 保存字体同名参数信息
            
            // default
            m_fnt.NameFarEast = "宋体";// 保存字体同名参数信息
            m_fnt.NameAscii = "宋体"; // "+西文正文";// 保存字体同名参数信息
            //m_fnt.NameOther = "+正文 CS 字体"; // "+西文正文";// 保存字体同名参数信息
            m_fnt.Name = "宋体";// 保存字体同名参数信息
            m_fnt.Size = 10.0f;// 保存字体同名参数信息

            m_fnt.Kerning = 1;// 保存字体同名参数信息
            m_fnt.Scaling = 100;// 保存字体同名参数信息


            // 保存段落同名参数信息
            m_paraFmt.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;// 保存段落同名参数信息
            m_paraFmt.CharacterUnitFirstLineIndent = 2;// 保存段落同名参数信息
            m_paraFmt.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;// 保存段落同名参数信息

            m_paraFmt.Hyphenation = -1; // true// 保存段落同名参数信息
            m_paraFmt.AutoAdjustRightIndent = -1; // true// 保存段落同名参数信息
            m_paraFmt.FarEastLineBreakControl = -1; // true// 保存段落同名参数信息
            m_paraFmt.WordWrap = -1;// true// 保存段落同名参数信息
            m_paraFmt.HangingPunctuation = -1; // true// 保存段落同名参数信息
            m_paraFmt.AddSpaceBetweenFarEastAndAlpha = -1; // true// 保存段落同名参数信息
            m_paraFmt.AddSpaceBetweenFarEastAndDigit = -1; // true// 保存段落同名参数信息
            m_paraFmt.LineSpacing = 12.0f;// 保存段落同名参数信息
            m_paraFmt.BaseLineAlignment = Word.WdBaselineAlignment.wdBaselineAlignAuto;// 保存段落同名参数信息
            m_paraFmt.ReadingOrder = Word.WdReadingOrder.wdReadingOrderLtr;// 保存段落同名参数信息
            m_paraFmt.KeepTogether = 0;// 保存段落同名参数信息
            m_paraFmt.KeepWithNext = -1;// 保存段落同名参数信息
            m_paraFmt.SpaceAfter = 3.0f;// 保存段落同名参数信息
            m_paraFmt.SpaceAfterAuto = 0;// 保存段落同名参数信息
            m_paraFmt.SpaceBefore = 12.0f;// 保存段落同名参数信息
            m_paraFmt.SpaceBeforeAuto = 0;// 保存段落同名参数信息

            // 
            m_bFntAssigned = false;// 初始化值
            m_bParaFmtAssigned = false;// 初始化值

            return;
        }

    }
}
