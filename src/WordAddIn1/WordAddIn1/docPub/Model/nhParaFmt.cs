﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OfficeTools.Common;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

using System.Xml;
using System.Xml.Serialization;
using System.Collections;
using System.Windows.Forms;

namespace OfficeAssist.docPub.Model
{
    [Serializable]
    //[XmlRootAttribute]
    [XmlType]
    public class nhParaFmt
    {
        //[XmlIgnoreAttribute]
        public enum euMembers
        {
            Zero = 0,
            AddSpaceBetweenFarEastAndAlpha,
            AddSpaceBetweenFarEastAndDigit,
            Alignment,
            AutoAdjustRightIndent,
            BaseLineAlignment,
            Borders,
            CharacterUnitFirstLineIndent,
            CharacterUnitLeftIndent,
            CharacterUnitRightIndent,
            DisableLineHeightGrid,
            Duplicate,
            FarEastLineBreakControl,
            FirstLineIndent,
            HalfWidthPunctuationOnTopOfLine,
            HangingPunctuation,
            Hyphenation,
            KeepTogether,
            KeepWithNext,
            LeftIndent,
            LineSpacing,
            LineSpacingRule,
            LineUnitAfter,
            LineUnitBefore,
            MirrorIndents,
            NoLineNumber,
            OutlineLevel,
            PageBreakBefore,
            ReadingOrder,
            RightIndent,
            Shading,
            SpaceAfter,
            SpaceAfterAuto,
            SpaceBefore,
            SpaceBeforeAuto,
            TabStops,
            TextboxTightWrap,
            WidowControl,
            WordWrap,
        }

        [XmlAttribute]
        public Boolean bEnable = true;

        [XmlIgnoreAttribute]
        public HashSet<int> setMembers = null;

        [XmlAttribute]
        public int AddSpaceBetweenFarEastAndAlpha;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int AddSpaceBetweenFarEastAndDigit;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int Alignment; // WdParagraphAlignment Alignment;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        // public Application Application { get; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int AutoAdjustRightIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int BaseLineAlignment;// WdBaselineAlignment BaseLineAlignment;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        //[XmlIgnoreAttribute]
        //public ClassBorders Borders = new ClassBorders();
        // public Borders Borders;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float CharacterUnitFirstLineIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float CharacterUnitLeftIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float CharacterUnitRightIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        // public int Creator;// { get; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int DisableLineHeightGrid;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        //[XmlIgnoreAttribute]
        //public ParagraphFormat Duplicate;// { get; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int FarEastLineBreakControl;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float FirstLineIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int HalfWidthPunctuationOnTopOfLine;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int HangingPunctuation;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int Hyphenation;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int KeepTogether;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int KeepWithNext;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float LeftIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float LineSpacing;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int LineSpacingRule;// WdLineSpacing LineSpacingRule;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float LineUnitAfter;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float LineUnitBefore;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int MirrorIndents;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int NoLineNumber;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int OutlineLevel;// WdOutlineLevel OutlineLevel;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int PageBreakBefore;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        // public dynamic Parent { get; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int ReadingOrder;// WdReadingOrder ReadingOrder;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float RightIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        //[XmlIgnoreAttribute]
        //public ClassShading Shading = new ClassShading();
        // public Shading Shading;// { get; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float SpaceAfter;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int SpaceAfterAuto;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public float SpaceBefore;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int SpaceBeforeAuto;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        //[XmlIgnoreAttribute]
        //public TabStops TabStops;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int TextboxTightWrap; // WdTextboxTightWrap TextboxTightWrap;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int WidowControl;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        [XmlAttribute]
        public int WordWrap;// { get; set; }// 同WORD.ParagraphFormat的同名参数


        public nhParaFmt()
        {
            Hyphenation = -1; // true
            AutoAdjustRightIndent = -1; // true
            FarEastLineBreakControl = -1; // true
            WordWrap = -1;// true
            HangingPunctuation = -1; // true
            AddSpaceBetweenFarEastAndAlpha = -1; // true
            AddSpaceBetweenFarEastAndDigit = -1; // true
            LineSpacing = 12.0f;
            LineSpacingRule = (int)Word.WdLineSpacing.wdLineSpaceSingle;
            BaseLineAlignment = (int)WdBaselineAlignment.wdBaselineAlignAuto;
            ReadingOrder = (int)WdReadingOrder.wdReadingOrderLtr;
            KeepTogether = 0;
            KeepWithNext = -1;
            SpaceAfter = 0.0f;
            SpaceAfterAuto = 0;
            SpaceBefore = 0.0f;
            SpaceBeforeAuto = 0;
            OutlineLevel = 10;

            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }
            else
            {
                setMembers.Clear();
            }

            return;
        }

        public int getSetCount()
        {
            if (setMembers == null)
            {
                return 0;
            }

            return setMembers.Count;
        }

        public Boolean isSet(euMembers enMem)
        {
            if (setMembers != null)
            {
                return setMembers.Contains((int)enMem);
            }

            return false;
        }


        public void ClearSelMember()
        {
            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }
            else
            {
                setMembers.Clear();
            }

            return;
        }


        public Boolean AddSelMember(int nParaFmtMember)
        {
            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }

            Boolean bRet = setMembers.Add(nParaFmtMember);

            return bRet;
        }


        public Boolean RemoveSelMember(int nParaFmtMember)
        {
            if (setMembers == null)
            {
                setMembers = new HashSet<int>();
            }

            Boolean bRet = setMembers.Remove(nParaFmtMember);

            return bRet;
        }


        public int SelCopy2(Word.ParagraphFormat wParaFmt)
        {
            if (setMembers == null)
            {
                return 0;
            }

            euMembers euItem = euMembers.Zero;

            foreach (int nItem in setMembers)
            {
                euItem = (euMembers)nItem;

                switch (euItem)
                {
                    case euMembers.AddSpaceBetweenFarEastAndAlpha:
                        wParaFmt.AddSpaceBetweenFarEastAndAlpha = AddSpaceBetweenFarEastAndAlpha;
                        break;

                    case euMembers.AddSpaceBetweenFarEastAndDigit:
                        wParaFmt.AddSpaceBetweenFarEastAndDigit = AddSpaceBetweenFarEastAndDigit;
                        break;

                    case euMembers.Alignment:
                        wParaFmt.Alignment = (Word.WdParagraphAlignment)Alignment;
                        break;

                    case euMembers.AutoAdjustRightIndent:
                        wParaFmt.AutoAdjustRightIndent = AutoAdjustRightIndent;
                        break;

                    case euMembers.BaseLineAlignment:
                        wParaFmt.BaseLineAlignment = (Word.WdBaselineAlignment)BaseLineAlignment;
                        break;

                    case euMembers.Borders:
                        // wParaFmt.Borders = Borders;
                        break;

                    case euMembers.CharacterUnitFirstLineIndent:
                        wParaFmt.CharacterUnitFirstLineIndent = CharacterUnitFirstLineIndent;
                        break;

                    case euMembers.CharacterUnitLeftIndent:
                        wParaFmt.CharacterUnitLeftIndent = CharacterUnitLeftIndent;
                        break;

                    case euMembers.CharacterUnitRightIndent:
                        wParaFmt.CharacterUnitRightIndent = CharacterUnitRightIndent;
                        break;

                    case euMembers.DisableLineHeightGrid:
                        wParaFmt.DisableLineHeightGrid = DisableLineHeightGrid;
                        break;

                    case euMembers.Duplicate:
                        // wParaFmt.Duplicate = Duplicate;
                        break;

                    case euMembers.FarEastLineBreakControl:
                        wParaFmt.FarEastLineBreakControl = FarEastLineBreakControl;
                        break;

                    case euMembers.FirstLineIndent:
                        wParaFmt.FirstLineIndent = FirstLineIndent;
                        break;

                    case euMembers.HalfWidthPunctuationOnTopOfLine:
                        wParaFmt.HalfWidthPunctuationOnTopOfLine = HalfWidthPunctuationOnTopOfLine;
                        break;

                    case euMembers.HangingPunctuation:
                        wParaFmt.HangingPunctuation = HangingPunctuation;
                        break;

                    case euMembers.Hyphenation:
                        wParaFmt.Hyphenation = Hyphenation;
                        break;

                    case euMembers.KeepTogether:
                        wParaFmt.KeepTogether = KeepTogether;
                        break;

                    case euMembers.KeepWithNext:
                        wParaFmt.KeepWithNext = KeepWithNext;
                        break;

                    case euMembers.LeftIndent:
                        wParaFmt.LeftIndent = LeftIndent;
                        break;

                    case euMembers.LineSpacing:
                        wParaFmt.LineSpacing = LineSpacing;
                        break;

                    case euMembers.LineSpacingRule:
                        wParaFmt.LineSpacingRule = (Word.WdLineSpacing)LineSpacingRule;
                        break;

                    case euMembers.LineUnitAfter:
                        wParaFmt.LineUnitAfter = LineUnitAfter;
                        break;

                    case euMembers.LineUnitBefore:
                        wParaFmt.LineUnitBefore = LineUnitBefore;
                        break;

                    case euMembers.MirrorIndents:
                        wParaFmt.MirrorIndents = MirrorIndents;
                        break;

                    case euMembers.NoLineNumber:
                        wParaFmt.NoLineNumber = NoLineNumber;
                        break;

                    case euMembers.OutlineLevel:
                        wParaFmt.OutlineLevel = (Word.WdOutlineLevel)OutlineLevel;
                        break;

                    case euMembers.PageBreakBefore:
                        wParaFmt.PageBreakBefore = PageBreakBefore;
                        break;

                    case euMembers.ReadingOrder:
                        wParaFmt.ReadingOrder = (Word.WdReadingOrder) ReadingOrder;
                        break;

                    case euMembers.RightIndent:
                        wParaFmt.RightIndent = RightIndent;
                        break;

                    case euMembers.Shading:
                        //wParaFmt.Shading = Shading;
                        break;

                    case euMembers.SpaceAfter:
                        wParaFmt.SpaceAfter = SpaceAfter;
                        break;

                    case euMembers.SpaceAfterAuto:
                        wParaFmt.SpaceAfterAuto = SpaceAfterAuto;
                        break;

                    case euMembers.SpaceBefore:
                        wParaFmt.SpaceBefore = SpaceBefore;
                        break;

                    case euMembers.SpaceBeforeAuto:
                        wParaFmt.SpaceBeforeAuto = SpaceBeforeAuto;
                        break;

                    case euMembers.TabStops:
                        //wParaFmt.TabStops = TabStops;
                        break;

                    case euMembers.TextboxTightWrap:
                        wParaFmt.TextboxTightWrap = (Word.WdTextboxTightWrap)TextboxTightWrap;
                        break;

                    case euMembers.WidowControl:
                        wParaFmt.WidowControl = WidowControl;
                        break;

                    case euMembers.WordWrap:
                        wParaFmt.WordWrap = WordWrap;
                        break;

                    default:
                        break;
                }
            }

            return setMembers.Count;
        }


        public int SelCopy2(ClassParagraphFormat wParaFmt)
        {
            if (setMembers == null)
            {
                return 0;
            }

            wParaFmt.ClearSelMember();

            euMembers euItem = euMembers.Zero;

            foreach (int nItem in setMembers)
            {
                wParaFmt.AddSelMember(nItem);

                euItem = (euMembers)nItem;

                switch (euItem)
                {
                    case euMembers.AddSpaceBetweenFarEastAndAlpha:
                        wParaFmt.AddSpaceBetweenFarEastAndAlpha = AddSpaceBetweenFarEastAndAlpha;
                        break;

                    case euMembers.AddSpaceBetweenFarEastAndDigit:
                        wParaFmt.AddSpaceBetweenFarEastAndDigit = AddSpaceBetweenFarEastAndDigit;
                        break;

                    case euMembers.Alignment:
                        wParaFmt.Alignment = (Word.WdParagraphAlignment)Alignment;
                        break;

                    case euMembers.AutoAdjustRightIndent:
                        wParaFmt.AutoAdjustRightIndent = AutoAdjustRightIndent;
                        break;

                    case euMembers.BaseLineAlignment:
                        wParaFmt.BaseLineAlignment = (Word.WdBaselineAlignment)BaseLineAlignment;
                        break;

                    case euMembers.Borders:
                        // wParaFmt.Borders = Borders;
                        break;

                    case euMembers.CharacterUnitFirstLineIndent:
                        wParaFmt.CharacterUnitFirstLineIndent = CharacterUnitFirstLineIndent;
                        break;

                    case euMembers.CharacterUnitLeftIndent:
                        wParaFmt.CharacterUnitLeftIndent = CharacterUnitLeftIndent;
                        break;

                    case euMembers.CharacterUnitRightIndent:
                        wParaFmt.CharacterUnitRightIndent = CharacterUnitRightIndent;
                        break;

                    case euMembers.DisableLineHeightGrid:
                        wParaFmt.DisableLineHeightGrid = DisableLineHeightGrid;
                        break;

                    case euMembers.Duplicate:
                        // wParaFmt.Duplicate = Duplicate;
                        break;

                    case euMembers.FarEastLineBreakControl:
                        wParaFmt.FarEastLineBreakControl = FarEastLineBreakControl;
                        break;

                    case euMembers.FirstLineIndent:
                        wParaFmt.FirstLineIndent = FirstLineIndent;
                        break;

                    case euMembers.HalfWidthPunctuationOnTopOfLine:
                        wParaFmt.HalfWidthPunctuationOnTopOfLine = HalfWidthPunctuationOnTopOfLine;
                        break;

                    case euMembers.HangingPunctuation:
                        wParaFmt.HangingPunctuation = HangingPunctuation;
                        break;

                    case euMembers.Hyphenation:
                        wParaFmt.Hyphenation = Hyphenation;
                        break;

                    case euMembers.KeepTogether:
                        wParaFmt.KeepTogether = KeepTogether;
                        break;

                    case euMembers.KeepWithNext:
                        wParaFmt.KeepWithNext = KeepWithNext;
                        break;

                    case euMembers.LeftIndent:
                        wParaFmt.LeftIndent = LeftIndent;
                        break;

                    case euMembers.LineSpacing:
                        wParaFmt.LineSpacing = LineSpacing;
                        break;

                    case euMembers.LineSpacingRule:
                        wParaFmt.LineSpacingRule = (Word.WdLineSpacing)LineSpacingRule;
                        break;

                    case euMembers.LineUnitAfter:
                        wParaFmt.LineUnitAfter = LineUnitAfter;
                        break;

                    case euMembers.LineUnitBefore:
                        wParaFmt.LineUnitBefore = LineUnitBefore;
                        break;

                    case euMembers.MirrorIndents:
                        wParaFmt.MirrorIndents = MirrorIndents;
                        break;

                    case euMembers.NoLineNumber:
                        wParaFmt.NoLineNumber = NoLineNumber;
                        break;

                    case euMembers.OutlineLevel:
                        wParaFmt.OutlineLevel = (Word.WdOutlineLevel)OutlineLevel;
                        break;

                    case euMembers.PageBreakBefore:
                        wParaFmt.PageBreakBefore = PageBreakBefore;
                        break;

                    case euMembers.ReadingOrder:
                        wParaFmt.ReadingOrder = (Word.WdReadingOrder)ReadingOrder;
                        break;

                    case euMembers.RightIndent:
                        wParaFmt.RightIndent = RightIndent;
                        break;

                    case euMembers.Shading:
                        //wParaFmt.Shading = Shading;
                        break;

                    case euMembers.SpaceAfter:
                        wParaFmt.SpaceAfter = SpaceAfter;
                        break;

                    case euMembers.SpaceAfterAuto:
                        wParaFmt.SpaceAfterAuto = SpaceAfterAuto;
                        break;

                    case euMembers.SpaceBefore:
                        wParaFmt.SpaceBefore = SpaceBefore;
                        break;

                    case euMembers.SpaceBeforeAuto:
                        wParaFmt.SpaceBeforeAuto = SpaceBeforeAuto;
                        break;

                    case euMembers.TabStops:
                        //wParaFmt.TabStops = TabStops;
                        break;

                    case euMembers.TextboxTightWrap:
                        wParaFmt.TextboxTightWrap = (Word.WdTextboxTightWrap)TextboxTightWrap;
                        break;

                    case euMembers.WidowControl:
                        wParaFmt.WidowControl = WidowControl;
                        break;

                    case euMembers.WordWrap:
                        wParaFmt.WordWrap = WordWrap;
                        break;

                    default:
                        break;
                }
            }

            return setMembers.Count;
        }

        public String encode2String()
        {
            String strRet = "";

            strRet += "[ParaFmt_Start:ParaFmt_Start]";

            strRet += "[ParaFmt_AddSpaceBetweenFarEastAndAlpha:" + AddSpaceBetweenFarEastAndAlpha + "]";
            strRet += "[ParaFmt_AddSpaceBetweenFarEastAndDigit:" + AddSpaceBetweenFarEastAndDigit + "]";
            strRet += "[ParaFmt_Alignment:" + (int)Alignment + "]";
            strRet += "[ParaFmt_AutoAdjustRightIndent:" + AutoAdjustRightIndent + "]";
            strRet += "[ParaFmt_BaseLineAlignment:" + (int)BaseLineAlignment + "]";
            strRet += "[ParaFmt_CharacterUnitFirstLineIndent:" + CharacterUnitFirstLineIndent + "]";
            strRet += "[ParaFmt_CharacterUnitLeftIndent:" + CharacterUnitLeftIndent + "]";
            strRet += "[ParaFmt_CharacterUnitRightIndent:" + CharacterUnitRightIndent + "]";
            strRet += "[ParaFmt_DisableLineHeightGrid:" + DisableLineHeightGrid + "]";
            strRet += "[ParaFmt_FarEastLineBreakControl:" + FarEastLineBreakControl + "]";
            strRet += "[ParaFmt_FirstLineIndent:" + FirstLineIndent + "]";
            strRet += "[ParaFmt_HalfWidthPunctuationOnTopOfLine:" + HalfWidthPunctuationOnTopOfLine + "]";
            strRet += "[ParaFmt_HangingPunctuation:" + HangingPunctuation + "]";
            strRet += "[ParaFmt_Hyphenation:" + Hyphenation + "]";
            strRet += "[ParaFmt_KeepTogether:" + KeepTogether + "]";
            strRet += "[ParaFmt_KeepWithNext:" + KeepWithNext + "]";
            strRet += "[ParaFmt_LeftIndent:" + LeftIndent + "]";
            strRet += "[ParaFmt_LineSpacing:" + LineSpacing + "]";
            strRet += "[ParaFmt_LineSpacingRule:" + (int)LineSpacingRule + "]";
            strRet += "[ParaFmt_LineUnitAfter:" + LineUnitAfter + "]";
            strRet += "[ParaFmt_LineUnitBefore:" + LineUnitBefore + "]";
            strRet += "[ParaFmt_MirrorIndents:" + MirrorIndents + "]";
            strRet += "[ParaFmt_NoLineNumber:" + NoLineNumber + "]";
            strRet += "[ParaFmt_OutlineLevel:" + (int)OutlineLevel + "]";
            strRet += "[ParaFmt_PageBreakBefore:" + PageBreakBefore + "]";
            strRet += "[ParaFmt_ReadingOrder:" + (int)ReadingOrder + "]";
            strRet += "[ParaFmt_RightIndent:" + RightIndent + "]";
            strRet += "[ParaFmt_SpaceAfter:" + SpaceAfter + "]";
            strRet += "[ParaFmt_SpaceAfterAuto:" + SpaceAfterAuto + "]";
            strRet += "[ParaFmt_SpaceBefore:" + SpaceBefore + "]";
            strRet += "[ParaFmt_SpaceBeforeAuto:" + SpaceBeforeAuto + "]";
            strRet += "[ParaFmt_TextboxTightWrap:" + (int)TextboxTightWrap + "]";
            strRet += "[ParaFmt_WidowControl:" + WidowControl + "]";
            strRet += "[ParaFmt_WordWrap:" + WordWrap + "]";

            strRet += "[ParaFmt_End:ParaFmt_End]";

            return strRet;
        }

        public int decodeFromString(Hashtable hashFeatures)
        {
            if (hashFeatures == null || hashFeatures.Count == 0)
            {
                return 1;
            }

            String strVal = "";
            int nVal = 0, nDefaultVal = 0;
            float fVal = 0.0f;

            strVal = (String)hashFeatures["ParaFmt_AddSpaceBetweenFarEastAndAlpha"];
            if (int.TryParse(strVal, out nVal))
            {
                AddSpaceBetweenFarEastAndAlpha = nVal;
            }
            else
            {
                AddSpaceBetweenFarEastAndAlpha = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_AddSpaceBetweenFarEastAndDigit"];
            if (int.TryParse(strVal, out nVal))
            {
                AddSpaceBetweenFarEastAndDigit = nVal;
            }
            else
            {
                AddSpaceBetweenFarEastAndDigit = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_Alignment"];
            if (int.TryParse(strVal, out nVal))
            {
                Alignment = nVal;
            }
            else
            {
                Alignment = (int)WdParagraphAlignment.wdAlignParagraphDistribute;
            }

            strVal = (String)hashFeatures["ParaFmt_AutoAdjustRightIndent"];
            if (int.TryParse(strVal, out nVal))
            {
                AutoAdjustRightIndent = nVal;
            }
            else
            {
                AutoAdjustRightIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_BaseLineAlignment"];
            if (int.TryParse(strVal, out nVal))
            {
                BaseLineAlignment = nVal;
            }
            else
            {
                BaseLineAlignment = (int)WdBaselineAlignment.wdBaselineAlignAuto;
            }

            strVal = (String)hashFeatures["ParaFmt_CharacterUnitFirstLineIndent"];
            if (float.TryParse(strVal, out fVal))
            {
                CharacterUnitFirstLineIndent = fVal;
            }
            else
            {
                CharacterUnitFirstLineIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_CharacterUnitLeftIndent"];
            if (float.TryParse(strVal, out fVal))
            {
                CharacterUnitLeftIndent = fVal;
            }
            else
            {
                CharacterUnitLeftIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_CharacterUnitRightIndent"];
            if (float.TryParse(strVal, out fVal))
            {
                CharacterUnitRightIndent = fVal;
            }
            else
            {
                CharacterUnitRightIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_DisableLineHeightGrid"];
            if (int.TryParse(strVal, out nVal))
            {
                DisableLineHeightGrid = nVal;
            }
            else
            {
                DisableLineHeightGrid = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_FarEastLineBreakControl"];
            if (int.TryParse(strVal, out nVal))
            {
                FarEastLineBreakControl = nVal;
            }
            else
            {
                FarEastLineBreakControl = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_FirstLineIndent"];
            if (float.TryParse(strVal, out fVal))
            {
                FirstLineIndent = fVal;
            }
            else
            {
                FirstLineIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_HalfWidthPunctuationOnTopOfLine"];
            if (int.TryParse(strVal, out nVal))
            {
                HalfWidthPunctuationOnTopOfLine = nVal;
            }
            else
            {
                HalfWidthPunctuationOnTopOfLine = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_HangingPunctuation"];
            if (int.TryParse(strVal, out nVal))
            {
                HangingPunctuation = nVal;
            }
            else
            {
                HangingPunctuation = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_Hyphenation"];
            if (int.TryParse(strVal, out nVal))
            {
                Hyphenation = nVal;
            }
            else
            {
                Hyphenation = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_KeepTogether"];
            if (int.TryParse(strVal, out nVal))
            {
                KeepTogether = nVal;
            }
            else
            {
                KeepTogether = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_KeepWithNext"];
            if (int.TryParse(strVal, out nVal))
            {
                KeepWithNext = nVal;
            }
            else
            {
                KeepWithNext = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_LeftIndent"];
            if (float.TryParse(strVal, out fVal))
            {
                LeftIndent = fVal;
            }
            else
            {
                LeftIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_LineSpacing"];
            if (float.TryParse(strVal, out fVal))
            {
                LineSpacing = fVal;
            }
            else
            {
                LineSpacing = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_LineSpacingRule"];
            if (int.TryParse(strVal, out nVal))
            {
                LineSpacingRule = nVal;
            }
            else
            {
                LineSpacingRule = (int)WdLineSpacing.wdLineSpaceSingle;
            }

            strVal = (String)hashFeatures["ParaFmt_LineUnitAfter"];
            if (float.TryParse(strVal, out fVal))
            {
                LineUnitAfter = fVal;
            }
            else
            {
                LineUnitAfter = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_LineUnitBefore"];
            if (float.TryParse(strVal, out fVal))
            {
                LineUnitBefore = fVal;
            }
            else
            {
                LineUnitBefore = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_MirrorIndents"];
            if (int.TryParse(strVal, out nVal))
            {
                MirrorIndents = nVal;
            }
            else
            {
                MirrorIndents = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_NoLineNumber"];
            if (int.TryParse(strVal, out nVal))
            {
                NoLineNumber = nVal;
            }
            else
            {
                NoLineNumber = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_OutlineLevel"];
            if (int.TryParse(strVal, out nVal))
            {
                OutlineLevel = nVal;
            }
            else
            {
                OutlineLevel = (int)WdOutlineLevel.wdOutlineLevelBodyText;
            }

            strVal = (String)hashFeatures["ParaFmt_PageBreakBefore"];
            if (int.TryParse(strVal, out nVal))
            {
                PageBreakBefore = nVal;
            }
            else
            {
                PageBreakBefore = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_ReadingOrder"];
            if (int.TryParse(strVal, out nVal))
            {
                ReadingOrder = nVal;
            }
            else
            {
                ReadingOrder = (int)WdReadingOrder.wdReadingOrderLtr;
            }

            strVal = (String)hashFeatures["ParaFmt_RightIndent"];
            if (float.TryParse(strVal, out fVal))
            {
                RightIndent = fVal;
            }
            else
            {
                RightIndent = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_SpaceAfter"];
            if (float.TryParse(strVal, out fVal))
            {
                SpaceAfter = fVal;
            }
            else
            {
                SpaceAfter = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_SpaceAfterAuto"];
            if (int.TryParse(strVal, out nVal))
            {
                SpaceAfterAuto = nVal;
            }
            else
            {
                SpaceAfterAuto = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_SpaceBefore"];
            if (float.TryParse(strVal, out fVal))
            {
                SpaceBefore = fVal;
            }
            else
            {
                SpaceBefore = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_SpaceBeforeAuto"];
            if (int.TryParse(strVal, out nVal))
            {
                SpaceBeforeAuto = nVal;
            }
            else
            {
                SpaceBeforeAuto = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_TextboxTightWrap"];
            if (int.TryParse(strVal, out nVal))
            {
                TextboxTightWrap = nVal;
            }
            else
            {
                TextboxTightWrap = (int)WdTextboxTightWrap.wdTightNone;
            }

            strVal = (String)hashFeatures["ParaFmt_WidowControl"];
            if (int.TryParse(strVal, out nVal))
            {
                WidowControl = nVal;
            }
            else
            {
                WidowControl = nDefaultVal;
            }

            strVal = (String)hashFeatures["ParaFmt_WordWrap"];
            if (int.TryParse(strVal, out nVal))
            {
                WordWrap = nVal;
            }
            else
            {
                WordWrap = nDefaultVal;
            }

            return 0;
        }

        public int decodeFromString(String strRet)
        {
            // 
            Hashtable hashFeatures = ClassOfficeCommon.Decode(strRet);

            if (hashFeatures == null || hashFeatures.Count == 0)
            {
                return 1;
            }

            int nRet = decodeFromString(hashFeatures);

            return nRet;
        }


        // 复制保存WORD.ParagraphFormat同名参数值
        public void clone(Word.ParagraphFormat srcParaFormat)
        {
            this.Alignment = (int)srcParaFormat.Alignment;// 复制保存WORD.ParagraphFormat同名参数值

            this.AutoAdjustRightIndent = srcParaFormat.AutoAdjustRightIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.BaseLineAlignment = (int)srcParaFormat.BaseLineAlignment;// 复制保存WORD.ParagraphFormat同名参数值

            //this.Borders.clone(srcParaFormat.Borders);// 复制保存WORD.ParagraphFormat同名参数值

            this.CharacterUnitFirstLineIndent = srcParaFormat.CharacterUnitFirstLineIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.CharacterUnitLeftIndent = srcParaFormat.CharacterUnitLeftIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.CharacterUnitRightIndent = srcParaFormat.CharacterUnitRightIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.DisableLineHeightGrid = srcParaFormat.DisableLineHeightGrid;// 复制保存WORD.ParagraphFormat同名参数值
            this.FarEastLineBreakControl = srcParaFormat.FarEastLineBreakControl;// 复制保存WORD.ParagraphFormat同名参数值
            this.FirstLineIndent = srcParaFormat.FirstLineIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.HalfWidthPunctuationOnTopOfLine = srcParaFormat.HalfWidthPunctuationOnTopOfLine;// 复制保存WORD.ParagraphFormat同名参数值
            this.HangingPunctuation = srcParaFormat.HangingPunctuation;// 复制保存WORD.ParagraphFormat同名参数值
            this.Hyphenation = srcParaFormat.Hyphenation;// 复制保存WORD.ParagraphFormat同名参数值
            this.KeepTogether = srcParaFormat.KeepTogether;// 复制保存WORD.ParagraphFormat同名参数值
            this.KeepWithNext = srcParaFormat.KeepWithNext;// 复制保存WORD.ParagraphFormat同名参数值
            this.LeftIndent = srcParaFormat.LeftIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.LineSpacing = srcParaFormat.LineSpacing;// 复制保存WORD.ParagraphFormat同名参数值
            this.LineSpacingRule = (int)srcParaFormat.LineSpacingRule;// 复制保存WORD.ParagraphFormat同名参数值
            this.LineUnitAfter = srcParaFormat.LineUnitAfter;// 复制保存WORD.ParagraphFormat同名参数值
            this.LineUnitBefore = srcParaFormat.LineUnitBefore;// 复制保存WORD.ParagraphFormat同名参数值
            this.MirrorIndents = srcParaFormat.MirrorIndents;// 复制保存WORD.ParagraphFormat同名参数值
            this.OutlineLevel = (int)srcParaFormat.OutlineLevel;// 复制保存WORD.ParagraphFormat同名参数值
            this.NoLineNumber = srcParaFormat.NoLineNumber;// 复制保存WORD.ParagraphFormat同名参数值
            this.PageBreakBefore = srcParaFormat.PageBreakBefore;// 复制保存WORD.ParagraphFormat同名参数值
            this.ReadingOrder = (int)srcParaFormat.ReadingOrder;// 复制保存WORD.ParagraphFormat同名参数值
            this.RightIndent = srcParaFormat.RightIndent;// 复制保存WORD.ParagraphFormat同名参数值

            // this.Shading = srcParaFormat.Shading;
            //this.Shading.clone(srcParaFormat.Shading);

            this.SpaceAfter = srcParaFormat.SpaceAfter;// 复制保存WORD.ParagraphFormat同名参数值
            this.SpaceAfterAuto = srcParaFormat.SpaceAfterAuto;// 复制保存WORD.ParagraphFormat同名参数值
            this.SpaceBefore = srcParaFormat.SpaceBefore;// 复制保存WORD.ParagraphFormat同名参数值
            this.SpaceBeforeAuto = srcParaFormat.SpaceBeforeAuto;// 复制保存WORD.ParagraphFormat同名参数值
            
            // this.TabStops = srcParaFormat.TabStops;

            this.TextboxTightWrap = (int)srcParaFormat.TextboxTightWrap;// 复制保存WORD.ParagraphFormat同名参数值
            this.WidowControl = srcParaFormat.WidowControl;// 复制保存WORD.ParagraphFormat同名参数值
            this.WordWrap = srcParaFormat.WordWrap;// 复制保存WORD.ParagraphFormat同名参数值

            return;
        }


        // 复制保存ClassParagraphFormat同名参数值
        public void clone(ClassParagraphFormat srcParaFormat)
        {
            this.Alignment = (int)srcParaFormat.Alignment;// 复制保存ClassParagraphFormat同名参数值

            this.AutoAdjustRightIndent = srcParaFormat.AutoAdjustRightIndent;// 复制保存ClassParagraphFormat同名参数值
            this.BaseLineAlignment = (int)srcParaFormat.BaseLineAlignment;// 复制保存ClassParagraphFormat同名参数值

            //this.Borders.clone(srcParaFormat.Borders);// 复制保存ClassParagraphFormat同名参数值

            this.CharacterUnitFirstLineIndent = srcParaFormat.CharacterUnitFirstLineIndent;// 复制保存ClassParagraphFormat同名参数值
            this.CharacterUnitLeftIndent = srcParaFormat.CharacterUnitLeftIndent;// 复制保存ClassParagraphFormat同名参数值
            this.CharacterUnitRightIndent = srcParaFormat.CharacterUnitRightIndent;// 复制保存ClassParagraphFormat同名参数值
            this.DisableLineHeightGrid = srcParaFormat.DisableLineHeightGrid;// 复制保存ClassParagraphFormat同名参数值
            this.FarEastLineBreakControl = srcParaFormat.FarEastLineBreakControl;// 复制保存ClassParagraphFormat同名参数值
            this.FirstLineIndent = srcParaFormat.FirstLineIndent;// 复制保存ClassParagraphFormat同名参数值
            this.HalfWidthPunctuationOnTopOfLine = srcParaFormat.HalfWidthPunctuationOnTopOfLine;// 复制保存ClassParagraphFormat同名参数值
            this.HangingPunctuation = srcParaFormat.HangingPunctuation;// 复制保存ClassParagraphFormat同名参数值
            this.Hyphenation = srcParaFormat.Hyphenation;// 复制保存ClassParagraphFormat同名参数值
            this.KeepTogether = srcParaFormat.KeepTogether;// 复制保存ClassParagraphFormat同名参数值
            this.KeepWithNext = srcParaFormat.KeepWithNext;// 复制保存ClassParagraphFormat同名参数值
            this.LeftIndent = srcParaFormat.LeftIndent;// 复制保存ClassParagraphFormat同名参数值
            this.LineSpacingRule = (int)srcParaFormat.LineSpacingRule;// 复制保存ClassParagraphFormat同名参数值
            this.LineSpacing = srcParaFormat.LineSpacing;// 复制保存ClassParagraphFormat同名参数值
            this.LineUnitAfter = srcParaFormat.LineUnitAfter;// 复制保存ClassParagraphFormat同名参数值
            this.LineUnitBefore = srcParaFormat.LineUnitBefore;// 复制保存ClassParagraphFormat同名参数值
            this.MirrorIndents = srcParaFormat.MirrorIndents;// 复制保存ClassParagraphFormat同名参数值
            this.OutlineLevel = (int)srcParaFormat.OutlineLevel;// 复制保存ClassParagraphFormat同名参数值
            this.NoLineNumber = srcParaFormat.NoLineNumber;// 复制保存ClassParagraphFormat同名参数值
            this.PageBreakBefore = srcParaFormat.PageBreakBefore;// 复制保存ClassParagraphFormat同名参数值
            this.ReadingOrder = (int)srcParaFormat.ReadingOrder;// 复制保存ClassParagraphFormat同名参数值
            this.RightIndent = srcParaFormat.RightIndent;// 复制保存ClassParagraphFormat同名参数值

            // this.Shading = srcParaFormat.Shading;
            //this.Shading.clone(srcParaFormat.Shading);// 复制保存ClassParagraphFormat同名参数值

            this.SpaceAfter = srcParaFormat.SpaceAfter;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceAfterAuto = srcParaFormat.SpaceAfterAuto;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceBefore = srcParaFormat.SpaceBefore;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceBeforeAuto = srcParaFormat.SpaceBeforeAuto;// 复制保存ClassParagraphFormat同名参数值

            // this.TabStops = srcParaFormat.TabStops;

            this.TextboxTightWrap = (int)srcParaFormat.TextboxTightWrap;// 复制保存ClassParagraphFormat同名参数值
            this.WidowControl = srcParaFormat.WidowControl;// 复制保存ClassParagraphFormat同名参数值
            this.WordWrap = srcParaFormat.WordWrap;// 复制保存ClassParagraphFormat同名参数值

            if (srcParaFormat.setMembers != null && this.setMembers != null)
            {
                this.setMembers.Clear();
                foreach (int nItem in srcParaFormat.setMembers)
                {
                    this.setMembers.Add(nItem);
                }
            }

            return;
        }


        // 复制到WORD.ParagraphFormat同名参数值
        public void copy2(Word.ParagraphFormat dstParaFmt)
        {
            dstParaFmt.Alignment = (Word.WdParagraphAlignment)this.Alignment;// 复制到WORD.ParagraphFormat同名参数值

            dstParaFmt.AutoAdjustRightIndent = this.AutoAdjustRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.BaseLineAlignment = (Word.WdBaselineAlignment)this.BaseLineAlignment;// 复制到WORD.ParagraphFormat同名参数值

            //Word.Borders bds = dstParaFmt.Borders;// 复制到WORD.ParagraphFormat同名参数值
            //this.Borders.copy2(ref bds);

            dstParaFmt.CharacterUnitFirstLineIndent = this.CharacterUnitFirstLineIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.CharacterUnitLeftIndent = this.CharacterUnitLeftIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.CharacterUnitRightIndent = this.CharacterUnitRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.DisableLineHeightGrid = this.DisableLineHeightGrid;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.FarEastLineBreakControl = this.FarEastLineBreakControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.FirstLineIndent = this.FirstLineIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.HalfWidthPunctuationOnTopOfLine = this.HalfWidthPunctuationOnTopOfLine;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.HangingPunctuation = this.HangingPunctuation;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.Hyphenation = this.Hyphenation;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.KeepTogether = this.KeepTogether;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.KeepWithNext = this.KeepWithNext;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LeftIndent = this.LeftIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineSpacingRule = (Word.WdLineSpacing)this.LineSpacingRule;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineSpacing = this.LineSpacing;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineUnitAfter = this.LineUnitAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineUnitBefore = this.LineUnitBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.MirrorIndents = this.MirrorIndents;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.OutlineLevel = (Word.WdOutlineLevel)this.OutlineLevel;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.NoLineNumber = this.NoLineNumber;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.PageBreakBefore = this.PageBreakBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.ReadingOrder = (Word.WdReadingOrder)this.ReadingOrder;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.RightIndent = this.RightIndent;// 复制到WORD.ParagraphFormat同名参数值

            // dstParaFmt.Shading = this.Shading;
            //Word.Shading shd = dstParaFmt.Shading;// 复制到WORD.ParagraphFormat同名参数值
            //this.Shading.copy2(ref shd);

            dstParaFmt.SpaceAfter = this.SpaceAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceAfterAuto = this.SpaceAfterAuto;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBefore = this.SpaceBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBeforeAuto = this.SpaceBeforeAuto;// 复制到WORD.ParagraphFormat同名参数值
            
            // dstParaFmt.TabStops = this.TabStops;

            dstParaFmt.TextboxTightWrap = (Word.WdTextboxTightWrap)this.TextboxTightWrap;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WidowControl = this.WidowControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WordWrap = this.WordWrap;// 复制到WORD.ParagraphFormat同名参数值

            return;
        }


        public void copy2(ClassParagraphFormat dstParaFmt)
        {
            dstParaFmt.Alignment = (Word.WdParagraphAlignment)this.Alignment;// 复制到WORD.ParagraphFormat同名参数值

            dstParaFmt.AutoAdjustRightIndent = this.AutoAdjustRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.BaseLineAlignment = (Word.WdBaselineAlignment)this.BaseLineAlignment;// 复制到WORD.ParagraphFormat同名参数值

            //Word.Borders bds = dstParaFmt.Borders;// 复制到WORD.ParagraphFormat同名参数值
            //this.Borders.copy2(ref bds);

            dstParaFmt.CharacterUnitFirstLineIndent = this.CharacterUnitFirstLineIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.CharacterUnitLeftIndent = this.CharacterUnitLeftIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.CharacterUnitRightIndent = this.CharacterUnitRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.DisableLineHeightGrid = this.DisableLineHeightGrid;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.FarEastLineBreakControl = this.FarEastLineBreakControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.FirstLineIndent = this.FirstLineIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.HalfWidthPunctuationOnTopOfLine = this.HalfWidthPunctuationOnTopOfLine;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.HangingPunctuation = this.HangingPunctuation;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.Hyphenation = this.Hyphenation;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.KeepTogether = this.KeepTogether;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.KeepWithNext = this.KeepWithNext;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LeftIndent = this.LeftIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineSpacingRule = (Word.WdLineSpacing)this.LineSpacingRule;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineSpacing = this.LineSpacing;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineUnitAfter = this.LineUnitAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineUnitBefore = this.LineUnitBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.MirrorIndents = this.MirrorIndents;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.OutlineLevel = (Word.WdOutlineLevel)this.OutlineLevel;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.NoLineNumber = this.NoLineNumber;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.PageBreakBefore = this.PageBreakBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.ReadingOrder = (Word.WdReadingOrder)this.ReadingOrder;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.RightIndent = this.RightIndent;// 复制到WORD.ParagraphFormat同名参数值

            // dstParaFmt.Shading = this.Shading;
            //Word.Shading shd = dstParaFmt.Shading;// 复制到WORD.ParagraphFormat同名参数值
            //this.Shading.copy2(ref shd);

            dstParaFmt.SpaceAfter = this.SpaceAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceAfterAuto = this.SpaceAfterAuto;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBefore = this.SpaceBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBeforeAuto = this.SpaceBeforeAuto;// 复制到WORD.ParagraphFormat同名参数值

            // dstParaFmt.TabStops = this.TabStops;

            dstParaFmt.TextboxTightWrap = (Word.WdTextboxTightWrap)this.TextboxTightWrap;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WidowControl = this.WidowControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WordWrap = this.WordWrap;// 复制到WORD.ParagraphFormat同名参数值

            if (dstParaFmt.setMembers != null && this.setMembers != null)
            {
                dstParaFmt.setMembers.Clear();
                foreach (int nItem in this.setMembers)
                {
                    dstParaFmt.setMembers.Add(nItem);
                }
            }

            return;
        }

        public void clone(nhParaFmt srcParaFormat)
        {
            this.bEnable = srcParaFormat.bEnable;
            this.Alignment = (int)srcParaFormat.Alignment;// 复制保存ClassParagraphFormat同名参数值

            this.AutoAdjustRightIndent = srcParaFormat.AutoAdjustRightIndent;// 复制保存ClassParagraphFormat同名参数值
            this.BaseLineAlignment = (int)srcParaFormat.BaseLineAlignment;// 复制保存ClassParagraphFormat同名参数值

            //this.Borders.clone(srcParaFormat.Borders);// 复制保存ClassParagraphFormat同名参数值

            this.CharacterUnitFirstLineIndent = srcParaFormat.CharacterUnitFirstLineIndent;// 复制保存ClassParagraphFormat同名参数值
            this.CharacterUnitLeftIndent = srcParaFormat.CharacterUnitLeftIndent;// 复制保存ClassParagraphFormat同名参数值
            this.CharacterUnitRightIndent = srcParaFormat.CharacterUnitRightIndent;// 复制保存ClassParagraphFormat同名参数值
            this.DisableLineHeightGrid = srcParaFormat.DisableLineHeightGrid;// 复制保存ClassParagraphFormat同名参数值
            this.FarEastLineBreakControl = srcParaFormat.FarEastLineBreakControl;// 复制保存ClassParagraphFormat同名参数值
            this.FirstLineIndent = srcParaFormat.FirstLineIndent;// 复制保存ClassParagraphFormat同名参数值
            this.HalfWidthPunctuationOnTopOfLine = srcParaFormat.HalfWidthPunctuationOnTopOfLine;// 复制保存ClassParagraphFormat同名参数值
            this.HangingPunctuation = srcParaFormat.HangingPunctuation;// 复制保存ClassParagraphFormat同名参数值
            this.Hyphenation = srcParaFormat.Hyphenation;// 复制保存ClassParagraphFormat同名参数值
            this.KeepTogether = srcParaFormat.KeepTogether;// 复制保存ClassParagraphFormat同名参数值
            this.KeepWithNext = srcParaFormat.KeepWithNext;// 复制保存ClassParagraphFormat同名参数值
            this.LeftIndent = srcParaFormat.LeftIndent;// 复制保存ClassParagraphFormat同名参数值
            this.LineSpacingRule = (int)srcParaFormat.LineSpacingRule;// 复制保存ClassParagraphFormat同名参数值
            this.LineSpacing = srcParaFormat.LineSpacing;// 复制保存ClassParagraphFormat同名参数值
            this.LineUnitAfter = srcParaFormat.LineUnitAfter;// 复制保存ClassParagraphFormat同名参数值
            this.LineUnitBefore = srcParaFormat.LineUnitBefore;// 复制保存ClassParagraphFormat同名参数值
            this.MirrorIndents = srcParaFormat.MirrorIndents;// 复制保存ClassParagraphFormat同名参数值
            this.OutlineLevel = (int)srcParaFormat.OutlineLevel;// 复制保存ClassParagraphFormat同名参数值
            this.NoLineNumber = srcParaFormat.NoLineNumber;// 复制保存ClassParagraphFormat同名参数值
            this.PageBreakBefore = srcParaFormat.PageBreakBefore;// 复制保存ClassParagraphFormat同名参数值
            this.ReadingOrder = (int)srcParaFormat.ReadingOrder;// 复制保存ClassParagraphFormat同名参数值
            this.RightIndent = srcParaFormat.RightIndent;// 复制保存ClassParagraphFormat同名参数值

            // this.Shading = srcParaFormat.Shading;
            //this.Shading.clone(srcParaFormat.Shading);// 复制保存ClassParagraphFormat同名参数值

            this.SpaceAfter = srcParaFormat.SpaceAfter;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceAfterAuto = srcParaFormat.SpaceAfterAuto;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceBefore = srcParaFormat.SpaceBefore;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceBeforeAuto = srcParaFormat.SpaceBeforeAuto;// 复制保存ClassParagraphFormat同名参数值

            // this.TabStops = srcParaFormat.TabStops;

            this.TextboxTightWrap = (int)srcParaFormat.TextboxTightWrap;// 复制保存ClassParagraphFormat同名参数值
            this.WidowControl = srcParaFormat.WidowControl;// 复制保存ClassParagraphFormat同名参数值
            this.WordWrap = srcParaFormat.WordWrap;// 复制保存ClassParagraphFormat同名参数值

            if (srcParaFormat.setMembers != null && this.setMembers != null)
            {
                this.setMembers.Clear();
                foreach (int nItem in srcParaFormat.setMembers)
                {
                    this.setMembers.Add(nItem);
                }
            }

            return;
        }


        public void copy2(nhParaFmt dstParaFmt)
        {
            dstParaFmt.bEnable = this.bEnable;
            dstParaFmt.Alignment = this.Alignment;// 复制到WORD.ParagraphFormat同名参数值

            dstParaFmt.AutoAdjustRightIndent = this.AutoAdjustRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.BaseLineAlignment = this.BaseLineAlignment;// 复制到WORD.ParagraphFormat同名参数值

            //Word.Borders bds = dstParaFmt.Borders;// 复制到WORD.ParagraphFormat同名参数值
            //this.Borders.copy2(ref bds);

            dstParaFmt.CharacterUnitFirstLineIndent = this.CharacterUnitFirstLineIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.CharacterUnitLeftIndent = this.CharacterUnitLeftIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.CharacterUnitRightIndent = this.CharacterUnitRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.DisableLineHeightGrid = this.DisableLineHeightGrid;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.FarEastLineBreakControl = this.FarEastLineBreakControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.FirstLineIndent = this.FirstLineIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.HalfWidthPunctuationOnTopOfLine = this.HalfWidthPunctuationOnTopOfLine;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.HangingPunctuation = this.HangingPunctuation;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.Hyphenation = this.Hyphenation;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.KeepTogether = this.KeepTogether;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.KeepWithNext = this.KeepWithNext;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LeftIndent = this.LeftIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineSpacingRule = this.LineSpacingRule;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineSpacing = this.LineSpacing;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineUnitAfter = this.LineUnitAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.LineUnitBefore = this.LineUnitBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.MirrorIndents = this.MirrorIndents;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.OutlineLevel = this.OutlineLevel;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.NoLineNumber = this.NoLineNumber;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.PageBreakBefore = this.PageBreakBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.ReadingOrder = this.ReadingOrder;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.RightIndent = this.RightIndent;// 复制到WORD.ParagraphFormat同名参数值

            // dstParaFmt.Shading = this.Shading;
            //Word.Shading shd = dstParaFmt.Shading;// 复制到WORD.ParagraphFormat同名参数值
            //this.Shading.copy2(ref shd);

            dstParaFmt.SpaceAfter = this.SpaceAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceAfterAuto = this.SpaceAfterAuto;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBefore = this.SpaceBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBeforeAuto = this.SpaceBeforeAuto;// 复制到WORD.ParagraphFormat同名参数值

            // dstParaFmt.TabStops = this.TabStops;

            dstParaFmt.TextboxTightWrap = this.TextboxTightWrap;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WidowControl = this.WidowControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WordWrap = this.WordWrap;// 复制到WORD.ParagraphFormat同名参数值

            if (dstParaFmt.setMembers != null && this.setMembers != null)
            {
                dstParaFmt.setMembers.Clear();
                foreach (int nItem in this.setMembers)
                {
                    dstParaFmt.setMembers.Add(nItem);
                }
            }

            return;
        }


        public int formatString(RichTextBox rchTxt, String strPreBlanks)
        {
            if (setMembers == null)
            {
                return -1;
            }

            String strRet = "";

            // 对齐
            if (isSet(euMembers.Alignment))
            {
                switch(Alignment)
                {
                    case (int)Word.WdParagraphAlignment.wdAlignParagraphLeft:
                        strRet += "对齐方式：居左\r\n";
                        break;

                    case (int)Word.WdParagraphAlignment.wdAlignParagraphRight:
                        strRet += "对齐方式：居右\r\n";
                        break;

                    case (int)Word.WdParagraphAlignment.wdAlignParagraphCenter:
                        strRet += "对齐方式：居中\r\n";
                        break;

                    default:
                        break;
                }
            }
            
            // 行间距
            if (isSet(euMembers.LineSpacingRule))
            {
                float fPonds = float.NaN;

                if (isSet(euMembers.LineSpacing))
                {
                    fPonds = LineSpacing;
                }

                switch(LineSpacingRule)
                {
                    case (int)Word.WdLineSpacing.wdLineSpace1pt5:
                        strRet += "行间距：1.5倍\r\n";
                        break;

                    case (int)Word.WdLineSpacing.wdLineSpaceAtLeast:
                        strRet += "行间距：最小值[" + fPonds.ToString("##.#") + "磅]\r\n";
                        break;

                    case (int)Word.WdLineSpacing.wdLineSpaceDouble:
                        strRet += "行间距：2倍\r\n";
                        break;

                    case (int)Word.WdLineSpacing.wdLineSpaceExactly:
                        strRet += "行间距：固定值[" + fPonds.ToString("##.#") + "磅]\r\n";
                        break;

                    case (int)Word.WdLineSpacing.wdLineSpaceMultiple:
                        strRet += "行间距：多倍[" + fPonds.ToString("##") + "倍]\r\n";
                        break;

                    case (int)Word.WdLineSpacing.wdLineSpaceSingle:
                        strRet += "行间距：单倍\r\n";
                        break;

                    default:
                        break;
                }
            }

            // 缩进
            if (isSet(euMembers.CharacterUnitFirstLineIndent))
            {
                int nVal = (int)Math.Abs(CharacterUnitFirstLineIndent);

                if (CharacterUnitFirstLineIndent > 0)
                {
                    strRet += "首行缩进：" + nVal + "字符\r\n";
                }
                else if(CharacterUnitFirstLineIndent < 0)
                {
                    strRet += "悬挂缩进：" + nVal + "字符\r\n";
                }

            }
            else if (isSet(euMembers.FirstLineIndent))
            {
                float fVal = Math.Abs(FirstLineIndent);

                if (FirstLineIndent > 0)
                {
                    strRet += "首行缩进：" + fVal.ToString("##") + "磅\r\n";
                }
                else if (FirstLineIndent < 0)
                {
                    strRet += "悬挂缩进：" + fVal.ToString("##") + "磅\r\n";
                }
            }


            if (isSet(euMembers.CharacterUnitLeftIndent))
            {
                strRet += "左缩进：" + CharacterUnitLeftIndent.ToString("##") + "字符\r\n";
            }
            else if (isSet(euMembers.LeftIndent))
            {
                strRet += "左缩进：" + LeftIndent.ToString("##") + "磅\r\n";
            }

            euMembers euItem = euMembers.Zero;

            foreach (int nItem in setMembers)
            {
                euItem = (euMembers)nItem;

                switch (euItem)
                {
                    case euMembers.AddSpaceBetweenFarEastAndAlpha:
                        break;

                    case euMembers.AddSpaceBetweenFarEastAndDigit:
                        break;

                    case euMembers.Alignment:
                        break;

                    case euMembers.AutoAdjustRightIndent:
                        break;

                    case euMembers.BaseLineAlignment:
                        break;

                    case euMembers.Borders:
                        // wParaFmt.Borders = Borders;
                        break;

                    case euMembers.CharacterUnitFirstLineIndent:
                        break;

                    case euMembers.CharacterUnitLeftIndent:
                        break;

                    case euMembers.CharacterUnitRightIndent:
                        break;

                    case euMembers.DisableLineHeightGrid:
                        break;

                    case euMembers.Duplicate:
                        // wParaFmt.Duplicate = Duplicate;
                        break;

                    case euMembers.FarEastLineBreakControl:
                        break;

                    case euMembers.FirstLineIndent:
                        break;

                    case euMembers.HalfWidthPunctuationOnTopOfLine:
                        break;

                    case euMembers.HangingPunctuation:
                        break;

                    case euMembers.Hyphenation:
                        break;

                    case euMembers.KeepTogether:
                        break;

                    case euMembers.KeepWithNext:
                        break;

                    case euMembers.LeftIndent:
                        break;

                    case euMembers.LineSpacing:
                        break;

                    case euMembers.LineSpacingRule:
                        break;

                    case euMembers.LineUnitAfter:
                        break;

                    case euMembers.LineUnitBefore:
                        break;

                    case euMembers.MirrorIndents:
                        break;

                    case euMembers.NoLineNumber:
                        break;

                    case euMembers.OutlineLevel:
                        break;

                    case euMembers.PageBreakBefore:
                        break;

                    case euMembers.ReadingOrder:
                        break;

                    case euMembers.RightIndent:
                        break;

                    case euMembers.Shading:
                        //wParaFmt.Shading = Shading;
                        break;

                    case euMembers.SpaceAfter:
                        strRet += "[段后间距：" + SpaceAfter.ToString("##.#") + "磅]";
                        break;

                    case euMembers.SpaceAfterAuto:
                        strRet += "[段后间距：自动]";
                        break;

                    case euMembers.SpaceBefore:
                        strRet += "[段前间距：" + SpaceBefore.ToString("##.#") + "磅]";
                        break;

                    case euMembers.SpaceBeforeAuto:
                        strRet += "[段前间距：自动]";
                        break;

                    case euMembers.TabStops:
                        //wParaFmt.TabStops = TabStops;
                        break;

                    case euMembers.TextboxTightWrap:
                        break;

                    case euMembers.WidowControl:
                        strRet += "[孤行控制]";
                        break;

                    case euMembers.WordWrap:
                        break;

                    default:
                        break;
                }
            }

            if (strRet.Length > 0)
            {
                strRet += "\r\n";
            }
            rchTxt.AppendText(strRet);

            return 0;
        }





    } // class

} // namespace
