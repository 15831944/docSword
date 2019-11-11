using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;


namespace OfficeAssist
{
    // 同WORD.ParagraphFormat的同名类，用于保存其同名参数值
    public class ClassParagraphFormat /*: ParagraphFormat*/
    {
        public int AddSpaceBetweenFarEastAndAlpha;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int AddSpaceBetweenFarEastAndDigit;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public WdParagraphAlignment Alignment;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        // public Application Application { get; }// 同WORD.ParagraphFormat的同名参数
        public int AutoAdjustRightIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public WdBaselineAlignment BaseLineAlignment;// { get; set; }// 同WORD.ParagraphFormat的同名参数

        public ClassBorders Borders = new ClassBorders();
        // public Borders Borders;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float CharacterUnitFirstLineIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float CharacterUnitLeftIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float CharacterUnitRightIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        // public int Creator;// { get; }// 同WORD.ParagraphFormat的同名参数
        public int DisableLineHeightGrid;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public ParagraphFormat Duplicate;// { get; }// 同WORD.ParagraphFormat的同名参数
        public int FarEastLineBreakControl;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float FirstLineIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int HalfWidthPunctuationOnTopOfLine;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int HangingPunctuation;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int Hyphenation;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int KeepTogether;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int KeepWithNext;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float LeftIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float LineSpacing;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public WdLineSpacing LineSpacingRule;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float LineUnitAfter;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float LineUnitBefore;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int MirrorIndents;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int NoLineNumber;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public WdOutlineLevel OutlineLevel;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int PageBreakBefore;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        // public dynamic Parent { get; }// 同WORD.ParagraphFormat的同名参数
        public WdReadingOrder ReadingOrder;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float RightIndent;// { get; set; }// 同WORD.ParagraphFormat的同名参数

        public ClassShading Shading = new ClassShading();
        // public Shading Shading;// { get; }// 同WORD.ParagraphFormat的同名参数

        public float SpaceAfter;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int SpaceAfterAuto;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public float SpaceBefore;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int SpaceBeforeAuto;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public TabStops TabStops;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public WdTextboxTightWrap TextboxTightWrap;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int WidowControl;// { get; set; }// 同WORD.ParagraphFormat的同名参数
        public int WordWrap;// { get; set; }// 同WORD.ParagraphFormat的同名参数


        public ClassParagraphFormat()
        {
//             Hyphenation = -1; // true
//             AutoAdjustRightIndent = -1; // true
//             FarEastLineBreakControl = -1; // true
//             WordWrap = -1;// true
//             HangingPunctuation = -1; // true
//             AddSpaceBetweenFarEastAndAlpha = -1; // true
//             AddSpaceBetweenFarEastAndDigit = -1; // true
//             LineSpacing = 12.0f;
//             BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto;
//             ReadingOrder = WdReadingOrder.wdReadingOrderLtr;
//             KeepTogether = 0;
//             KeepWithNext = -1;
//             SpaceAfter = 3.0f;
//             SpaceAfterAuto = 0;
//             SpaceBefore = 12.0f;
//             SpaceBeforeAuto = 0;

            return;
        }

        // 复制保存WORD.ParagraphFormat同名参数值
        public void clone(Word.ParagraphFormat srcParaFormat)
        {
            this.Alignment = srcParaFormat.Alignment;// 复制保存WORD.ParagraphFormat同名参数值

            this.AutoAdjustRightIndent = srcParaFormat.AutoAdjustRightIndent;// 复制保存WORD.ParagraphFormat同名参数值
            this.BaseLineAlignment = srcParaFormat.BaseLineAlignment;// 复制保存WORD.ParagraphFormat同名参数值

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
            this.LineSpacingRule = srcParaFormat.LineSpacingRule;// 复制保存WORD.ParagraphFormat同名参数值
            this.LineUnitAfter = srcParaFormat.LineUnitAfter;// 复制保存WORD.ParagraphFormat同名参数值
            this.LineUnitBefore = srcParaFormat.LineUnitBefore;// 复制保存WORD.ParagraphFormat同名参数值
            this.MirrorIndents = srcParaFormat.MirrorIndents;// 复制保存WORD.ParagraphFormat同名参数值
            this.OutlineLevel = srcParaFormat.OutlineLevel;// 复制保存WORD.ParagraphFormat同名参数值
            this.NoLineNumber = srcParaFormat.NoLineNumber;// 复制保存WORD.ParagraphFormat同名参数值
            this.PageBreakBefore = srcParaFormat.PageBreakBefore;// 复制保存WORD.ParagraphFormat同名参数值
            this.ReadingOrder = srcParaFormat.ReadingOrder;// 复制保存WORD.ParagraphFormat同名参数值
            this.RightIndent = srcParaFormat.RightIndent;// 复制保存WORD.ParagraphFormat同名参数值

            // this.Shading = srcParaFormat.Shading;
            //this.Shading.clone(srcParaFormat.Shading);

            this.SpaceAfter = srcParaFormat.SpaceAfter;// 复制保存WORD.ParagraphFormat同名参数值
            this.SpaceAfterAuto = srcParaFormat.SpaceAfterAuto;// 复制保存WORD.ParagraphFormat同名参数值
            this.SpaceBefore = srcParaFormat.SpaceBefore;// 复制保存WORD.ParagraphFormat同名参数值
            this.SpaceBeforeAuto = srcParaFormat.SpaceBeforeAuto;// 复制保存WORD.ParagraphFormat同名参数值
            
            // this.TabStops = srcParaFormat.TabStops;

            this.TextboxTightWrap = srcParaFormat.TextboxTightWrap;// 复制保存WORD.ParagraphFormat同名参数值
            this.WidowControl = srcParaFormat.WidowControl;// 复制保存WORD.ParagraphFormat同名参数值
            this.WordWrap = srcParaFormat.WordWrap;// 复制保存WORD.ParagraphFormat同名参数值

            return;
        }


        // 复制保存ClassParagraphFormat同名参数值
        public void clone(ClassParagraphFormat srcParaFormat)
        {
            this.Alignment = srcParaFormat.Alignment;// 复制保存ClassParagraphFormat同名参数值

            this.AutoAdjustRightIndent = srcParaFormat.AutoAdjustRightIndent;// 复制保存ClassParagraphFormat同名参数值
            this.BaseLineAlignment = srcParaFormat.BaseLineAlignment;// 复制保存ClassParagraphFormat同名参数值

            this.Borders.clone(srcParaFormat.Borders);// 复制保存ClassParagraphFormat同名参数值

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
            this.LineSpacingRule = srcParaFormat.LineSpacingRule;// 复制保存ClassParagraphFormat同名参数值
            this.LineSpacing = srcParaFormat.LineSpacing;// 复制保存ClassParagraphFormat同名参数值
            this.LineUnitAfter = srcParaFormat.LineUnitAfter;// 复制保存ClassParagraphFormat同名参数值
            this.LineUnitBefore = srcParaFormat.LineUnitBefore;// 复制保存ClassParagraphFormat同名参数值
            this.MirrorIndents = srcParaFormat.MirrorIndents;// 复制保存ClassParagraphFormat同名参数值
            this.OutlineLevel = srcParaFormat.OutlineLevel;// 复制保存ClassParagraphFormat同名参数值
            this.NoLineNumber = srcParaFormat.NoLineNumber;// 复制保存ClassParagraphFormat同名参数值
            this.PageBreakBefore = srcParaFormat.PageBreakBefore;// 复制保存ClassParagraphFormat同名参数值
            this.ReadingOrder = srcParaFormat.ReadingOrder;// 复制保存ClassParagraphFormat同名参数值
            this.RightIndent = srcParaFormat.RightIndent;// 复制保存ClassParagraphFormat同名参数值

            // this.Shading = srcParaFormat.Shading;
            this.Shading.clone(srcParaFormat.Shading);// 复制保存ClassParagraphFormat同名参数值

            this.SpaceAfter = srcParaFormat.SpaceAfter;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceAfterAuto = srcParaFormat.SpaceAfterAuto;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceBefore = srcParaFormat.SpaceBefore;// 复制保存ClassParagraphFormat同名参数值
            this.SpaceBeforeAuto = srcParaFormat.SpaceBeforeAuto;// 复制保存ClassParagraphFormat同名参数值

            // this.TabStops = srcParaFormat.TabStops;

            this.TextboxTightWrap = srcParaFormat.TextboxTightWrap;// 复制保存ClassParagraphFormat同名参数值
            this.WidowControl = srcParaFormat.WidowControl;// 复制保存ClassParagraphFormat同名参数值
            this.WordWrap = srcParaFormat.WordWrap;// 复制保存ClassParagraphFormat同名参数值

            return;
        }


        // 复制到WORD.ParagraphFormat同名参数值
        public void copy2(Word.ParagraphFormat dstParaFmt)
        {
            dstParaFmt.Alignment = this.Alignment;// 复制到WORD.ParagraphFormat同名参数值

            dstParaFmt.AutoAdjustRightIndent = this.AutoAdjustRightIndent;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.BaseLineAlignment = this.BaseLineAlignment;// 复制到WORD.ParagraphFormat同名参数值

            Word.Borders bds = dstParaFmt.Borders;// 复制到WORD.ParagraphFormat同名参数值
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
            Word.Shading shd = dstParaFmt.Shading;// 复制到WORD.ParagraphFormat同名参数值
            //this.Shading.copy2(ref shd);

            dstParaFmt.SpaceAfter = this.SpaceAfter;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceAfterAuto = this.SpaceAfterAuto;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBefore = this.SpaceBefore;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.SpaceBeforeAuto = this.SpaceBeforeAuto;// 复制到WORD.ParagraphFormat同名参数值
            
            // dstParaFmt.TabStops = this.TabStops;

            dstParaFmt.TextboxTightWrap = this.TextboxTightWrap;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WidowControl = this.WidowControl;// 复制到WORD.ParagraphFormat同名参数值
            dstParaFmt.WordWrap = this.WordWrap;// 复制到WORD.ParagraphFormat同名参数值

            return;
        }

    }
}
