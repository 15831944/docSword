using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;


namespace OfficeAssist
{
    // 参照Word.Font的对象
    public class ClassFont /*: Font*/
    {
        public int AllCaps;// { get; set; }
        public WdAnimation Animation;// { get; set; }
        // public Application Application;// { get; }
        public int Bold;// { get; set; }
        public int BoldBi;// { get; set; }
        // 对应Word.Borders的对象
        public ClassBorders Borders = new ClassBorders();
        //public Borders Borders;// { get; set; }

        public WdColor Color;// { get; set; }
        public WdColorIndex ColorIndex;// { get; set; }
        public WdColorIndex ColorIndexBi;// { get; set; }
        // public int Creator;// { get; }
        public WdColor DiacriticColor;// { get; set; }
        public bool DisableCharacterSpaceGrid;// { get; set; }
        public int DoubleStrikeThrough;// { get; set; }
        public Font Duplicate;// { get; }
        public int Emboss;// { get; set; }
        public WdEmphasisMark EmphasisMark;// { get; set; }
        public int Engrave;// { get; set; }
        public int Hidden;// { get; set; }
        public int Italic;// { get; set; }
        public int ItalicBi;// { get; set; }
        public float Kerning;// { get; set; }
        public string Name = "";// { get; set; }
        public string NameAscii = "";// { get; set; }
        public string NameBi = "";// { get; set; }
        public string NameFarEast = "";// { get; set; }
        //public string NameOther = "";// { get; set; }
        public int Outline;// { get; set; }
        // public dynamic Parent { get; }
        public int Position;// { get; set; }
        public int Scaling;// { get; set; }

        public ClassShading Shading = new ClassShading();
        // public Shading Shading;// { get; }
        
        public int Shadow;// { get; set; }
        public float Size;// { get; set; }
        public float SizeBi;// { get; set; }
        public int SmallCaps;// { get; set; }
        public float Spacing;// { get; set; }
        public int StrikeThrough;// { get; set; }
        public int Subscript;// { get; set; }
        public int Superscript;// { get; set; }
        public WdUnderline Underline;// { get; set; }
        public WdColor UnderlineColor;// { get; set; }

        // for Font Format Dialog
        //public int CharacterWidthGrid;
        //public int ColorDialog;
        //public String KerningMin = "";
        //public String PointsBi = "";

        public ClassFont()
        {
            setDefault();
            return;
        }

        // 设置缺省初始值
        public void setDefault()
        {
            int nInit = (int)Word.WdConstants.wdUndefined;// 进行赋值

            AllCaps = nInit;// 进行赋值
            Animation = (WdAnimation)nInit;//.wdAnimationNone;// 进行赋值
            Bold = nInit;// 进行赋值
            BoldBi = nInit;// 进行赋值

            Color = (WdColor)nInit;//.wdColorAutomatic;// 进行赋值
            ColorIndex = (WdColorIndex)nInit;//.wdAuto;// 进行赋值
            ColorIndexBi = (WdColorIndex)nInit;//.wdAuto;// 进行赋值
            // Creator = nInit;
            DiacriticColor = (WdColor)nInit;//.wdColorAutomatic;// 进行赋值
            DisableCharacterSpaceGrid = false;// 进行赋值
            DoubleStrikeThrough = nInit;// 进行赋值
            Emboss = nInit;
            EmphasisMark = (WdEmphasisMark)nInit;//.wdEmphasisMarkNone;// 进行赋值

            Engrave = nInit;// 进行赋值
            Hidden = nInit;// 进行赋值
            Italic = nInit;// 进行赋值
            ItalicBi = nInit;// 进行赋值
            Kerning = (float)nInit;// 进行赋值

            Name = "";// 进行赋值
            NameAscii = "";// 进行赋值
            NameBi = "";// 进行赋值
            NameFarEast = "";// 进行赋值
            //NameOther = "";// 进行赋值
            Outline = nInit;// 进行赋值

            Position = nInit;// 进行赋值
            Scaling = nInit;// 进行赋值

            Shadow = nInit;// 进行赋值
            Size = (float)nInit;// 进行赋值
            SizeBi = (float)nInit;// 进行赋值

            SmallCaps = nInit;// 进行赋值
            Spacing = (float)nInit;// 进行赋值
            StrikeThrough = nInit;// 进行赋值
            Subscript = nInit;// 进行赋值
            Superscript = nInit;// 进行赋值
            Underline = (WdUnderline)nInit;//WdUnderline.wdUnderlineNone;// 进行赋值
            UnderlineColor = (WdColor)nInit; //.wdColorAutomatic;// 进行赋值

            // CharacterWidthGrid = 0;
            DisableCharacterSpaceGrid = false;// 进行赋值
            
            return;
        }

        // 比较2个类的值，是否异同
        public Boolean diff(Word.Font srcFont)
        {
            if (this.AllCaps != srcFont.AllCaps) // 比较AllCaps
                return false;

            if (this.Animation != srcFont.Animation)// 比较Animation
                return false;

            if (this.Bold != srcFont.Bold)// 比较Bold
                return false;

            if (this.BoldBi != srcFont.BoldBi)// 比较BoldBi
                return false;

            // copyBorders(srcFont.Borders, this.Borders);
            // this.Borders.clone(srcFont.Borders);

            if (this.Color != srcFont.Color)// 比较Color
                return false;

            if (this.ColorIndex != srcFont.ColorIndex)// 比较ColorIndex
                return false;

            if (this.ColorIndexBi != srcFont.ColorIndexBi)// 比较ColorIndexBi
                return false;

            if (this.DiacriticColor != srcFont.DiacriticColor)// 比较DiacriticColor
                return false;

            if (this.DisableCharacterSpaceGrid != srcFont.DisableCharacterSpaceGrid)// 比较DisableCharacterSpaceGrid
                return false;

            if (this.DoubleStrikeThrough != srcFont.DoubleStrikeThrough)// 比较DoubleStrikeThrough
                return false;

            if (this.Emboss != srcFont.Emboss)// 比较Emboss
                return false;

            if (this.EmphasisMark != srcFont.EmphasisMark)// 比较EmphasisMark
                return false;

            if (this.Engrave != srcFont.Engrave)// 比较Engrave
                return false;

            if (this.Hidden != srcFont.Hidden)// 比较Hidden
                return false;

            if (this.Italic != srcFont.Italic)// 比较Italic
                return false;

            if (this.ItalicBi != srcFont.ItalicBi)// 比较ItalicBi
                return false;

            if (this.Kerning != srcFont.Kerning)// 比较Kerning
                return false;

            if (this.Name != srcFont.Name)// 比较Name
                return false;

            if (this.NameAscii != srcFont.NameAscii)// 比较NameAscii
                return false;

            if (this.NameBi != srcFont.NameBi)// 比较NameBi
                return false;

            if (this.NameFarEast != srcFont.NameFarEast)// 比较NameFarEast
                return false;

            //            if (this.NameOther != srcFont.NameOther)// 比较NameOther
//                return false;

            if (this.Outline != srcFont.Outline)// 比较Outline
                return false;

            if (this.Position != srcFont.Position)// 比较Position
                return false;

            if (this.Scaling != srcFont.Scaling)// 比较Scaling
                return false;

            // this.Shading = srcFont.Shading;// 比较Shading
            // this.Shading.clone(srcFont.Shading);

            if (this.Shadow != srcFont.Shadow)// 比较Shadow
                return false;

            if (this.Size != srcFont.Size)// 比较Size
                return false;

            if (this.SizeBi != srcFont.SizeBi)// 比较SizeBi
                return false;

            if (this.SmallCaps != srcFont.SmallCaps)// 比较SmallCaps
                return false;

            if (this.Spacing != srcFont.Spacing)// 比较Spacing
                return false;

            if (this.StrikeThrough != srcFont.StrikeThrough)// 比较StrikeThrough
                return false;

            if (this.Subscript != srcFont.Subscript)// 比较Subscript
                return false;

            if (this.Superscript != srcFont.Superscript)// 比较Superscript
                return false;

            if (this.Underline != srcFont.Underline)// 比较Underline
                return false;

            if (this.UnderlineColor != srcFont.UnderlineColor)// 比较UnderlineColor
                return false;

            return true;
        }

        // 比较2个类的值，是否异同
        public Boolean diff(ClassFont srcFont)
        {
            if (this.AllCaps != srcFont.AllCaps) // 比较成员的异同，不同则返回
                return false;

            if (this.Animation != srcFont.Animation) // 比较成员的异同，不同则返回
                return false;

            if (this.Bold != srcFont.Bold) // 比较成员的异同，不同则返回
                return false;

            if (this.BoldBi != srcFont.BoldBi) // 比较成员的异同，不同则返回
                return false;

            // copyBorders(srcFont.Borders, this.Borders); // 比较成员的异同，不同则返回
            // this.Borders.clone(srcFont.Borders);

            if (this.Color != srcFont.Color) // 比较成员的异同，不同则返回
                return false;

            if (this.ColorIndex != srcFont.ColorIndex) // 比较成员的异同，不同则返回
                return false;

            if (this.ColorIndexBi != srcFont.ColorIndexBi) // 比较成员的异同，不同则返回
                return false;

            if (this.DiacriticColor != srcFont.DiacriticColor) // 比较成员的异同，不同则返回
                return false;

            if (this.DisableCharacterSpaceGrid != srcFont.DisableCharacterSpaceGrid) // 比较成员的异同，不同则返回
                return false;

            if (this.DoubleStrikeThrough != srcFont.DoubleStrikeThrough) // 比较成员的异同，不同则返回
                return false;

            if (this.Emboss != srcFont.Emboss) // 比较成员的异同，不同则返回
                return false;

            if (this.EmphasisMark != srcFont.EmphasisMark) // 比较成员的异同，不同则返回
                return false;

            if (this.Engrave != srcFont.Engrave) // 比较成员的异同，不同则返回
                return false;

            if (this.Hidden != srcFont.Hidden) // 比较成员的异同，不同则返回
                return false;

            if (this.Italic != srcFont.Italic) // 比较成员的异同，不同则返回
                return false;

            if (this.ItalicBi != srcFont.ItalicBi) // 比较成员的异同，不同则返回
                return false;

            if (this.Kerning != srcFont.Kerning) // 比较成员的异同，不同则返回
                return false;

            if (this.Name != srcFont.Name) // 比较成员的异同，不同则返回
                return false;

            if (this.NameAscii != srcFont.NameAscii) // 比较成员的异同，不同则返回
                return false;

            if (this.NameBi != srcFont.NameBi) // 比较成员的异同，不同则返回
                return false;

            if (this.NameFarEast != srcFont.NameFarEast) // 比较成员的异同，不同则返回
                return false;

            //            if(this.NameOther != srcFont.NameOther) // 比较成员的异同，不同则返回
//                return false;

            if (this.Outline != srcFont.Outline) // 比较成员的异同，不同则返回
                return false;

            if (this.Position != srcFont.Position) // 比较成员的异同，不同则返回
                return false;

            if (this.Scaling != srcFont.Scaling) // 比较成员的异同，不同则返回
                return false;

            // this.Shading = srcFont.Shading;
            // this.Shading.clone(srcFont.Shading);

            if (this.Shadow != srcFont.Shadow) // 比较成员的异同，不同则返回
                return false;

            if (this.Size != srcFont.Size) // 比较成员的异同，不同则返回
                return false;

            if (this.SizeBi != srcFont.SizeBi) // 比较成员的异同，不同则返回
                return false;

            if (this.SmallCaps != srcFont.SmallCaps) // 比较成员的异同，不同则返回
                return false;

            if (this.Spacing != srcFont.Spacing) // 比较成员的异同，不同则返回
                return false;

            if (this.StrikeThrough != srcFont.StrikeThrough) // 比较成员的异同，不同则返回
                return false;

            if (this.Subscript != srcFont.Subscript) // 比较成员的异同，不同则返回
                return false;

            if (this.Superscript != srcFont.Superscript) // 比较成员的异同，不同则返回
                return false;

            if (this.Underline != srcFont.Underline) // 比较成员的异同，不同则返回
                return false;

            if (this.UnderlineColor != srcFont.UnderlineColor) // 比较成员的异同，不同则返回
                return false;

            return true;
        }


        // 复制WORD.FONT的内容到本类
        public void clone(Word.Font srcFont)
        {
            this.AllCaps = srcFont.AllCaps;// 复制WORD.FONT的成员到本类成员
            this.Animation = srcFont.Animation;// 复制WORD.FONT的成员到本类成员

            this.Bold = srcFont.Bold;// 复制WORD.FONT的成员到本类成员
            this.BoldBi = srcFont.BoldBi;// 复制WORD.FONT的成员到本类成员

            //this.Borders.clone(srcFont.Borders);

            this.Color = srcFont.Color;// 复制WORD.FONT的成员到本类成员
            this.ColorIndex = srcFont.ColorIndex;// 复制WORD.FONT的成员到本类成员
            this.ColorIndexBi = srcFont.ColorIndexBi;// 复制WORD.FONT的成员到本类成员

            this.DiacriticColor = srcFont.DiacriticColor;// 复制WORD.FONT的成员到本类成员
            this.DisableCharacterSpaceGrid = srcFont.DisableCharacterSpaceGrid;// 复制WORD.FONT的成员到本类成员
            this.DoubleStrikeThrough = srcFont.DoubleStrikeThrough;// 复制WORD.FONT的成员到本类成员

            this.Emboss = srcFont.Emboss;// 复制WORD.FONT的成员到本类成员
            this.EmphasisMark = srcFont.EmphasisMark;// 复制WORD.FONT的成员到本类成员
            this.Engrave = srcFont.Engrave;// 复制WORD.FONT的成员到本类成员
            this.Hidden = srcFont.Hidden;// 复制WORD.FONT的成员到本类成员
            this.Italic = srcFont.Italic;// 复制WORD.FONT的成员到本类成员
            this.ItalicBi = srcFont.ItalicBi;// 复制WORD.FONT的成员到本类成员
            this.Kerning = srcFont.Kerning;// 复制WORD.FONT的成员到本类成员
            this.Name = srcFont.Name;// 复制WORD.FONT的成员到本类成员
            this.NameAscii = srcFont.NameAscii;// 复制WORD.FONT的成员到本类成员
            this.NameBi = srcFont.NameBi;// 复制WORD.FONT的成员到本类成员
            this.NameFarEast = srcFont.NameFarEast;// 复制WORD.FONT的成员到本类成员
            //this.NameOther = srcFont.NameOther;// 复制WORD.FONT的成员到本类成员
            this.Outline = srcFont.Outline;// 复制WORD.FONT的成员到本类成员
            this.Position = srcFont.Position;// 复制WORD.FONT的成员到本类成员
            this.Scaling = srcFont.Scaling;// 复制WORD.FONT的成员到本类成员

            //this.Shading.clone(srcFont.Shading);// 复制WORD.FONT的成员到本类成员

            this.Shadow = srcFont.Shadow;// 复制WORD.FONT的成员到本类成员
            this.Size = srcFont.Size;// 复制WORD.FONT的成员到本类成员
            this.SizeBi = srcFont.SizeBi;// 复制WORD.FONT的成员到本类成员
            this.SmallCaps = srcFont.SmallCaps;// 复制WORD.FONT的成员到本类成员
            this.Spacing = srcFont.Spacing;// 复制WORD.FONT的成员到本类成员
            this.StrikeThrough = srcFont.StrikeThrough;// 复制WORD.FONT的成员到本类成员
            this.Subscript = srcFont.Subscript;// 复制WORD.FONT的成员到本类成员
            this.Superscript = srcFont.Superscript;// 复制WORD.FONT的成员到本类成员
            this.Underline = srcFont.Underline;// 复制WORD.FONT的成员到本类成员
            this.UnderlineColor = srcFont.UnderlineColor;// 复制WORD.FONT的成员到本类成员
           
            return;
        }


        // 复制本类到WORD.FONT类
        public void copy2(Word.Font dstFnt)
        {
            dstFnt.AllCaps = this.AllCaps;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Animation = this.Animation;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.Bold = this.Bold;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.BoldBi = this.BoldBi;// 复制本类成员到WORD.FONT类同名成员

            Word.Borders bds = dstFnt.Borders;// 复制本类成员到WORD.FONT类同名成员
            //this.Borders.copy2(ref bds);

            dstFnt.Color = this.Color;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.ColorIndex = this.ColorIndex;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.ColorIndexBi = this.ColorIndexBi;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.DiacriticColor = this.DiacriticColor;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.DisableCharacterSpaceGrid = this.DisableCharacterSpaceGrid;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制本类成员到WORD.FONT类同名成员

            dstFnt.Emboss = this.Emboss;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.EmphasisMark = this.EmphasisMark;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Engrave = this.Engrave;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Hidden = this.Hidden;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Italic = this.Italic;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.ItalicBi = this.ItalicBi;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Kerning = this.Kerning;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Name = this.Name;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.NameAscii = this.NameAscii;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.NameBi = this.NameBi;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.NameFarEast = this.NameFarEast;// 复制本类成员到WORD.FONT类同名成员
            //dstFnt.NameOther = this.NameOther;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Outline = this.Outline;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Position = this.Position;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Scaling = this.Scaling;// 复制本类成员到WORD.FONT类同名成员

            Word.Shading shd = dstFnt.Shading;// 复制本类成员到WORD.FONT类同名成员
            //this.Shading.copy2(ref shd);

            dstFnt.Shadow = this.Shadow;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Size = this.Size;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.SizeBi = this.SizeBi;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.SmallCaps = this.SmallCaps;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Spacing = this.Spacing;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.StrikeThrough = this.StrikeThrough;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Subscript = this.Subscript;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Superscript = this.Superscript;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.Underline = this.Underline;// 复制本类成员到WORD.FONT类同名成员
            dstFnt.UnderlineColor = this.UnderlineColor;// 复制本类成员到WORD.FONT类同名成员
            
            return;
        }

        // 复制ClassFont到本类
        public void clone(ClassFont srcFont)
        {
            this.AllCaps = srcFont.AllCaps;// 复制ClassFont类同名成员到本类成员
            this.Animation = srcFont.Animation;// 复制ClassFont类同名成员到本类成员

            this.Bold = srcFont.Bold;// 复制ClassFont类同名成员到本类成员
            this.BoldBi = srcFont.BoldBi;// 复制ClassFont类同名成员到本类成员

            this.Borders.clone(srcFont.Borders);// 复制ClassFont类同名成员到本类成员

            this.Color = srcFont.Color;// 复制ClassFont类同名成员到本类成员
            this.ColorIndex = srcFont.ColorIndex;// 复制ClassFont类同名成员到本类成员
            this.ColorIndexBi = srcFont.ColorIndexBi;// 复制ClassFont类同名成员到本类成员

            this.DiacriticColor = srcFont.DiacriticColor;// 复制ClassFont类同名成员到本类成员
            this.DisableCharacterSpaceGrid = srcFont.DisableCharacterSpaceGrid;// 复制ClassFont类同名成员到本类成员
            this.DoubleStrikeThrough = srcFont.DoubleStrikeThrough;// 复制ClassFont类同名成员到本类成员

            this.Emboss = srcFont.Emboss;// 复制ClassFont类同名成员到本类成员
            this.EmphasisMark = srcFont.EmphasisMark;// 复制ClassFont类同名成员到本类成员
            this.Engrave = srcFont.Engrave;// 复制ClassFont类同名成员到本类成员
            this.Hidden = srcFont.Hidden;// 复制ClassFont类同名成员到本类成员
            this.Italic = srcFont.Italic;// 复制ClassFont类同名成员到本类成员
            this.ItalicBi = srcFont.ItalicBi;// 复制ClassFont类同名成员到本类成员
            this.Kerning = srcFont.Kerning;// 复制ClassFont类同名成员到本类成员
            this.Name = srcFont.Name;// 复制ClassFont类同名成员到本类成员
            this.NameAscii = srcFont.NameAscii;// 复制ClassFont类同名成员到本类成员
            this.NameBi = srcFont.NameBi;// 复制ClassFont类同名成员到本类成员
            this.NameFarEast = srcFont.NameFarEast;// 复制ClassFont类同名成员到本类成员
            //this.NameOther = srcFont.NameOther;// 复制ClassFont类同名成员到本类成员
            this.Outline = srcFont.Outline;// 复制ClassFont类同名成员到本类成员
            this.Position = srcFont.Position;// 复制ClassFont类同名成员到本类成员
            this.Scaling = srcFont.Scaling;// 复制ClassFont类同名成员到本类成员

            this.Shading.clone(srcFont.Shading);// 复制ClassFont类同名成员到本类成员

            this.Shadow = srcFont.Shadow;// 复制ClassFont类同名成员到本类成员
            this.Size = srcFont.Size;// 复制ClassFont类同名成员到本类成员
            this.SizeBi = srcFont.SizeBi;// 复制ClassFont类同名成员到本类成员
            this.SmallCaps = srcFont.SmallCaps;// 复制ClassFont类同名成员到本类成员
            this.Spacing = srcFont.Spacing;// 复制ClassFont类同名成员到本类成员
            this.StrikeThrough = srcFont.StrikeThrough;// 复制ClassFont类同名成员到本类成员
            this.Subscript = srcFont.Subscript;// 复制ClassFont类同名成员到本类成员
            this.Superscript = srcFont.Superscript;// 复制ClassFont类同名成员到本类成员
            this.Underline = srcFont.Underline;// 复制ClassFont类同名成员到本类成员
            this.UnderlineColor = srcFont.UnderlineColor;// 复制ClassFont类同名成员到本类成员

            return;
        }

        // 复制本类到ClassFont类
        public void copy2(ClassFont dstFnt)
        {
            dstFnt.AllCaps = this.AllCaps;// 复制本类成员到ClassFont类同名成员
            dstFnt.Animation = this.Animation;// 复制本类成员到ClassFont类同名成员

            dstFnt.Bold = this.Bold;// 复制本类成员到ClassFont类同名成员
            dstFnt.BoldBi = this.BoldBi;// 复制本类成员到ClassFont类同名成员

            ClassBorders bds = dstFnt.Borders;// 复制本类成员到ClassFont类同名成员
            this.Borders.copy2(ref bds);// 复制本类成员到ClassFont类同名成员

            dstFnt.Color = this.Color;// 复制本类成员到ClassFont类同名成员
            dstFnt.ColorIndex = this.ColorIndex;// 复制本类成员到ClassFont类同名成员
            dstFnt.ColorIndexBi = this.ColorIndexBi;// 复制本类成员到ClassFont类同名成员

            dstFnt.DiacriticColor = this.DiacriticColor;// 复制本类成员到ClassFont类同名成员
            dstFnt.DisableCharacterSpaceGrid = this.DisableCharacterSpaceGrid;// 复制本类成员到ClassFont类同名成员
            dstFnt.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制本类成员到ClassFont类同名成员

            dstFnt.Emboss = this.Emboss;// 复制本类成员到ClassFont类同名成员
            dstFnt.EmphasisMark = this.EmphasisMark;// 复制本类成员到ClassFont类同名成员
            dstFnt.Engrave = this.Engrave;// 复制本类成员到ClassFont类同名成员
            dstFnt.Hidden = this.Hidden;// 复制本类成员到ClassFont类同名成员
            dstFnt.Italic = this.Italic;// 复制本类成员到ClassFont类同名成员
            dstFnt.ItalicBi = this.ItalicBi;// 复制本类成员到ClassFont类同名成员
            dstFnt.Kerning = this.Kerning;// 复制本类成员到ClassFont类同名成员
            dstFnt.Name = this.Name;// 复制本类成员到ClassFont类同名成员
            dstFnt.NameAscii = this.NameAscii;// 复制本类成员到ClassFont类同名成员
            dstFnt.NameBi = this.NameBi;// 复制本类成员到ClassFont类同名成员
            dstFnt.NameFarEast = this.NameFarEast;// 复制本类成员到ClassFont类同名成员
            //dstFnt.NameOther = this.NameOther;// 复制本类成员到ClassFont类同名成员
            dstFnt.Outline = this.Outline;// 复制本类成员到ClassFont类同名成员
            dstFnt.Position = this.Position;// 复制本类成员到ClassFont类同名成员
            dstFnt.Scaling = this.Scaling;// 复制本类成员到ClassFont类同名成员

            ClassShading shd = dstFnt.Shading;// 复制本类成员到ClassFont类同名成员
            this.Shading.copy2(ref shd);// 复制本类成员到ClassFont类同名成员

            dstFnt.Shadow = this.Shadow;// 复制本类成员到ClassFont类同名成员
            dstFnt.Size = this.Size;// 复制本类成员到ClassFont类同名成员
            dstFnt.SizeBi = this.SizeBi;// 复制本类成员到ClassFont类同名成员
            dstFnt.SmallCaps = this.SmallCaps;// 复制本类成员到ClassFont类同名成员
            dstFnt.Spacing = this.Spacing;// 复制本类成员到ClassFont类同名成员
            dstFnt.StrikeThrough = this.StrikeThrough;// 复制本类成员到ClassFont类同名成员
            dstFnt.Subscript = this.Subscript;// 复制本类成员到ClassFont类同名成员
            dstFnt.Superscript = this.Superscript;// 复制本类成员到ClassFont类同名成员
            dstFnt.Underline = this.Underline;// 复制本类成员到ClassFont类同名成员
            dstFnt.UnderlineColor = this.UnderlineColor;// 复制本类成员到ClassFont类同名成员

            return;
        }


    }
}
