using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Collections;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using OfficeTools.Common;


namespace OfficeAssist
{

    /// <summary>
    /// 描述字体参数的对话框类，用于保存对话框设置的内容
    /// </summary>
    public class ClassFontDialogItems
    {
        public int AllCaps; // 同WORD.FONT类的同名成员
        public int Animations; // 同WORD.FONT类的同名成员
        public int Bold; // 同WORD.FONT类的同名成员
        public int BoldBi; // 同WORD.FONT类的同名成员
        public int CharAccent; // 同WORD.FONT类的同名成员
        public int CharacterWidthGrid; // 同WORD.FONT类的同名成员
        public int Color; // 同WORD.FONT类的同名成员
        public int ColorBi; // 同WORD.FONT类的同名成员
        public String ColorRGB = "";//-16777216,  // 同WORD.FONT类的同名成员
        public int DoubleStrikeThrough; // 同WORD.FONT类的同名成员
        public int Emboss; // 同WORD.FONT类的同名成员
        public int Engrave; // 同WORD.FONT类的同名成员
        public String Font = "+中文正文";//+中文正文, // 同WORD.FONT类的同名成员
        public String FontHighAnsi = "+西文正文";//:+西文正文, // 同WORD.FONT类的同名成员
        public String FontLowAnsi = "+西文正文";//:+西文正文, // 同WORD.FONT类的同名成员
        public String FontMajor = "+中文正文";//:+中文正文, // 同WORD.FONT类的同名成员
        public String FontNameBi = "+正文 CS 字体";//:+正文 CS 字体, // 同WORD.FONT类的同名成员
        public int Hidden; // 同WORD.FONT类的同名成员
        public int Italic; // 同WORD.FONT类的同名成员
        public int ItalicBi; // 同WORD.FONT类的同名成员
        public int Kerning; // 同WORD.FONT类的同名成员
        public String KerningMin = "八号"; //String, // 同WORD.FONT类的同名成员
        public int Outline; // 同WORD.FONT类的同名成员
        public String Points = ""; // 初号, // 同WORD.FONT类的同名成员
        public String PointsBi = "11";//11, // 同WORD.FONT类的同名成员
        public String Position = "";//:0 磅, // 同WORD.FONT类的同名成员
        public String Scale = "";//:100%, // 同WORD.FONT类的同名成员
        public int Shadow;//:0, // 同WORD.FONT类的同名成员
        public int SmallCaps;//:0, // 同WORD.FONT类的同名成员
        public String Spacing = "";//:0 磅, // 同WORD.FONT类的同名成员
        public int StrikeThrough;//:0, // 同WORD.FONT类的同名成员
        public int Subscript;//:0, // 同WORD.FONT类的同名成员
        public int Superscript;//:0, // 同WORD.FONT类的同名成员
        public int Underline;//:0, // 同WORD.FONT类的同名成员
        public String UnderlineColor = "";//:-16777216, String, // 同WORD.FONT类的同名成员


        Hashtable m_hashUnderlineDialog2WordFont = new Hashtable();
        Hashtable m_hashUnderlineWordFont2Dialog = new Hashtable();

        Hashtable m_hashPointsDialog2Size = new Hashtable();
        Hashtable m_hashPointsSize2Dialog = new Hashtable();


        public ClassFontDialogItems()
        {
            // 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(0, Word.WdUnderline.wdUnderlineNone);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(2, Word.WdUnderline.wdUnderlineWords);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(1, Word.WdUnderline.wdUnderlineSingle);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(3, Word.WdUnderline.wdUnderlineDouble);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(5, Word.WdUnderline.wdUnderlineThick);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(4, Word.WdUnderline.wdUnderlineDotted);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(10, Word.WdUnderline.wdUnderlineDottedHeavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(6, Word.WdUnderline.wdUnderlineDash);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(11, Word.WdUnderline.wdUnderlineDashHeavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(14, Word.WdUnderline.wdUnderlineDashLong);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(15, Word.WdUnderline.wdUnderlineDashLongHeavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(7, Word.WdUnderline.wdUnderlineDotDash);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(12, Word.WdUnderline.wdUnderlineDotDashHeavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(8, Word.WdUnderline.wdUnderlineDotDotDash);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(13, Word.WdUnderline.wdUnderlineDotDotDashHeavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(9, Word.WdUnderline.wdUnderlineWavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(16, Word.WdUnderline.wdUnderlineWavyHeavy);// 加载初始化值到查表中
            m_hashUnderlineDialog2WordFont.Add(17, Word.WdUnderline.wdUnderlineWavyDouble);// 加载初始化值到查表中

            // 
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineNone, 0);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineWords, 2);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineSingle, 1);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDouble, 3);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineThick, 5);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDotted, 4);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDottedHeavy, 10);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDash, 6);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDashHeavy, 11);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDashLong, 14);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDashLongHeavy, 15);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDotDash, 7);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDotDashHeavy, 12);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDotDotDash, 8);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineDotDotDashHeavy, 13);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineWavy, 9);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineWavyHeavy, 16);// 加载初始化值到查表中
            m_hashUnderlineWordFont2Dialog.Add(Word.WdUnderline.wdUnderlineWavyDouble, 17);// 加载初始化值到查表中

            // 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("初号", 42f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小初", 36f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("一号", 26f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小一", 24f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("二号", 22f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小二", 18f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("三号", 16f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小三", 15f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("四号", 14f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小四", 12f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("五号", 10f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小五", 9f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("六号", 7f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("小六", 6f);// 加载字体字号初始化值到查表中
            m_hashPointsDialog2Size.Add("七号", 5f);// 加载字体字号初始化值到查表中
            //m_hashPointsDialog2Size.Add("八号", 5f);// 加载字体字号初始化值到查表中
            
            // 
            m_hashPointsSize2Dialog.Add(42f, "初号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(36f, "小初");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(26f, "一号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(24f, "小一");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(22f, "二号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(18f, "小二");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(16f, "三号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(15f, "小三");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(14f, "四号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(12f, "小四");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(10f, "五号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(9f, "小五");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(7f, "六号");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(6f, "小六");// 加载字体字号初始化值到查表中
            m_hashPointsSize2Dialog.Add(5f, "七号");// 加载字体字号初始化值到查表中
            //m_hashPointsSize2Dialog.Add(5f, "八号");// 加载字体字号初始化值到查表中

            return;
        }

        // 复制WORD自带的FONT对话框的参数
        public void clone(dynamic fntDialog) // dynamic fntDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatFont];
        {
            this.AllCaps = fntDialog.AllCaps;// 复制WORD自带的FONT对话框的参数
            this.Animations = fntDialog.Animations;// 复制WORD自带的FONT对话框的参数
            this.Bold = fntDialog.Bold;// 复制WORD自带的FONT对话框的参数
            this.BoldBi = fntDialog.BoldBi;// 复制WORD自带的FONT对话框的参数
            this.CharAccent = fntDialog.CharAccent;// 复制WORD自带的FONT对话框的参数
            this.CharacterWidthGrid = fntDialog.CharacterWidthGrid;// 复制WORD自带的FONT对话框的参数
            this.Color = fntDialog.Color;// 复制WORD自带的FONT对话框的参数
            this.ColorBi = fntDialog.ColorBi;// 复制WORD自带的FONT对话框的参数
            // 复制WORD自带的FONT对话框的参数
            this.ColorRGB = fntDialog.ColorRGB;
            // 复制WORD自带的FONT对话框的参数
            this.DoubleStrikeThrough = fntDialog.DoubleStrikeThrough;
            this.Emboss = fntDialog.Emboss;// 复制WORD自带的FONT对话框的参数
            this.Engrave = fntDialog.Engrave;// 复制WORD自带的FONT对话框的参数
            this.Font = fntDialog.Font;//+中文正文,// 复制WORD自带的FONT对话框的参数
            this.FontHighAnsi = fntDialog.FontHighAnsi;//:+西文正文,// 复制WORD自带的FONT对话框的参数
            this.FontLowAnsi = fntDialog.FontLowAnsi;//:+西文正文,// 复制WORD自带的FONT对话框的参数
            this.FontMajor = fntDialog.FontMajor;//:+中文正文,// 复制WORD自带的FONT对话框的参数
            this.FontNameBi = fntDialog.FontNameBi;//:+正文 CS 字体,// 复制WORD自带的FONT对话框的参数
            this.Hidden = fntDialog.Hidden;// 复制WORD自带的FONT对话框的参数
            this.Italic = fntDialog.Italic;// 复制WORD自带的FONT对话框的参数
            this.ItalicBi = fntDialog.ItalicBi;// 复制WORD自带的FONT对话框的参数

            this.Kerning = fntDialog.Kerning;// 复制WORD自带的FONT对话框的参数
            this.KerningMin = fntDialog.KerningMin;// 复制WORD自带的FONT对话框的参数

            this.Outline = fntDialog.Outline;// 复制WORD自带的FONT对话框的参数
            this.Points = fntDialog.Points; // 初号,// 复制WORD自带的FONT对话框的参数
            this.PointsBi = fntDialog.PointsBi;//11,// 复制WORD自带的FONT对话框的参数
            this.Position = fntDialog.Position;//:0 磅,// 复制WORD自带的FONT对话框的参数
            this.Scale = fntDialog.Scale;//:100%,// 复制WORD自带的FONT对话框的参数
            this.Shadow = fntDialog.Shadow;//:0,// 复制WORD自带的FONT对话框的参数
            this.SmallCaps = fntDialog.SmallCaps;//:0,// 复制WORD自带的FONT对话框的参数
            this.Spacing = fntDialog.Spacing;//:0 磅,// 复制WORD自带的FONT对话框的参数
            this.StrikeThrough = fntDialog.StrikeThrough;//:0,// 复制WORD自带的FONT对话框的参数
            this.Subscript = fntDialog.Subscript;//:0,// 复制WORD自带的FONT对话框的参数
            this.Superscript = fntDialog.Superscript;//:0,// 复制WORD自带的FONT对话框的参数
            this.Underline = fntDialog.Underline;//:0,// 复制WORD自带的FONT对话框的参数

            this.UnderlineColor = fntDialog.UnderlineColor;// 复制WORD自带的FONT对话框的参数

            return;
        }

        // 复制自编的FONT对话框的参数
        public void clone(ClassFontDialogItems fntDialog) // dynamic fntDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatFont];
        {
            this.AllCaps = fntDialog.AllCaps;// 复制自编的FONT对话框的参数
            this.Animations = fntDialog.Animations;// 复制自编的FONT对话框的参数
            this.Bold = fntDialog.Bold;// 复制自编的FONT对话框的参数
            this.BoldBi = fntDialog.BoldBi;// 复制自编的FONT对话框的参数
            this.CharAccent = fntDialog.CharAccent;// 复制自编的FONT对话框的参数
            this.CharacterWidthGrid = fntDialog.CharacterWidthGrid;// 复制自编的FONT对话框的参数
            this.Color = fntDialog.Color;// 复制自编的FONT对话框的参数
            this.ColorBi = fntDialog.ColorBi;// 复制自编的FONT对话框的参数

            this.ColorRGB = fntDialog.ColorRGB;// 复制自编的FONT对话框的参数

            this.DoubleStrikeThrough = fntDialog.DoubleStrikeThrough;// 复制自编的FONT对话框的参数
            this.Emboss = fntDialog.Emboss;// 复制自编的FONT对话框的参数
            this.Engrave = fntDialog.Engrave;// 复制自编的FONT对话框的参数
            this.Font = fntDialog.Font;//+中文正文,// 复制自编的FONT对话框的参数
            this.FontHighAnsi = fntDialog.FontHighAnsi;//:+西文正文,// 复制自编的FONT对话框的参数
            this.FontLowAnsi = fntDialog.FontLowAnsi;//:+西文正文,// 复制自编的FONT对话框的参数
            this.FontMajor = fntDialog.FontMajor;//:+中文正文,// 复制自编的FONT对话框的参数
            this.FontNameBi = fntDialog.FontNameBi;//:+正文 CS 字体,// 复制自编的FONT对话框的参数
            this.Hidden = fntDialog.Hidden;// 复制自编的FONT对话框的参数
            this.Italic = fntDialog.Italic;// 复制自编的FONT对话框的参数
            this.ItalicBi = fntDialog.ItalicBi;// 复制自编的FONT对话框的参数

            this.Kerning = fntDialog.Kerning;// 复制自编的FONT对话框的参数
            this.KerningMin = fntDialog.KerningMin;// 复制自编的FONT对话框的参数

            this.Outline = fntDialog.Outline;// 复制自编的FONT对话框的参数
            this.Points = fntDialog.Points; // 初号,// 复制自编的FONT对话框的参数
            this.PointsBi = fntDialog.PointsBi;//11,// 复制自编的FONT对话框的参数
            this.Position = fntDialog.Position;//:0 磅,// 复制自编的FONT对话框的参数
            this.Scale = fntDialog.Scale;//:100%,// 复制自编的FONT对话框的参数
            this.Shadow = fntDialog.Shadow;//:0,// 复制自编的FONT对话框的参数
            this.SmallCaps = fntDialog.SmallCaps;//:0,// 复制自编的FONT对话框的参数
            this.Spacing = fntDialog.Spacing;//:0 磅,// 复制自编的FONT对话框的参数
            this.StrikeThrough = fntDialog.StrikeThrough;//:0,// 复制自编的FONT对话框的参数
            this.Subscript = fntDialog.Subscript;//:0,// 复制自编的FONT对话框的参数
            this.Superscript = fntDialog.Superscript;//:0,// 复制自编的FONT对话框的参数
            this.Underline = fntDialog.Underline;//:0,// 复制自编的FONT对话框的参数

            this.UnderlineColor = fntDialog.UnderlineColor;// 复制自编的FONT对话框的参数

            return;
        }

        // 复制ClassFONT的参数
        public void clone(ClassFont cFnt)
        {
//             if (cFnt.Name.Equals(""))
//                 return;

            this.AllCaps = cFnt.AllCaps;// 复制ClassFONT的参数
            this.Animations = (int)cFnt.Animation;// wdAnimationNone,// 复制ClassFONT的参数
            this.Bold = cFnt.Bold;// 复制ClassFONT的参数

            this.ColorRGB = "" + (int)cFnt.Color;// wdColorAutomatic,// 复制ClassFONT的参数
            this.DoubleStrikeThrough = cFnt.DoubleStrikeThrough;// 复制ClassFONT的参数
            this.Emboss = cFnt.Emboss;// 复制ClassFONT的参数
            this.Engrave = cFnt.Engrave;// 复制ClassFONT的参数
            this.FontHighAnsi = cFnt.NameAscii;//"Arial Unicode MS",// 复制ClassFONT的参数
            //this.FontLowAnsi = cFnt.NameOther;//"Arial Unicode MS",// 复制ClassFONT的参数
            this.FontMajor = cFnt.NameFarEast;//"微软雅黑",// 复制ClassFONT的参数
            this.Font = cFnt.Name;// 复制ClassFONT的参数

            this.Hidden = cFnt.Hidden;// 复制ClassFONT的参数
            this.Italic = cFnt.Italic;// 复制ClassFONT的参数

            if (m_hashPointsSize2Dialog.Contains(cFnt.Kerning))// 复制ClassFONT的参数
            {
                this.Kerning = 1;// 复制ClassFONT的参数
                this.KerningMin = (String)m_hashPointsSize2Dialog[cFnt.Kerning];// this.Kerning,// 复制ClassFONT的参数
            }
            else
            {
                this.Kerning = 0;// 复制ClassFONT的参数
                this.KerningMin = "";// 复制ClassFONT的参数
            }

            this.Outline = cFnt.Outline;// 复制ClassFONT的参数

            this.Position = cFnt.Position + " 磅";// 复制ClassFONT的参数

            this.Scale = cFnt.Scaling + "%";// 复制ClassFONT的参数

            this.Shadow = cFnt.Shadow;// 复制ClassFONT的参数
            this.SmallCaps = cFnt.SmallCaps;// 复制ClassFONT的参数

            this.Spacing = cFnt.Spacing + " 磅";// 复制ClassFONT的参数

            this.StrikeThrough = cFnt.StrikeThrough;// 复制ClassFONT的参数
            this.Subscript = cFnt.Subscript;// 复制ClassFONT的参数
            this.Superscript = cFnt.Superscript;// 复制ClassFONT的参数

            this.UnderlineColor = "" + (int)cFnt.UnderlineColor;// wdColorAutomatic,// 复制ClassFONT的参数

            this.Underline = (int)WdUnderline.wdUnderlineNone;// 复制ClassFONT的参数
            if (m_hashUnderlineWordFont2Dialog.Contains(cFnt.Underline))// 复制ClassFONT的参数
            {
                this.Underline = (int)m_hashUnderlineWordFont2Dialog[cFnt.Underline];// wdUnderlineNone,// 复制ClassFONT的参数
            }

            if (m_hashPointsSize2Dialog.Contains(cFnt.Size))// 复制ClassFONT的参数
            {
                this.Points = (String)m_hashPointsSize2Dialog[cFnt.Size];// cFnt.point,// 复制ClassFONT的参数
            }
            else
            {
                this.Points = "" + cFnt.Size;// 复制ClassFONT的参数
            }

            this.CharacterWidthGrid = (cFnt.DisableCharacterSpaceGrid ? 1:0);// 复制ClassFONT的参数
            //this.Color = cFnt.ColorDialog;

            //this.PointsBi = "11";

            return;
        }

        // 复制到Word内置的FONT对话框的参数，进行设置
        public void copy2(dynamic fntDialog) // dynamic fntDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatFont];
        {
            fntDialog.AllCaps = this.AllCaps;// 复制到Word内置的FONT对话框的参数
            fntDialog.Animations = this.Animations;// 复制到Word内置的FONT对话框的参数
            fntDialog.Bold = this.Bold;// 复制到Word内置的FONT对话框的参数
            fntDialog.BoldBi = this.BoldBi;// 复制到Word内置的FONT对话框的参数
            fntDialog.CharAccent = this.CharAccent;// 复制到Word内置的FONT对话框的参数
            fntDialog.CharacterWidthGrid = this.CharacterWidthGrid;// 复制到Word内置的FONT对话框的参数
            fntDialog.Color = this.Color;// 复制到Word内置的FONT对话框的参数
            fntDialog.ColorBi = this.ColorBi;// 复制到Word内置的FONT对话框的参数
            fntDialog.ColorRGB = this.ColorRGB;//-16777216，// 复制到Word内置的FONT对话框的参数
            fntDialog.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制到Word内置的FONT对话框的参数
            fntDialog.Emboss = this.Emboss;// 复制到Word内置的FONT对话框的参数
            fntDialog.Engrave = this.Engrave;// 复制到Word内置的FONT对话框的参数
            fntDialog.Font = this.Font;//+中文正文，// 复制到Word内置的FONT对话框的参数
            fntDialog.FontHighAnsi = this.FontHighAnsi;//:+西文正文，// 复制到Word内置的FONT对话框的参数
            fntDialog.FontLowAnsi = this.FontLowAnsi;//:+西文正文，// 复制到Word内置的FONT对话框的参数
            fntDialog.FontMajor = this.FontMajor;//:+中文正文，// 复制到Word内置的FONT对话框的参数
            fntDialog.FontNameBi = this.FontNameBi;//:+正文 CS 字体，// 复制到Word内置的FONT对话框的参数
            fntDialog.Hidden = this.Hidden;// 复制到Word内置的FONT对话框的参数
            fntDialog.Italic = this.Italic;// 复制到Word内置的FONT对话框的参数
            fntDialog.ItalicBi = this.ItalicBi;// 复制到Word内置的FONT对话框的参数
            fntDialog.Kerning = this.Kerning;// 复制到Word内置的FONT对话框的参数
            fntDialog.KerningMin = this.KerningMin;// 复制到Word内置的FONT对话框的参数
            fntDialog.Outline = this.Outline;// 复制到Word内置的FONT对话框的参数
            fntDialog.Points = this.Points; // 初号，// 复制到Word内置的FONT对话框的参数
            fntDialog.PointsBi = this.PointsBi;//11，// 复制到Word内置的FONT对话框的参数
            fntDialog.Position = this.Position;//:0 磅，// 复制到Word内置的FONT对话框的参数
            fntDialog.Scale = this.Scale;//:100%，// 复制到Word内置的FONT对话框的参数
            fntDialog.Shadow = this.Shadow;//:0，// 复制到Word内置的FONT对话框的参数
            fntDialog.SmallCaps = this.SmallCaps;//:0，// 复制到Word内置的FONT对话框的参数
            fntDialog.Spacing = this.Spacing;//:0 磅，// 复制到Word内置的FONT对话框的参数
            fntDialog.StrikeThrough = this.StrikeThrough;//:0，// 复制到Word内置的FONT对话框的参数
            fntDialog.Subscript = this.Subscript;//:0，// 复制到Word内置的FONT对话框的参数
            fntDialog.Superscript = this.Superscript;//:0，// 复制到Word内置的FONT对话框的参数
            fntDialog.Underline = this.Underline;//:0，// 复制到Word内置的FONT对话框的参数
            fntDialog.UnderlineColor = this.UnderlineColor;//:-16777216，// 复制到Word内置的FONT对话框的参数

            return;
        }

        // 复制到自定义FONT对话框的参数
        public void copy2(ClassFontDialogItems fntDialog) // dynamic fntDialog = app.Dialogs[Word.WdWordDialog.wdDialogFormatFont];
        {
            fntDialog.AllCaps = this.AllCaps;// 复制到自定义FONT对话框的参数
            fntDialog.Animations = this.Animations;// 复制到自定义FONT对话框的参数
            fntDialog.Bold = this.Bold;// 复制到自定义FONT对话框的参数
            fntDialog.BoldBi = this.BoldBi;// 复制到自定义FONT对话框的参数
            fntDialog.CharAccent = this.CharAccent;// 复制到自定义FONT对话框的参数
            fntDialog.CharacterWidthGrid = this.CharacterWidthGrid;// 复制到自定义FONT对话框的参数
            fntDialog.Color = this.Color;// 复制到自定义FONT对话框的参数
            fntDialog.ColorBi = this.ColorBi;// 复制到自定义FONT对话框的参数
            fntDialog.ColorRGB = this.ColorRGB;//-16777216，// 复制到自定义FONT对话框的参数
            fntDialog.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制到自定义FONT对话框的参数
            fntDialog.Emboss = this.Emboss;// 复制到自定义FONT对话框的参数
            fntDialog.Engrave = this.Engrave;// 复制到自定义FONT对话框的参数
            fntDialog.Font = this.Font;//+中文正文，// 复制到自定义FONT对话框的参数
            fntDialog.FontHighAnsi = this.FontHighAnsi;//:+西文正文，// 复制到自定义FONT对话框的参数
            fntDialog.FontLowAnsi = this.FontLowAnsi;//:+西文正文，// 复制到自定义FONT对话框的参数
            fntDialog.FontMajor = this.FontMajor;//:+中文正文，// 复制到自定义FONT对话框的参数
            fntDialog.FontNameBi = this.FontNameBi;//:+正文 CS 字体，// 复制到自定义FONT对话框的参数
            fntDialog.Hidden = this.Hidden;// 复制到自定义FONT对话框的参数
            fntDialog.Italic = this.Italic;// 复制到自定义FONT对话框的参数
            fntDialog.ItalicBi = this.ItalicBi;// 复制到自定义FONT对话框的参数
            fntDialog.Kerning = this.Kerning;// 复制到自定义FONT对话框的参数
            fntDialog.KerningMin = this.KerningMin;// 复制到自定义FONT对话框的参数
            fntDialog.Outline = this.Outline;// 复制到自定义FONT对话框的参数
            fntDialog.Points = this.Points; // 初号，// 复制到自定义FONT对话框的参数
            fntDialog.PointsBi = this.PointsBi;//11，// 复制到自定义FONT对话框的参数
            fntDialog.Position = this.Position;//:0 磅，// 复制到自定义FONT对话框的参数
            fntDialog.Scale = this.Scale;//:100%，// 复制到自定义FONT对话框的参数
            fntDialog.Shadow = this.Shadow;//:0，// 复制到自定义FONT对话框的参数
            fntDialog.SmallCaps = this.SmallCaps;//:0，// 复制到自定义FONT对话框的参数
            fntDialog.Spacing = this.Spacing;//:0 磅，// 复制到自定义FONT对话框的参数
            fntDialog.StrikeThrough = this.StrikeThrough;//:0，// 复制到自定义FONT对话框的参数
            fntDialog.Subscript = this.Subscript;//:0，// 复制到自定义FONT对话框的参数
            fntDialog.Superscript = this.Superscript;//:0，// 复制到自定义FONT对话框的参数
            fntDialog.Underline = this.Underline;//:0，// 复制到自定义FONT对话框的参数
            fntDialog.UnderlineColor = this.UnderlineColor;//:-16777216，// 复制到自定义FONT对话框的参数

            return;
        }

        // 复制到自定义FONT的参数
        public void copy2(ClassFont fnt)
        {

            fnt.AllCaps = this.AllCaps;// 复制到自定义FONT的参数
            fnt.Animation = (Word.WdAnimation)this.Animations;// wdAnimationNone// 复制到自定义FONT的参数
            fnt.Bold = this.Bold;// 复制到自定义FONT的参数
            fnt.BoldBi = this.BoldBi;// 复制到自定义FONT的参数

            int nTmp = 0;
            fnt.Color = WdColor.wdColorAutomatic;// 复制到自定义FONT的参数
            if (int.TryParse(this.ColorRGB, out nTmp))
            {
                fnt.Color = (WdColor)nTmp;// 复制到自定义FONT的参数
            }

            fnt.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制到自定义FONT的参数
            fnt.Emboss = this.Emboss;// 复制到自定义FONT的参数
            fnt.Engrave = this.Engrave;// 复制到自定义FONT的参数
            fnt.NameAscii = this.FontHighAnsi;//"Arial Unicode MS"// 复制到自定义FONT的参数
            //fnt.NameOther = this.FontLowAnsi;//"Arial Unicode MS"// 复制到自定义FONT的参数
            fnt.NameFarEast = this.FontMajor;//"微软雅黑"// 复制到自定义FONT的参数
            fnt.Name = this.FontMajor;// 复制到自定义FONT的参数

            fnt.Hidden = this.Hidden;// 复制到自定义FONT的参数
            fnt.Italic = this.Italic;// 复制到自定义FONT的参数
            fnt.ItalicBi = this.ItalicBi;// 复制到自定义FONT的参数

            if (this.Kerning > 0)
            {
                if (m_hashPointsDialog2Size.Contains(this.KerningMin))
                {
                    fnt.Kerning = (float)m_hashPointsDialog2Size[this.KerningMin];// this.point// 复制到自定义FONT的参数
                }
                else
                {
                    if (float.TryParse(this.KerningMin, out fnt.Kerning))
                    {

                    }
                    else
                    {
                        fnt.Kerning = 0.0f;// 复制到自定义FONT的参数
                    }
                }
            }
            else
            {
                fnt.Kerning = this.Kerning; // 0.0f// 复制到自定义FONT的参数
            }


            fnt.Outline = this.Outline;// 复制到自定义FONT的参数

            if (int.TryParse(this.Position, out nTmp))// 复制到自定义FONT的参数
            {
                fnt.Position = nTmp;// 复制到自定义FONT的参数
            }
            else
            {
                fnt.Position = nTmp;// 复制到自定义FONT的参数
            }

            if (int.TryParse(this.Scale.Replace("%", ""), out nTmp))// 复制到自定义FONT的参数
            {
                fnt.Scaling = nTmp;// 复制到自定义FONT的参数
            }
            else
            {
                fnt.Scaling = 100;// 复制到自定义FONT的参数
            }

            fnt.Shadow = this.Shadow;// 复制到自定义FONT的参数
            fnt.SmallCaps = this.SmallCaps;// 复制到自定义FONT的参数

            float fTmp = 0.0f;
            if (float.TryParse(this.Spacing, out fTmp))// 复制到自定义FONT的参数
            {
                fnt.Spacing = fTmp;// 复制到自定义FONT的参数
            }
            else
            {
                fnt.Spacing = 0.0f;// 复制到自定义FONT的参数
            }

            fnt.StrikeThrough = this.StrikeThrough;// 复制到自定义FONT的参数
            fnt.Subscript = this.Subscript;// 复制到自定义FONT的参数
            fnt.Superscript = this.Superscript;// 复制到自定义FONT的参数

            fnt.UnderlineColor = WdColor.wdColorAutomatic;// 复制到自定义FONT的参数
            if (int.TryParse(this.UnderlineColor, out nTmp))// 复制到自定义FONT的参数
            {
                fnt.UnderlineColor = (WdColor)nTmp;// 复制到自定义FONT的参数
            }

            fnt.Underline = WdUnderline.wdUnderlineNone;// 复制到自定义FONT的参数
            if (m_hashUnderlineDialog2WordFont.Contains(this.Underline))// 复制到自定义FONT的参数
            {
                fnt.Underline = (WdUnderline)m_hashUnderlineDialog2WordFont[this.Underline];// wdUnderlineNone
            }

            if (m_hashPointsDialog2Size.Contains(this.Points))// 复制到自定义FONT的参数
            {
                fnt.Size = (float)m_hashPointsDialog2Size[this.Points];// this.point
            }
            else
            {
                float.TryParse(this.Points, out fnt.Size);// 复制到自定义FONT的参数
            }

            if (m_hashPointsDialog2Size.Contains(this.PointsBi))// 复制到自定义FONT的参数
            {
                fnt.SizeBi = (float)m_hashPointsDialog2Size[this.PointsBi];// this.PointsBi
            }
            else
            {
                float.TryParse(this.PointsBi, out fnt.SizeBi);// 复制到自定义FONT的参数
            }


            fnt.DisableCharacterSpaceGrid = (this.CharacterWidthGrid != 0);// 复制到自定义FONT的参数

            return;
        }

        // 复制到WORD内置定义FONT的参数
        public void copy2(Word.Font fnt)
        {
            fnt.AllCaps = this.AllCaps;// 复制到WORD内置定义FONT的参数
            fnt.Animation = (Word.WdAnimation)this.Animations;// wdAnimationNone// 复制到WORD内置定义FONT的参数
            fnt.Bold = this.Bold;// 复制到WORD内置定义FONT的参数
            fnt.BoldBi = this.BoldBi;// 复制到WORD内置定义FONT的参数

            int nTmp = 0;
            fnt.Color = WdColor.wdColorAutomatic;// 复制到WORD内置定义FONT的参数
            if (int.TryParse(this.ColorRGB, out nTmp))
            {
                fnt.Color = (WdColor)nTmp;// 复制到WORD内置定义FONT的参数
            }

            fnt.DoubleStrikeThrough = this.DoubleStrikeThrough;// 复制到WORD内置定义FONT的参数
            fnt.Emboss = this.Emboss;// 复制到WORD内置定义FONT的参数
            fnt.Engrave = this.Engrave;// 复制到WORD内置定义FONT的参数
            fnt.NameAscii = this.FontHighAnsi;//"Arial Unicode MS"// 复制到WORD内置定义FONT的参数
            fnt.NameOther = this.FontLowAnsi;//"Arial Unicode MS"// 复制到WORD内置定义FONT的参数
            fnt.NameFarEast = this.FontMajor;//"微软雅黑"// 复制到WORD内置定义FONT的参数
            fnt.Name = this.FontMajor;// 复制到WORD内置定义FONT的参数

            fnt.Hidden = this.Hidden;// 复制到WORD内置定义FONT的参数
            fnt.Italic = this.Italic;// 复制到WORD内置定义FONT的参数
            fnt.ItalicBi = this.ItalicBi;// 复制到WORD内置定义FONT的参数

            float fTmp = 0.0f;
            if (this.Kerning > 0)
            {
                if (m_hashPointsDialog2Size.Contains(this.KerningMin))// 复制到WORD内置定义FONT的参数
                {
                    fnt.Kerning = (float)m_hashPointsDialog2Size[this.KerningMin];// this.point// 复制到WORD内置定义FONT的参数
                }
                else
                {
                    if (float.TryParse(this.KerningMin, out fTmp))// 复制到WORD内置定义FONT的参数
                    {
                        fnt.Kerning = fTmp;// 复制到WORD内置定义FONT的参数
                    }
                    else
                    {
                        fnt.Kerning = 0.0f;// 复制到WORD内置定义FONT的参数
                    }
                }
            }
            else
            {
                fnt.Kerning = this.Kerning; // 0.0f// 复制到WORD内置定义FONT的参数
            }

            fnt.Outline = this.Outline;// 复制到WORD内置定义FONT的参数

            if (int.TryParse(this.Position, out nTmp))// 复制到WORD内置定义FONT的参数
            {
                fnt.Position = nTmp;// 复制到WORD内置定义FONT的参数
            }
            else
            {
                fnt.Position = 0;// 复制到WORD内置定义FONT的参数
            }

            if (int.TryParse(this.Scale.Replace("%", ""), out nTmp))// 复制到WORD内置定义FONT的参数
            {
                fnt.Scaling = nTmp;// 复制到WORD内置定义FONT的参数
            }
            else
            {
                fnt.Scaling = 100;// 复制到WORD内置定义FONT的参数
            }

            fnt.Shadow = this.Shadow;// 复制到WORD内置定义FONT的参数
            fnt.SmallCaps = this.SmallCaps;// 复制到WORD内置定义FONT的参数


            if (float.TryParse(this.Spacing, out fTmp))// 复制到WORD内置定义FONT的参数
            {
                fnt.Spacing = fTmp;// 复制到WORD内置定义FONT的参数
            }
            else
            {
                fnt.Spacing = 0.0f;// 复制到WORD内置定义FONT的参数
            }


            fnt.StrikeThrough = this.StrikeThrough;// 复制到WORD内置定义FONT的参数
            fnt.Subscript = this.Subscript;// 复制到WORD内置定义FONT的参数
            fnt.Superscript = this.Superscript;// 复制到WORD内置定义FONT的参数

            fnt.UnderlineColor = WdColor.wdColorAutomatic;// 复制到WORD内置定义FONT的参数
            if (int.TryParse(this.UnderlineColor, out nTmp))// 复制到WORD内置定义FONT的参数
            {
                fnt.UnderlineColor = (WdColor)nTmp;// 复制到WORD内置定义FONT的参数
            }

            fnt.Underline = WdUnderline.wdUnderlineNone;// 复制到WORD内置定义FONT的参数
            if (m_hashUnderlineDialog2WordFont.Contains(this.Underline))// 复制到WORD内置定义FONT的参数
            {
                fnt.Underline = (WdUnderline)m_hashUnderlineDialog2WordFont[this.Underline];// wdUnderlineNone// 复制到WORD内置定义FONT的参数
            }

            if (m_hashPointsDialog2Size.Contains(this.Points))// 复制到WORD内置定义FONT的参数
            {
                fnt.Size = (float)m_hashPointsDialog2Size[this.Points];// this.point// 复制到WORD内置定义FONT的参数
            }
            else
            {
                if (float.TryParse(this.Points, out fTmp))// 复制到WORD内置定义FONT的参数
                {
                    fnt.Size = fTmp;// 复制到WORD内置定义FONT的参数
                }
                else
                {
                    fnt.Size = 14.0f;// 复制到WORD内置定义FONT的参数
                }
            }

            if (m_hashPointsDialog2Size.Contains(this.PointsBi))// 复制到WORD内置定义FONT的参数
            {
                fnt.SizeBi = (float)m_hashPointsDialog2Size[this.PointsBi];// this.PointsBi// 复制到WORD内置定义FONT的参数
            }
            else
            {
                if (float.TryParse(this.PointsBi, out fTmp))// 复制到WORD内置定义FONT的参数
                {
                    fnt.SizeBi = fTmp;// 复制到WORD内置定义FONT的参数
                }
                else
                {
                    fnt.SizeBi = 14.0f;// 复制到WORD内置定义FONT的参数
                }
            }

            fnt.DisableCharacterSpaceGrid = (this.CharacterWidthGrid != 0);// 复制到WORD内置定义FONT的参数

            return;
        }



    }
}
