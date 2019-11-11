using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace OfficeAssist
{
    // 自定义的章节序号类,用于保存ListLevel的信息
    public class ClassListLevel /*: ListLevel*/
    {
        // 同ListLevel类同名的参数的信息
        public WdListLevelAlignment Alignment;// { get; set; }
        // public Application Application;// { get; }
        public int Creator;// { get; }// 同ListLevel类同名的参数的信息

        // public Font Font;// { get; set; }// 同ListLevel类同名的参数的信息
        public ClassFont Font;

        public int Index;// { get; }// 同ListLevel类同名的参数的信息
        public string LinkedStyle = "";// { get; set; }// 同ListLevel类同名的参数的信息
        public string NumberFormat = "";// { get; set; }// 同ListLevel类同名的参数的信息
        public float NumberPosition;// { get; set; }// 同ListLevel类同名的参数的信息
        public WdListNumberStyle NumberStyle;// { get; set; }// 同ListLevel类同名的参数的信息

        public WdListNumberStyle NumberStyleSel;// { get; set; }// 同ListLevel类同名的参数的信息

        //public  dynamic Parent;// { get; }
        public InlineShape PictureBullet;// { get; }// 同ListLevel类同名的参数的信息
        public int ResetOnHigher;// { get; set; }// 同ListLevel类同名的参数的信息
        public bool ResetOnHigherOld;// { get; set; }// 同ListLevel类同名的参数的信息
        public int StartAt;// { get; set; }// 同ListLevel类同名的参数的信息
        public float TabPosition;// { get; set; }// 同ListLevel类同名的参数的信息
        public float TextPosition;// { get; set; }// 同ListLevel类同名的参数的信息
        public WdTrailingCharacter TrailingCharacter;// { get; set; }// 同ListLevel类同名的参数的信息

        // 复制ListLevel类同名的参数的信息
        public void clone(Word.ListLevel lvl)
        {
            this.Alignment = lvl.Alignment;// 同ListLevel类同名的参数的信息
            // this.Application = lvl.Application;
            this.Creator = lvl.Creator;// 同ListLevel类同名的参数的信息

            this.Font.clone(lvl.Font); // // 同ListLevel类同名的参数的信息

            this.Index = lvl.Index;// 同ListLevel类同名的参数的信息
            this.LinkedStyle = lvl.LinkedStyle;// 同ListLevel类同名的参数的信息
            this.NumberFormat = lvl.NumberFormat;// 同ListLevel类同名的参数的信息
            this.NumberPosition = lvl.NumberPosition;// 同ListLevel类同名的参数的信息
            this.NumberStyle = lvl.NumberStyle;// 同ListLevel类同名的参数的信息

            // this.PictureBullet = lvl.PictureBullet;// 同ListLevel类同名的参数的信息

            this.ResetOnHigher = lvl.ResetOnHigher;// 同ListLevel类同名的参数的信息
            this.ResetOnHigherOld = lvl.ResetOnHigherOld;// 同ListLevel类同名的参数的信息
            this.StartAt = lvl.StartAt;// 同ListLevel类同名的参数的信息
            this.TabPosition = lvl.TabPosition;// 同ListLevel类同名的参数的信息
            this.TextPosition = lvl.TextPosition;// 同ListLevel类同名的参数的信息
            this.TrailingCharacter = lvl.TrailingCharacter;// 同ListLevel类同名的参数的信息

            return;
        }

        // 复制到ListLevel类同名的参数的信息
        public void copy2(ref Word.ListLevel dstLvl)
        {
            dstLvl.Alignment = this.Alignment;// 复制到ListLevel类同名的参数的信息
            // dstLvl.Application = this.Application;
            // dstLvl.Creator = this.Creator;

            this.Font.copy2(dstLvl.Font);// 复制到ListLevel类同名的参数的信息
            // dstLvl.Font = this.Font.Duplicate; // 

            // dstLvl.Index = this.Index;
            dstLvl.LinkedStyle = this.LinkedStyle;// 复制到ListLevel类同名的参数的信息
            dstLvl.NumberFormat = this.NumberFormat;// 复制到ListLevel类同名的参数的信息
            dstLvl.NumberPosition = this.NumberPosition;// 复制到ListLevel类同名的参数的信息
            dstLvl.NumberStyle = this.NumberStyle;// 复制到ListLevel类同名的参数的信息

            // dstLvl.PictureBullet = this.PictureBullet;// 复制到ListLevel类同名的参数的信息

            dstLvl.ResetOnHigher = this.ResetOnHigher;// 复制到ListLevel类同名的参数的信息
            dstLvl.ResetOnHigherOld = this.ResetOnHigherOld;// 复制到ListLevel类同名的参数的信息
            dstLvl.StartAt = this.StartAt;// 复制到ListLevel类同名的参数的信息
            dstLvl.TabPosition = this.TabPosition;// 复制到ListLevel类同名的参数的信息
            dstLvl.TextPosition = this.TextPosition;// 复制到ListLevel类同名的参数的信息
            dstLvl.TrailingCharacter = this.TrailingCharacter;// 复制到ListLevel类同名的参数的信息

            return;
        }

        // 复制ClassListLevel类同名的参数的信息
        public void clone(ClassListLevel lvl)
        {
            this.Alignment = lvl.Alignment;// 复制ClassListLevel类同名的参数的信息
            // this.Application = lvl.Application;// 复制ClassListLevel类同名的参数的信息
            this.Creator = lvl.Creator;// 复制ClassListLevel类同名的参数的信息

            this.Font.clone(lvl.Font); // 复制ClassListLevel类同名的参数的信息 

            this.Index = lvl.Index;// 复制ClassListLevel类同名的参数的信息
            this.LinkedStyle = lvl.LinkedStyle;// 复制ClassListLevel类同名的参数的信息
            this.NumberFormat = lvl.NumberFormat;// 复制ClassListLevel类同名的参数的信息
            this.NumberPosition = lvl.NumberPosition;// 复制ClassListLevel类同名的参数的信息
            this.NumberStyle = lvl.NumberStyle;// 复制ClassListLevel类同名的参数的信息

            this.NumberStyleSel = lvl.NumberStyleSel;// 复制ClassListLevel类同名的参数的信息

            // this.PictureBullet = lvl.PictureBullet;// 复制ClassListLevel类同名的参数的信息

            this.ResetOnHigher = lvl.ResetOnHigher;// 复制ClassListLevel类同名的参数的信息
            this.ResetOnHigherOld = lvl.ResetOnHigherOld;// 复制ClassListLevel类同名的参数的信息
            this.StartAt = lvl.StartAt;// 复制ClassListLevel类同名的参数的信息
            this.TabPosition = lvl.TabPosition;// 复制ClassListLevel类同名的参数的信息
            this.TextPosition = lvl.TextPosition;// 复制ClassListLevel类同名的参数的信息
            this.TrailingCharacter = lvl.TrailingCharacter;// 复制ClassListLevel类同名的参数的信息

            return;
        }


        // 复制到ClassListLevel类同名的参数的信息
        public void copy2(ref ClassListLevel dstLvl)
        {
            dstLvl.Alignment = this.Alignment;// 复制到ClassListLevel类同名的参数的信息
            // dstLvl.Application = this.Application;
            // dstLvl.Creator = this.Creator;

            this.Font.copy2(dstLvl.Font);// 复制到ClassListLevel类同名的参数的信息
            // dstLvl.Font = this.Font.Duplicate; // 

            // dstLvl.Index = this.Index;
            dstLvl.LinkedStyle = this.LinkedStyle;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.NumberFormat = this.NumberFormat;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.NumberPosition = this.NumberPosition;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.NumberStyle = this.NumberStyle;// 复制到ClassListLevel类同名的参数的信息

            dstLvl.NumberStyleSel = this.NumberStyleSel;// 复制到ClassListLevel类同名的参数的信息

            // dstLvl.PictureBullet = this.PictureBullet;// 复制到ClassListLevel类同名的参数的信息

            dstLvl.ResetOnHigher = this.ResetOnHigher;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.ResetOnHigherOld = this.ResetOnHigherOld;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.StartAt = this.StartAt;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.TabPosition = this.TabPosition;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.TextPosition = this.TextPosition;// 复制到ClassListLevel类同名的参数的信息
            dstLvl.TrailingCharacter = this.TrailingCharacter;// 复制到ClassListLevel类同名的参数的信息

            return;
        }


    }
}
