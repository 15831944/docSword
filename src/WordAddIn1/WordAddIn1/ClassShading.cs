using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace OfficeAssist
{
    // 复制WORD.Shading同名参数值
    public class ClassShading /*: Shading*/
    {
        // public Application Application { get; } // WORD.Shading同名参数值
        public WdColor BackgroundPatternColor;// { get; set; }// WORD.Shading同名参数值
        public WdColorIndex BackgroundPatternColorIndex;// { get; set; }// WORD.Shading同名参数值
        public int Creator;// { get; }// WORD.Shading同名参数值
        public WdColor ForegroundPatternColor;// { get; set; }// WORD.Shading同名参数值
        public WdColorIndex ForegroundPatternColorIndex;// { get; set; }// WORD.Shading同名参数值
        // public dynamic Parent { get; }// WORD.Shading同名参数值
        public WdTextureIndex Texture;// { get; set; }// WORD.Shading同名参数值

        // 复制WORD.Shading同名参数值
        public void clone(Word.Shading shd)
        {
            this.BackgroundPatternColor = shd.BackgroundPatternColor;// 复制WORD.Shading同名参数值
            this.BackgroundPatternColorIndex = shd.BackgroundPatternColorIndex;// 复制WORD.Shading同名参数值
            this.Creator = shd.Creator;// 复制WORD.Shading同名参数值

            this.ForegroundPatternColor = shd.ForegroundPatternColor;// 复制WORD.Shading同名参数值
            this.ForegroundPatternColorIndex = shd.ForegroundPatternColorIndex;// 复制WORD.Shading同名参数值

            this.Texture = shd.Texture;// 复制WORD.Shading同名参数值

            return;
        }

        // 复制到WORD.Shading同名参数值
        public void copy2(ref Word.Shading oShd)
        {
            oShd.BackgroundPatternColor = this.BackgroundPatternColor;// 复制到WORD.Shading同名参数值
            oShd.BackgroundPatternColorIndex = this.BackgroundPatternColorIndex;// 复制到WORD.Shading同名参数值
            //oShd.Creator = this.Creator;

            oShd.ForegroundPatternColor = this.ForegroundPatternColor;// 复制到WORD.Shading同名参数值
            oShd.ForegroundPatternColorIndex = this.ForegroundPatternColorIndex;// 复制到WORD.Shading同名参数值

            oShd.Texture = this.Texture;// 复制到WORD.Shading同名参数值

            return;
        }

        //////////////////////////////////////////
        // 复制ClassShading同名参数值
        public void clone(ClassShading shd)
        {
            this.BackgroundPatternColor = shd.BackgroundPatternColor;// 复制ClassShading同名参数值
            this.BackgroundPatternColorIndex = shd.BackgroundPatternColorIndex;// 复制ClassShading同名参数值
            //this.Creator = shd.Creator;// 复制ClassShading同名参数值

            this.ForegroundPatternColor = shd.ForegroundPatternColor;// 复制ClassShading同名参数值
            this.ForegroundPatternColorIndex = shd.ForegroundPatternColorIndex;// 复制ClassShading同名参数值

            this.Texture = shd.Texture;// 复制ClassShading同名参数值

            return;
        }

        // 复制到ClassShading同名参数值
        public void copy2(ref ClassShading oShd)
        {
            oShd.BackgroundPatternColor = this.BackgroundPatternColor;// 复制到ClassShading同名参数值
            oShd.BackgroundPatternColorIndex = this.BackgroundPatternColorIndex;// 复制到ClassShading同名参数值
            //oShd.Creator = this.Creator;

            oShd.ForegroundPatternColor = this.ForegroundPatternColor;// 复制到ClassShading同名参数值
            oShd.ForegroundPatternColorIndex = this.ForegroundPatternColorIndex;// 复制到ClassShading同名参数值

            oShd.Texture = this.Texture;// 复制到ClassShading同名参数值

            return;
        }


    }
}
