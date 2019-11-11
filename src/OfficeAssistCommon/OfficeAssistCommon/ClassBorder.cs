using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace OfficeTools.Common
{
    public class ClassBorder /*: Border*/
    {
        // public Application Application { get; }
        public WdPageBorderArt ArtStyle;// { get; set; }
        public int ArtWidth;// { get; set; }
        public WdColor Color;// { get; set; }
        public WdColorIndex ColorIndex;// { get; set; }
        public int Creator;// { get; }
        public bool Inside;// { get; }
        public WdLineStyle LineStyle;// { get; set; }
        public WdLineWidth LineWidth;// { get; set; }
        // public dynamic Parent { get; }
        public bool Visible;// { get; set; }

        // copy
        public void clone(Word.Border obd)
        {
            this.ArtStyle = obd.ArtStyle; // 进行赋值
            this.ArtWidth = obd.ArtWidth; // 进行赋值
            this.Color = obd.Color; // 进行赋值
            this.ColorIndex = obd.ColorIndex;// 进行赋值
            this.Creator = obd.Creator;// 进行赋值
            this.Inside = obd.Inside;// 进行赋值
            this.LineStyle = obd.LineStyle;// 进行赋值
            this.LineWidth = obd.LineWidth;// 进行赋值
            this.Visible = obd.Visible;// 进行赋值
     
            return;
        }

        // 复制到对象
        public void copy2(ref Word.Border extBorder)
        {
            extBorder.ArtStyle = this.ArtStyle; // 进行赋值
            extBorder.ArtWidth = this.ArtWidth;// 进行赋值
            extBorder.Color = this.Color;// 进行赋值
            extBorder.ColorIndex = this.ColorIndex;// 进行赋值
            //extBorder.Creator = this.Creator;
            //extBorder.Inside = this.Inside;
            extBorder.LineStyle = this.LineStyle;// 进行赋值
            extBorder.LineWidth = this.LineWidth;// 进行赋值
            extBorder.Visible = this.Visible;// 进行赋值

            return;
        }

        ////////////////////////////////////////////
        // 复制ClassBorder对象的内容
        public void clone(ClassBorder obd)
        {
            this.ArtStyle = obd.ArtStyle;// 进行赋值
            this.ArtWidth = obd.ArtWidth;// 进行赋值
            this.Color = obd.Color;// 进行赋值
            this.ColorIndex = obd.ColorIndex;// 进行赋值
            this.Creator = obd.Creator;// 进行赋值
            this.Inside = obd.Inside;// 进行赋值
            this.LineStyle = obd.LineStyle;// 进行赋值
            this.LineWidth = obd.LineWidth;// 进行赋值
            this.Visible = obd.Visible;// 进行赋值

            return;
        }

        // 复制到对象
        public void copy2(ref ClassBorder extBorder)
        {
            extBorder.ArtStyle = this.ArtStyle; // 进行赋值
            extBorder.ArtWidth = this.ArtWidth;// 进行赋值
            extBorder.Color = this.Color;// 进行赋值
            extBorder.ColorIndex = this.ColorIndex;// 进行赋值
            //extBorder.Creator = this.Creator;
            //extBorder.Inside = this.Inside;
            extBorder.LineStyle = this.LineStyle;// 进行赋值
            extBorder.LineWidth = this.LineWidth;// 进行赋值
            extBorder.Visible = this.Visible;// 进行赋值

            return;
        }




    }
}
