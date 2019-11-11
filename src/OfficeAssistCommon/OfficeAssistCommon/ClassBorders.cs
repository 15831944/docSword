using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Collections;

namespace OfficeTools.Common
{
    // 参照WORD.Borders的对象
    public class ClassBorders /*: Borders*/
    {
        public bool AlwaysInFront;// { get; set; }
        //public Application Application { get; }
        public int Count;// { get; }
        public int Creator;// { get; }
        public WdBorderDistanceFrom DistanceFrom;// { get; set; }
        public int DistanceFromBottom;// { get; set; }
        public int DistanceFromLeft;// { get; set; }
        public int DistanceFromRight;// { get; set; }
        public int DistanceFromTop;// { get; set; }
        public int Enable;// { get; set; }
        public bool EnableFirstPageInSection;// { get; set; }
        public bool EnableOtherPagesInSection;// { get; set; }
        public bool HasHorizontal;// { get; }
        public bool HasVertical;// { get; }
        public WdColor InsideColor;// { get; set; }
        public WdColorIndex InsideColorIndex;// { get; set; }
        public WdLineStyle InsideLineStyle;// { get; set; }
        public WdLineWidth InsideLineWidth;// { get; set; }
        public bool JoinBorders;// { get; set; }
        public WdColor OutsideColor;// { get; set; }
        public WdColorIndex OutsideColorIndex;// { get; set; }
        public WdLineStyle OutsideLineStyle;// { get; set; }
        public WdLineWidth OutsideLineWidth;// { get; set; }
        //public dynamic Parent { get; }
        public bool Shadow;// { get; set; }
        public bool SurroundFooter;// { get; set; }
        public bool SurroundHeader;// { get; set; }

        // 根据索引取值
        public ClassBorder this[WdBorderType Index] 
        { 
            get
            {
                return (ClassBorder)m_hashBorder[Index];
            }
        }

        private Hashtable m_hashBorder = new Hashtable();

        // 复制Word.Borders的对象到本类
        public void clone(Word.Borders obds)
        {
            this.AlwaysInFront = obds.AlwaysInFront; // 进行赋值
            this.Count = obds.Count;// 进行赋值
            this.Creator = obds.Creator;// 进行赋值

            try
            {
                this.DistanceFrom = obds.DistanceFrom;// 进行赋值
                this.DistanceFromBottom = obds.DistanceFromBottom;// 进行赋值
                this.DistanceFromLeft = obds.DistanceFromLeft;// 进行赋值
                this.DistanceFromRight = obds.DistanceFromRight;// 进行赋值
                this.DistanceFromTop = obds.DistanceFromTop;// 进行赋值
            }
            catch (System.Exception ex)
            {
           	    
            }
            finally
            {
                
            }

            this.Enable = obds.Enable;// 进行赋值
            this.EnableFirstPageInSection = obds.EnableFirstPageInSection;// 进行赋值
            this.EnableOtherPagesInSection = obds.EnableOtherPagesInSection;// 进行赋值
            this.HasHorizontal = obds.HasHorizontal;// 进行赋值
            this.HasVertical = obds.HasVertical;// 进行赋值
            this.InsideColor = obds.InsideColor;// 进行赋值
            this.InsideColorIndex = obds.InsideColorIndex;// 进行赋值
            this.InsideLineStyle = obds.InsideLineStyle;// 进行赋值
            this.InsideLineWidth = obds.InsideLineWidth;// 进行赋值
            this.JoinBorders = obds.JoinBorders;// 进行赋值
            this.OutsideColor = obds.OutsideColor;// 进行赋值
            this.OutsideColorIndex = obds.OutsideColorIndex;// 进行赋值
            this.OutsideLineStyle = obds.OutsideLineStyle;// 进行赋值
            this.OutsideLineWidth = obds.OutsideLineWidth;// 进行赋值
            this.Shadow = obds.Shadow;// 进行赋值
            this.SurroundFooter = obds.SurroundFooter;// 进行赋值
            this.SurroundHeader = obds.SurroundHeader;// 进行赋值

            this.m_hashBorder.Clear();

            Border bd = null;

            // 将Border的各成员进行赋值
            for (int i = (int)WdBorderType.wdBorderVertical; i <= (int)WdBorderType.wdBorderTop; i++)
            {
                try
                {
                	bd = (Border)obds[(WdBorderType)i];
                }
                catch (System.Exception ex)
                {
                    continue;
                }

                if(bd != null)
                {
                    ClassBorder classBd = new ClassBorder();
                    classBd.clone(bd);

                    m_hashBorder[i] = classBd;
                }

            }
            
            return;
        }


        public void copy2(ref Word.Borders extBorders)
        {
            extBorders.AlwaysInFront = this.AlwaysInFront;
            //extBorders.Count = this.Count;
            //extBorders.Creator = this.Creator;
            
            
            try
            {
	            extBorders.DistanceFrom = this.DistanceFrom;
	            extBorders.DistanceFromBottom = this.DistanceFromBottom;
	            extBorders.DistanceFromLeft = this.DistanceFromLeft;
	            extBorders.DistanceFromRight = this.DistanceFromRight;
	            extBorders.DistanceFromTop = this.DistanceFromTop;
            }
            catch (System.Exception ex)
            {
           	
            }
            finally
            {
            }


            extBorders.Enable = this.Enable;
            extBorders.EnableFirstPageInSection = this.EnableFirstPageInSection;
            extBorders.EnableOtherPagesInSection = this.EnableOtherPagesInSection;
            //extBorders.HasHorizontal = this.HasHorizontal;
            //extBorders.HasVertical = this.HasVertical;
            extBorders.InsideColor = this.InsideColor;
            extBorders.InsideColorIndex = this.InsideColorIndex;
            extBorders.InsideLineStyle = this.InsideLineStyle;
            extBorders.InsideLineWidth = this.InsideLineWidth;
            extBorders.JoinBorders = this.JoinBorders;
            extBorders.OutsideColor = this.OutsideColor;
            extBorders.OutsideColorIndex = this.OutsideColorIndex;
            extBorders.OutsideLineStyle = this.OutsideLineStyle;
            extBorders.OutsideLineWidth = this.OutsideLineWidth;
            extBorders.Shadow = this.Shadow;
            extBorders.SurroundFooter = this.SurroundFooter;
            extBorders.SurroundHeader = this.SurroundHeader;


            ClassBorder classbd = null;
            for (int i = (int)WdBorderType.wdBorderVertical; i <= (int)WdBorderType.wdBorderTop; i++)
            {
                classbd = (ClassBorder)m_hashBorder[(WdBorderType)i];

                if (classbd != null)
                {
                    Word.Border bd = extBorders[(WdBorderType)i];

                    classbd.copy2(ref bd);

                }

            }

            return;
        }

        //////////////////////////
        public void clone(ClassBorders obds)
        {
            this.AlwaysInFront = obds.AlwaysInFront;
            this.Count = obds.Count;
            this.Creator = obds.Creator;
            this.DistanceFrom = obds.DistanceFrom;
            this.DistanceFromBottom = obds.DistanceFromBottom;
            this.DistanceFromLeft = obds.DistanceFromLeft;
            this.DistanceFromRight = obds.DistanceFromRight;
            this.DistanceFromTop = obds.DistanceFromTop;
            this.Enable = obds.Enable;
            this.EnableFirstPageInSection = obds.EnableFirstPageInSection;
            this.EnableOtherPagesInSection = obds.EnableOtherPagesInSection;
            this.HasHorizontal = obds.HasHorizontal;
            this.HasVertical = obds.HasVertical;
            this.InsideColor = obds.InsideColor;
            this.InsideColorIndex = obds.InsideColorIndex;
            this.InsideLineStyle = obds.InsideLineStyle;
            this.InsideLineWidth = obds.InsideLineWidth;
            this.JoinBorders = obds.JoinBorders;
            this.OutsideColor = obds.OutsideColor;
            this.OutsideColorIndex = obds.OutsideColorIndex;
            this.OutsideLineStyle = obds.OutsideLineStyle;
            this.OutsideLineWidth = obds.OutsideLineWidth;
            this.Shadow = obds.Shadow;
            this.SurroundFooter = obds.SurroundFooter;
            this.SurroundHeader = obds.SurroundHeader;

            this.m_hashBorder.Clear();


            ClassBorder bd = null;
            for (int i = (int)WdBorderType.wdBorderVertical; i <= (int)WdBorderType.wdBorderTop; i++)
            {
                bd = (ClassBorder)obds[(WdBorderType)i];

                if (bd != null)
                {
                    ClassBorder classBd = new ClassBorder();
                    classBd.clone(bd);

                    m_hashBorder[i] = classBd;
                }

            }

            return;
        }


        public void copy2(ref ClassBorders extBorders)
        {
            extBorders.AlwaysInFront = this.AlwaysInFront;
            //extBorders.Count = this.Count;
            //extBorders.Creator = this.Creator;


            try
            {
                extBorders.DistanceFrom = this.DistanceFrom;
                extBorders.DistanceFromBottom = this.DistanceFromBottom;
                extBorders.DistanceFromLeft = this.DistanceFromLeft;
                extBorders.DistanceFromRight = this.DistanceFromRight;
                extBorders.DistanceFromTop = this.DistanceFromTop;
            }
            catch (System.Exception ex)
            {

            }
            finally
            {
            }


            extBorders.Enable = this.Enable;
            extBorders.EnableFirstPageInSection = this.EnableFirstPageInSection;
            extBorders.EnableOtherPagesInSection = this.EnableOtherPagesInSection;
            //extBorders.HasHorizontal = this.HasHorizontal;
            //extBorders.HasVertical = this.HasVertical;
            extBorders.InsideColor = this.InsideColor;
            extBorders.InsideColorIndex = this.InsideColorIndex;
            extBorders.InsideLineStyle = this.InsideLineStyle;
            extBorders.InsideLineWidth = this.InsideLineWidth;
            extBorders.JoinBorders = this.JoinBorders;
            extBorders.OutsideColor = this.OutsideColor;
            extBorders.OutsideColorIndex = this.OutsideColorIndex;
            extBorders.OutsideLineStyle = this.OutsideLineStyle;
            extBorders.OutsideLineWidth = this.OutsideLineWidth;
            extBorders.Shadow = this.Shadow;
            extBorders.SurroundFooter = this.SurroundFooter;
            extBorders.SurroundHeader = this.SurroundHeader;


            ClassBorder classbd = null;
            for (int i = (int)WdBorderType.wdBorderVertical; i <= (int)WdBorderType.wdBorderTop; i++)
            {
                classbd = (ClassBorder)m_hashBorder[(WdBorderType)i];

                if (classbd != null)
                {
                    ClassBorder bd = extBorders[(WdBorderType)i];

                    classbd.copy2(ref bd);

                }

            }

            return;
        }


    }
}
