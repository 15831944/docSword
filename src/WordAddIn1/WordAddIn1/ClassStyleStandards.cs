using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OfficeTools.Common;


namespace OfficeAssist
{
    public class ClassStyleStandards
    {
        // 章节目录
            // 总体
        public ClassFont                headingTocTopFnt       = new ClassFont();
        public ClassParagraphFormat     headingTocTopParaFmt   = new ClassParagraphFormat();
            // 1-9级
        public ClassFont[]              headingTocFnts     = new ClassFont[9];
        public ClassParagraphFormat[]   headingTocParaFmts = new ClassParagraphFormat[9];
            // 字体、段落  （中英文字体名称、字号、字形、段间距，其余的不应该修改？）


        // 图文目录
            // 总体
        public ClassFont                tuwenTocTopFnt = new ClassFont();
        public ClassParagraphFormat     tuwenTocTopParaFmt = new ClassParagraphFormat();
            // 图：字体、段落（）
        public ClassFont                tuwenTocTuFnt = new ClassFont();
        public ClassParagraphFormat tuwenTocTuParaFmt = new ClassParagraphFormat();
            // 表：字体、段落（）
        public ClassFont                tuwenTocTableFnt = new ClassFont();
        public ClassParagraphFormat tuwenTocTableParaFmt = new ClassParagraphFormat();

        // 章节
        public ClassFont[]              headingFonts = new ClassFont[9];
        public ClassParagraphFormat[] headingParaFmts = new ClassParagraphFormat[9];

        public Boolean bHeadingListLevels = false;
        public ClassListLevel[]  headingListLevels = new ClassListLevel[9];    

        // 表格
        public ClassFont tableFnt = new ClassFont();
        public ClassParagraphFormat tableParaFmt = new ClassParagraphFormat();

        // 题注
        public ClassFont tizhuFnt = new ClassFont();
        public ClassParagraphFormat tizhuParaFmt = new ClassParagraphFormat();

        // 序号段落
            // 总体
        public ClassFont listParaTopFnt = new ClassFont();
        public ClassParagraphFormat listParaTopParaFmt = new ClassParagraphFormat();
            // 1-9级
        public ClassFont[] listParaFnts = new ClassFont[9];
        public ClassParagraphFormat[] listParaParaFmts = new ClassParagraphFormat[9];

        // 正文区
        public ClassFont textBodyFnt = new ClassFont();
        public ClassParagraphFormat textBodyParaFmt = new ClassParagraphFormat();

        // 节，页眉页脚


        ClassStyleStandards()
        {
            // headingTocFnts
            for (int i = 0; i < 9; i++)
            {
                headingTocFnts[i] = new ClassFont();
            }

            // headingTocParaFmts
            for (int i = 0; i < 9; i++)
            {
                headingTocParaFmts[i] = new ClassParagraphFormat();
            }

            // headingFonts
            for (int i = 0; i < 9; i++)
            {
                headingFonts[i] = new ClassFont();
            }

            // headingParaFmts
            for (int i = 0; i < 9; i++)
            {
                headingParaFmts[i] = new ClassParagraphFormat();
            }

            // headingListLevels
            for (int i = 0; i < 9; i++)
            {
                headingListLevels[i] = new ClassListLevel();
            }
            
            // listParaFnts
            for (int i = 0; i < 9; i++)
            {
                listParaFnts[i] = new ClassFont();
            }

            // listParaParaFmts
            for (int i = 0; i < 9; i++)
            {
                listParaParaFmts[i] = new ClassParagraphFormat();
            }

            return;

        }



    }// class


}// namespace
