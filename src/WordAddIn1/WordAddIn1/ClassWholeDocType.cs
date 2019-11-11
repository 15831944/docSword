using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OfficeTools.Common;

using System.Collections;
using System.Collections.Specialized;


namespace OfficeAssist
{
    public class ClassWholeDocType
    {
	    public class sepPart
	    {
		    public int 				        nIndex      = -1; 	                        // 序号，0代表“其余”
            public String                   strName     = "";                           // 显示名称
		    public ClassFont  				cFont       = new ClassFont();  	        // 字体
		    public ClassParagraphFormat 	cParaFmt    = new ClassParagraphFormat(); 	// 段落
            public Boolean                  bHeader     = false;                        // 页眉
            public Boolean                  bFooter     = false;                        // 页脚

            public void clone(sepPart osep)
            {
                nIndex = osep.nIndex;
                strName = osep.strName;
                bHeader = osep.bHeader;
                bFooter = osep.bFooter;

                osep.cFont.SelCopy2(cFont);
                osep.cParaFmt.SelCopy2(cParaFmt);

                return;
            }

	    }

        // 封面
	    // ArrayList(sepParts)
        public Boolean b1stPageEnable = false; // 是否启用
        public ArrayList arrs1stPagePart = new ArrayList(); // 部分列表，1--n表示是第1--n部分，其余部分--0
        public Hashtable hsh1stPagePart = new Hashtable(); // k/v: name/sepParts
	
        // 章节目录
        public Boolean bHeadingTocEnable = false;
        // 0--total, 1-9 levels
        public ClassFont[] arrsHeadingTocFnt = new ClassFont[10];
        public ClassParagraphFormat[] arrsHeadingTocParaFmt = new ClassParagraphFormat[10];
	
        // 图文目录
        public Boolean bTuWenTocEnable = false;
        public ClassFont tuWenTocTotalFnt = new ClassFont();
        public ClassParagraphFormat tuWenTocTotalParaFmt = new ClassParagraphFormat();
	
        // 章节
        public Boolean bHeadingEnable = false;
        public String headingStyleSchemeName = "";
        public String headingSnSchemeName = "";

        // 表格
        public Boolean bTableEnable = false;
        public ClassFont tableTotalFont = new ClassFont();
        public ClassParagraphFormat tableTotalParaFmt = new ClassParagraphFormat();
        public Boolean bClearIndent = false;

        // 题注
        public Boolean bTizhuEnable = false;
        public ClassFont tizhuFont = new ClassFont();
        public ClassParagraphFormat tizhuParaFmt = new ClassParagraphFormat();
	
        // 正文区
        public Boolean bTextBodyZoneEnable = false;
        public ClassFont textbodyZoneFont = new ClassFont();
        public ClassParagraphFormat textbodyZoneParaFmt = new ClassParagraphFormat();

        // 节和页眉页脚
        // ArrayList(sepParts)
        public Boolean bSectionEnable = false;
        public ArrayList arrsSections = new ArrayList(); // 部分列表，1--n表示是第1--n部分，其余部分--0
        public Hashtable hshSectionPart = new Hashtable(); // k/v: name/sepParts

        public ClassWholeDocType()
        {
            for(int i = 0; i < 10; i++)
            {
                arrsHeadingTocFnt[i] = new ClassFont();
                arrsHeadingTocParaFmt[i] = new ClassParagraphFormat();
            }

            return;
        }



    }// class

}// namespace 
