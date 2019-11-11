using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAssist
{
    // 字体尺寸排序对比类
    public class CFontSizeComparer : IComparer
    {
        public int Compare(object x, object y)
        {
            Word.Paragraph px = (Word.Paragraph)x; // 转换
            Word.Paragraph py = (Word.Paragraph)y; // 转换

            return (int)(px.Range.Font.Size - py.Range.Font.Size); // 比较大小
        }
    }
}
