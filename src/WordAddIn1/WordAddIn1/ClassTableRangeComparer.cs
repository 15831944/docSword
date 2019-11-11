using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace OfficeAssist
{
    class ClassTableRangeComparer : IComparer
    {
        public int Compare(object x, object y)
        {
            Word.Table px = (Word.Table)x;
            Word.Table py = (Word.Table)y;

            int xStart = px.Range.Start;
            int xEnd = px.Range.End;

            int yStart = py.Range.Start;
            int yEnd = py.Range.End;

            if (xStart < yStart && xEnd < yEnd)
            {
                return -1;
            }
            else if (xStart > yStart && xEnd > yEnd)
            {
                return 1;
            }
            else
            {

            }


            return 0;
        }

    }
}
