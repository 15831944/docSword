using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace OfficeAssist
{
    // 字符串比较类
    public class ClassStringComparer : IComparer
    {
        public int Compare(object x, object y)
        {
            String strX = (String)x;
            String strY = (String)y;

            char[] chx = strX.ToCharArray();
            char[] chy = strY.ToCharArray();

            String strXprefix = "", strYprefix = "";
            String strXnum = "", strYnum = "";
            int nx = chx.Length - 1, ny = chy.Length - 1;

            while (nx > 0 && chx[nx] >= '0' && chx[nx] <= '9') // 先判断是否数字
            {
                nx--;
            }

            while (ny > 0 && chy[ny] >= '0' && chy[ny] <= '9')// 先判断是否数字
            {
                ny--;
            }

            strXprefix = strX.Substring(0, nx + 1); // 数字前缀
            strYprefix = strY.Substring(0, ny + 1); // 数字前缀

            int nRet = 0;

            nRet = strXprefix.CompareTo(strYprefix);// 数字前缀比较

            if (nRet != 0)
            {
                return nRet;
            }

            strXnum = strX.Substring(nx + 1);
            strYnum = strY.Substring(ny + 1);

            int nNumx = 0;
            int nNumy = 0;

            Boolean bx = int.TryParse(strXnum, out nNumx);
            Boolean by = int.TryParse(strYnum, out nNumy);

            if (!bx && !by)
            {
                nRet = strX.CompareTo(strY);
            }
            else
            {
                nRet = nNumx - nNumy;
            }

            return nRet;
        }

    }
}
