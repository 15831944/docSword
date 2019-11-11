using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OfficeAssist
{
    public class CRelTreeNodeComparer : IComparer
    {
        /// <summary>
        /// 关联树节点比较类
        /// </summary>
        /// <param name="x">TreeNode</param>
        /// <param name="y">TreeNode</param>
        /// <returns></returns>
        public int Compare(object x, object y)
        {
            TreeNode trX = (TreeNode)x;
            TreeNode trY = (TreeNode)y;


            String strX = trX.Name;
            String strY = trY.Name;

            char[] chx = strX.ToCharArray();
            char[] chy = strY.ToCharArray();

            String strXprefix = "", strYprefix = "";
            String strXnum = "", strYnum = "";
            int nx = chx.Length-1, ny = chy.Length-1;

            while (nx > 0 && chx[nx] >= '0' && chx[nx] <= '9') // 先判断是否是数字
            {
                nx--;
            }

            while (ny > 0 && chy[ny] >= '0' && chy[ny] <= '9') // 先判断是否是数字
            {
                ny--;
            }

            strXprefix = strX.Substring(0, nx + 1); // 取前缀
            strYprefix = strY.Substring(0, ny + 1);// 取前缀

            int nRet = 0;

            nRet = strXprefix.CompareTo(strYprefix); // 前缀比较

            if (nRet != 0)
            {
                return nRet;
            }

            strXnum = strX.Substring(nx + 1); // 取数字
            strYnum = strY.Substring(ny + 1); // 取数字

            int nNumx = 0;
            int nNumy = 0;

            Boolean bx = int.TryParse(strXnum, out nNumx); // 转换
            Boolean by = int.TryParse(strYnum, out nNumy); // 转换

            if (!bx && !by)
            {
                nRet = strX.CompareTo(strY); // 比较
            }
            else
            {
                nRet = nNumx - nNumy;
            }

            return nRet;
        }

    }
}
