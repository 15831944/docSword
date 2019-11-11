using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace OfficeAssist
{
    public class ClassShareLibTreeNodeComparer : IComparer
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

            ShareContributorOper.TypeNode ndXtag = (ShareContributorOper.TypeNode)trX.Tag;
            ShareContributorOper.TypeNode ndYtag = (ShareContributorOper.TypeNode)trY.Tag;


            String strX = trX.Name;
            String strY = trY.Name;

            String strXType = "";
            String strYType = "";


            if (trX.Level == 0 && trY.Level == 0)
            {
                return 0; // unchange
            }


            if (ndXtag != null)
            {
                strXType = ndXtag.type;// strXtag.Substring(0, 1);
            }

            if (ndYtag != null)
            {
                strYType = ndYtag.type;// strYtag.Substring(0, 1);
            }

            if (ndXtag == null)
            {
                if (ndYtag == null)
                {
                    return 0;
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                if (ndYtag == null)
                {
                    return 1;
                }
                else
                {
                    if (!(strXType.Equals("1") || strXType.Equals("2") ))
                    {
                        if (!(strYType.Equals("1") || strYType.Equals("2")))
                        {
                            return trX.Name.CompareTo(trY.Name);
                        }
                        else if (strYType.Equals("1") || strYType.Equals("2"))
                        {
                            return -1;
                        }
                        else
                        {
                            return -1;
                            // return trX.Name.CompareTo(trY.Name);
                        }
                    }
                    else if (strXType.Equals("1") || strXType.Equals("2"))
                    {
                        if (!(strYType.Equals("1") || strYType.Equals("2")))
                        {
                            return 1;
                        }
                        else if (strYType.Equals("1") || strYType.Equals("2"))
                        {
                            return trX.Name.CompareTo(trY.Name);
                        }
                        else
                        {
                            return 1;
                            // return trX.Name.CompareTo(trY.Name);
                        }
                    }
                    else
                    {
                        return trX.Name.CompareTo(trY.Name);
                    }
                }
            }

        }//

    }//
}
