using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeAssist.localDB.Model
{
    public class tblListLevelSchemes
    {

        /// <summary>
        /// ID
        /// </summary>
        public virtual Guid ID
        {
            get;
            set;
        }
        /// <summary>
        /// schemeName
        /// </summary>
        public virtual string schemeName
        {
            get;
            set;
        }
        /// <summary>
        /// isPreBuiltIn
        /// </summary>
        public virtual bool isPreBuiltIn
        {
            get;
            set;
        }
        /// <summary>
        /// bVisible
        /// </summary>
        public virtual bool bVisible
        {
            get;
            set;
        }
        /// <summary>
        /// nOrderSn
        /// </summary>
        public virtual int nOrderSn
        {
            get;
            set;
        }

    }
}
