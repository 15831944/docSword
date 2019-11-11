using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeAssist.localDB.Model
{
    public class tblUniformStyleHistoryDocs
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
        /// fullPathDoc
        /// </summary>
        public virtual string fullPathDoc
        {
            get;
            set;
        }

    }
}
