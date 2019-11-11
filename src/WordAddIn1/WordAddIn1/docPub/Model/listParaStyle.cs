using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Serialization;
using System.Windows.Forms;

namespace OfficeAssist.docPub.Model
{
    [Serializable]
    [XmlType]
    public class listParaStyle
    {
        [XmlAttribute]
        public Boolean bEnable = true;

        [XmlAttribute]
        public int nLevel = 0;
        [XmlElement]
        public nhFont Fnt { get; set; }//= new nhFont();
        [XmlElement]
        public nhParaFmt ParaFmt { get; set; } // = new nhParaFmt();

        public listParaStyle()
        {
            Fnt = new nhFont();
            ParaFmt = new nhParaFmt();

            return;
        }

        public void clone(listParaStyle oth)
        {
            bEnable = oth.bEnable;
            nLevel = oth.nLevel;

            Fnt.clone(oth.Fnt);
            ParaFmt.clone(oth.ParaFmt);

            return;
        }

        public void copy2(listParaStyle oth)
        {
            oth.bEnable = bEnable;
            oth.nLevel = nLevel;

            oth.Fnt.clone(Fnt);
            oth.ParaFmt.clone(ParaFmt);

            return;
        }


        public int formatString(RichTextBox rchTxt, String strPreBlanks)
        {
            if (bEnable)
            {
                rchTxt.AppendText("启用\r\n");

                Fnt.formatString(rchTxt, strPreBlanks);
                ParaFmt.formatString(rchTxt, strPreBlanks);
            }
            else
            {
                rchTxt.AppendText("停用\r\n");
            }

            return 0;
        }

    }
}
