using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OfficeAssist
{
    public partial class FormAllRelTxtOut : Form
    {
        public FormAllRelTxtOut()
        {
            InitializeComponent();
        }

        public void SetContent(String strCnt)
        {
            txtAllRelsOut.Text = strCnt;
        }


    }
}
