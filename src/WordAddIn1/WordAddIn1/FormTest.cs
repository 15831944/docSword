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
    public partial class FormTest : Form
    {
        public String strFileName = "";

        public FormTest()
        {
            InitializeComponent();
        }

        private void FormTest_Load(object sender, EventArgs e)
        {
            if (!strFileName.Equals(""))
            {
                webBrowser1.Navigate(strFileName);
            }

            return;
        }


    }
}
