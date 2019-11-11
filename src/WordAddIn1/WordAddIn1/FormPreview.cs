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
    public partial class FormPreview : Form
    {
        public FormPreview()
        {
            InitializeComponent();
        }


        private void FormPreview_Load(object sender, EventArgs e)
        {
            richTextBoxCnt.Top = this.ClientRectangle.Top;
            richTextBoxCnt.Left = this.ClientRectangle.Left;

            richTextBoxCnt.Width = this.ClientSize.Width;
            richTextBoxCnt.Height = this.ClientSize.Height;

            return;
        }


        private void FormPreview_ClientSizeChanged(object sender, EventArgs e)
        {
            richTextBoxCnt.Width = this.ClientSize.Width;
            richTextBoxCnt.Height = this.ClientSize.Height;

            return;
        }
    }
}
