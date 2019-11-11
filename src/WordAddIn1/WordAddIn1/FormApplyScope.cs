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
    public partial class FormApplyScope : Form
    {
        public Boolean m_bScopeAllDoc = true;
        public Boolean m_bIgnoreToc = true;
        public Boolean m_bIgnoreTable = true;
        public Boolean m_bIgnorePages = false;
        public uint m_nIgnoredPages = 1;
        public Boolean m_bIgnoreHeadings = false;
        public Boolean m_bIgnoreTextBody = true;
        public Boolean m_bIgnoreFont = false;
        public Boolean m_bIgnoreParaFmt = false;


        public FormApplyScope()
        {
            InitializeComponent();
        }

        private void FormApplyScope_Load(object sender, EventArgs e)
        {
            chkIgnoreParaFormat.Checked = m_bIgnoreParaFmt;
            chkIgnoreFont.Checked = m_bIgnoreFont;
            
            // chkIgnoreTextBody.Checked = m_bIgnoreTextBody;
            // chkIgnoreHeadings.Checked = m_bIgnoreHeadings;

            chkIgnoreTextBody.Checked = true;
            chkIgnoreHeadings.Checked = false;

            chkIgnorePages.Checked = m_bIgnorePages;
            txtIgnorePages.Enabled = chkIgnorePages.Checked;
            txtIgnorePages.Text = ""+m_nIgnoredPages;

            chkIgnoreTable.Checked = m_bIgnoreTable;
            chkIgnoreTOC.Checked = m_bIgnoreToc;

            radioBtnStyleSelection.Checked = !m_bScopeAllDoc;
            radioBtnStyleAllDoc.Checked = m_bScopeAllDoc;

            return;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            m_bIgnoreParaFmt = chkIgnoreParaFormat.Checked;
            m_bIgnoreFont = chkIgnoreFont.Checked;

            m_bIgnoreTextBody = chkIgnoreTextBody.Checked;
            m_bIgnoreHeadings = chkIgnoreHeadings.Checked;

            m_bIgnorePages = chkIgnorePages.Checked;

            if (m_bIgnorePages)
            {
                if (!uint.TryParse(txtIgnorePages.Text, out m_nIgnoredPages))
                {
                    m_nIgnoredPages = 1;
                }
            }

            m_bIgnoreTable = chkIgnoreTable.Checked;
            m_bIgnoreToc = chkIgnoreTOC.Checked;

            m_bScopeAllDoc = radioBtnStyleAllDoc.Checked;
            m_bScopeAllDoc = !radioBtnStyleSelection.Checked;

            return;
        }

        private void chkIgnorePages_CheckedChanged(object sender, EventArgs e)
        {
            txtIgnorePages.Enabled = chkIgnorePages.Checked;

            return;
        }


        private void txtIgnorePages_Leave(object sender, EventArgs e)
        {
            uint nTmp = 1;

            if (!uint.TryParse(txtIgnorePages.Text,out nTmp))
            {
                MessageBox.Show("请输入正整数");
                return;
            }

            return;
        }







    }
}
