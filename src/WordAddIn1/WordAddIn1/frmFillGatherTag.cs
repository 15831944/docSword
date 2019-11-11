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
    public partial class frmFillGatherTag : Form
    {
        public frmFillGatherTag()
        {
            InitializeComponent();
        }

        public Boolean m_bTagFullPath = false;
        public Boolean m_bTagShortFileName = false;
        public Boolean m_bTagOnlyDirectory = false;
        public Boolean m_bTagSelfFill = false;
        public String m_strSelfFillTxt = "";
        public Boolean m_bTagTableSn = false;
        public Boolean m_bTagAbsPageNum = false;
        public Boolean m_bTagShortFileNameNoExt = false;

        private void btnTagDialogOK_Click(object sender, EventArgs e)
        {
            if (rdBtnTagFullPath.Checked)
            {
                m_bTagFullPath = true;
            }
            else if(rdBtnShortFileName.Checked)
            {
                m_bTagShortFileName = true;
            }
            else if (rdBtnOnlyDirectory.Checked)
            {
                m_bTagOnlyDirectory = true;
            }
            else if (rdBtnSelfFill.Checked)
            {
                m_bTagSelfFill = true;
                m_strSelfFillTxt = txtTagDialogSelfFill.Text.Trim();

                if (m_strSelfFillTxt.Equals(""))
                {
                    MessageBox.Show("自填内容不能为空!");
                }
            }
            else if (rdBtnTblSn.Checked)
            {
                m_bTagTableSn = true;
            }
            else if(rdBtnTagAbsPageNum.Checked)
            {
                m_bTagAbsPageNum = true;
            }
            else if(rdBtnFileShortNameNoExt.Checked)
            {
                m_bTagShortFileNameNoExt = true;
            }
            else
            {
                m_bTagFullPath = true;
                // MessageBox.Show("SHOULD NOT SEE IT");
            }

            return;
        }
    }
}
