using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeTools.Common;
using System.IO;
using System.Collections;


namespace OfficeAssist
{
    public partial class FormLoadLicFile : Form
    {
        public Boolean m_bLicLegal = false;

        ClassOfficeCommon m_cmnTools = null;
        classMultiEditionCenter m_edtCenter = null;


        public FormLoadLicFile()
        {
            InitializeComponent();
        }


        public void setItems(ClassOfficeCommon oCmnTools,
                             classMultiEditionCenter edtCenter)
        {
            m_cmnTools = oCmnTools;
            m_edtCenter = edtCenter;

            return;
        }


        private void btnSelectLicFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofDialog = new OpenFileDialog();

            ofDialog.Filter = "Text files(*.txt)|*.txt|Text files(*.lic)|*.lic|All files(*.*)|*.*";

            DialogResult res = ofDialog.ShowDialog();

            if (res == DialogResult.Cancel)
            {
                return;
            }

            String strFilePath = ofDialog.FileName;
            txtBoxSelectedLicFileLoc.Text = strFilePath;


            // Parse

            StreamReader rd = new StreamReader(strFilePath);
            String strCnt = rd.ReadToEnd();
            rd.Close();


            String strMachineID_seg = "";
            String strInfo = "";

            Hashtable hashPerm = new Hashtable();
            DateTime dt = DateTime.MinValue;

            int nRet = m_cmnTools.DecodeLic(strCnt, ref strMachineID_seg, ref hashPerm, ref dt, false);

            if (nRet != 0)
            {
                strInfo = "许可文件非法";
                m_bLicLegal = false;
            }
            else
            {
                m_bLicLegal = true;

                double dbDays = m_cmnTools.DateDiff(DateTime.Now, dt);
                int nDays = (dbDays < 0) ? (int)Math.Floor(dbDays) : (int)Math.Ceiling(dbDays);

                if (nDays > 0)
                {
                    strInfo = "已过期" + nDays + "天，到期日期：" + dt.ToString("yyyy年MM月dd日");
                }
                else
                {
                    strInfo = "还余" + (-1 * nDays) + "天到期，到期日期：" + dt.ToString("yyyy年MM月dd日");
                }
            }

            txtBoxDecodeInfo.Text = strInfo;

            return;
        }

        private void btnLoadInto_Click(object sender, EventArgs e)
        {
            if (!m_bLicLegal)
            {
                MessageBox.Show("许可文件非法，不能导入");
            }

            return;
        }

        private void btnBackupCurLicFile_Click(object sender, EventArgs e)
        {
            if (!m_edtCenter.IsExistCurSoloLic())
            {
                MessageBox.Show("当前无单机版许可");
                return;
            }


            SaveFileDialog svDialog = new SaveFileDialog();

            svDialog.Filter = "Text files(*.txt)|*.txt|Text files(*.lic)|*.lic|All files(*.*)|*.*";

            DialogResult res = svDialog.ShowDialog();

            if (res == DialogResult.Cancel)
            {
                return;
            }
            String strDestLicFile = svDialog.FileName;

            int nRet = m_edtCenter.BackupCurSoloLic(strDestLicFile);

            if (nRet == -1)
            {
                MessageBox.Show("当前无单机版许可");
                return;
            }
            else if (nRet == -2)
            {
                MessageBox.Show("保存到的文件被占用不能覆盖：" + strDestLicFile);
                return;
            }
            else if (nRet == -3)
            {
                MessageBox.Show("备份时出错，请确保选择的文件的磁盘有充分空间或权限：" + strDestLicFile);
                return;
            }

            MessageBox.Show("完成：" + strDestLicFile);

            return;
        }




    }
}
