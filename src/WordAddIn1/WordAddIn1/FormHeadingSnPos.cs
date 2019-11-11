using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeAssist
{
    public partial class FormHeadingSnPos : Form
    {
        public WdTrailingCharacter m_chAfterHeadingSn; // 编号之后
        public WdListLevelAlignment m_alignment; // alignment

        public float m_NumberPosition; // align pos对齐位置
        public float m_TextPosition; // 文本缩进位置
        public float m_TabPosition; // 制表位添加位置

        public uint m_StartAt;
        public uint m_ResetOnHigher = 0;

        public int m_nLevel = 0;


        public FormHeadingSnPos()
        {
            InitializeComponent(); // 初始化

            m_chAfterHeadingSn = WdTrailingCharacter.wdTrailingTab; // 初始化赋值
            m_alignment = WdListLevelAlignment.wdListLevelAlignLeft;// 初始化赋值

            m_NumberPosition = 0.0f;// 初始化赋值
            m_TextPosition = 0.0f;// 初始化赋值
            m_TabPosition = 0.0f;// 初始化赋值

            m_StartAt = 1;// 初始化赋值
            m_ResetOnHigher = 0;
            m_nLevel = 0;

            return;
        }


        private void chkHeadingSnTabPos_CheckedChanged(object sender, EventArgs e)
        {
            numHeadingSnTabPos.Enabled = chkHeadingSnTabPos.Checked;// 初始化赋值
            return;
        }


        private void FormHeadingSnPos_Load(object sender, EventArgs e)
        {
            cmbHeadingSnAlign.SelectedIndex = (int)m_alignment;// 显示前赋值
            cmbHeadingSnBehindSn.SelectedIndex = (int)m_chAfterHeadingSn;// 显示前赋值

            numHeadingSnAlignPos.Value = (decimal)m_NumberPosition;// 显示前赋值
            numHeadingSnTextIndentPos.Value = (decimal)m_TextPosition;// 显示前赋值

            if (m_TabPosition != (float)Word.WdConstants.wdUndefined)
            {
                chkHeadingSnTabPos.Checked = true;// 显示前赋值
                numHeadingSnTabPos.Enabled = true;/// 显示前赋值
                numHeadingSnTabPos.Value = (decimal)m_TabPosition;// 显示前赋值
            }
            else
            {
                chkHeadingSnTabPos.Checked = false;// 显示前赋值
                numHeadingSnTabPos.Enabled = false;// 显示前赋值
                numHeadingSnTabPos.Value = (decimal)0.0;// 显示前赋值
            }

            numStartAt.Value = m_StartAt;// 显示前赋值

            for (int i = 0; i < m_nLevel; i++)
            {
                cmbResetOnHigher.Items.Add("级别" + (i + 1));
            }

            if (m_nLevel == 0)
            {
                chkResetOnHigher.Checked = false;
                chkResetOnHigher.Enabled = false;
                cmbResetOnHigher.Enabled = false;
            }
            else if (m_ResetOnHigher == 0)
            {
                chkResetOnHigher.Checked = false;
                chkResetOnHigher.Enabled = true;
                cmbResetOnHigher.Enabled = false;
            }
            else
            {
                chkResetOnHigher.Checked = true;
                chkResetOnHigher.Enabled = true;
                cmbResetOnHigher.Enabled = true;

                if (m_ResetOnHigher > 0 && m_ResetOnHigher <= m_nLevel)
                {
                    cmbResetOnHigher.SelectedIndex = (int)(m_ResetOnHigher-1);
                }
            }
            return;
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            m_alignment = (WdListLevelAlignment)cmbHeadingSnAlign.SelectedIndex; // 保存用户选择的值
            m_chAfterHeadingSn = (WdTrailingCharacter)cmbHeadingSnBehindSn.SelectedIndex;// 保存用户选择的值

            m_NumberPosition = (float)numHeadingSnAlignPos.Value;// 保存用户选择的值
            m_TextPosition = (float)numHeadingSnTextIndentPos.Value;// 保存用户选择的值

            if (chkHeadingSnTabPos.Checked)
            {
                m_TabPosition = (float)numHeadingSnTabPos.Value;// 保存用户选择的值
            }
            else
            {
                m_TabPosition = 0.0f;// 保存用户选择的值
            }

            m_StartAt = (uint)numStartAt.Value;// 保存用户选择的值

            if (!chkResetOnHigher.Checked)
            {
                m_ResetOnHigher = 0;
            }
            else
            {
                m_ResetOnHigher = 0;
                if (cmbResetOnHigher.SelectedIndex != -1)
                {
                    m_ResetOnHigher = (uint)(cmbResetOnHigher.SelectedIndex + 1);
                }
            }

            return;
        }


        private void chkResetOnHigher_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkResetOnHigher.Checked)
            {
                m_ResetOnHigher = 0;
                cmbResetOnHigher.Enabled = false;
            }
            else
            {
                cmbResetOnHigher.Enabled = true;
            }

            return;
        }


    }
}
