using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeTools.Common;
using Word = Microsoft.Office.Interop.Word;
using System.Collections;
using System.Collections.Specialized;

namespace OfficeAssist
{
    public partial class frmFontSetting : Form
    {
        Word.Application app = null;
        Word.Document doc = null;
        ClassFont inFont = new ClassFont();
        String m_strNoChange = "(保持原值不变)";

        public void setFeatures(Word.Application oApp, ClassFont oFont = null, String strNoChange = "(保持原值不变)")
        {
            app = oApp;
            try
            {
                doc = app.ActiveDocument;
            }
            catch (System.Exception ex)
            {
                doc = null;
            }
            finally
            {

            }

            m_strNoChange = strNoChange;

            if (oFont != null)
            {
                oFont.SelCopy2(inFont);
                // inFont.clone(oFont);
            }

            return;
        }


        public frmFontSetting()
        {
            InitializeComponent();
        }

        public ClassFont getFont()
        {
            return inFont;
        }


        private void frmFontSetting_Load(object sender, EventArgs e)
        {
            ArrayList arrFontNames = new ArrayList();
            //String strNoChange = m_strNoChange;
            //
            arrFontNames.Add(m_strNoChange);

            if (app != null)
            {
                foreach (String strFntNameItem in app.FontNames)
                {
                    arrFontNames.Add(strFntNameItem);
                }
            }

            foreach (String strFntNameItem in arrFontNames)
            {
                cmbChineseFonts.Items.Add(strFntNameItem);
                cmbAsciiFonts.Items.Add(strFntNameItem);
            }

            ArrayList arrWordFontSize = new ArrayList();

            arrWordFontSize.Add(m_strNoChange);

            ArrayList arrFontSizes = Globals.ThisAddIn.getFontSizes();

            foreach (String strSize in arrFontSizes)
            {
                arrWordFontSize.Add(strSize);
            }

            foreach (String strFntSize in arrWordFontSize)
            {
                cmbFontSize.Items.Add(strFntSize);
            }

            Data2UI();

            return;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            return;
        }


        private void btnOK_Click(object sender, EventArgs e)
        {
            Ui2Data();

            return;
        }


        private void Data2UI(ClassFont oFont = null)
        {
            if (oFont != null)
            {
                oFont.SelCopy2(inFont);
            }

            if (inFont.isSet(ClassFont.euMembers.NameFarEast))
            {
                cmbChineseFonts.Text = inFont.NameFarEast;
            }
            else
            {
                cmbChineseFonts.Text = m_strNoChange;
            }

            if (inFont.isSet(ClassFont.euMembers.NameAscii))
            {
                cmbAsciiFonts.Text = inFont.NameAscii;
            }
            else
            {
                cmbAsciiFonts.Text = m_strNoChange;
            }

            if (inFont.isSet(ClassFont.euMembers.Size))
            {
                cmbFontSize.Text = "" + inFont.Size;
            }
            else
            {
                cmbFontSize.Text = m_strNoChange;
            }

            // 
            if (inFont.isSet(ClassFont.euMembers.Bold))
            {
                if (inFont.Bold != 0) // true
                {
                    chkFontBold.CheckState = CheckState.Checked;
                }
                else
                {
                    chkFontBold.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkFontBold.CheckState = CheckState.Indeterminate;
            }

            // 
            if (inFont.isSet(ClassFont.euMembers.Italic))
            {
                if (inFont.Italic != 0) // true
                {
                    chkFontItalic.CheckState = CheckState.Checked;
                }
                else
                {
                    chkFontItalic.CheckState = CheckState.Unchecked;
                }
            }
            else
            {
                chkFontItalic.CheckState = CheckState.Indeterminate;
            }

            return;
        }


        private void Ui2Data()
        {
            String strItem = "";

            strItem = cmbChineseFonts.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                inFont.RemoveSelMember((int)ClassFont.euMembers.NameFarEast);
            }
            else
            {
                inFont.AddSelMember((int)ClassFont.euMembers.NameFarEast);
                inFont.NameFarEast = strItem;
            }


            strItem = cmbAsciiFonts.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                inFont.RemoveSelMember((int)ClassFont.euMembers.NameAscii);
            }
            else
            {
                inFont.AddSelMember((int)ClassFont.euMembers.NameAscii);
                inFont.NameAscii = strItem;
            }

            strItem = cmbFontSize.Text;
            if (String.IsNullOrWhiteSpace(strItem) || strItem.Equals(m_strNoChange))
            {
                inFont.RemoveSelMember((int)ClassFont.euMembers.Size);
            }
            else
            {
                float fSize = Globals.ThisAddIn.m_commTools.str2float(strItem);

                if (!float.IsNaN(fSize) && fSize > 0.0f)
                {
                    inFont.AddSelMember((int)ClassFont.euMembers.Size);
                    inFont.Size = fSize;
                }
                else
                {
                    inFont.RemoveSelMember((int)ClassFont.euMembers.Size);
                }
            }

            // 
            switch (chkFontBold.CheckState)
            {
                case CheckState.Checked:
                    inFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    inFont.Bold = -1;
                    break;

                case CheckState.Unchecked:
                    inFont.AddSelMember((int)ClassFont.euMembers.Bold);
                    inFont.Bold = 0;
                    break;

                case CheckState.Indeterminate:
                default:
                    inFont.RemoveSelMember((int)ClassFont.euMembers.Bold);
                    break;
            }


            // 
            switch (chkFontItalic.CheckState)
            {
                case CheckState.Checked:
                    inFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    inFont.Italic = -1;
                    break;

                case CheckState.Unchecked:
                    inFont.AddSelMember((int)ClassFont.euMembers.Italic);
                    inFont.Italic = 0;
                    break;

                case CheckState.Indeterminate:
                default:
                    inFont.RemoveSelMember((int)ClassFont.euMembers.Italic);
                    break;
            }

            return;
        }




    }
}
