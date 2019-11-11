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
    public partial class frmRegister : Form
    {
        public frmRegister()
        {
            InitializeComponent();
        }

        private void btnRegisterStart_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(txtRegisterAccount.Text))
            {
                MessageBox.Show("请填写账号");
                return;
            }

            if (String.IsNullOrWhiteSpace(txtActivateSn.Text))
            {
                MessageBox.Show("请输入激活码");
                return;
            }


            return;
        }
    }
}
