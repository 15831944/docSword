using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System;
using System.IO;


namespace OfficeAssist
{
    public partial class FormConfig : Form
    {
        public FormConfig()
        {
            InitializeComponent();
        }

        public String m_configDbUrl = "";
        public String m_configTempLoc = "";

        private void btnConfigSave_Click(object sender, EventArgs e)
        {
            String dbUrl   = txtConfigDbUrl.Text.Trim();
            String tempLoc = txtConfigTempLoc.Text.Trim();

            if (dbUrl.Equals("") || tempLoc.Equals(""))
            {
                MessageBox.Show("配置项不能为空");
                return;
            }

            m_configDbUrl = dbUrl;
            m_configTempLoc = tempLoc;

            // save into text
            using (FileStream fs = new FileStream(".\\config.txt", FileMode.Create))
            {
                //lock (fs)
                {
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine("dbUrl=" + m_configDbUrl);
                    sw.WriteLine("tempLoc=" + m_configTempLoc);
                    //sw.Dispose();
                    sw.Close();
                }
            }

        }

        private void FormConfig_Load(object sender, System.EventArgs e)
        {

            StreamReader sr = new StreamReader(".\\config.txt", Encoding.Default);
            String line;
            while ((line = sr.ReadLine()) != null)
            {
                String[] strArr = line.Split('=');

                if (strArr.GetLength(0) == 2)
                {
                    if (strArr[0].Equals("dbUrl"))
                        m_configDbUrl = strArr[1];

                    if (strArr[0].Equals("tempLoc"))
                        m_configTempLoc = strArr[1];
                }

            }

            sr.Close();

            txtConfigDbUrl.Text = m_configDbUrl;
            txtConfigTempLoc.Text = m_configTempLoc;


        }


    }
}
