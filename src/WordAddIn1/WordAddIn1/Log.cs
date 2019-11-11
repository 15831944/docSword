using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeAssist
{
    /// <summary>
    /// 应用程序运行日志
    /// </summary>
    public static class Log
    {
        private static readonly Object LOCK = new object();
        private static int saveDay = 7;

        /// <summary>
        /// 文件保存的天数，默认7天
        /// </summary>
        public static int SaveDay
        {
            get { return Log.saveDay; }
            set { Log.saveDay = value; }
        }

        /// <summary>
        /// 记录一条日志记录，会引发删除过期日志文件
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="path">环境路径</param>
        /// <param name="fileName">指定文件名</param>
        public static void WriteLog(string text, string fileName)
        {
            lock (LOCK)
            {
                string path = AppDomain.CurrentDomain.BaseDirectory + @"\log\";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                StreamWriter sw = null;
                DateTime dt = DateTime.Now;
                fileName = fileName + "[" + dt.Year + "_" + dt.Month + "_" + dt.Day + "].log";
                string textName = path + fileName;
                try
                {
                    sw = new StreamWriter(textName, true);
                }
                catch
                {
                    return;
                }

                // MessageBox.Show(text);

                sw.WriteLine("[" + DateTime.Now.ToString("yyyyMMdd_hhmmss:ffff") + "]" + text);
                sw.Flush();
                sw.Close();
                foreach (string file in Directory.GetFiles(path))
                {
                    FileInfo fi = new FileInfo(file);
                    DateTime creatTime = fi.CreationTime;
                    TimeSpan ts = dt - creatTime;
                    if (ts.TotalDays > SaveDay)
                    {
                        File.Delete(file);
                    }
                }
            }
        }

        /// <summary>
        /// 记录一条日志记录，默认文件名为Exception
        /// </summary>
        /// <param name="text">文本</param>
        public static void WriteLog(string text)
        {
            WriteLog(text, "运行日志");
        }
    }
}
