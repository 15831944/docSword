using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows;
using System.Drawing;
using System.Reflection;
using System.Collections;


namespace OfficeAssist
{
    public class ColorComboBox : ComboBox
    {
        private Hashtable m_hashName2Color = null;
        private ArrayList m_arrColorNames = null;


        public void setColors(ArrayList arrColorNames, Hashtable hashName2Color)
        {
            m_arrColorNames = arrColorNames;
            m_hashName2Color = hashName2Color;

            if (m_arrColorNames != null && m_arrColorNames.Count > 0 && 
                m_arrColorNames.Count == m_hashName2Color.Count)
            {
                this.Items.Clear();

                foreach (String strColorName in m_arrColorNames)
                {
                    this.Items.Add(strColorName);
                }

                this.Text = (String)m_arrColorNames[0];
            }
        }


        public ColorComboBox()
        {
            this.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.DropDownStyle = ComboBoxStyle.DropDownList;
            // this.ItemHeight = 20;

            //PropertyInfo[] propInfoList = typeof(Color).GetProperties(BindingFlags.Static |
            //    BindingFlags.DeclaredOnly | BindingFlags.Public);

            if (m_arrColorNames == null)
            {
                this.Items.Clear();

                //foreach (PropertyInfo c in propInfoList)
                //{
                //    this.Items.Add(c.Name);
                //}

                // this.Text = "Black";
            }
            else
            {
                if (m_arrColorNames.Count > 0)
                {
                    this.Items.Clear();

                    foreach (String strColorName in m_arrColorNames)
                    {
                        this.Items.Add(strColorName);
                    }

                    this.Text = (String)m_arrColorNames[0];
                }

            }

            return;
        }


        protected override void  OnDrawItem(DrawItemEventArgs e)
        {
            Rectangle rect = e.Bounds;

            if (e.Index >= 0)
            {
                String colorName = this.Items[e.Index].ToString();
                Color c = Color.White;

                if (m_arrColorNames == null)
                {
                    c = Color.FromName(colorName);
                }
                else
                {
                    c = (Color)m_hashName2Color[colorName];
                }

                using (Brush b = new SolidBrush(c))
                {
                    // e.Graphics.FillRectangle(b, rect.X, rect.Y + 2, rect.Width, rect.Height - 4);
                    e.Graphics.FillRectangle(b, rect.X, rect.Y, rect.Width, rect.Height);

                    if (e.Index == 0)
                    {
                        using (Pen p = new Pen(Color.Black))
                        {
                            e.Graphics.DrawLine(p, new Point(rect.X, rect.Y + rect.Height / 2), new Point(rect.X + rect.Width, rect.Y + rect.Height / 2));
                            // e.Graphics.DrawLine(p, new Point(rect.X, rect.Y + rect.Height), new Point(rect.X + rect.Width, rect.Y));
                        }
                    }
                }
            }

            return;
        }
        


    }
}
