﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections.Specialized;
using System.Collections;


namespace OfficeAssist
{
    public class ConfigReader
    {
        // private NameValueCollection m_nameValues= new NameValueCollection();
        
        public Hashtable getConfigItems(String strCfgFile)
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreComments = true;

            // String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strCfgDir = strCfgFile;

            XmlReader reader = XmlReader.Create(strCfgDir, settings);
            xmlDoc.Load(reader);

            Hashtable nameValues = new Hashtable();

            nameValues.Clear();

            XmlNode xn = xmlDoc.SelectSingleNode("configItems");

            XmlNodeList xnl = xn.ChildNodes;
            String strName = "", strValue = "";

            foreach (XmlNode xn1 in xnl)
            {
                XmlElement xe = (XmlElement)xn1;
                strName = xe.GetAttribute("name").ToString();
                strValue = xe.GetAttribute("value").ToString();

                nameValues.Add(strName, strValue);
            }

            reader.Close();

            return nameValues;
        }

        public Hashtable getConfigItems()
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreComments = true;

            String strBaseDir = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            String strCfgDir = strBaseDir + @"config\config.xml";
            XmlReader reader = XmlReader.Create(strCfgDir, settings);
            xmlDoc.Load(reader);

            Hashtable nameValues = new Hashtable();

            nameValues.Clear();

            XmlNode xn = xmlDoc.SelectSingleNode("configItems");
 
            XmlNodeList xnl = xn.ChildNodes;
            String strName = "", strValue = "";

            foreach (XmlNode xn1 in xnl)
            {
                XmlElement xe = (XmlElement)xn1;
                strName = xe.GetAttribute("name").ToString();
                strValue = xe.GetAttribute("value").ToString();

                nameValues.Add(strName, strValue);
            }

            reader.Close();

            return nameValues;
        }

    }
}