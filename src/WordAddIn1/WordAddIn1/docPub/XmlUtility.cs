using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Xml;

namespace OfficeAssist.docPub
{
    public class XmlUtility
   {
         /// <summary>
         /// 将自定义对象序列化为XML字符串
         /// </summary>
         /// <param name="myObject">自定义对象实体</param>
         /// <returns>序列化后的XML字符串</returns>
         public static String SerializeToXml<T>(T myObject)
         {
             if (myObject != null)
             {
                 XmlSerializer xs = new XmlSerializer(typeof(T));
 
                 MemoryStream stream = new MemoryStream();
                 XmlTextWriter writer = new XmlTextWriter(stream, Encoding.UTF8);
                 writer.Formatting = Formatting.Indented;//Formatting.None;//缩进
                 xs.Serialize(writer, myObject);
 
                 stream.Position = 0;
                 StringBuilder sb = new StringBuilder();
                 using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                 {
                     string line;
                     while ((line = reader.ReadLine()) != null)
                     {
                         sb.Append(line);
                     }
                     reader.Close();
                 }
                 writer.Close();
                 return sb.ToString();
             }
             return string.Empty;
         }
 
         /// <summary>
         /// 将XML字符串反序列化为对象
         /// </summary>
         /// <typeparam name="T">对象类型</typeparam>
         /// <param name="xml">XML字符</param>
         /// <returns></returns>
         public static T DeserializeToObject<T>(string xml)
         {
             T myObject;
             XmlSerializer serializer = new XmlSerializer(typeof(T));
             StringReader reader = new StringReader(xml);
             myObject = (T)serializer.Deserialize(reader);
             reader.Close();
             return myObject;
         }


     }// class

}// namespace
