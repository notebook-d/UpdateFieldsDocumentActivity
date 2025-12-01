using System;
using System.IO;
using System.Xml.Serialization;

namespace UpdateFieldsDocumentActivity.Services
{
    public static class SerializeService
    {
        public static string SerializeToString<T>(T value)
        {
            XmlSerializer formatter = new XmlSerializer(typeof(T));
            using (StringWriter textWriter = new StringWriter())
            {
                try
                {
                    formatter.Serialize(textWriter, value);
                    return textWriter.ToString();
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
        }

        public static T DeserializeFromString<T>(string value)
        {
            if (value == null) return default(T);
            XmlSerializer formatter = new XmlSerializer(typeof(T));
            using (TextReader textReader = new StringReader(value))
            {
                try
                {
                    return (T)formatter.Deserialize(textReader);
                }
                catch (Exception ex)
                {
                    return default(T);
                }
            }
        }

        //public static void SerializeToFile<T>(T serializableObject, string fileName)
        //{
        //    var dir = Path.GetDirectoryName(fileName);
        //    if (!Directory.Exists(dir))
        //        Directory.CreateDirectory(dir);

        //    if (serializableObject == null) { return; }

        //    try
        //    {
        //        XmlDocument xmlDocument = new XmlDocument();
        //        XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
        //        using (MemoryStream stream = new MemoryStream())
        //        {
        //            serializer.Serialize(stream, serializableObject);
        //            stream.Position = 0;
        //            xmlDocument.Load(stream);
        //            xmlDocument.Save(fileName);
        //            stream.Close();
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }
        //}

        //public static T DeSerializeFromFile<T>(string fileName)
        //{
        //    if (string.IsNullOrEmpty(fileName)) { return default(T); }

        //    T objectOut = default(T);

        //    try
        //    {
        //        XmlDocument xmlDocument = new XmlDocument();
        //        xmlDocument.Load(fileName);
        //        string xmlString = xmlDocument.OuterXml;

        //        using (StringReader read = new StringReader(xmlString))
        //        {
        //            Type outType = typeof(T);

        //            XmlSerializer serializer = new XmlSerializer(outType);
        //            using (XmlReader reader = new XmlTextReader(read))
        //            {
        //                objectOut = (T)serializer.Deserialize(reader);
        //                reader.Close();
        //            }

        //            read.Close();
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }

        //    return objectOut;
        //}
    }
}
