using System.Collections.Generic;
using System.Xml.Serialization;

namespace UpdateFieldsDocumentActivity.Settings
{
    [XmlRoot("AutoFillSettings")]
    public class AutoFillSettings
    {
        [XmlElement("Type")]
        public List<DocType> Types { get; set; }
    }

    public class DocType
    {
        [XmlAttribute("Name")]
        public string Name { get; set; }

        [XmlElement("Field")]
        public List<FieldMapping> Fields { get; set; }
    }

    public class FieldMapping
    {
        [XmlAttribute("PilotFileName")]
        public string PilotFileName { get; set; }

        [XmlAttribute("FileField")]
        public string FileField { get; set; }

        [XmlAttribute("NewNameTemplate")]
        public string NewNameTemplate { get; set; }
    }
}
