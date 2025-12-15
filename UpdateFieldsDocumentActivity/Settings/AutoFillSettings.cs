using System.Collections.Generic;
using System.Xml.Serialization;

namespace UpdateFieldsDocumentActivity.Settings
{
    [XmlRoot("AutoFillSettings")]
    public class AutoFillSettings
    {
        [XmlElement("Type")]
        public List<DocType> Types { get; set; } = new List<DocType>();
    }

    public class DocType
    {
        [XmlAttribute("Name")]
        public string Name { get; set; }

        // Для старого формата
        [XmlElement("Field")]
        public List<FieldMapping> Fields { get; set; } = new List<FieldMapping>();

        // Для нового формата
        [XmlElement("Fields")]
        public FieldsContainer FieldsContainer { get; set; }

        // Универсальное свойство для доступа к расширениям
        [XmlIgnore]
        public string Extensions => FieldsContainer?.Extensions;

        // Универсальное свойство для доступа ко всем полям
        [XmlIgnore]
        public List<FieldMapping> AllFields
        {
            get
            {
                var allFields = new List<FieldMapping>();

                // Добавляем поля из старого формата
                if (Fields != null)
                    allFields.AddRange(Fields);

                // Добавляем поля из нового формата
                if (FieldsContainer?.FieldMappings != null)
                    allFields.AddRange(FieldsContainer.FieldMappings);

                return allFields;
            }
        }
    }

    public class FieldsContainer
    {
        [XmlAttribute("Extensions")]
        public string Extensions { get; set; }

        [XmlElement("Field")]
        public List<FieldMapping> FieldMappings { get; set; } = new List<FieldMapping>();
    }

    public class FieldMapping
    {
        // Атрибуты из старого формата
        [XmlAttribute("PilotFileName")]
        public string PilotFileName { get; set; }

        // Атрибуты из нового формата
        [XmlAttribute("PilotAttr")]
        public string PilotAttr { get; set; }

        // Общие атрибуты
        [XmlAttribute("FileField")]
        public string FileField { get; set; }

        [XmlAttribute("NewNameTemplate")]
        public string NewNameTemplate { get; set; }

        // Универсальное свойство для получения Pilot атрибута
        [XmlIgnore]
        public string PilotAttribute =>
            !string.IsNullOrEmpty(PilotAttr) ? PilotAttr : PilotFileName;
    }

    //[XmlRoot("AutoFillSettings")]
    //public class AutoFillSettings
    //{
    //    [XmlElement("Type")]
    //    public List<DocType> Types { get; set; }
    //}

    //public class DocType
    //{
    //    [XmlAttribute("Name")]
    //    public string Name { get; set; }

    //    [XmlElement("Field")]
    //    public List<FieldMapping> Fields { get; set; }
    //}

    //public class FieldMapping
    //{
    //    [XmlAttribute("PilotFileName")]
    //    public string PilotFileName { get; set; }

    //    [XmlAttribute("FileField")]
    //    public string FileField { get; set; }

    //    [XmlAttribute("NewNameTemplate")]
    //    public string NewNameTemplate { get; set; }
    //}
}
