using System;
using System.Xml.Serialization;

namespace SpreadImporter.Mapper
{
    [Serializable()]
    public class DbSource
    {
        [XmlAttribute(AttributeName = "src")]
        public string DbSrc { get; set; }

        [XmlAttribute(AttributeName = "name")]
        public string DbName { get; set; }

        [XmlAttribute(AttributeName = "user")]
        public string DbUser { get; set; }

        [XmlAttribute(AttributeName = "password")]
        public string DbPassword { get; set; }
    }
}
