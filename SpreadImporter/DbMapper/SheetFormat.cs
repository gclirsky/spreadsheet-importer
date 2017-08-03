using System;
using System.Xml.Serialization;

namespace SpreadImporter.Mapper
{
    [Serializable()]
    public class SheetFormat
    {
        [XmlAttribute(AttributeName = "header_row", DataType = "int")]
        public int HeaderRow { get; set; }

        [XmlAttribute(AttributeName = "body_start_row", DataType = "int")]
        public int BodyStartRow { get; set; }
    }
}
