using System;
using System.Linq;
using System.Xml.Serialization;

namespace SpreadImporter.Mapper
{
    [Serializable()]
    public class DbTable
    {
        [XmlAttribute(AttributeName = "name")]
        public string Name { get; set; }

        [XmlArray("DbFields")]
        [XmlArrayItem("DbField", typeof(DbField))]
        public DbField[] DbFields { get; set; }

        public DbField getDbField(int sheeColumnIndex)
        {
            return DbFields.FirstOrDefault(f => f.FieldMappingColumnIndex.Equals(sheeColumnIndex));
        }
    }


}
