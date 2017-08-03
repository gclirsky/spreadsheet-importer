using System;
using System.Linq;
using System.Xml.Serialization;

namespace SpreadImporter.Mapper
{
    [Serializable()]
    public class DbInfo
    {
        [XmlElement("ImportSheetFormat")]
        public SheetFormat SheetFormat { get; set; }

        [XmlElement("DbSource")]
        public DbSource DbSource { get; set; }

        [XmlArray("DbTables")]
        [XmlArrayItem("DbTable", typeof(DbTable))]
        public DbTable[] DbTables { get; set; }

        public DbTable getDbTable(string tableName)
        {
            return DbTables.FirstOrDefault(t => t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
