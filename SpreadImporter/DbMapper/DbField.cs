using System;
using System.Data;
using System.Data.SqlClient;
using System.Xml.Serialization;

namespace SpreadImporter.Mapper
{
    [Serializable()]
    public class DbField
    {
        [XmlAttribute(AttributeName = "name")]
        public string FieldName { get; set; }

        [XmlAttribute(AttributeName = "type")]
        public string FieldType { get; set; }

        [XmlAttribute(AttributeName = "sheet_column_index", DataType = "int")]
        public int FieldMappingColumnIndex { get; set; }

        public object dataFormatting(object value)
        {
            DbFieldType dbFieldType = mapToDbFieldType(this.FieldType);
            switch (dbFieldType)
            {
                case DbFieldType.INT:
                case DbFieldType.LONG:
                case DbFieldType.VARCHAR:
                    return value;
                case DbFieldType.DATETIME:
                    return Convert.ToDateTime(value);
                default:
                    return value;
            }
        }

        public SqlParameter createParameter(DataRow row)
        {
            SqlParameter parameter = null;
            DbFieldType dbFieldType = mapToDbFieldType(this.FieldType);
            switch (dbFieldType)
            {
                case DbFieldType.INT:
                    parameter = new SqlParameter(this.FieldName, SqlDbType.Int);
                    parameter.Value = row[this.FieldName];
                    break;
                case DbFieldType.LONG:
                    parameter = new SqlParameter(this.FieldName, SqlDbType.BigInt);
                    parameter.Value = row[this.FieldName];
                    break;
                case DbFieldType.VARCHAR:
                    parameter = new SqlParameter(this.FieldName, SqlDbType.VarChar);
                    parameter.Value = row[this.FieldName];
                    break;
                case DbFieldType.DATETIME:
                    parameter = new SqlParameter(this.FieldName, SqlDbType.DateTime);
                    parameter.Value = row[this.FieldName];
                    break;
                default:
                    break;
            }

            return parameter;
        }

        private DbFieldType mapToDbFieldType(string fieldType)
        {
            DbFieldType dbFieldType;
            if (Enum.TryParse(fieldType, out dbFieldType))
            {
                return dbFieldType;
            }
            else
            {
                return DbFieldType.UNKNOWN;
            }
        }
    }
}
