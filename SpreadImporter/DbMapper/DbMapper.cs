using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace SpreadImporter.Mapper
{
    internal enum DbFieldType
    {
        INT,
        LONG,
        VARCHAR,
        DATETIME,
        UNKNOWN
    }

    public class Lookup
    {
        public int Id { get; set; }
        public string Type { get; set; }
        public int Counter { get; set; }
        public string ComponentValue { get; set; }
        public string ComponentDescription { get; set; }
    }

    public class DbMapper
    {
        public List<Lookup> Lookups { get; private set; }
        private readonly string xmlPath;

        private bool isDbConnected = false;

        public SqlConnection DbConnection { get; private set; }
        //public SqlCommand InsertCommand { get; private set; }
        //public SqlCommand DeleteCommand { get; private set; }
        //public SqlCommand UpdateCommand { get; private set; }

        public DbInfo DbInfo { get; private set; }

        public DbMapper()
        {
            this.xmlPath = "DbMapper.xml";
        }

        public DbMapper(string xmlPath)
        {
            this.xmlPath = xmlPath;
        }

        public void init()
        {
            try
            {
                if (!File.Exists(xmlPath))
                {
                    throw new FileNotFoundException("DbMapper xml configuration not found");
                }

                XmlSerializer serializer = new XmlSerializer(typeof(DbInfo));
                using (StreamReader reader = new StreamReader(xmlPath))
                {
                    DbInfo = (DbInfo)serializer.Deserialize(reader);
                }

                Lookups = new List<Lookup>();
            }
            catch (XmlException ex)
            {
                throw new InvalidDataException("Invalid DbMapper xml", ex);
            }
        }

        public void loadLookupsFromDb()
        {
            using (var connection = openDbConnection())
            {
                queryLookup(connection);
            }
        }

        public SqlConnection openDbConnection()
        {
            if (!isDbConnected)
            {
                var connection = string.Format(@"{0}Initial Catalog={1};User ID={2};Password={3}",
                                            DbInfo.DbSource.DbSrc,
                                            DbInfo.DbSource.DbName,
                                            DbInfo.DbSource.DbUser,
                                            DbInfo.DbSource.DbPassword);

                DbConnection = new SqlConnection(connection);
                isDbConnected = true;
            }

            if (DbConnection.State == ConnectionState.Closed)
            {
                DbConnection.Open();
            }

            return DbConnection;
        }

        public void closeDbConnection(bool dispose = false)
        {
            if (DbConnection.State != System.Data.ConnectionState.Closed)
            {
                DbConnection.Close();
                isDbConnected = false;
            }

            if (dispose)
            {
                DbConnection.Dispose();
                DbConnection = null;
            }
        }

        private void queryLookup(SqlConnection connection)
        {
            string sql = string.Format(@"SELECT   lu.Id, 
                                                    lu.Type, 
                                                    lu.Counter, 
                                                    lu.ComponentValue,
                                                    ld.ComponentDescription
                                              FROM  dbo.Lookup lu
                                         LEFT JOIN  dbo.LookupDescription ld
                                                ON  lu.Type=ld.Type AND lu.Counter=ld.Counter
                                             WHERE  ld.LanguageId = @LangId");

            using (SqlCommand queryCommand = new SqlCommand(sql, connection))
            {
                queryCommand.Parameters.Add(new SqlParameter("LangId", 2052)); // zh-cn

                SqlDataReader reader = queryCommand.ExecuteReader();
                while (reader.Read())
                {
                    Lookups.Add(new Lookup
                    {
                        Id = reader.GetInt32(0),
                        Type = reader.GetString(1),
                        Counter = reader.GetInt32(2),
                        ComponentValue = reader.GetString(3),
                        ComponentDescription = reader.GetString(4)
                    });
                }
            }
        }

        public int insert(SqlConnection connection, DataTable dataTable)
        {
            var transaction = connection.BeginTransaction();

            try
            {
                int execResult = 0;
                using (SqlCommand insertCommand = new SqlCommand())
                {
                    insertCommand.Connection = connection;
                    insertCommand.Transaction = transaction;

                    foreach (DataRow row in dataTable.Rows)
                    {
                        combineParameter(this.DbInfo.getDbTable(dataTable.TableName), row, insertCommand);

                        var insertSqlBuilder = new StringBuilder();
                        insertSqlBuilder.AppendFormat(@"INSERT INTO {0} VALUES (", dataTable.TableName);

                        foreach (SqlParameter param in insertCommand.Parameters)
                        {
                            insertSqlBuilder.AppendFormat(@"@{0},", param.ParameterName);
                        }

                        insertSqlBuilder.Replace(",", ")", insertSqlBuilder.Length - 1, 1);
                        var insertSqlText = insertSqlBuilder.ToString();

                        insertCommand.CommandText = insertSqlText;
                        execResult += insertCommand.ExecuteNonQuery();
                    }
                }

                transaction.Commit();

                return execResult;
            }
            catch (SqlException ex)
            {
                transaction.Rollback();
                throw new System.Exception("Data insertion to database failed", ex);
            }
        }

        private void combineParameter(DbTable dbTable, DataRow dtRow, SqlCommand sqlCommand)
        {
            foreach (var field in dbTable.DbFields)
            {
                sqlCommand.Parameters.Add(field.createParameter(dtRow));
            }
        }
    }
}
