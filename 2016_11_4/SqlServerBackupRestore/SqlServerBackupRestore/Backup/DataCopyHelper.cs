using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Common;

namespace SqlServerBackupRestore
{
    public class DataCopyHelper
    {
        Server sqlServer;

        public DataCopyHelper(String serverName, String userName, String password)
        {
            sqlServer = new Server(new ServerConnection(serverName, userName, password));
        }

        public void CopyData(String sourceDatabase, String destinationDataBase)
        {
            Database dbSource = sqlServer.Databases[sourceDatabase];
            Database dbDestination = sqlServer.Databases[destinationDataBase];

            if (dbDestination == null || dbSource == null)
                throw new Exception("Specified Database not found the server " + sqlServer.Name);
            StringBuilder sqlScript = new StringBuilder("");
            sqlScript.AppendLine("USE " + destinationDataBase + ";");
            sqlScript.AppendLine("");

            foreach (Table dataTable in dbSource.Tables)
            {
                if (!dbDestination.Tables.Contains(dataTable.Name, dataTable.Schema))
                    continue;
                sqlScript.AppendFormat("INSERT INTO {0} \n SELECT * FROM {0}", dataTable.Name);
                sqlScript.AppendLine();
            }

            dbDestination.ExecuteNonQuery(sqlScript.ToString());
        }

        public void TrancateData(String[] tableNames, String databaseName)
        {
            StringBuilder sqlScript = new StringBuilder("");
            sqlScript.AppendFormat("USE {0};", databaseName);
            sqlScript.AppendLine();
            foreach (String tableName in tableNames)
            {
                sqlScript.AppendFormat("TRUNCATE TABLE {0}", tableName);
                sqlScript.AppendLine();
            }

            Database db = sqlServer.Databases[databaseName];
            db.ExecuteNonQuery(sqlScript.ToString());
        }

        public void TruncateDatabase(String databaseName)
        {
            Database db = sqlServer.Databases[databaseName];
            if(db==null)
                throw new Exception("Specified Database not found the server " + sqlServer.Name);
            List<String> tables = new List<string>();
            foreach (Table dataTable in db.Tables)
            {
                tables.Add(dataTable.Name);
            }

            this.TrancateData(tables.ToArray(), databaseName);
        }
    }
}
