using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.SqlServer.Management;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Common;

namespace SqlServerBackupRestore
{
    public class BackupHelper
    {
        public BackupHelper()
        {

        }

        public void BackupDatabase(String databaseName, String userName, String password, String serverName, String destinationPath)
        {
            Backup sqlBackup = new Backup();
            
            sqlBackup.Action = BackupActionType.Database;
            sqlBackup.BackupSetDescription = "ArchiveDataBase:" + DateTime.Now.ToShortDateString();
            sqlBackup.BackupSetName = "Archive";

            sqlBackup.Database = databaseName;

            BackupDeviceItem deviceItem = new BackupDeviceItem(destinationPath, DeviceType.File);
            ServerConnection connection = new ServerConnection(serverName, userName, password);
            Server sqlServer = new Server(connection);
            
            Database db = sqlServer.Databases[databaseName];
            
            sqlBackup.Initialize = true;
            sqlBackup.Checksum = true;
            sqlBackup.ContinueAfterError = true;
            
            sqlBackup.Devices.Add(deviceItem);
            sqlBackup.Incremental = false;

            sqlBackup.ExpirationDate = DateTime.Now.AddDays(3);
            sqlBackup.LogTruncation = BackupTruncateLogType.Truncate;

            sqlBackup.FormatMedia = false;

            sqlBackup.SqlBackup(sqlServer);
        }
    }
}
