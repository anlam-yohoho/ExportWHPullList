using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportWHPullList
{
    internal class MSSQL
    {
        // Declaration:
        public static string DatabaseServer = @"vnmsrv601.dl.net\SIPLACE_2008R2EX";
        public string cnnDLMDB = string.Empty;
        public string cnnDLVNDB = string.Empty;
        public string cnnPackingFFC = string.Empty;
        public string cnnAssyRecord = string.Empty;
        //public string cnnDLMDB = "Provider=SQLOLEDB;Data Source=" + DatabaseServer + ";Initial Catalog= DLMDB;Persist Security Info=False;User ID=sa;Password=Siplace.1";
        //public string cnnDLVNDB = "Provider=SQLOLEDB;Data Source=" + DatabaseServer + ";Initial Catalog= DLVNDB;Persist Security Info=False;User ID=sa;Password=Siplace.1";
        //public string cnnPackingFFC = "Provider=SQLOLEDB;Data Source=" + DatabaseServer + ";Initial Catalog= FFCPacking;Persist Security Info=False;User ID=sa;Password=Siplace.1";
        //public string cnnAssyRecord = "Provider=SQLOLEDB;Data Source=" + DatabaseServer + ";Initial Catalog= FinalAssy;Persist Security Info=False;User ID=sa;Password=Siplace.1";
        public MSSQL()
        {
            cnnDLMDB = BuildConnectionString("DLMDB");
            cnnDLVNDB = BuildConnectionString("DLVNDB");
            cnnPackingFFC = BuildConnectionString("FFCPacking");
            cnnAssyRecord = BuildConnectionString("FinalAssy");
        }

        public string BuildConnectionString(string catalog)
        {
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = DatabaseServer,       // server name or "server,port"
                InitialCatalog = catalog,
                UserID = "sa",
                Password = "Siplace.1",
                PersistSecurityInfo = false,
                // Optional:
                // IntegratedSecurity = true, // if you want Windows authentication
                // MultipleActiveResultSets = true,
                // ConnectTimeout = 30
            };
            return builder.ConnectionString;
        }
    }
}
