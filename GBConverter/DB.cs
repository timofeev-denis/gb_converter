using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess.Client;

namespace GBConverter {
    class DB {
        private static OracleConnection conn = null;
        private static string oradb = "";
        private static OracleTransaction transaction = null;

        private DB() {}
        public static OracleConnection GetConnection() {
            if (conn != null) {
                if (conn.State == System.Data.ConnectionState.Open) {
                    return conn;
                }
            }
            string DBName = "RA00C000";
            string DBUser = "voshod";
            string DBPass = "voshod";
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 3) {
                DBName = args[1];
                DBUser = args[2];
                DBPass = args[3];
            }
            oradb = "Data Source=" + DBName + ";User Id=" + DBUser + ";Password=" + DBPass + ";";
            try {
                conn = new OracleConnection(oradb);
                conn.Open();
                transaction = conn.BeginTransaction();
                return conn;
            } catch (Exception) {
                return null;
            }
        }
        public static void Commit() {
            if (conn != null) {
                transaction.Commit();
                transaction = conn.BeginTransaction();
            }
        }
        public static void CloseConnection() {
            Rollback();
            if (conn != null) {
                conn.Dispose();
            }
        }
        public static void Rollback() {
            if(transaction != null) {
                try {
                    transaction.Rollback();
                    transaction = conn.BeginTransaction();
                } catch (Exception) {
                    // Подавляем исключения.
                }
            }
        }
    }
}
