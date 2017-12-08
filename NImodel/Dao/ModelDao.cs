using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace NImodel.Dao
{
    class ModelDao
    {

        public bool isTableExist(string filename, string table)
        {
            //数据源链接
            OleDbConnectionStringBuilder oleString = new OleDbConnectionStringBuilder();
            oleString.Provider = "Microsoft.ACE.OleDB.15.0";
            oleString.DataSource = filename;

            OleDbConnection conn = new OleDbConnection(oleString.ToString());
            conn.Open();

            //DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            //if (schemaTable != null)
            //{
            //    for (int i = 0; i < schemaTable.Rows.Count; i++)
            //    {
            //        string tableName = schemaTable.Rows[i]["TABLE_NAME"].ToString();
            //        if (table == tableName)
            //        {
            //            return true;
            //        }
            //    }
            //    return false;
            //}
            //else
            //{
            //    return false;
            //}

            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,new object[]{null, null, table,null});
            if (schemaTable.Rows.Count > 0)
            {
                conn.Close();
                return true;
            }
            else
            {
                conn.Close();
                return false;
            }
        }

        public bool createTable(string filename, string table)
        {
            OleDbConnectionStringBuilder olestring = new OleDbConnectionStringBuilder();
            olestring.Provider = "Microsoft.ACE.OleDB.15.0";
            olestring.DataSource = filename;

            OleDbConnection conn = new OleDbConnection(olestring.ToString());
            conn.Open();

            String sql = "create table " + table + "(id autoincrement primary key)";
            OleDbCommand command = new OleDbCommand(sql,conn);
            try
            {
                command.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                return false;
            }
        }
        /**
         * 返回值：1表示字段存在，2表示表不存在，0表示表中不存在该字段
         */
        public int isFieldExist(string filename, string table, string field)
        {
            OleDbConnectionStringBuilder oleString = new OleDbConnectionStringBuilder();
            oleString.Provider = "Microsoft.ACE.OleDB.15.0";
            oleString.DataSource = filename;

            OleDbConnection conn = new OleDbConnection(oleString.ToString());
            conn.Open();


            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, table, field });
            if (schemaTable.Rows.Count > 0)
            {
                conn.Close();
                return 1;
            }
            else if (!isTableExist(filename, table))
            {
                conn.Close();
                return 2;
            }
            else 
            {
                conn.Close();
                return 0;
            }
        }

        public bool createField(string filename, string table, string field, string dataType)
        {
            OleDbConnectionStringBuilder oleString = new OleDbConnectionStringBuilder();
            oleString.Provider = "Microsoft.ACE.OleDB.15.0";
            oleString.DataSource = filename;

            OleDbConnection conn = new OleDbConnection(oleString.ToString());
            conn.Open();

            OleDbCommand sql = new OleDbCommand();
            sql.Connection = conn;
            sql.CommandText = "alter table " + table + " add " + field + " " + dataType + ";";
            try
            {
                sql.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                return false;
            }
        }

        public void insert(string dataSource, string tablename, List<string> columnNames, List<List<string>> columnsValues)
        {
            foreach(List<string> strlist in columnsValues)
            {
                write(dataSource,tablename,columnNames,strlist);
            }
        }

        public void write(string dataSource, string table, List<string> columnNames, List<string> columnValues)
        {
            //数据源链接
            OleDbConnectionStringBuilder oleString = new OleDbConnectionStringBuilder();
            oleString.Provider = "Microsoft.ACE.OleDB.15.0";
            oleString.DataSource = dataSource;

            //数据库链接
            OleDbConnection conn = new OleDbConnection(oleString.ToString());
            conn.Open();

            OleDbCommand sql = new OleDbCommand();
            sql.Connection = conn;

            sql.CommandText = "insert into " + table + "(";
            foreach (String columnName in columnNames)
            {
                sql.CommandText += columnName + ",";
            }
            sql.CommandText = sql.CommandText.Substring(0, sql.CommandText.Length - 1);
            sql.CommandText += ") values(";

            foreach (String columnName in columnNames)
            {
                sql.CommandText += "@" + columnName + ",";
            }
            sql.CommandText = sql.CommandText.Substring(0, sql.CommandText.Length - 1);
            sql.CommandText += ")";

            for (int i = 0; i < columnNames.Count; i++)
            {
                sql.Parameters.AddWithValue("@" + columnNames[i], columnValues[i] == null ? "" : columnValues[i]);
            }
            Console.WriteLine(sql.CommandText);

            try
            {
                sql.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }

            conn.Close();
        }

        public int getIdByName(string dataSource, string tableName, string modelName)
        {
            //数据源链接
            OleDbConnectionStringBuilder oleString = new OleDbConnectionStringBuilder();
            oleString.Provider = "Microsoft.ACE.OleDB.15.0";
            oleString.DataSource = dataSource;

            //数据库链接
            OleDbConnection conn = new OleDbConnection(oleString.ToString());
            conn.Open();

            OleDbCommand sql = new OleDbCommand();
            sql.Connection = conn;
            sql.CommandText = "select id from " + tableName + " where Name = @name";
            sql.Parameters.AddWithValue("@name", modelName);
            Console.WriteLine(sql.CommandText);
            try
            {
                OleDbDataReader dr = sql.ExecuteReader();
                while (dr.HasRows)
                {
                    dr.NextResult();
                }
                //Console.WriteLine(id);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
            
            
            return 0;
        }
    }
}
