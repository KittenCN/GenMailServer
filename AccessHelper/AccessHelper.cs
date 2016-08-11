using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AccessHelper
{
    public class AccessHelper
    {
        #region[字段]
        private string accesspath = string.Empty;
        private string accesspwd = string.Empty;

        private string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=product.mdb";
        //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\1.mdb;User ID='admin';Password='rr';
        OleDbConnection connect = null;
        OleDbCommand command = null;
        #endregion
        #region[属性]
        /// <summary>
        /// Access数据库连接字符串
        /// </summary>
        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; }
        }

        /// <summary>
        /// Access数据库路径
        /// </summary>
        public string AccessPath
        {
            get { return accesspath; }
            set { accesspath = value; }
        }
        #endregion
        #region[构造函数]
        /// <summary>
        /// 构造函数，初始化
        /// </summary>
        /// <param name="FilePath">Access文件路径</param>
        public AccessHelper(string FilePath)
        {
            ToFullRootPath(FilePath);
            ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AccessPath;
            connect = new OleDbConnection(ConnectionString);
        }
        /// <summary>
        /// 构造函数，初始化
        /// </summary>
        /// <param name="FilePath">Access文件路径</param>
        /// <param name="PassWord">密码</param>
        public AccessHelper(string FilePath, string PassWord)
        {
            accesspwd = PassWord;
            ToFullRootPath(FilePath);
            ConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};User ID=Admin;Jet OLEDB:Database Password={1};", AccessPath, accesspwd);
            connect = new OleDbConnection(ConnectionString);
        }
        #endregion
        #region[私有函数]
        private void Open()
        {
            try
            {
                if (connect.State != System.Data.ConnectionState.Open)
                {
                    connect.Open();
                }

            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message));
            }
        }

        private void Close()
        {
            try
            {
                if (connect.State != System.Data.ConnectionState.Closed)
                {
                    connect.Close();
                }

            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message));
            }
        }

        private void ToFullRootPath(string path)
        {
            if (System.IO.Path.IsPathRooted(path))
            {
                AccessPath = path;
            }
            else
            {
                AccessPath = AppDomain.CurrentDomain.BaseDirectory + path;
            }
        }
        #endregion

        #region[数据库简单操作]
        /// <summary>
        /// 测试是否能够连通
        /// </summary>
        /// <returns>布尔值</returns>
        public bool ConnectTest()
        {
            try
            {
                connect.Open();
            }
            catch(Exception ex)
            {
                connect.Close();
                return false;
            }
            return true;
        }

        /// <summary>
        /// 执行无返回的Sql语句，如插入，删除，更新
        /// </summary>
        /// <param name="sqlstr">SQL语句</param>
        /// <returns>受影响的条数,出错则产生异常</returns>
        public int ExecuteNonQuery(string sqlstr)
        {
            try
            {
                Open();
                command = new OleDbCommand(sqlstr, connect);
                int num = command.ExecuteNonQuery();
                command.Parameters.Clear();
                Close();
                return num;
            }
            catch
            { throw; }
        }

        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="sqlstr">Sql</param>
        /// <returns>DataSet数据集,出错则产生异常</returns>
        public DataSet ReturnDataSet(string sqlstr)
        {
            DataSet ds = new DataSet();
            try
            {
                OleDbDataAdapter adapter = new OleDbDataAdapter(sqlstr, connect);//adapter可自动打开数据库连接，不用Open()
                adapter.Fill(ds, "Obj");
                adapter.Dispose();
            }
            catch (Exception)
            {
                throw;
            }
            return ds;

        }

        /// <summary>
        /// 执行查询语句，返回DataTable
        /// </summary>
        /// <param name="sqlstr">Sqk</param>
        /// <returns>DataTable数据表,出错则产生异常</returns>
        public DataTable ReturnDataTable(string sqlstr)
        {
            try
            {
                return ReturnDataSet(sqlstr).Tables[0];
            }
            catch { throw; }
        }

        /// <summary>
        /// 执行查询语句，返回DataReader
        /// </summary>
        /// <param name="sqlstr">Sql</param>
        /// <returns>DataReader,出错则产生异常</returns>
        public OleDbDataReader ReturnDataReader(string sqlstr)
        {
            try
            {
                Open();
                command = new OleDbCommand(sqlstr, connect);
                OleDbDataReader myReader = command.ExecuteReader();
                command.Parameters.Clear();
                Close();
                return myReader;
            }
            catch (System.Data.SqlClient.SqlException e)
            {
                throw new Exception(e.Message, e);
            }

        }

        /// <summary>
        /// 执行事务,出错则产生异常
        /// </summary>
        /// <param name="SQLStringList">事务集合</param>
        public void ExecuteSqlTran(ArrayList SQLStringList)
        {

            Open();
            command = new OleDbCommand();
            command.Connection = connect;
            OleDbTransaction tx = connect.BeginTransaction();
            command.Transaction = tx;
            try
            {
                for (int n = 0; n < SQLStringList.Count; n++)
                {
                    string strsql = SQLStringList[n].ToString();
                    if (strsql.Trim().Length > 1)
                    {
                        command.CommandText = strsql;
                        command.ExecuteNonQuery();
                    }
                }
                tx.Commit();
                Close();
            }
            catch (Exception)
            {
                tx.Rollback();
                Close();
                throw;
            }
        }

        /// <summary>
        /// Access获取数据库中的所有表
        /// </summary>
        /// <returns>表集合</returns>
        public string[] GetTables()
        {
            List<string> Ls = new List<string>();

            Open();
            DataTable dt = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "table" }); //利用oledbconnection的getoledbschematable来获得数据库的结构
            Close();
            foreach (DataRow dr in dt.Rows)
            {
                Ls.Add((string)dr["table_name"]);
            }

            return Ls.ToArray();
        }

        /// <summary>
        /// 获取指定表的所有列
        /// </summary>
        /// <param name="TableName">表名</param>
        /// <returns>列集合</returns>
        public string[] GetColumns(string TableName)
        {
            Open();
            DataTable dt = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, TableName, null });
            Close();
            List<string> Ls = new List<string>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Ls.Add(dt.Rows[i]["COLUMN_NAME"].ToString());//["COLUMN_NAME"]);
            }
            return Ls.ToArray();

        }
        #endregion

        #region[数据库整体操作]
        /// <summary>
        /// 创建Access2003版本的数据库，可设密码
        /// 
        /// 
        /// Exception
        /// 通常创建数据库失败时会抛出基本类型的Exception，请注意在使用该方法时提供接收异常的容器
        /// </summary>
        /// <param name="mdbFilePath">Access数据库地址</param>
        /// <param name="password">密码</param>
        /// <returns></returns>
        public static bool CreateMDB(string mdbFilePath, string password)
        {
            if (System.IO.File.Exists(mdbFilePath))
            {
                throw (new Exception("当前目录下已存在该数据库，为保证安全性，请先删除后在新建！"));
            }
            try
            {
                string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;";
                if (password == null || password.Trim() == "")
                {
                    connStr += "Data Source=" + mdbFilePath;
                }
                else
                {
                    connStr += "Jet OLEDB:Database Password=" + password + ";Data Source=" + mdbFilePath;
                }
                object objCatalog = Activator.CreateInstance(Type.GetTypeFromProgID("ADOX.Catalog"));
                object[] oParams = new object[] { connStr };
                objCatalog.GetType().InvokeMember("Create", BindingFlags.InvokeMethod, null, objCatalog, oParams);
                Marshal.ReleaseComObject(objCatalog);
                objCatalog = null;
                return true;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// 创建Access2003版本的数据库
        /// 
        /// 
        /// Exception
        /// 通常创建数据库失败时会抛出基本类型的Exception，请注意在使用该方法时提供接收异常的容器
        /// </summary>
        /// <param name="mdbFilePath">Access数据库地址</param>
        /// <returns></returns>
        public static bool CreateMDB(string mdbFilePath)
        {
            try
            {
                CreateMDB(mdbFilePath, null);
                return true;
            }
            catch
            {
                throw;
            }
        }

        /// <summary>
        /// 设置或修改Access数据库的访问密码
        /// 
        /// Exception
        /// 设置失败时会弹出异常，请注意提供容器
        /// </summary>
        /// <param name="mdbFilePath">数据库文件路径</param>
        /// <param name="oldPwd">旧密码</param>
        /// <param name="newPwd">新密码</param>
        /// <returns></returns>
        public static bool SetMDBPassword(string mdbFilePath, string oldPwd, string newPwd)
        {
            string connStr = string.Concat("Provider=Microsoft.Jet.OLEDB.4.0;", "Mode=Share Deny Read|Share Deny Write;", /*独占模式*/"Jet OLEDB:Database Password=" + oldPwd + ";Data Source=" + mdbFilePath);
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                try
                {
                    conn.Open();
                    //如果密码为空时，请不要写方括号，只写一个null即可
                    string sqlOldPwd = (oldPwd == null || oldPwd.Trim() == "") ? "null" : "[" + oldPwd + "]";
                    string sqlNewPwd = (newPwd == null || newPwd.Trim() == "") ? "null" : "[" + newPwd + "]";
                    OleDbCommand cmd = new OleDbCommand(string.Concat("ALTER DATABASE PASSWORD ", sqlNewPwd, " ", sqlOldPwd), conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    return true;
                }
                catch
                {
                    throw;
                }
            }
        }

        /// <summary>
        /// 设置或修改Access数据库的访问密码
        /// 
        /// Exception
        /// 设置失败时会弹出异常，请注意提供容器
        /// </summary>
        /// <param name="newPwd">新密码</param>
        /// <returns></returns>
        public bool SetMDBPassword(string newPwd)
        {
            string connStr = string.Concat("Provider=Microsoft.Jet.OLEDB.4.0;", "Mode=Share Deny Read|Share Deny Write;", /*独占模式*/"Jet OLEDB:Database Password=" + accesspwd + ";Data Source=" + AccessPath);
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                try
                {
                    conn.Open();
                    //如果密码为空时，请不要写方括号，只写一个null即可
                    string sqlOldPwd = (accesspwd == null || accesspwd.Trim() == "") ? "null" : "[" + accesspwd + "]";
                    string sqlNewPwd = (newPwd == null || newPwd.Trim() == "") ? "null" : "[" + newPwd + "]";
                    OleDbCommand cmd = new OleDbCommand(string.Concat("ALTER DATABASE PASSWORD ", sqlNewPwd, " ", sqlOldPwd), conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    return true;
                }
                catch
                {
                    throw;
                }
            }
        }
        #endregion

    }
}
