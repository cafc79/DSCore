using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DeltaCore.Utilities;
using Npgsql;

namespace DeltaCore.DataAccess.DBConnect
{
    public class PgConexion : IDBConnect
    {
        #region "Elementos de la Clase"
        protected NpgsqlCommand cmd = null;
        protected NpgsqlConnection cnx = null;
        #endregion

        #region "Métodos Base"
        /// <summary>
        /// Constructor, inicializa la clase creando el objeto de conexion con la cadena de conexión, un command y los relaciona.
        /// Crea una nueva instancia del la clase “conexion” y se conecta por default a la base de datos “bdnsar”
        /// </summary>    
        public void DBConnect(string DB)
        {
            cnx = new NpgsqlConnection();
            cmd = new NpgsqlCommand { Connection = cnx };

            try
            {
                cnx.ConnectionString = DB;
            }
            catch (NpgsqlException sqlExcep)
            {
                throw new Exception(sqlExcep.Message, sqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        /// <summary>
        /// Este método, verifica si la conexion con la base de datos esta disponible
        /// </summary>
        public bool Monitor()
        {
            try
            {
                if (cnx.State == System.Data.ConnectionState.Closed)
                {
                    cnx.Open();
                }
                cmd.CommandType = CommandType.Text; //Solo existen select en PostgreSQL, un stp es un function que regresa o no parametros
                return true;
            }
            catch (NpgsqlException sqlExcep)
            {
                throw new Exception(sqlExcep.Message, sqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        public bool Closer()
        {
            try
            {
                if (cnx.State == System.Data.ConnectionState.Open)
                    cnx.Close();
                return true;
            }
            catch (NpgsqlException sqlExcep)
            {
                throw new Exception(sqlExcep.Message, sqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }
        #endregion

        #region Wizard
        private string LoadParameters(string select, Dictionary<string, object> sqlParams)
        {
            var swap = sqlParams.Select(param => param.Value).ToList();
            var sb = new StringBuilder();
            foreach (var sqlParam in sqlParams)
            {
                if (sqlParam.Value == null)
                    sb.Append("null,");
                else
                    sb.AppendFormat("'{0}',", sqlParam.Value);
            }
            return "(" + sb.ToString().Substring(0, sb.ToString().Length - 1) + ")";
        }

        private string LoadParameters(string select, Dictionary<string, string> sqlParams)
        {
            var swap = sqlParams.Select(param => param.Value).ToList();
            var sb = new StringBuilder();
            foreach (var sqlParam in sqlParams)
            {
                if (sqlParam.Value == null)
                    sb.Append("null,");
                else
                    sb.AppendFormat("'{0}',", sqlParam.Value);
            }
            return "(" + sb.ToString().Substring(0, sb.ToString().Length - 1) + ")";
        }

        private DataTable Reader2Table()
        {
            var dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dr.HasRows)
            {
                var dt = new DataTable();
                dt.Load(dr);
                return dt;
            }
            return null;
        }

        private List<object> Query2List(object objBase)
        {
            var dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            if (dr.HasRows)
            {
                var lista = new List<object>();
                var swap = SystemHelper.GetPropiedades(objBase);
                while (dr.Read())
                {
                    var registro = new List<object>();
                    for (int i = 0; i <= swap.Count - 1; i++)
                        registro.Add(dr.GetValue(dr.GetOrdinal(swap[i])));
                    lista.Add(registro);
                }
                return lista;
            }
            return null;
        }



        private DataTable StpAdapter2Table()
        {
            var da = new NpgsqlDataAdapter();
            var dt = new DataTable();
            da.SelectCommand = cmd;
            da.Fill(dt);
            return dt;
        }

        private DataTable Adapter2Table()
        {
            var da = new NpgsqlDataAdapter(cmd);
            var dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        private void Refresh()
        {
            cmd.CommandText = string.Empty;
            cmd.Parameters.Clear();
        }
        #endregion

        #region DB Methods
        public void ContinuesQry(string stp, Dictionary<string, object> SqlParams)
        {
            Refresh();
            try
            {
                cmd.CommandText = "select " + LoadParameters(stp, SqlParams);
                cmd.ExecuteNonQuery();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        public void ContinuesQry(string stp, Dictionary<string, string> SqlParams)
        {
            Refresh();
            try
            {
                cmd.CommandText = "select " + LoadParameters(stp, SqlParams);
                cmd.ExecuteNonQuery();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        public void ContinuesQry(string stp)
        {
            Refresh();
            try
            {
                cmd.CommandText = stp;
                cmd.ExecuteNonQuery();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
        }

        public void ExecQry(string stp, Dictionary<string, object> SqlParams)
        {
            Refresh();
            try
            {
                cmd.CommandText = "select " + stp + LoadParameters(stp, SqlParams);
                Monitor();
                cmd.ExecuteNonQuery();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                Closer();
            }
        }

        public void ExecQry(string stp)
        {
            Refresh();
            try
            {
                cmd.CommandText = stp;
                Monitor();
                cmd.ExecuteNonQuery();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                Closer();
            }
        }

        public List<object> ExecQry(string Query, Dictionary<string, object> SqlParams, object objBase)
        {
            Refresh();
            try
            {
                cmd.CommandText = "select * from " + Query + LoadParameters(Query, SqlParams);
                Monitor();
                return Query2List(objBase);
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                if (cmd.Connection.State == System.Data.ConnectionState.Open)
                {
                    cmd.Connection.Close();
                }
            }
        }

        public void ExecStp(string stp, Dictionary<string, object> SqlParams, NpgsqlTransaction Trans)
        {
            Refresh();
            try
            {
                cmd.CommandText = "select " + stp + LoadParameters(stp, SqlParams);
                cmd.Transaction = Trans;
                Monitor();
                cmd.ExecuteNonQuery();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                Closer();
            }
        }

        public DataTable ExecStp(string stp, Dictionary<string, object> SqlParams)
        {
            Refresh();
            try
            {
                cmd.CommandText = "select * from " + stp + LoadParameters(stp, SqlParams);
                Monitor();
                return StpAdapter2Table();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                Closer();
            }
        }

        public DataTable GetSelect(string Program)
        {
            Refresh();
            try
            {
                cmd.CommandText = Program;
                Monitor();
                return Adapter2Table();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                Closer();
            }
        }

        public object ExecScalar(string sql)
        {
            Refresh();
            try
            {
                cmd.CommandText = sql;
                cmd.CommandType = CommandType.Text;
                Monitor();
                return cmd.ExecuteScalar();
            }
            catch (NpgsqlException SqlExcep)
            {
                throw new Exception(SqlExcep.Message, SqlExcep);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                Closer();
            }
        }

        public DataRow BusquedaReg(string SQL, int Index)
        {
            Refresh();
            try
            {
                cmd.CommandText = SQL;
                Monitor();
                return Reader2Table().Rows[Index];
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                if (cmd.Connection.State == ConnectionState.Open)
                {
                    cmd.Connection.Close();
                }
            }
        }
        #endregion
    }
}