using DeltaCore.Utilities.Research;
using System;
using System.Collections.Generic;
using System.Data;                     

namespace DeltaCore.DataAccess.DBConnect
{
    public  interface IDBConnect
    {
        #region "Métodos Base"

        [Title("Default Settings")]
        [Category("1. Getting Server")]
        [Description("We can use Interface que crea un cliente para interactuar con base de datos SQL. If we do not specify any arguement then it connects with mongodb instance running on localhost on default port [27017]")]
        [Code("Console", "mongo")]
        [Code("C#", "new DeltaCore.DataAccess.DBConnect\n\t.GetServer()\n\t.Connect()")]
        void DBConnect(string db);

        /// <summary>
        /// Este método, verifica si la conexion con la base de datos esta disponible
        /// </summary>
        bool Monitor();

        bool Closer();
        
        #endregion

        #region DB Methods

        void ContinuesQry(string stp, Dictionary<string, object> SqlParams);

        void ContinuesQry(string stp);

        List<object> ExecQry(string Query, Dictionary<string, object> SqlParams, object objBase);

        void ExecQry(string stp, Dictionary<string, object> SqlParams);

        void ExecQry(string stp);
        
        DataTable ExecStp(string stp, Dictionary<string, object> SqlParams);

        DataTable GetSelect(string program);

	    Object ExecScalar(string program);

        DataRow BusquedaReg(string SQL, int Index);


        #endregion
    }
}