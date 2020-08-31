using System;
using Db4objects.Db4o;
using Db4objects.Db4o.Config;

namespace DeltaCore.DBConnect
{
	public class EmbeddedServer
	{
		private string ObjectDB;
		//private static IObjectContainer cnx;

		#region Server
		public EmbeddedServer(string objectDB)
		{
			ObjectDB = objectDB;
		}

		public IObjectContainer OpenDB(IObjectContainer cnx)
		{
			try
			{
				return cnx ?? (cnx = Db4oEmbedded.OpenFile(ObjectDB));
			}
			catch (Exception ex)
			{
				throw new Exception(ex.Message + "; Error en db4o - No se puede establecer instancia", ex);
			}
		}

		public IObjectContainer OpenDB()
		{
			IEmbeddedConfiguration configuration = Db4oEmbedded.NewConfiguration();
			return Db4oEmbedded.OpenFile(configuration, ObjectDB);
		}

		public void Connect(ref IObjectContainer cnx)
		{
			if (cnx.Ext().IsClosed()) 
				cnx.Ext().OpenSession();
		}

		public void Disconnect(ref IObjectContainer cnx)
		{
			if (cnx == null) return;
			cnx.Close();
			cnx.Dispose();
			cnx = null;
		}

		public void Confirm(ref IObjectContainer cnx, bool commit)
		{
			if (commit)
				cnx.Commit();
			else
				cnx.Rollback();
		}
		#endregion
	}
}