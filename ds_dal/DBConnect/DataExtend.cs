using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;

namespace DeltaCore.DataAccess.DBConnect
{
    public static class DataExtend
    {
        public static IEnumerable<T> Distintos<T>(this DataTable tabla, string nombre)
        {
            return (from row in tabla.AsEnumerable() select row.Field<T>(nombre)).Distinct();
        }

        public static Dictionary<TKey, TRow> TableToDictionary<TKey, TRow>(
                this DataTable table,
                Func<DataRow, TKey> getKey,
                Func<DataRow, TRow> getRow)
        {
            return table
                .Rows
                .OfType<DataRow>()
                .ToDictionary(getKey, getRow);
        }
    }
}
