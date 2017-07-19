using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

namespace Acai.NewsStatistics.Common
{
    public class Utils
    {
        public static DataTable ListToTable<T>(IList<T> list, string[] arrCols = null)
        {
            if (list == null || list.Count == 0)
                return null;

            DataTable dt = new DataTable("dt");
            var type = typeof(T);
            var propertys = type.GetProperties().ToList();
            if (arrCols.Any())
                propertys = propertys.Where(x => arrCols.Contains(x.Name)).ToList();

            propertys.ForEach(x => dt.Columns.Add(new DataColumn(x.Name)));
            foreach (var item in list)
            {
                DataRow row = dt.NewRow();
                propertys.ForEach(x => row[x.Name] = x.GetValue(item, null));
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}