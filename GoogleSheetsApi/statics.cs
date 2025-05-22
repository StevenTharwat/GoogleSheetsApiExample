using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsApi
{
    static public class statics
    {

        public static List<List<object>> DataWithoutFlags = new List<List<object>>();

        static DataTable table = new DataTable();

        public static DataTable convertToTable(IList<IList<object>> ApiData)
        {
            table.Clear();
            table.Columns.Clear();

            
            for (int i = 0; i < ApiData[0].Count; i++) // to set headers
            {
                table.Columns.Add(ApiData[0][i].ToString());

            }
            
            for (int i = 1; i < ApiData.Count; i++) // show every row
            {
                DataRow row = table.NewRow();
                for (int j = 0; j < ApiData[0].Count; j++) // show every cell in a row
                {
                    var o = ApiData[i][j];
                    row[ApiData[0][j].ToString()] = ApiData[i][j];
                }
                table.Rows.Add(row);
            }
            
            return table;
        }

        public static List<List<object>> ConvertFromIListToList(IList<IList<object>> From)
        {
            List<List<object>> TO = new List<List<object>>();
            for (int i = 0; i < From.Count; i++)
            {
                TO.Add(From[i] as List<object>);

            }
            return TO;
        }

        public static string Equalstring(this string str)
        {
            return str.ToLower().Trim();
        }
    }
}
