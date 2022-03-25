using System.Data;
using System.Text;

namespace SorianaCCIncomesReportFunction.Utilities
{
    public static class Extensions
    {
        public static string ToCSV(this DataTable table, string delimiter)
        {
            var result = new StringBuilder();
            for (int i = 0; i < table.Columns.Count; i++)
            {
                result.Append(table.Columns[i].ColumnName);
                result.Append(i == table.Columns.Count - 1 ? "\n" : delimiter);
            }

            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    result.Append(row[i].ToString());
                    result.Append(i == table.Columns.Count - 1 ? "\n" : delimiter);
                }
            }

            return result.ToString();
        }
    }
}
