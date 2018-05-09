using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace ConsoleApplication7
{
    public class ExcelToClass
    {
        private string SheetName { get; set; }
        private string ConnectionString { get; set; }
        private DataTable dataTable { get; set; }
        /// <summary>
        /// Constructor to set  xlsx file path and the sheet name to fill your class model
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public ExcelToClass(string filePath,string sheetName)
        {
            ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath}; Extended Properties=Excel 12.0;";
            SheetName = sheetName;
        }

        public void ReadFromExcel()
        {
            var adapter = new OleDbDataAdapter($"SELECT * FROM [{SheetName}$]", ConnectionString);
            var ds = new DataSet();
            adapter.Fill(ds, SheetName);
            dataTable = ds.Tables[SheetName];
        }

        public List<T> LoadFromDataTable<T>() where T : class, new()
        {
            List<T> output = new List<T>();
            T entry = new T();
            var cols = entry.GetType().GetProperties();
            var headers = dataTable.Columns.Cast<DataColumn>().Select(s => s.ColumnName).ToArray();
            foreach (DataRow row in dataTable.Rows)
            {
                entry = new T();
                for (int i = 0; i < headers.Length; i++)
                {
                    foreach(var col in cols)
                    {
                        if (col.Name == headers[i])
                        {
                            col.SetValue(entry, Convert.ChangeType(row[i], col.PropertyType));
                        }
                    }
                }
                output.Add(entry);
            }
            
            return output;

        }

    }
}
