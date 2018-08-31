using System;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace xlsConnector
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "Files")))
                Directory.CreateDirectory(Path.Combine(Directory.GetCurrentDirectory(), "Files"));

            var files = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Files"));
            Console.WriteLine("Введите название таблицы!");

            var sheet = Console.ReadLine();

            if (files.Length>1)
            {
                Console.WriteLine(files[0]);
                var ds = getDS(files[0], sheet);
                addRowBegin(ds, files[0],ds.Tables[0].TableName);

                for (int i = 1; i < files.Length; i++)
                {
                    Console.WriteLine(files[i]);
                    var tempDs = getDS(files[i], sheet);
                    addRowBegin(tempDs, files[i], ds.Tables[0].TableName);
                    var rows = tempDs.Tables[0].Rows;
                    foreach (DataRow item in rows)
                    {
                        ds.Tables[0].Rows.Add(item.ItemArray);
                    }
                }
                Console.WriteLine("Сохранение!");
                ExportToExcel(ds, Path.Combine(Directory.GetCurrentDirectory(), "Files", "result.xls"));
                Console.WriteLine("Завершено!");
            }
            else
            {
                Console.WriteLine("Добавьте файлы .xls в папку"+ Path.Combine(Directory.GetCurrentDirectory(), "Files"));
            }
            Console.ReadLine();

        }

        public static void addRowBegin(DataSet ds, string value, string sheet)
        {
            var nr = ds.Tables[0].NewRow();
            nr[0] = value+ sheet.Replace("'", "").Replace("$", "");
            ds.Tables[0].Rows.InsertAt(nr, 0);

        }

        public static DataSet getDS(string fileName, string sheet)
        {
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            DataSet ds = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (sheetName.Contains(sheet) && !sheetName.Contains("Print"))
                    {
                        Console.WriteLine(sheetName);
                        cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                        DataTable dt = new DataTable();
                        dt.TableName = sheetName;

                        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                        da.Fill(dt);

                        ds.Tables.Add(dt);
                    }
                }

                cmd = null;
                conn.Close();
            }
            return ds;
        }
        public static void ExportToExcel(DataSet dataSet, string filePath, bool overwiteFile = true)
        {
            if (File.Exists(filePath) && overwiteFile)
            {
                File.Delete(filePath);
            }

            foreach (DataTable dataTable in dataSet.Tables)
            {
                dataTable.ExportToExcel(filePath, false);
            }
        }
    }
}