using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using ReaderSituacionFiscal.Dto;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ReaderSituacionFiscal
{
    internal class Program
    {
        private static string pathDir = @"c:\ConstanciasFiscalRev";
        

        [STAThreadAttribute]
        static void Main(string[] args)
        {

           

            if (Directory.Exists(pathDir))
            {
                Console.WriteLine("That path exists already.");
            }
            else {
                DirectoryInfo di = Directory.CreateDirectory(pathDir);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(pathDir));
            }

            string sqlcon = "Data Source = 192.168.11.31; Initial Catalog = Conalep; User ID = conalep; Password = Carlos1085#";
            clearTable(sqlcon);
         
            List<DtoCFiscal> dataCFiscal = new List<DtoCFiscal>();
            List<string> lineas = new List<string>();
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                var path = fbd.SelectedPath;
                var fileName = path + @"\cfdiInfo.txt";

                if (File.Exists(fileName))
                {
                    File.Delete(fileName);
                }

                // Create a new file     
                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    if (Directory.Exists(path))
                    {
                        // This path is a directory
                        ProcessDirectory(path, writer, dataCFiscal);
                    }
                    else
                    {
                        writer.WriteLine($"{0} is not directory.", path);
                    }

                }



            }



            using (var copy = new SqlBulkCopy(sqlcon))
            {
                copy.DestinationTableName = "dbo.DtoCFiscal";
                // Add mappings so that the column order doesn't matter
               
                copy.ColumnMappings.Add(nameof(DtoCFiscal.DtoCFiscalId), "DtoCFiscalId");
                copy.ColumnMappings.Add(nameof(DtoCFiscal.Rfc), "Rfc");
                copy.ColumnMappings.Add(nameof(DtoCFiscal.Cp), "Cp");

                copy.WriteToServer(ToDataTable(dataCFiscal));
            }
            Console.WriteLine("Finalizado");
            Console.ReadLine();
        }

        public static void ProcessDirectory(string targetDirectory, StreamWriter writer, List<DtoCFiscal> dataCFiscal)
        {

            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory, "*.pdf", SearchOption.TopDirectoryOnly);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName, writer, dataCFiscal);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory, writer, dataCFiscal);
        }

        private static void ProcessFile(string fileName, StreamWriter writer, List<DtoCFiscal> dataCFiscal)
        {
            var filterText =  @"([A - ZÑ &]{ 3,4}) ?(?: - ?) ? (\d{ 2} (?: 0[1 - 9] | 1[0 - 2])(?:0[1 - 9] |[12]\d | 3[01])) ?(?: - ?) ? ([A - Z\d]{ 2})([A\d])"; 
            Regex searchTerm = new Regex(filterText);

            using (var pdfDocument = new PdfDocument( new PdfReader(fileName)))
            {
                var strategy = new LocationTextExtractionStrategy();
                StringBuilder processed = new StringBuilder();

                for (int i = 1; i <= pdfDocument.GetNumberOfPages(); ++i)
                {
                    var page = pdfDocument.GetPage(i);
                    string text = PdfTextExtractor.GetTextFromPage(page, strategy);
                    if (text != null)
                    {
                       var separador = Convert.ToChar("\n");
                       var lineas =  text.ToString().Split(separador);

                        Match m = searchTerm.Match(text);

                        var cp = from x in lineas
                                 where x.Contains("Código Postal")
                                 select x;


                        DtoCFiscal cFiscal = new DtoCFiscal()
                        {
                            DtoCFiscalId = Guid.NewGuid().ToString(),
                            Rfc = "",
                            Cp =  cp.First().ToString().Substring(14,5),

                        };
                        
                    }
                    else
                    {
                        
                    }
                    processed.Append(text);
                }
                pdfDocument.Close();
            }
        }

        private static void clearTable(string sqlcon)
        {
            using (var sql = new SqlConnection(sqlcon))
            {

                sql.Open();
                SqlCommand command = sql.CreateCommand();
                command.CommandText = "Delete from dbo.DataCFiscal";
                command.ExecuteNonQuery();
                sql.Close();
            }
        }

        public static DataTable ToDataTable<T>(IEnumerable<T> list)
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            dataTable.TableName = typeof(T).FullName;
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }

            foreach (T entity in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }
      

    }


}
