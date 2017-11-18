using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;


namespace ExcelReader
{

    public class ExcelFileParser
    {
        //TODO: Take these values from CONFIG
        public static string m_FilePath = "C:\\Users\\sven.LAWTRUST\\Desktop\\Sample_Data_Import.xlsx";
        public static string m_FileExtension = ".xlsx";
        public static bool m_IgnoreFirstRow = true;

        public ExcelFileParser(string m_FilePath, bool IgnoreFirstRow)
        {
        }


        /// <summary>
        /// This method performs the parsing of an Excel file.
        /// </summary>
        /// <param name="numberOfRows">The number of rows to return. If not specified, then all rows will be returned.</param>
        /// <returns>A list of CSV strings.</returns>
        public Dictionary<string, List<string>> ParseFile()
        {
            Dictionary<string, List<string>> rawRows = new Dictionary<string, List<string>>();

            using (FileStream fs = File.Open(m_FilePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader excelDataReader = null;

                switch (m_FileExtension)
                {
                    case ".xls":
                        excelDataReader = ExcelReaderFactory.CreateBinaryReader(fs);
                        break;
                    case ".xlsx":
                        excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                        break;
                }

                if (excelDataReader != null)
                {
                    using (DataSet ds = excelDataReader.AsDataSet())
                    {
                        if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                        {
                            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                            {
                                rawRows.Add(Convert.ToString(ds.Tables[0].Rows[0][i]), new List<string>());
                                rawRows[Convert.ToString(ds.Tables[0].Rows[0][i])].AddRange(ds.Tables[0].AsEnumerable()
                                    .Skip(1).Select(dr => dr[i].ToString()).ToArray());
                            }
                        }
                    }
                }
            }

            return rawRows;
        }


        static void Main(string[] args)
        {
            ExcelFileParser efp = new ExcelFileParser(m_FilePath, m_IgnoreFirstRow);

            var excelDictionary = efp.ParseFile();
            var keys = excelDictionary.Keys.ToList();
            var v = 0; // ValuePair values index
            var totalRowsCount = excelDictionary.Values.Sum(x => x.Count);
            var rowsProcessedCount = 1;

            for (int k = 0; k < excelDictionary.Keys.Count; k++)
            {
                Console.WriteLine($"{keys[k]}: {excelDictionary[keys[k]][v]}");

                if (totalRowsCount == rowsProcessedCount)
                {
                    break;
                }
                else
                {
                    v = v == excelDictionary[keys[k]].Count ? 0 : (k == excelDictionary.Keys.Count - 1) ? v += 1 : v;
                    k = k == excelDictionary.Keys.Count - 1 ? -1 : k; // Set to -1 after values check, because post-incrementer in loop constructor will set to zero again
                }

                rowsProcessedCount++;
            }

            Console.ReadLine();
        }

    }
}
