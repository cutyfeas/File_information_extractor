using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace FileExtractor
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Processing...");
            new Program().readDirectory();
            Console.WriteLine("Success !!");
            Console.ReadLine();
        }

        private void readDirectory()
        {
            DataTable dt = new DataTable();
            string ImageDirectory = ConfigurationManager.AppSettings["FolderPath"].ToString();
            DirectoryInfo di = new DirectoryInfo(ImageDirectory);
            FileInfo[] files = di.GetFiles();
            dt.Columns.Add("Image_Name");
            DataRow dr = null;
            foreach (var x in files)
            {
                dr = dt.NewRow();
                dr["Image_Name"] = x.Name.ToString();
                dt.Rows.Add(dr);
            }
            ExportExcel(dt, "Image_Name.xlsx");
        }

        private void ExportExcel(DataTable dt, string fileame)
        {
            string FileSavePath = ConfigurationManager.AppSettings["ExcelSavePath"].ToString();
            string tempFile = System.IO.Path.Combine(FileSavePath, fileame);
            using (XLWorkbook wb = new XLWorkbook())
            {
                IXLWorksheet ws = wb.Worksheets.Add("Export Data");
                int z = 1;
                foreach (DataColumn column in dt.Columns)
                {
                    ws.Cell(1, z).Value = column.ColumnName;
                    ws.Cell(1, z).Style.Font.Bold = true;
                    ws.Cell(1, z).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    ws.Cell(1, z).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    z += 1;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ws.Cell(i + 2, j + 1).Value = dt.Rows[i][j];
                        ws.Cell(i + 2, j + 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        ws.Cell(i + 2, j + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        ws.Cell(i + 2, 6).Style.DateFormat.Format = "dd-MMM-yyyy hh:mm:ss";
                    }
                }
                ws.Columns().AdjustToContents();
                wb.SaveAs(tempFile);
                Console.WriteLine(tempFile);
            }
        }
    }
}
