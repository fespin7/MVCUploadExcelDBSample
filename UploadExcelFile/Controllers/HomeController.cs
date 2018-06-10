using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using System.Configuration;

namespace UploadExcelFile.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
                try
                {
                    if (Path.GetExtension(file.FileName).Equals(".xlsx"))
                    {
                        var connectionString = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
                        var excel = new ExcelPackage(file.InputStream);
                        var dt = excel.ToDataTable();
                        var table = "People";
                        using (var conn = new SqlConnection(connectionString))
                        {
                            var truncateTable = new SqlCommand("TRUNCATE TABLE People", conn);
                            var bulkCopy = new SqlBulkCopy(conn);
                            bulkCopy.DestinationTableName = table;
                            conn.Open();
                            //truncate table
                            truncateTable.ExecuteNonQuery();
                            //start bulk copy operation
                            var schema = conn.GetSchema("Columns", new[] { null, null, table, null });
                            foreach (DataColumn sourceColumn in dt.Columns)
                            {
                                foreach (DataRow row in schema.Rows)
                                {
                                    if (string.Equals(sourceColumn.ColumnName, (string)row["COLUMN_NAME"], StringComparison.OrdinalIgnoreCase))
                                    {
                                        bulkCopy.ColumnMappings.Add(sourceColumn.ColumnName, (string)row["COLUMN_NAME"]);
                                        break;
                                    }
                                }
                            }
                            bulkCopy.WriteToServer(dt);
                            ViewBag.Message = "File loaded successfully";
                        }
                    }
                    else
                    {
                        ViewBag.Message = "Only excel files";
                    }
                }
                catch (Exception ex)
                {
                    ViewBag.Message = "ERROR:" + ex.Message.ToString();
                }
            else
            {
                ViewBag.Message = "You have not specified a file.";
            }
            return View();
        }



        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }

    public static class ExcelPackageExtensions
    {
        public static DataTable ToDataTable(this ExcelPackage package)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable table = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                table.Columns.Add(firstRowCell.Text);
            }
            for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = table.NewRow();
                foreach (var cell in row)
                {
                    //var format = cell.Style.Numberformat.Format;
                    //newRow[cell.Start.Column - 1] = cell.Text;
                    newRow[cell.Start.Column - 1] = cell.Value;
                }
                table.Rows.Add(newRow);
            }
            return table;
        }
    }
}