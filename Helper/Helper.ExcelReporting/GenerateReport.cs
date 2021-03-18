namespace Helper.ExcelReporting
{
    using OfficeOpenXml;
    using OfficeOpenXml.Style;
    using System;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Web;

    public static class GenerateReport
    {
        public static void Generate(DataTable dt)
        {
            try
            {
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ClearContent();
                HttpContext.Current.Response.ClearHeaders();
                HttpContext.Current.Response.Buffer = true;
                HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8;
                HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
                HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=WellQuailtyTracker_Forecast.xlsx");


                using (ExcelPackage pack = new ExcelPackage())
                {
                    ExcelWorksheet ws = pack.Workbook.Worksheets.Add("ExcelReport");

                    var modelCells = ws.Cells["A1"];
                    var modelRows = dt.Rows.Count + 1;
                    string modelRange = "A1:AP" + modelRows.ToString();
                    var modelTable = ws.Cells[modelRange];

                    var UnLockCells = ws.Cells["C2"];
                    string UnLockRange = "C2:AP" + modelRows.ToString();


                    // Assign borders
                    modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


                    // Fill worksheet with data to export
                    modelCells.LoadFromDataTable(dt, true);
                    modelTable.AutoFitColumns();

                    ws.Protection.IsProtected = true;
                    ws.Cells[2, 3, modelRows, 42].Style.Locked = false;

                    var ms = new System.IO.MemoryStream();
                    pack.SaveAs(ms);
                    ms.WriteTo(HttpContext.Current.Response.OutputStream);
                }

                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.End();
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static DataTable ExcelToDataTable(Stream bytes)
        {
            ExcelPackage package = new ExcelPackage(bytes);
            return package.ToDataTable();
        }

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
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                table.Rows.Add(newRow);
            }
            return table;
        }

    }
}
