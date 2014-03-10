using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;

namespace SharePointListitemManager
{
    public interface IExcelService
    {
        void WriteListToExcel(string fileName, SharepointList list);
        SharepointList ReadListFromExcel(string filename);
    }

    public class ExcelService : IExcelService
    {
        public SharepointList ReadListFromExcel(string filename)
        {
            var newFile = new FileInfo(filename);
            var list = new SharepointList();
            if (!newFile.Exists)
            {
                return null;
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // Get worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                // Get columns
                list.Columns = new List<SharepointList.Column>();
                var col = 1;
                while (worksheet.Cells[1, col].Value != null)
                {
                    list.Columns.Add(new SharepointList.Column() { Name = worksheet.Cells[1, col].Value.ToString() });
                    col++;
                }

                var totalCols = col;

                // Get rows
                list.Rows = new List<SharepointList.Row>();
                var row = 2;
                while (worksheet.Cells[row, 1].Value != null)
                {
                    var spRow = new SharepointList.Row() { Values = new List<object>() };
                    for (col = 1; col < totalCols; col++)
                    {
                        Debug.WriteLine("Add " + col);
                        spRow.Values.Add(worksheet.Cells[row, col].Value);
                    }
                    list.Rows.Add(spRow);
                    row++;
                }
            }

            return list;
        }

        public void WriteListToExcel(string fileName, SharepointList list)
        {
            var newFile = new FileInfo(fileName);
			if (newFile.Exists)
			{
				newFile.Delete();
				newFile = new FileInfo(fileName);
			}

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");

                //Add the headers
                for (var i = 0; i < list.Columns.Count; i++)
                {
                    var column = list.Columns[i];
                    worksheet.Cells[1, i+1].Value = column.Name;
                }

                //Add the items
                for (var i = 0; i < list.Rows.Count; i++)
                {
                    var row = list.Rows[i];
                    for (var j = 0; j < row.Values.Count; j++)
                    {
                        var value = row.Values[j] != null ? row.Values[j].ToString() : row.Values[j];
                        worksheet.Cells[i+2, j+1].Value = value;
                    }
                }

                package.Save();
            }
        }
    }
}