// 参考: https://qiita.com/hukatama024e/items/37427f2578a8987645dd

using System;
using System.IO;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelCombiner
{
    static class NPOIUtility
    {
        public static IWorkbook CreateNewBook(string filePath)
        {
            IWorkbook book;
            var extension = Path.GetExtension(filePath);
            if (extension == Defines._EXTENSION_XLS)
            {
                book = new HSSFWorkbook();
            }
            else if (extension == Defines._EXTENSION_XLSX)
            {
                book = new XSSFWorkbook();
            }
            else
            {
                throw new ApplicationException("CreateNewBook: invalid extension");
            }

            return book;
        }

        public static IWorkbook LoadBook(string path)
        {
            var book = WorkbookFactory.Create(path);
            if (book == null)
            {
                Logger.WriteLine("Failed to load : " + path);
            }

            return book;
        }

        public static void WriteToCell(ISheet sheet, int columnIndex, int rowIndex, string value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            cell.SetCellType( CellType.String );
        }

        public static void WriteToCell(ISheet sheet, int columnIndex, int rowIndex, int value)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);

            cell.SetCellValue(value);
            cell.SetCellType(CellType.Numeric);
        }

        public static void WriteCellStyle(ISheet sheet, int columnIndex, int rowIndex, ICellStyle style)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);           
            cell.CellStyle = style;
        }

        public static ICell GetCell(ISheet sheet, int columnIndex, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            return row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        }
    }
}
