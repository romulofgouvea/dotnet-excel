using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace pocExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo fileInfo = new FileInfo(@"C:\workspace\pocs\pocExcel\pocExcel\Modelos\NovaContratacao.xlsm");

            //ReadAllSheet(fileInfo, 2);
            //ReadColumnOrLineSheetFromString(fileInfo, 1, "C2:C5");
            //ReadColumnOrLineSheetFromString(fileInfo, 1, "C2:I3");
            //ReadColumnOrLineSheetFromInt(fileInfo, 1, 3, 3);
            UpdateValueCell(fileInfo, 1, "C3", "123");
        }

        public static void ReadAllSheet(FileInfo fileInfo, int sheet)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                        }
                    }
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("Posição inexistente");
            }
        }

        public static void ReadColumnOrLineSheetFromString(FileInfo fileInfo, int sheet, string column)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count

                    foreach (var cell in worksheet.Cells[column])
                    {
                        Console.WriteLine("Column: " + cell.Value.ToString());
                    }
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("Posição inexistente");
            }
        }

        public static void ReadColumnOrLineSheetFromInt(FileInfo fileInfo, int sheet, int Row, int column)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count

                    foreach (var cell in worksheet.Cells[Row, column])
                    {
                        Console.WriteLine("Column: " + cell.Value.ToString());
                    }
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("Posição inexistente");
            }
        }

        public static void UpdateValueCell(FileInfo fileInfo, int sheet, string cell, string value)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count

                    Console.WriteLine("Antes: " + worksheet.Cells[cell].Value);
                    worksheet.Cells[cell].Value = value.Trim();
                    Console.WriteLine("Depois: " + worksheet.Cells[cell].Value);

                    var savePath = Directory.GetCurrentDirectory() + "/Exportacao/"+ fileInfo.Name;
                    Console.WriteLine("Path: " + savePath);

                    //Cria o arquivo
                    var fileStream = File.Create(savePath);
                    fileStream.Close();

                    //escreve no arquivo
                    File.WriteAllBytes(savePath, package.GetAsByteArray());
                }
            }
            catch (IndexOutOfRangeException)
            {
                Console.WriteLine("Posição inexistente");
            }
        }
    }
}
