using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.Style;

public class ExcelDraw
{
    public byte[] Draw(FileStream stream)
    {
        var package = new ExcelPackage();
        package.Load(stream);


        var sheet = package.Workbook.Worksheets.Add("List1");
        //var sheet = package.Workbook.Worksheets[0];

        int MaxColumn = 255;
        int MaxRow = 255;
        int Offset = 200;

        int column = MaxColumn;
        int row = MaxColumn;

        for (int j = 1; j < column; j++)
        {
            DrawRow(row, j, sheet);
        }

        return package.GetAsByteArray();
    }

    private async Task DrawRow(int row, int column, ExcelWorksheet sheet)
    {
        Random rand = new Random();
        for (int i = 1; i < row; i++)
        {
            //sheet.Cells[j, i].Value = i + j;
            //sheet.Cells[j, i].Value = n;
            sheet.Cells[column, i].Style.Fill.PatternType = ExcelFillStyle.Solid;

            sheet.Cells[column, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(
                255 / row * i > 255 ? 255 : 255 / row * i,
                255 / column * i > 255 ? 255 : 255 / column * i,
                255 / column * column > 255 ? 255 : 255 / column * column));
        }
    }
}