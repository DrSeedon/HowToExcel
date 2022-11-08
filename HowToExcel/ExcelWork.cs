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

namespace HowToExcel
{
    public class ExcelWork
    {
        public BudgetData BudgetData = new BudgetData();


        public byte[] Generate(MarketReport report)
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets
                .Add("Market Report");

            sheet.Cells["B2"].Value = "Company:";
            sheet.Cells[2, 3].Value = report.Company.Name;
            sheet.Cells["B3"].Value = "Location:";
            sheet.Cells["C3"].Value = $"{report.Company.Address}, " +
                                      $"{report.Company.City}, " +
                                      $"{report.Company.Country}";
            sheet.Cells["B4"].Value = "Sector:";
            sheet.Cells["C4"].Value = report.Company.Sector;
            sheet.Cells["B5"].Value = report.Company.Description;

            sheet.Cells[8, 2, 8, 4].LoadFromArrays(new object[][] {new[] {"Capitalization", "SharePrice", "Date"}});
            var row = 9;
            var column = 2;
            foreach (var item in report.History)
            {
                sheet.Cells[row, column].Value = item.Capitalization;
                sheet.Cells[row, column + 1].Value = item.SharePrice;
                sheet.Cells[row, column + 2].Value = item.Date;
                row++;
            }

            sheet.Cells[1, 1, row, column + 2].AutoFitColumns();
            sheet.Column(2).Width = 14;
            sheet.Column(3).Width = 12;

            sheet.Cells[9, 4, 9 + report.History.Length, 4].Style.Numberformat.Format = "yyyy";
            sheet.Cells[9, 2, 9 + report.History.Length, 2].Style.Numberformat.Format = "### ### ### ##0";

            sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[8, 3, 8 + report.History.Length, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            sheet.Cells[8, 2, 8, 4].Style.Font.Bold = true;
            sheet.Cells["B2:C4"].Style.Font.Bold = true;

            sheet.Cells[8, 2, 8 + report.History.Length, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[8, 2, 8, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            var capitalizationChart =
                sheet.Drawings.AddChart("FindingsChart", OfficeOpenXml.Drawing.Chart.eChartType.Line);
            capitalizationChart.Title.Text = "Capitalization";
            capitalizationChart.SetPosition(7, 0, 5, 0);
            capitalizationChart.SetSize(800, 400);
            var capitalizationData =
                (ExcelChartSerie) (capitalizationChart.Series.Add(sheet.Cells["B9:B28"], sheet.Cells["D9:D28"]));
            var capitalizationData2 =
                (ExcelChartSerie) (capitalizationChart.Series.Add(sheet.Cells["C9:C28"], sheet.Cells["B9:B28"]));
            capitalizationData.Header = report.Company.Currency;
            capitalizationData2.Header = report.Company.Currency + "123123";

            //sheet.Protection.IsProtected = true;
            return package.GetAsByteArray();
        }

        public void ReadSimpleData(FileStream stream)
        {
            var package = new ExcelPackage();
            package.Load(stream);
        }

        public string[] SplitWord = new string[]
        {
            "Прочие операции",
            "Рестораны и кафе",
            "Супермаркеты",
            "Перевод с карты",
            "Неизвестная категория(-)",
            "Внесение наличных",
            "Транспорт",
            "Здоровье и красота",
            "Одежда и аксессуары",
            "Все для дома",
            "Отдых и развлечения",
            "Прочие расходы",
            "Возврат, отмена операции",
            "все для дома",
            "Выдача наличных",
            "Комунальные платежи, связь, интернет.",
        };
        public string separator = "ДАТА ОПЕРАЦИИ (МСК) Дата обработки1 и код авторизации";
        public string separator2 = "Дата обработки3 авторизации";
        public string endText = "Дата формирования";
        public byte[] ReadSberData(FileStream stream)
        {
            var package = new ExcelPackage();
            package.Load(stream);

            var sheet = package.Workbook.Worksheets[0];
            var sheetEdit = package.Workbook.Worksheets.Add("Edit");
            Console.Write(sheet.Name);


            for (int i = 1; i < 2000; i++)
            {
                string a = (string) sheet.Cells[i, 1].Value;
                if (a == null)
                    continue;

                if (separator == a)
                    ParseData(ref i);
                if (separator2 == a)
                    ParseData(ref i);

                if (endText == a)
                    break;
            }

            sheetEdit.Cells.AutoFitColumns();
            return package.GetAsByteArray();

            void ParseData(ref int row)
            {
                row++;
                bool isFirstIteration = true;
                bool isNullData = false;
                for (;; row++)
                {
                    DateTime DateOperation = new DateTime();
                    DateTime DateProcessing = new DateTime();
                    string OperationName = "";
                    string Category = "";
                    double Amount = 0;
                    double AccountBalance = 0;
                    string AuthorizationCode = "";

                    string a = (string) sheet.Cells[row, 1].Value;
                    if (a == null)
                        return;

                    string nullCell = "";
                    if (isFirstIteration)
                    {
                        nullCell = (string) sheet.Cells[row + 1, 5].Value;
                        if(nullCell == null)
                            isNullData = true;
                        isFirstIteration = false;
                    }

                    if (row > 900)
                    {
                        var dsf = 2;
                    }
                    
                    if (nullCell == null || isNullData)
                    {
                        var columnData1 = sheet.Cells[row, 1].Value.ToString();
                        var columnData2 = sheet.Cells[row + 1, 1].Value.ToString();
                        var columnTime1 = sheet.Cells[row, 2].Value.ToString();
                        var columnTime2 = sheet.Cells[row + 1, 2].Value.ToString();

                        var time = $"{columnData1} {columnTime1}";
                        DateOperation = DateTime.Parse(time);
                        DateProcessing = Convert.ToDateTime(columnData2);
                        AuthorizationCode = columnTime2;
                        Category = sheet.Cells[row, 3].Value.ToString();
                        OperationName = sheet.Cells[row + 1, 3].Value.ToString();

                        Amount = Convert.ToDouble(sheet.Cells[row, 4].Value.ToString());
                        AccountBalance = Convert.ToDouble(sheet.Cells[row, 5].Value.ToString());
                        row++;
                    }
                    else
                    {
                        var columnData = sheet.Cells[row, 1].Value.ToString().Split(" ");
                        var columnTime = sheet.Cells[row, 2].Value.ToString().Split(" ");

                        var time = $"{columnData[0]} {columnTime[0]}";
                        DateOperation = DateTime.Parse(time);
                        DateProcessing = Convert.ToDateTime(columnData[1]);
                        AuthorizationCode = columnTime[1];
                        
                        var rawColumnCategory = sheet.Cells[row, 3].Value.ToString();
                        
                        foreach (var splitWord in SplitWord)
                        {
                            if (rawColumnCategory.Contains(splitWord))
                            {
                                var word = rawColumnCategory.Split(splitWord);
                                OperationName = word[1];
                                Category = splitWord;
                                break;
                            }
                        }

                        double isNegative = -1;
                        string amount = sheet.Cells[row, 4].Value.ToString();
                        if (amount.Substring(0, 1) == "+")
                        {
                            isNegative = 1;
                        }
                        Amount = isNegative * Convert.ToDouble(sheet.Cells[row, 4].Value.ToString());
                        AccountBalance = Convert.ToDouble(sheet.Cells[row, 5].Value.ToString());
                    }

                    var procedureData = new ProcedureData
                    {
                        DateOperation = DateOperation,
                        DateProcessing = DateProcessing,
                        AuthorizationCode = AuthorizationCode,
                        Category = Category,
                        OperationName = OperationName.Trim(),
                        Amount = Amount,
                        AccountBalance = AccountBalance
                    };
                    BudgetData.ProcedureDatas.Add(procedureData);
                    if (procedureData.Category == "")
                        Console.WriteLine(row);

                    for (int i = 2; i < BudgetData.ProcedureDatas.Count; i++)
                    {
                        Write(sheet.Cells[i, 1], BudgetData.ProcedureDatas[i].DateOperation.ToString());
                        Write(sheet.Cells[i, 2], BudgetData.ProcedureDatas[i].Category);
                        Write(sheet.Cells[i, 3], BudgetData.ProcedureDatas[i].AuthorizationCode.ToString());
                        Write(sheet.Cells[i, 4], BudgetData.ProcedureDatas[i].OperationName);
                        Write(sheet.Cells[i, 5], BudgetData.ProcedureDatas[i].Amount.ToString());
                        Write(sheet.Cells[i, 6], BudgetData.ProcedureDatas[i].AccountBalance.ToString());
                    }
                }
            }

            void Write(ExcelRange cell, string value)
            {
                sheetEdit.Cells[cell.End.Row, cell.End.Column].Value = value;
            }
        }

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
}

public class BudgetData
{
    public List<ProcedureData> ProcedureDatas = new List<ProcedureData>();
}

public class ProcedureData
{
    public DateTime DateOperation;
    public DateTime DateProcessing;
    public string AuthorizationCode;
    public string Category;
    public string OperationName;
    public double Amount;
    public double AccountBalance;
}