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

    public class ExcelParser
    {
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
        public byte[] Read(FileStream stream)
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

                    var procedureData = new Data.ProcedureData
                    {
                        DateOperation = DateOperation,
                        DateProcessing = DateProcessing,
                        AuthorizationCode = AuthorizationCode,
                        Category = Category,
                        OperationName = OperationName.Trim(),
                        Amount = Amount,
                        AccountBalance = AccountBalance
                    };
                    Data.budgetData.ProcedureDatas.Add(procedureData);
                    if (procedureData.Category == "")
                        Console.WriteLine(row);

                    for (int i = 2; i < Data.budgetData.ProcedureDatas.Count; i++)
                    {
                        Write(sheet.Cells[i, 1], Data.budgetData.ProcedureDatas[i].DateOperation.ToString());
                        Write(sheet.Cells[i, 2], Data.budgetData.ProcedureDatas[i].Category);
                        Write(sheet.Cells[i, 3], Data.budgetData.ProcedureDatas[i].AuthorizationCode.ToString());
                        Write(sheet.Cells[i, 4], Data.budgetData.ProcedureDatas[i].OperationName);
                        Write(sheet.Cells[i, 5], Data.budgetData.ProcedureDatas[i].Amount.ToString());
                        Write(sheet.Cells[i, 6], Data.budgetData.ProcedureDatas[i].AccountBalance.ToString());
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
