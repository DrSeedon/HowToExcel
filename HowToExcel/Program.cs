using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace HowToExcel
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            string dir = "C:/Users/User/Desktop/HowToExcel/Files/";
            //string nameDir = "C:/Users/User/Desktop/HowToExcel/Files/test.xlsx";
            //string nameDir = CreateRandomFile(dir);
            string nameDir = dir + "source.xlsx";

            byte[] tableData;
            ExcelWork excelWork = new ExcelWork();
            using (FileStream fs = File.Open(nameDir, FileMode.OpenOrCreate))
            {
                tableData = excelWork.ReadSberData(fs);
            }

            string editDir = dir + "maximEdit.xlsx";
            WriteDataFile(editDir, tableData);
            OpenFile(editDir);
        }

        /// <summary>
        /// Создает рандомное имя файла XLSX в директории
        /// </summary>
        /// <param name="dir">Директория</param>
        /// <returns>Рандомное имя</returns>
        private static string CreateRandomFile(string dir)
        {
            var rand = new Random();
            string nameDir = dir + "Table " + rand.Next() + ".xlsx";
            return nameDir;
        }

        /// <summary>
        /// Открывает файл приложением в OS
        /// </summary>
        /// <param name="nameDir">Путь к файлу</param>
        private static void OpenFile(string nameDir)
        {
            using (var p = new Process())
            {
                p.StartInfo = new ProcessStartInfo(@nameDir)
                {
                    UseShellExecute = true
                };
                p.Start();
            }
        }

        /// <summary>
        /// Записывает данные таблицы в файл
        /// </summary>
        /// <param name="nameDir">Путь к файлу</param>
        /// <param name="tableData">Данные таблицы</param>
        private static void WriteDataFile(string nameDir, byte[] tableData)
        {
            File.WriteAllBytes(nameDir, tableData);
        }
    }
}