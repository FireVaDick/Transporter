using Bytescout.Spreadsheet;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Transporter
{
    public partial class MainWindow : Window
    {
        private string filePathFrom;
        private string fileNameFrom;
        private string filePathTo;
        private string fileNameTo;

        private int driverColumnFrom;
        private int moneyColumnFrom;
        private int worksheetColumnFrom;
        private int driverColumnTo;
        private int moneyColumnTo;
        private int worksheetColumnTo;

        Dictionary<string, decimal> driverMoneyDictionary;


        #region Конструктор
        public MainWindow()
        {
            InitializeComponent();

            this.Left = (SystemParameters.FullPrimaryScreenWidth - this.Width) / 2;
            this.Top = (SystemParameters.FullPrimaryScreenHeight - this.Height) / 2;

            DataContext = this;
        }
        #endregion



        #region Горячие клавиши
        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            //switch (e.Key)
            //{
            //    case Key.X:
            //        SwapX_MouseLeftButtonDown(null, null);
            //        break;
            //}
        }
        #endregion



        #region Главные методы
        private void FileFrom_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Read(ref filePathFrom, ref fileNameFrom, FileInfoFromTb);
        }

        private void FileTo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Read(ref filePathTo, ref fileNameTo, FileInfoToTb);
        }

        private void Read(ref string filePath, ref string fileName, TextBlock fileInfoTb)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
            fileDialog.Title = "Выберите файл";
            bool? success = fileDialog.ShowDialog();

            if (success == true)
            {
                filePath = fileDialog.FileName;
                fileName = fileDialog.SafeFileName;

                fileInfoTb.Text = fileName;
            }
        }

        private void Transfer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (!AreFilesChosen()) return;
                if (!AreColumnsGood()) return;

                ShowMessage("Чтение исходного файла...", true);
                driverMoneyDictionary = ReadSourceFile(filePathFrom, driverColumnFrom, moneyColumnFrom, worksheetColumnFrom);

                if (driverMoneyDictionary.Count == 0)
                {
                    ShowMessage("Исходный файл не содержит данных", false);
                    return;
                }

                ShowMessage("Обновление целевого файла...", true);
                int updatedRows = UpdateTargetFile(filePathTo, driverColumnTo, moneyColumnTo, worksheetColumnTo);

                ShowMessage($"Успешно! Обновлено {updatedRows} строк. Файл сохранен.", true);

                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{filePathTo}\"");
            }
            catch (Exception ex)
            {
                ShowMessage($"Ошибка: {ex.Message}", false);
            }
        }
        #endregion



        #region Чтение из файла и заполнение словаря
        private Dictionary<string, decimal> ReadSourceFile(string filePath, int driverColumn, int moneyColumn, int worksheetIndex)
        {
            var dictionary = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

            Spreadsheet document = new Spreadsheet();

            try
            {
                document.LoadFromFile(filePath);
                Worksheet worksheet = document.Workbook.Worksheets[worksheetIndex];

                int rowCount = worksheet.UsedRangeRowMax;

                for (int row = 1; row <= rowCount; row++)
                {
                    string driverName = worksheet.Cell(row, driverColumn).Value?.ToString()?.Trim();
                    object moneyValueObj = worksheet.Cell(row, moneyColumn).Value;
                    string moneyText = moneyValueObj?.ToString()?.Trim();

                    if (!string.IsNullOrEmpty(driverName) && !string.IsNullOrEmpty(moneyText))
                    {
                        if (TryParseDecimal(moneyText, out decimal moneyValue))
                        {
                            if (dictionary.ContainsKey(driverName))
                            {
                                // Суммируем, если перевозчик уже встречался
                                dictionary[driverName] += moneyValue;
                            }
                            else
                            {
                                dictionary[driverName] = moneyValue;
                            }
                        }
                    }
                }
            }
            finally
            {
                document.Close();
            }

            return dictionary;
        }
        #endregion



        #region Обновление суммы в целевом файле
        private int UpdateTargetFile(string filePath, int driverColumn, int moneyColumn, int worksheetIndex)
        {
            int updatedRows = 0;

            Spreadsheet document = new Spreadsheet();

            try
            {
                document.LoadFromFile(filePath);

                Worksheet worksheet = document.Workbook.Worksheets[worksheetIndex];

                int rowCount = worksheet.UsedRangeRowMax;

                for (int row = 1; row <= rowCount; row++)
                {
                    string driverName = worksheet.Cell(row, driverColumn).Value?.ToString()?.Trim();

                    if (!string.IsNullOrEmpty(driverName) && driverMoneyDictionary.TryGetValue(driverName, out decimal moneyValue))
                    {
                        worksheet.Cell(row, moneyColumn).Value = moneyValue;
                        updatedRows++;
                    }
                }

                document.SaveAs(filePath);
            }
            finally
            {
                document.Close();
            }

            return updatedRows;
        }
        #endregion



        #region Вспомогательные методы
        private bool AreFilesChosen()
        {
            if (string.IsNullOrEmpty(filePathFrom) || !File.Exists(filePathFrom))
            {
                ShowMessage("Пожалуйста, выберите исходный файл", false);
                return false;
            }

            if (string.IsNullOrEmpty(filePathTo) || !File.Exists(filePathTo))
            {
                ShowMessage("Пожалуйста, выберите целевой файл", false);
                return false;
            }

            return true;
        }

        private bool AreColumnsGood()
        {
            driverColumnFrom = Convert.ToInt32(DriverColumnFromTB.Text) - 1;
            moneyColumnFrom = Convert.ToInt32(MoneyColumnFromTB.Text) - 1;
            worksheetColumnFrom = Convert.ToInt32(WorksheetColumnFromTB.Text) - 1;

            driverColumnTo = Convert.ToInt32(DriverColumnToTB.Text) - 1;
            moneyColumnTo = Convert.ToInt32(MoneyColumnToTB.Text) - 1;
            worksheetColumnTo = Convert.ToInt32(WorksheetColumnToTB.Text) - 1;

            if (driverColumnFrom < 0 || moneyColumnFrom < 0 || worksheetColumnFrom < 0 
                || driverColumnTo < 0 || moneyColumnTo < 0 || worksheetColumnTo < 0)
            {
                ShowMessage("Номера столбцов должны быть положительными числами (от 1)", false);
                return false;
            }

            return true;
        }
        private bool TryParseDecimal(string text, out decimal value)
        {
            // Убираем возможные пробелы и заменяем разделители
            text = text.Trim();
            text = text.Replace(",", ".");  // Заменяем запятую на точку
            text = text.Replace(" ", "");   // Убираем пробелы (например, в "1 000.50")

            return decimal.TryParse(text, System.Globalization.NumberStyles.Any,
                                  System.Globalization.CultureInfo.InvariantCulture, out value);
        }

        private void ShowMessage(string message, bool isSuccess)
        {
            if (isSuccess)
            {
                AddSuccess.Text = message;
                AddSuccess.Opacity = 1;
                AddError.Opacity = 0;
            }
            else
            {
                AddError.Text = message;
                AddError.Opacity = 1;
                AddSuccess.Opacity = 0;
            }

            // Автоматически скрываем сообщение через 5 секунд
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(5);
            timer.Tick += (s, args) =>
            {
                if (isSuccess)
                    AddSuccess.Opacity = 0;
                else
                    AddError.Opacity = 0;
                timer.Stop();
            };
            timer.Start();
        }
        #endregion
    }
}