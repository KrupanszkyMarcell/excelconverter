using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.ComponentModel;
using System.Security.Claims;
using System.Threading;

namespace excelconverter
{
    public partial class MainWindow : Window
    {
        private DataTable resultsTable = new DataTable();
        private string filePath;

        public MainWindow()
        {
            InitializeComponent();

            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                filePath = openFileDialog.FileName;
                txtFilePath.Text = filePath;
                LoadSheetNames();
            }
        }

        private void LoadSheetNames()
        {
            try
            {
                cboInputSheet.Items.Clear();
                cboOutputSheet.Items.Clear();
                cboDateColumn.Items.Clear();
                cboEonColumn.Items.Clear();
                cboSolarColumn.Items.Clear();

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var sheet in package.Workbook.Worksheets)
                    {
                        cboInputSheet.Items.Add(sheet.Name);
                        cboOutputSheet.Items.Add(sheet.Name);
                    }

                    if (cboInputSheet.Items.Count > 0)
                    {
                        cboInputSheet.SelectedIndex = 0;
                        cboOutputSheet.SelectedIndex = Math.Min(1, cboOutputSheet.Items.Count - 1);
                        LoadColumnNames();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading sheet names: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadColumnNames()
        {
            try
            {
                cboDateColumn.Items.Clear();
                cboEonColumn.Items.Clear();
                cboSolarColumn.Items.Clear();

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[cboInputSheet.SelectedItem.ToString()];

                    // Get column names from the first row
                    int colCount = worksheet.Dimension.End.Column;
                    for (int col = 1; col <= colCount; col++)
                    {
                        var colName = worksheet.Cells[1, col].Text;

                        // If cell is empty, use column letter as name
                        if (string.IsNullOrWhiteSpace(colName))
                        {
                            colName = GetExcelColumnName(col);
                        }

                        cboDateColumn.Items.Add(colName);
                        cboEonColumn.Items.Add(colName);
                        cboSolarColumn.Items.Add(colName);
                    }

                    // Try to auto-select appropriate columns
                    if (cboDateColumn.Items.Count > 0)
                    {
                        // Try to find date/time column
                        var dateIndex = FindColumnIndexByPattern(cboDateColumn.Items, new[] { "date", "time", "timestamp", "időpont" });
                        cboDateColumn.SelectedIndex = dateIndex >= 0 ? dateIndex : 0;

                        // Try to find E.ON column
                        var eonIndex = FindColumnIndexByPattern(cboEonColumn.Items, new[] { "e.on", "eon", "érték" });
                        cboEonColumn.SelectedIndex = eonIndex >= 0 ? eonIndex : Math.Min(1, cboEonColumn.Items.Count - 1);

                        // Try to find Solar column
                        var solarIndex = FindColumnIndexByPattern(cboSolarColumn.Items, new[] { "solar", "sun", "nap" });
                        cboSolarColumn.SelectedIndex = solarIndex >= 0 ? solarIndex : Math.Min(2, cboSolarColumn.Items.Count - 1);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading column names: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int FindColumnIndexByPattern(System.Collections.IList items, string[] patterns)
        {
            for (int i = 0; i < items.Count; i++)
            {
                string item = items[i].ToString().ToLower();
                if (patterns.Any(p => item.Contains(p.ToLower())))
                {
                    return i;
                }
            }
            return -1;
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(filePath) || cboInputSheet.SelectedItem == null ||
                cboDateColumn.SelectedItem == null || cboEonColumn.SelectedItem == null ||
                cboSolarColumn.SelectedItem == null)
            {
                MessageBox.Show("Please select all required fields.", "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Read and convert the data
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[cboInputSheet.SelectedItem.ToString()];
                    var data = ReadExcelData(worksheet);
                    var hourlyData = ConvertToHourlyData(data);

                    // Display results
                    DisplayResults(hourlyData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during conversion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private List<DataPoint> ReadExcelData(ExcelWorksheet worksheet)
        {
            var data = new List<DataPoint>();

            // Get column indices
            int dateColumnIndex = cboDateColumn.SelectedIndex + 1;
            int eonColumnIndex = cboEonColumn.SelectedIndex + 1;
            int solarColumnIndex = cboSolarColumn.SelectedIndex + 1;

            // Read data rows (start from row 2, assuming row 1 has headers)
            int rowCount = worksheet.Dimension.End.Row;
            for (int row = 2; row <= rowCount; row++)
            {
                var dateStr = worksheet.Cells[row, dateColumnIndex].Text;
                var eonStr = worksheet.Cells[row, eonColumnIndex].Text;
                var solarStr = worksheet.Cells[row, solarColumnIndex].Text;

                if (!string.IsNullOrWhiteSpace(dateStr))
                {
                    try
                    {
                        // Parse date string based on selected format
                        var dateFormatStr = cboDateFormat.Text;
                        DateTime date;

                        // Try to parse the date with the selected format
                        if (!DateTime.TryParseExact(dateStr, dateFormatStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                        {
                            // If parsing fails, try a few common formats
                            string[] commonFormats = new[]
                            {
                                "MM/dd/yy HH:mm",
                                "dd/MM/yyyy HH:mm",
                                "yyyy-MM-dd HH:mm",
                                "MM/dd/yyyy HH:mm",
                                "MM/dd/yy H:mm",
                                "MM/dd/yyyy H:mm:ss"
                            };

                            // Try each format
                            if (!DateTime.TryParseExact(dateStr, commonFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                            {
                                // Last resort: try general parsing
                                if (!DateTime.TryParse(dateStr, out date))
                                {
                                    // Skip invalid dates
                                    continue;
                                }
                            }
                        }

                        // Parse E.ON and Solar values
                        double eonValue = 0;
                        double solarValue = 0;

                        if (!double.TryParse(eonStr, NumberStyles.Any, CultureInfo.InvariantCulture, out eonValue))
                        {
                            // Try with comma as decimal separator
                            double.TryParse(eonStr.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out eonValue);
                        }

                        if (!double.TryParse(solarStr, NumberStyles.Any, CultureInfo.InvariantCulture, out solarValue))
                        {
                            // Try with comma as decimal separator
                            double.TryParse(solarStr.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out solarValue);
                        }

                        data.Add(new DataPoint
                        {
                            Timestamp = date,
                            EonValue = eonValue,
                            SolarValue = solarValue
                        });
                    }
                    catch (Exception ex)
                    {
                        // Skip invalid data
                        Console.WriteLine($"Error parsing row {row}: {ex.Message}");
                    }
                }
            }

            return data;
        }

        private List<HourlyDataPoint> ConvertToHourlyData(List<DataPoint> data)
        {
            var hourlyData = new List<HourlyDataPoint>();
            var aggregationMethod = ((ComboBoxItem)cboAggregation.SelectedItem).Content.ToString();

            // Group the data by date (day and hour)
            var groupedData = data.GroupBy(d => new DateTime(
                d.Timestamp.Year,
                d.Timestamp.Month,
                d.Timestamp.Day,
                d.Timestamp.Hour,
                0, 0));

            foreach (var group in groupedData)
            {
                var hourlyPoint = new HourlyDataPoint
                {
                    Timestamp = group.Key
                };

                // Apply aggregation method
                switch (aggregationMethod)
                {
                    case "Average":
                        hourlyPoint.EonValue = group.Average(d => d.EonValue);
                        hourlyPoint.SolarValue = group.Average(d => d.SolarValue);
                        break;
                    case "Sum":
                        hourlyPoint.EonValue = group.Sum(d => d.EonValue);
                        hourlyPoint.SolarValue = group.Sum(d => d.SolarValue);
                        break;
                    case "Max":
                        hourlyPoint.EonValue = group.Max(d => d.EonValue);
                        hourlyPoint.SolarValue = group.Max(d => d.SolarValue);
                        break;
                    case "Min":
                        hourlyPoint.EonValue = group.Min(d => d.EonValue);
                        hourlyPoint.SolarValue = group.Min(d => d.SolarValue);
                        break;
                }

                hourlyData.Add(hourlyPoint);
            }

            // Sort by timestamp
            return hourlyData.OrderBy(d => d.Timestamp).ToList();
        }

        private void DisplayResults(List<HourlyDataPoint> hourlyData)
        {
            // Create a DataTable to hold the results
            resultsTable = new DataTable();
            resultsTable.Columns.Add("Timestamp", typeof(DateTime));
            resultsTable.Columns.Add("E.ON [kW]", typeof(double));
            resultsTable.Columns.Add("Solar [kW]", typeof(double));

            // Add rows to the table
            foreach (var point in hourlyData)
            {
                resultsTable.Rows.Add(
                    point.Timestamp,
                    Math.Round(point.EonValue, 2),
                    Math.Round(point.SolarValue, 2)
                );
            }

            // Bind the table to the DataGrid
            dgResults.ItemsSource = resultsTable.DefaultView;
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (resultsTable.Rows.Count == 0)
            {
                MessageBox.Show("No data to export. Please convert the data first.", "Export Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                // Show save file dialog
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Save Hourly Data",
                    FileName = "Hourly_Data.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var package = new ExcelPackage())
                    {
                        // Create a new worksheet
                        var worksheet = package.Workbook.Worksheets.Add("Hourly Data");

                        // Add headers
                        worksheet.Cells[1, 1].Value = "Timestamp";
                        worksheet.Cells[1, 2].Value = "E.ON [kW]";
                        worksheet.Cells[1, 3].Value = "Solar [kW]";

                        // Add data
                        for (int i = 0; i < resultsTable.Rows.Count; i++)
                        {
                            var row = resultsTable.Rows[i];
                            worksheet.Cells[i + 2, 1].Value = (DateTime)row["Timestamp"];
                            worksheet.Cells[i + 2, 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm";
                            worksheet.Cells[i + 2, 2].Value = (double)row["E.ON [kW]"];
                            worksheet.Cells[i + 2, 3].Value = (double)row["Solar [kW]"];
                        }

                        // Auto-fit columns
                        worksheet.Cells.AutoFitColumns();

                        // Save the file
                        package.SaveAs(new FileInfo(saveFileDialog.FileName));

                        MessageBox.Show("Data exported successfully!", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting data: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class DataPoint
    {
        public DateTime Timestamp { get; set; }
        public double EonValue { get; set; }
        public double SolarValue { get; set; }
    }

    public class HourlyDataPoint
    {
        public DateTime Timestamp { get; set; }
        public double EonValue { get; set; }
        public double SolarValue { get; set; }
    }
}

// App.xaml
< Application x: Class = "EnergyDataConverter.App"
             xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns: x = "http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns: local = "clr-namespace:EnergyDataConverter"
             StartupUri = "MainWindow.xaml" >
    < Application.Resources >


    </ Application.Resources >
</ Application >

// App.xaml.cs
using System;
using System.Windows;

namespace EnergyDataConverter
{
    public partial class App : Application
    {
    }
}

// Project file (EnergyDataConverter.csproj)
< Project Sdk = "Microsoft.NET.Sdk" >

  < PropertyGroup >
    < OutputType > WinExe </ OutputType >
    < TargetFramework > net6.0 - windows </ TargetFramework >
    < UseWPF > true </ UseWPF >
  </ PropertyGroup >

  < ItemGroup >
    < PackageReference Include = "EPPlus" Version = "6.2.10" />
  </ ItemGroup >

</ Project >