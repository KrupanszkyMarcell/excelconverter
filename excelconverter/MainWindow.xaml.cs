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
            ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
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

        private void InputSheet_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cboInputSheet.SelectedItem != null && !string.IsNullOrEmpty(filePath))
            {
                LoadColumnNames();
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

                        // Auto-detect date format based on the selected date column
                        if (cboDateColumn.SelectedIndex >= 0)
                        {
                            DetectAndSetDateFormat(worksheet, cboDateColumn.SelectedIndex + 1);
                        }
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
                // Validate input and output intervals
                var inputInterval = ((ComboBoxItem)cboInputInterval.SelectedItem).Content.ToString();
                var outputInterval = ((ComboBoxItem)cboOutputInterval.SelectedItem).Content.ToString();

                // Check if the output interval is larger than or equal to the input interval
                if (!IsOutputIntervalValid(inputInterval, outputInterval))
                {
                    MessageBox.Show("Output time interval must be equal to or larger than input time interval.",
                        "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Read and convert the data
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[cboInputSheet.SelectedItem.ToString()];
                    var data = ReadExcelData(worksheet);
                    var convertedData = ConvertData(data, inputInterval, outputInterval);

                    // Display results
                    DisplayResults(convertedData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during conversion: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool IsOutputIntervalValid(string inputInterval, string outputInterval)
        {
            // Define time intervals in ascending order (smaller to larger)
            var intervals = new List<string>
            {
                "15 Minutes",
                "1 Hour",
                "1 Day",
                "1 Week",
                "1 Month",
                "Quarter Year"
            };

            int inputIndex = intervals.IndexOf(inputInterval);
            int outputIndex = intervals.IndexOf(outputInterval);

            // Output interval should be equal to or larger than input interval
            return outputIndex >= inputIndex;
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
                        // Get the selected date format
                        var selectedDateFormat = cboDateFormat.Text;
                        DateTime date;
                        bool parsed = false;

                        // First try with the user-selected format
                        if (DateTime.TryParseExact(dateStr, selectedDateFormat,
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                        {
                            parsed = true;
                        }
                        else
                        {
                            // Define a larger set of formats to try
                            string[] commonFormats = new[]
                            {
                        "MM/dd/yyyy HH:mm:ss", // Add this format for your second image
                        "dd/MM/yyyy HH:mm:ss",
                        "MM/dd/yy HH:mm",
                        "dd/MM/yyyy HH:mm",
                        "yyyy-MM-dd HH:mm",
                        "MM/dd/yyyy HH:mm",
                        "MM/dd/yy H:mm",
                        "MM/dd/yyyy H:mm:ss",
                        "M/d/yyyy HH:mm:ss",  // Add this format for flexibility
                        "d/M/yyyy HH:mm:ss",
                        "yyyy-MM-dd"
                    };

                            // Try each format
                            foreach (var format in commonFormats)
                            {
                                if (DateTime.TryParseExact(dateStr, format,
                                    CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                                {
                                    parsed = true;
                                    break;
                                }
                            }

                            // If all formats fail, try general parsing as last resort
                            if (!parsed && DateTime.TryParse(dateStr, out date))
                            {
                                parsed = true;
                            }
                        }

                        if (!parsed)
                        {
                            // Skip this row if date couldn't be parsed
                            continue;
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
                        // Log the error but continue processing other rows
                        Console.WriteLine($"Error parsing row {row}: {ex.Message}");
                    }
                }
            }

            return data;
        }
        private List<DataPoint> ConvertData(List<DataPoint> data, string inputInterval, string outputInterval)
        {
            var convertedData = new List<DataPoint>();
            var aggregationMethod = ((ComboBoxItem)cboAggregation.SelectedItem).Content.ToString();

            // Check if we have any data to convert
            if (data.Count == 0)
            {
                MessageBox.Show("No data found or could not parse dates. Please check your date format selection.",
                    "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return convertedData;
            }

            // Log some data for debugging
            Console.WriteLine($"Converting {data.Count} data points");
            if (data.Count > 0)
            {
                Console.WriteLine($"First data point: {data[0].Timestamp} - E.ON: {data[0].EonValue}, Solar: {data[0].SolarValue}");
                if (data.Count > 1)
                    Console.WriteLine($"Second data point: {data[1].Timestamp} - E.ON: {data[1].EonValue}, Solar: {data[1].SolarValue}");
            }

            try
            {
                // Group the data by the appropriate time interval
                var groupedData = GroupDataByInterval(data, outputInterval);

                foreach (var group in groupedData)
                {
                    var newPoint = new DataPoint
                    {
                        Timestamp = group.Key
                    };

                    // Skip empty groups
                    if (!group.Any())
                        continue;

                    // Apply aggregation method
                    switch (aggregationMethod)
                    {
                        case "Average":
                            newPoint.EonValue = group.Average(d => d.EonValue);
                            newPoint.SolarValue = group.Average(d => d.SolarValue);
                            break;
                        case "Sum":
                            newPoint.EonValue = group.Sum(d => d.EonValue);
                            newPoint.SolarValue = group.Sum(d => d.SolarValue);
                            break;
                        case "Max":
                            newPoint.EonValue = group.Max(d => d.EonValue);
                            newPoint.SolarValue = group.Max(d => d.SolarValue);
                            break;
                        case "Min":
                            newPoint.EonValue = group.Min(d => d.EonValue);
                            newPoint.SolarValue = group.Min(d => d.SolarValue);
                            break;
                    }

                    convertedData.Add(newPoint);
                }

                // Sort by timestamp
                return convertedData.OrderBy(d => d.Timestamp).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during data conversion: {ex.Message}", "Conversion Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<DataPoint>();
            }
        }
        private IEnumerable<IGrouping<DateTime, DataPoint>> GroupDataByInterval(List<DataPoint> data, string interval)
        {
            switch (interval)
            {
                case "15 Minutes":
                    return data.GroupBy(d => new DateTime(
                        d.Timestamp.Year,
                        d.Timestamp.Month,
                        d.Timestamp.Day,
                        d.Timestamp.Hour,
                        15 * (d.Timestamp.Minute / 15), // Round to nearest 15 min
                        0));

                case "1 Hour":
                    return data.GroupBy(d => new DateTime(
                        d.Timestamp.Year,
                        d.Timestamp.Month,
                        d.Timestamp.Day,
                        d.Timestamp.Hour,
                        0, 0));

                case "1 Day":
                    return data.GroupBy(d => new DateTime(
                        d.Timestamp.Year,
                        d.Timestamp.Month,
                        d.Timestamp.Day,
                        0, 0, 0));

                case "1 Week":
                    return data.GroupBy(d => {
                        // Calculate the start of the week (Sunday)
                        int diff = (7 + (d.Timestamp.DayOfWeek - DayOfWeek.Sunday)) % 7;
                        return d.Timestamp.Date.AddDays(-1 * diff);
                    });

                case "1 Month":
                    return data.GroupBy(d => new DateTime(
                        d.Timestamp.Year,
                        d.Timestamp.Month,
                        1, 0, 0, 0));

                case "Quarter Year":
                    return data.GroupBy(d => {
                        int quarter = (d.Timestamp.Month - 1) / 3;
                        return new DateTime(d.Timestamp.Year, quarter * 3 + 1, 1, 0, 0, 0);
                    });

                default:
                    // Default to hourly if something unexpected happens
                    return data.GroupBy(d => new DateTime(
                        d.Timestamp.Year,
                        d.Timestamp.Month,
                        d.Timestamp.Day,
                        d.Timestamp.Hour,
                        0, 0));
            }
        }

        private void DisplayResults(List<DataPoint> convertedData)
        {
            // Create a DataTable to hold the results
            resultsTable = new DataTable();
            resultsTable.Columns.Add("Timestamp", typeof(DateTime));
            resultsTable.Columns.Add("E.ON [kW]", typeof(double));
            resultsTable.Columns.Add("Solar [kW]", typeof(double));

            // Add rows to the table
            foreach (var point in convertedData)
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
                    Title = "Save Converted Data",
                    FileName = "Converted_Data.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (var package = new ExcelPackage())
                    {
                        // Create a new worksheet or use the specified output sheet
                        var worksheetName = cboOutputSheet.SelectedItem?.ToString() ?? "Converted Data";
                        var worksheet = package.Workbook.Worksheets.Add(worksheetName);

                        // Add headers
                        worksheet.Cells[1, 1].Value = "Timestamp";
                        worksheet.Cells[1, 2].Value = "E.ON [kW]";
                        worksheet.Cells[1, 3].Value = "Solar [kW]";

                        // Add data
                        for (int i = 0; i < resultsTable.Rows.Count; i++)
                        {
                            var row = resultsTable.Rows[i];
                            worksheet.Cells[i + 2, 1].Value = (DateTime)row["Timestamp"];
                            worksheet.Cells[i + 2, 1].Style.Numberformat.Format = GetDateFormatForExport();
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

        private string GetDateFormatForExport()
        {
            var outputInterval = ((ComboBoxItem)cboOutputInterval.SelectedItem).Content.ToString();

            switch (outputInterval)
            {
                case "15 Minutes":
                case "1 Hour":
                    return "yyyy-MM-dd HH:mm";
                case "1 Day":
                    return "yyyy-MM-dd";
                case "1 Week":
                    return "yyyy-MM-dd 'Week'";
                case "1 Month":
                    return "yyyy-MM";
                case "Quarter Year":
                    return "yyyy 'Q'Q";
                default:
                    return "yyyy-MM-dd HH:mm";
            }
        }

        private void DetectAndSetDateFormat(ExcelWorksheet worksheet, int dateColumnIndex)
        {
            try
            {
                // Get a few sample dates from the column
                List<string> sampleDates = new List<string>();
                int maxSamples = 5;
                int count = 0;

                for (int row = 2; row <= worksheet.Dimension.End.Row && count < maxSamples; row++)
                {
                    string dateText = worksheet.Cells[row, dateColumnIndex].Text;
                    if (!string.IsNullOrWhiteSpace(dateText))
                    {
                        sampleDates.Add(dateText);
                        count++;
                    }
                }

                if (sampleDates.Count == 0)
                    return;

                // Define formats to test
                string[] dateFormats = new[]
                {
            "MM/dd/yyyy HH:mm:ss",
            "dd/MM/yyyy HH:mm:ss",
            "MM/dd/yy HH:mm",
            "dd/MM/yyyy HH:mm",
            "yyyy-MM-dd HH:mm",
            "MM/dd/yyyy HH:mm",
        };

                // Count successful parses for each format
                Dictionary<string, int> formatSuccessCount = new Dictionary<string, int>();
                foreach (var format in dateFormats)
                {
                    formatSuccessCount[format] = 0;
                }

                // Test each sample with each format
                foreach (string dateStr in sampleDates)
                {
                    foreach (string format in dateFormats)
                    {
                        if (DateTime.TryParseExact(dateStr, format, CultureInfo.InvariantCulture,
                            DateTimeStyles.None, out _))
                        {
                            formatSuccessCount[format]++;
                        }
                    }
                }

                // Find the format with most successful parses
                string bestFormat = dateFormats[0];
                int maxSuccesses = 0;

                foreach (var kvp in formatSuccessCount)
                {
                    if (kvp.Value > maxSuccesses)
                    {
                        maxSuccesses = kvp.Value;
                        bestFormat = kvp.Key;
                    }
                }

                // If we found a likely format, select it in the combobox
                if (maxSuccesses > 0)
                {
                    foreach (ComboBoxItem item in cboDateFormat.Items)
                    {
                        if (item.Content.ToString() == bestFormat)
                        {
                            cboDateFormat.SelectedItem = item;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Just log the error but don't disrupt the flow
                Console.WriteLine($"Error detecting date format: {ex.Message}");
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
}