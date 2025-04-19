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
    // Define the DynamicDataPoint class at the top of the file, immediately after the namespace declaration
    public class DynamicDataPoint
    {
        public DateTime Timestamp { get; set; }
        public Dictionary<string, double> Values { get; set; }

        public DynamicDataPoint()
        {
            Values = new Dictionary<string, double>();
        }
    }

    public partial class MainWindow : Window
    {
        private DataTable resultsTable = new DataTable();
        private string filePath;
        private List<string> selectedDataColumns = new List<string>(); // Store selected data column names
        private Dictionary<string, ComboBox> dataColumnControls = new Dictionary<string, ComboBox>(); // For tracking UI elements

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
                cboAddDataColumn.Items.Clear();

                // Clear existing data column selections
                ClearDataColumns();

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
                cboAddDataColumn.Items.Clear();

                // Clear existing data column selections
                ClearDataColumns();

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
                        cboAddDataColumn.Items.Add(colName);
                    }

                    // Try to auto-select appropriate columns
                    if (cboDateColumn.Items.Count > 0)
                    {
                        // Try to find date/time column
                        var dateIndex = FindColumnIndexByPattern(cboDateColumn.Items, new[] { "date", "time", "timestamp", "időpont" });
                        cboDateColumn.SelectedIndex = dateIndex >= 0 ? dateIndex : 0;

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

        private void ClearDataColumns()
        {
            // Clear list of selected columns
            selectedDataColumns.Clear();

            // Clear dictionary of UI controls
            dataColumnControls.Clear();

            // Clear UI panel
            pnlDataColumns.Children.Clear();
        }

        private void AddDataColumn_Click(object sender, RoutedEventArgs e)
        {
            if (cboAddDataColumn.SelectedItem == null)
            {
                MessageBox.Show("Please select a column to add.", "Selection Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string selectedColumn = cboAddDataColumn.SelectedItem.ToString();

            // Check if this column is already selected
            if (selectedDataColumns.Contains(selectedColumn))
            {
                MessageBox.Show("This column is already added.", "Duplicate Column", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check if this is the date column
            if (cboDateColumn.SelectedItem != null && selectedColumn == cboDateColumn.SelectedItem.ToString())
            {
                MessageBox.Show("This column is already selected as the Date Column.", "Duplicate Column", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Add to the list of selected columns
            selectedDataColumns.Add(selectedColumn);

            // Create UI for this column
            StackPanel panel = new StackPanel();
            panel.Orientation = Orientation.Horizontal;
            panel.Margin = new Thickness(0, 5, 0, 5);

            // Create remove button
            Button removeButton = new Button();
            removeButton.Content = "Remove";
            removeButton.Padding = new Thickness(5, 2, 5, 2);
            removeButton.Margin = new Thickness(0, 0, 10, 0);
            removeButton.Tag = selectedColumn;
            removeButton.Click += RemoveDataColumn_Click;

            // Create column label
            Label columnLabel = new Label();
            columnLabel.Content = selectedColumn;
            columnLabel.Width = 200;
            columnLabel.VerticalAlignment = VerticalAlignment.Center;

            // Add controls to panel
            panel.Children.Add(removeButton);
            panel.Children.Add(columnLabel);

            // Add panel to the main UI container
            pnlDataColumns.Children.Add(panel);

            // Optionally, remove the column from the dropdown to prevent reselection
            cboAddDataColumn.Items.Remove(selectedColumn);
            if (cboAddDataColumn.Items.Count > 0)
            {
                cboAddDataColumn.SelectedIndex = 0;
            }
        }

        private void RemoveDataColumn_Click(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            if (button != null)
            {
                string columnName = button.Tag.ToString();

                // Remove from the list
                selectedDataColumns.Remove(columnName);

                // Find and remove the UI row
                for (int i = 0; i < pnlDataColumns.Children.Count; i++)
                {
                    StackPanel panel = pnlDataColumns.Children[i] as StackPanel;
                    if (panel != null)
                    {
                        Button removeBtn = panel.Children[0] as Button;
                        if (removeBtn != null && removeBtn.Tag.ToString() == columnName)
                        {
                            pnlDataColumns.Children.RemoveAt(i);
                            break;
                        }
                    }
                }

                // Add back to the dropdown
                if (!cboAddDataColumn.Items.Contains(columnName))
                {
                    cboAddDataColumn.Items.Add(columnName);
                }
            }
        }

        private void FinishAddingColumns_Click(object sender, RoutedEventArgs e)
        {
            // Optional: Validate that at least one data column has been selected
            if (selectedDataColumns.Count == 0)
            {
                MessageBox.Show("Please add at least one data column before finishing.",
                    "Data Column Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Visually indicate that column selection is complete
            btnAddDataColumn.IsEnabled = false;
            cboAddDataColumn.IsEnabled = false;
            btnFinishAdding.IsEnabled = false;

            MessageBox.Show("Data column selection completed. Click 'Convert Data' when ready to proceed.",
                "Selection Complete", MessageBoxButton.OK, MessageBoxImage.Information);
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
                cboDateColumn.SelectedItem == null || selectedDataColumns.Count == 0)
            {
                MessageBox.Show("Please select all required fields and at least one data column.",
                    "Validation Error", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private List<DynamicDataPoint> ReadExcelData(ExcelWorksheet worksheet)
        {
            var data = new List<DynamicDataPoint>();

            // Get column indices
            int dateColumnIndex = cboDateColumn.SelectedIndex + 1;

            // Create a dictionary to map column names to column indices
            Dictionary<string, int> dataColumnIndices = new Dictionary<string, int>();

            // Find the column indices for all selected data columns
            foreach (string columnName in selectedDataColumns)
            {
                // Find the column index by name
                int columnIndex = -1;
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var colName = worksheet.Cells[1, col].Text;
                    if (string.IsNullOrWhiteSpace(colName))
                    {
                        colName = GetExcelColumnName(col);
                    }

                    if (colName == columnName)
                    {
                        columnIndex = col;
                        break;
                    }
                }

                if (columnIndex > 0)
                {
                    dataColumnIndices.Add(columnName, columnIndex);
                }
            }

            // Read data rows (start from row 2, assuming row 1 has headers)
            int rowCount = worksheet.Dimension.End.Row;

            // Debug info - log first few rows for diagnostics
            Console.WriteLine($"Reading data from {rowCount} rows");
            Console.WriteLine($"Date column: {dateColumnIndex}, Data columns: {string.Join(", ", dataColumnIndices.Select(kvp => $"{kvp.Key}:{kvp.Value}"))}");

            // Create a comprehensive list of date formats to try
            string[] commonFormats = new[]
            {
                "MM/dd/yyyy HH:mm:ss",
                "dd/MM/yyyy HH:mm:ss",
                "MM/dd/yy HH:mm",
                "dd/MM/yyyy HH:mm",
                "yyyy-MM-dd HH:mm",
                "MM/dd/yyyy HH:mm",
                "yyyy-MM-dd HH:mm:ss",
                "dd/MM/yyyy H:mm",
                "MM/dd/yyyy H:mm",
                "M/d/yyyy HH:mm:ss",
                "d/M/yyyy HH:mm:ss",
                "d/M/yyyy H:mm:ss",
                "dd/MM/yyyy",
                "MM/dd/yyyy",
                "yyyy-MM-dd",
                "dd.MM.yyyy HH:mm",
                "dd.MM.yyyy HH:mm:ss"
            };

            int successfullyParsed = 0;
            int failedToParse = 0;

            for (int row = 2; row <= rowCount; row++)
            {
                var dateCell = worksheet.Cells[row, dateColumnIndex];

                // First check if the date cell contains a real Excel date
                DateTime timestamp;
                bool dateValid = false;

                if (dateCell.Value is DateTime)
                {
                    timestamp = (DateTime)dateCell.Value;
                    dateValid = true;
                }
                else
                {
                    // Try to parse as text
                    var dateStr = dateCell.Text;

                    if (!string.IsNullOrWhiteSpace(dateStr))
                    {
                        // First try with the user-selected format
                        var selectedDateFormat = cboDateFormat.Text;

                        // Try the selected format first
                        if (DateTime.TryParseExact(dateStr, selectedDateFormat,
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out timestamp))
                        {
                            dateValid = true;
                        }
                        else
                        {
                            // Try each format in our comprehensive list
                            foreach (var format in commonFormats)
                            {
                                if (DateTime.TryParseExact(dateStr, format,
                                    CultureInfo.InvariantCulture, DateTimeStyles.None, out timestamp))
                                {
                                    dateValid = true;
                                    break;
                                }
                            }

                            // If all specific formats fail, try general parsing as last resort
                            if (!dateValid && DateTime.TryParse(dateStr, out timestamp))
                            {
                                dateValid = true;
                            }
                        }
                    }
                    else
                    {
                        timestamp = DateTime.MinValue; // Default value
                    }
                }

                if (dateValid)
                {
                    // Create a new dynamic data point with the timestamp
                    var dataPoint = new DynamicDataPoint { Timestamp = timestamp };

                    // Read all the data values for this row
                    foreach (var columnPair in dataColumnIndices)
                    {
                        string columnName = columnPair.Key;
                        int columnIndex = columnPair.Value;

                        var dataCell = worksheet.Cells[row, columnIndex];
                        double value = ParseNumericValue(dataCell);

                        // Add this value to the data point
                        dataPoint.Values[columnName] = value;
                    }

                    data.Add(dataPoint);
                    successfullyParsed++;
                }
                else
                {
                    failedToParse++;
                    // Log the problematic date string for diagnostics
                    if (failedToParse < 10)  // Limit logging to avoid console overflow
                    {
                        Console.WriteLine($"Failed to parse date at row {row}: '{dateCell.Text}'");
                    }
                }
            }

            Console.WriteLine($"Successfully parsed {successfullyParsed} rows, failed to parse {failedToParse} rows");

            // Log some sample data for verification
            if (data.Count > 0)
            {
                Console.WriteLine("First 5 parsed data points:");
                for (int i = 0; i < Math.Min(5, data.Count); i++)
                {
                    Console.WriteLine($"{data[i].Timestamp:yyyy-MM-dd HH:mm:ss} - " +
                        string.Join(", ", data[i].Values.Select(kvp => $"{kvp.Key}: {kvp.Value}")));
                }
            }

            return data;
        }

        // Helper method to parse numeric values with various formats
        private double ParseNumericValue(ExcelRange cell)
        {
            // If the cell contains a numeric value directly, use it
            if (cell.Value is double doubleValue)
            {
                return doubleValue;
            }

            // Try to parse the text value
            string valueStr = cell.Text;
            double result = 0;

            // First try standard parsing with invariant culture
            if (!string.IsNullOrWhiteSpace(valueStr) &&
                !double.TryParse(valueStr, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
            {
                // Try with different decimal separators
                if (valueStr.Contains(',') && !valueStr.Contains('.'))
                {
                    // European format with comma as decimal separator
                    double.TryParse(valueStr.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                }
                else if (valueStr.Contains('.') && valueStr.Contains(',') && valueStr.IndexOf(',') > valueStr.IndexOf('.'))
                {
                    // Format like "1.234,56"
                    double.TryParse(valueStr.Replace(".", "").Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                }
                else
                {
                    // Try to remove any non-numeric characters except decimal separators
                    string cleanValue = new string(valueStr.Where(c => char.IsDigit(c) || c == '.' || c == ',' || c == '-').ToArray());
                    double.TryParse(cleanValue.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                }
            }

            return result;
        }

        private IEnumerable<IGrouping<DateTime, DynamicDataPoint>> GroupDataByInterval(List<DynamicDataPoint> data, string interval)
        {
            // First sort the data by timestamp to ensure proper grouping
            var sortedData = data.OrderBy(d => d.Timestamp).ToList();

            switch (interval)
            {
                case "15 Minutes":
                    return sortedData.GroupBy(d => new DateTime(
                        d.Timestamp.Year,
                        d.Timestamp.Month,
                        d.Timestamp.Day,
                        d.Timestamp.Hour,
                        (d.Timestamp.Minute / 15) * 15, // Round down to nearest 15 min
                        0));

                case "1 Hour":
                    return sortedData.GroupBy(d => {
                        // Special case: if already exactly on the hour (00:00), keep it at that hour
                        if (d.Timestamp.Minute == 0 && d.Timestamp.Second == 0)
                        {
                            return new DateTime(
                                d.Timestamp.Year,
                                d.Timestamp.Month,
                                d.Timestamp.Day,
                                d.Timestamp.Hour,
                                0, 0);
                        }
                        else
                        {
                            // Otherwise round up to the next hour
                            return new DateTime(
                                d.Timestamp.Year,
                                d.Timestamp.Month,
                                d.Timestamp.Day,
                                d.Timestamp.Hour,
                                0, 0).AddHours(1);
                        }
                    });

                case "1 Day":
                    return sortedData.GroupBy(d => {
                        // If already at midnight (00:00:00), keep it on that day
                        if (d.Timestamp.Hour == 0 && d.Timestamp.Minute == 0 && d.Timestamp.Second == 0)
                        {
                            return d.Timestamp.Date;
                        }
                        else
                        {
                            return d.Timestamp.Date.AddDays(1);
                        }
                    });

                case "1 Week":
                    return sortedData.GroupBy(d => {
                        // If already exactly on Sunday at 00:00:00, keep it on that week
                        if (d.Timestamp.DayOfWeek == DayOfWeek.Sunday &&
                            d.Timestamp.Hour == 0 && d.Timestamp.Minute == 0 && d.Timestamp.Second == 0)
                        {
                            return d.Timestamp.Date;
                        }
                        else
                        {
                            // Calculate days until next Sunday
                            int daysUntilNextSunday = ((int)DayOfWeek.Sunday - (int)d.Timestamp.DayOfWeek + 7) % 7;
                            if (daysUntilNextSunday == 0) daysUntilNextSunday = 7;  // If already Sunday but not at 00:00:00
                            return d.Timestamp.Date.AddDays(daysUntilNextSunday);
                        }
                    });

                case "1 Month":
                    return sortedData.GroupBy(d => {
                        // If already on the 1st of the month at 00:00:00, keep it on that month
                        if (d.Timestamp.Day == 1 && d.Timestamp.Hour == 0 &&
                            d.Timestamp.Minute == 0 && d.Timestamp.Second == 0)
                        {
                            return new DateTime(d.Timestamp.Year, d.Timestamp.Month, 1);
                        }
                        else
                        {
                            // Otherwise go to first day of next month
                            return new DateTime(
                                d.Timestamp.Month == 12 ? d.Timestamp.Year + 1 : d.Timestamp.Year,
                                d.Timestamp.Month == 12 ? 1 : d.Timestamp.Month + 1,
                                1);
                        }
                    });

                case "Quarter Year":
                    return sortedData.GroupBy(d => {
                        int currentQuarter = (d.Timestamp.Month - 1) / 3;
                        // If already on the first day of the quarter at 00:00:00
                        if (d.Timestamp.Month == currentQuarter * 3 + 1 && d.Timestamp.Day == 1 &&
                            d.Timestamp.Hour == 0 && d.Timestamp.Minute == 0 && d.Timestamp.Second == 0)
                        {
                            return new DateTime(d.Timestamp.Year, currentQuarter * 3 + 1, 1);
                        }
                        else
                        {
                            // Otherwise go to first day of next quarter
                            int nextQuarterFirstMonth = (currentQuarter + 1) % 4 * 3 + 1;
                            int yearAdjustment = (currentQuarter == 3) ? 1 : 0; // If Q4, move to Q1 of next year

                            return new DateTime(d.Timestamp.Year + yearAdjustment, nextQuarterFirstMonth, 1);
                        }
                    });

                default:
                    // Default to hourly with special case handling
                    return sortedData.GroupBy(d => {
                        if (d.Timestamp.Minute == 0 && d.Timestamp.Second == 0)
                        {
                            return new DateTime(
                                d.Timestamp.Year,
                                d.Timestamp.Month,
                                d.Timestamp.Day,
                                d.Timestamp.Hour,
                                0, 0);
                        }
                        else
                        {
                            return new DateTime(
                                d.Timestamp.Year,
                                d.Timestamp.Month,
                                d.Timestamp.Day,
                                d.Timestamp.Hour,
                                0, 0).AddHours(1);
                        }
                    });
            }
        }

        private List<DynamicDataPoint> ConvertData(List<DynamicDataPoint> data, string inputInterval, string outputInterval)
        {
            var convertedData = new List<DynamicDataPoint>();
            var aggregationMethod = ((ComboBoxItem)cboAggregation.SelectedItem).Content.ToString();

            // Check if we have any data to convert
            if (data.Count == 0)
            {
                MessageBox.Show("No data found or could not parse dates. Please check your date format selection.",
                    "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
                return convertedData;
            }

            try
            {
                // Group the data by the appropriate time interval
                var groupedData = GroupDataByInterval(data, outputInterval);
                int groupCount = 0;

                foreach (var group in groupedData)
                {
                    groupCount++;

                    // Skip empty groups
                    if (!group.Any())
                        continue;

                    // Create a new data point for this group with the timestamp from the group key
                    var newPoint = new DynamicDataPoint { Timestamp = group.Key };

                    // Get all unique column names from all data points in this group
                    var allColumnNames = new HashSet<string>();
                    foreach (var dataPoint in group)
                    {
                        foreach (var columnName in dataPoint.Values.Keys)
                        {
                            allColumnNames.Add(columnName);
                        }
                    }

                    // Apply the selected aggregation method for each column
                    foreach (var columnName in allColumnNames)
                    {
                        // Get all values for this column in the current group
                        var columnValues = group
                            .Where(dp => dp.Values.ContainsKey(columnName))
                            .Select(dp => dp.Values[columnName])
                            .ToList();

                        if (columnValues.Count > 0)
                        {
                            // Apply the selected aggregation method
                            switch (aggregationMethod)
                            {
                                case "Average":
                                    newPoint.Values[columnName] = columnValues.Average();
                                    break;
                                case "Sum":
                                    newPoint.Values[columnName] = columnValues.Sum();
                                    break;
                                case "Max":
                                    newPoint.Values[columnName] = columnValues.Max();
                                    break;
                                case "Min":
                                    newPoint.Values[columnName] = columnValues.Min();
                                    break;
                            }
                        }
                    }

                    convertedData.Add(newPoint);
                }

                // Sort by timestamp and return
                return convertedData.OrderBy(d => d.Timestamp).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during data conversion: {ex.Message}\n\nStack trace: {ex.StackTrace}",
                     "Conversion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<DynamicDataPoint>();
            }
        }

        private void DisplayResults(List<DynamicDataPoint> convertedData)
        {
            // Create a DataTable to hold the results
            resultsTable = new DataTable();
            resultsTable.Columns.Add("Timestamp", typeof(DateTime));

            // Add a column for each data value
            if (convertedData.Count > 0)
            {
                // Get all unique column names from all data points
                var columnNames = new HashSet<string>();
                foreach (var dataPoint in convertedData)
                {
                    foreach (var columnName in dataPoint.Values.Keys)
                    {
                        if (!columnNames.Contains(columnName))
                        {
                            columnNames.Add(columnName);
                            resultsTable.Columns.Add($"{columnName} [kW]", typeof(double));
                        }
                    }
                }

                // Add rows to the table
                foreach (var point in convertedData)
                {
                    DataRow row = resultsTable.NewRow();
                    row["Timestamp"] = point.Timestamp;

                    // Add each value from this data point
                    foreach (var columnName in columnNames)
                    {
                        if (point.Values.ContainsKey(columnName))
                        {
                            row[$"{columnName} [kW]"] = Math.Round(point.Values[columnName], 2);
                        }
                        else
                        {
                            // If this point doesn't have a value for this column, set to 0 or null
                            row[$"{columnName} [kW]"] = DBNull.Value;
                        }
                    }

                    resultsTable.Rows.Add(row);
                }
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

                        // Add headers - first column is always timestamp
                        worksheet.Cells[1, 1].Value = "Timestamp";

                        // Add all other column headers
                        for (int col = 1; col < resultsTable.Columns.Count; col++)
                        {
                            worksheet.Cells[1, col + 1].Value = resultsTable.Columns[col].ColumnName;
                        }

                        // Add data
                        for (int i = 0; i < resultsTable.Rows.Count; i++)
                        {
                            var row = resultsTable.Rows[i];

                            // Add timestamp
                            worksheet.Cells[i + 2, 1].Value = (DateTime)row["Timestamp"];
                            worksheet.Cells[i + 2, 1].Style.Numberformat.Format = GetDateFormatForExport();

                            // Add all other columns
                            for (int col = 1; col < resultsTable.Columns.Count; col++)
                            {
                                var colName = resultsTable.Columns[col].ColumnName;
                                if (row[colName] != DBNull.Value)
                                {
                                    worksheet.Cells[i + 2, col + 1].Value = (double)row[colName];
                                }
                            }
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
}