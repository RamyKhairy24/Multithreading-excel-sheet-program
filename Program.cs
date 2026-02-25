using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Windows.Forms;

namespace MultithreadingPhonenumberChecker
{
    internal class Program
    {
        static readonly ConcurrentQueue<string> logBuffer = new ConcurrentQueue<string>();
        static string logFilePath = "scan.log";

        static void LogBuffered(string message)
        {
            logBuffer.Enqueue($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}");
        }

        static void LogImmediate(string message)
        {
            Console.WriteLine(message);
            File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}{Environment.NewLine}");
        }

        static (bool IsValid, List<string> FailedCriteria) ValidatePhoneNumber(string number)
        {
            var failed = new List<string>();
            if (string.IsNullOrWhiteSpace(number))
            {
                failed.Add("Null or empty");
                return (false, failed);
            }

            string n = number.Trim().Replace(" ", "");

            // Egyptian mobile prefixes
            string egyptPrefixes = "(10|11|12|15)";

            // Regex patterns for each format
            var patterns = new[]
            {
                // 0020XXXXXXXXXX (Egypt)
                $@"^0020{egyptPrefixes}\d{{8}}$",
                // +20XXXXXXXXXX (Egypt)
                $@"^\+20{egyptPrefixes}\d{{8}}$",
                // 20XXXXXXXXXX (Egypt)
                $@"^20{egyptPrefixes}\d{{8}}$",
                // 0XXXXXXXXXX (Egypt)
                $@"^0{egyptPrefixes}\d{{8}}$",
                // XXXXXXXXXX (Egypt, 10 digits)
                $@"^{egyptPrefixes}\d{{8}}$",
                // 00[1-9]XXXXXXX... (International, 8-17 digits)
                @"^00[1-9]\d{5,14}$",
                // +[1-9]XXXXXXX... (International, 7-16 digits)
                @"^\+[1-9]\d{5,14}$"
            };

            bool isValid = patterns.Any(p => Regex.IsMatch(n, p));

            if (!isValid)
            {
                failed.Add("Does not match any allowed MSISDN format");
            }

            return (isValid, failed);
        }

        static List<(int Row, string Number)> ReadPhoneNumbersFromExcel(string filePath)
        {
            var phoneNumbers = new List<(int, string)>();
            LogBuffered("Opening workbook with ClosedXML...");
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); 
                LogBuffered("Reading rows...");
                foreach (var row in rows)
                {
                    var cell = row.Cell(1).GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(cell))
                        phoneNumbers.Add((row.RowNumber(), cell));
                }
                LogBuffered("Finished reading rows.");
            }
            return phoneNumbers;
        }

        [STAThread]
        static void Main(string[] args)
        {
            var stopwatch = Stopwatch.StartNew();

            File.WriteAllText(logFilePath, ""); 

            string excelFilePath = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Select the Excel file to scan"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = openFileDialog.FileName;
            }
            else
            {
                LogImmediate("No file selected. Exiting.");
                return;
            }

            if (!File.Exists(excelFilePath))
            {
                LogImmediate($"File not found: {excelFilePath}");
                return;
            }

            List<(int Row, string Number)> phoneNumbers;
            try
            {
                LogBuffered($"Starting scan of file: {excelFilePath}");
                phoneNumbers = ReadPhoneNumbersFromExcel(excelFilePath);
                LogBuffered($"Loaded {phoneNumbers.Count} phone numbers from Excel.");
            }
            catch (IOException ioEx)
            {
                LogImmediate($"Blocked from scanning. IO error: {ioEx.Message}");
                return;
            }
            catch (Exception ex)
            {
                LogImmediate("Error reading Excel file: " + ex.Message);
                return;
            }

            var invalidNumbers = new ConcurrentBag<(int Row, string Number, List<string> FailedCriteria)>();
            ParallelOptions options = new ParallelOptions { MaxDegreeOfParallelism = 16 };

            Parallel.ForEach(phoneNumbers, options, entry =>
            {
                LogBuffered($"Scanned row {entry.Row}: {entry.Number}");
                var validation = ValidatePhoneNumber(entry.Number);
                if (!validation.IsValid)
                {
                    string criteria = string.Join("; ", validation.FailedCriteria);
                    LogBuffered($"Invalid number at row {entry.Row}: {entry.Number} | Failed criteria: {criteria}");
                    invalidNumbers.Add((entry.Row, entry.Number, validation.FailedCriteria));
                }
            });

            if (invalidNumbers.Any())
            {
                LogBuffered("failed file");
                LogBuffered($"Total invalid numbers: {invalidNumbers.Count}");
            }
            else
            {
                LogBuffered("succeeded file");
            }

            stopwatch.Stop();
            LogBuffered($"Total time taken: {stopwatch.Elapsed}");
            LogBuffered("Scan complete.");

            File.WriteAllLines(logFilePath, logBuffer);
            foreach (var line in logBuffer)
                Console.WriteLine(line);

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
