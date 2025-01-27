using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace CtdConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            // ----------------------------------------------------------------
            // 1) Validate and parse input
            // ----------------------------------------------------------------
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: DomainChecker <path-to-excel-file>");
                Console.WriteLine("Example: DomainChecker C:\\domains.xlsx");
                return;
            }

            string filePath = args[0];

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found at '{filePath}'.");
                return;
            }

            // ----------------------------------------------------------------
            // 2) Read Excel file and parse domain names and extensions
            // ----------------------------------------------------------------
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // For EPPlus in .NET 5+

            var domainCheckResults = new Dictionary<string, Dictionary<string, string>>();
            string[] extensions;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    Console.WriteLine("Error: No worksheet found in the Excel file.");
                    return;
                }

                // Read extensions from the first row (excluding the first column)
                extensions = worksheet.Cells[1, 2, 1, worksheet.Dimension.End.Column]
                                      .Select(cell => cell.Text.Trim().ToLower())
                                      .Where(ext => !string.IsNullOrWhiteSpace(ext))
                                      .ToArray();

                if (extensions.Length == 0)
                {
                    Console.WriteLine("No valid extensions found in the Excel file.");
                    return;
                }

                // Read domain names and existing statuses
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    string domain = worksheet.Cells[row, 1].Text.Trim();
                    if (string.IsNullOrWhiteSpace(domain)) continue;

                    domainCheckResults[domain] = new Dictionary<string, string>();

                    for (int col = 2; col <= worksheet.Dimension.End.Column; col++)
                    {
                        string ext = extensions[col - 2];
                        string status = worksheet.Cells[row, col].Text.Trim();
                        domainCheckResults[domain][ext] = status;
                    }
                }
            }

            // ----------------------------------------------------------------
            // 3) For each domain, check each extension if not already checked.
            // ----------------------------------------------------------------
            foreach (var domain in domainCheckResults.Keys.ToList())
            {
                foreach (var ext in extensions)
                {
                    if (string.IsNullOrWhiteSpace(domainCheckResults[domain][ext]))
                    {
                        string fullDomain = $"{domain}.{ext}";
                        bool isAvailable = CheckDomainAvailability(fullDomain);
                        domainCheckResults[domain][ext] = isAvailable ? "Not Registered" : "Registered";
                        Console.WriteLine($"[{(isAvailable ? "NOT REGISTERED" : "REGISTERED")}] {fullDomain}");
                    }
                }
            }

            // ----------------------------------------------------------------
            // 4) Update Excel file with the results
            // ----------------------------------------------------------------
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    Console.WriteLine("Error: No worksheet found in the Excel file.");
                    return;
                }

                int row = 2;
                var lastRow = worksheet.Dimension.End.Row;
                for(var rowIndex = row;rowIndex<=lastRow;rowIndex++)
                {
                    var domain = worksheet.Cells[rowIndex, 1].Text.Trim();
                    for (int col = 2; col <= extensions.Length + 1; col++)
                    {
                        string ext = extensions[col - 2];
                        worksheet.Cells[rowIndex, col].Value = domainCheckResults[domain][ext];
                    }
                }

                try
                {
                    package.Save();
                }
                catch (Exception ex)
                {
                    package.Save();
                }

                Console.WriteLine($"Excel file updated: {Path.GetFullPath(filePath)}");
            }
        }

        private static bool CheckDomainAvailability(string fullDomain)
        {
            Random random = new Random();
            int delay = random.Next(30, 501);
            System.Threading.Thread.Sleep(delay);

            try
            {
                Dns.GetHostEntry(fullDomain);
                return false;
            }
            catch (SocketException)
            {
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERROR CHECKING] {fullDomain}: {ex.Message}");
                return false;
            }
        }
    }
}
