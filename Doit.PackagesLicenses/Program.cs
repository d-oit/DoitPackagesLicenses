using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using NuGet.Common;
using NuGet.Protocol;
using PackageLicenses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Doit.PackagesLicenses
{
    internal class Logger : ILogger
    {
        public void Log(LogLevel level, string data) => $"{level.ToString().ToUpper()}: {data}".Dump();

        public void Log(ILogMessage message) => Task.FromResult(0);

        public Task LogAsync(LogLevel level, string data) => Task.FromResult(0);

        public Task LogAsync(ILogMessage message) => throw new NotImplementedException();

        public void LogDebug(string data) => $"DEBUG: {data}".Dump();

        public void LogError(string data) => $"ERROR: {data}".Dump();

        public void LogInformation(string data) => $"INFORMATION: {data}".Dump();

        public void LogInformationSummary(string data) => $"SUMMARY: {data}".Dump();

        public void LogMinimal(string data) => $"MINIMAL: {data}".Dump();

        public void LogVerbose(string data) => $"VERBOSE: {data}".Dump();

        public void LogWarning(string data) => $"WARNING: {data}".Dump();
    }

    internal static class LogExtension
    {
        public static void Dump(this string value) => Console.WriteLine(value);
    }

    internal class Program
    {
        private const int ERROR_PACKAGE_PATH = 0xA0;
        private static IConfigurationRoot _configuration = null;
        private static string _outputPath = "";

        private static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
              .SetBasePath(Directory.GetCurrentDirectory())
              .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            string env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");

            if (string.IsNullOrWhiteSpace(env))
            {
                env = "Development";
            }

            if (env == "Development")
            {
                builder.AddUserSecrets<Program>();
            }
            _configuration = builder.Build();

            Console.Write("Change product and company name in the applicationSettings.json");
            Console.WriteLine("----");
            var path = _configuration["Path"];

            if (!Directory.Exists(path))
            {
                Console.Write("Path not Found: " + path);
                Environment.ExitCode = ERROR_PACKAGE_PATH;
                return;
            }

            _outputPath = _configuration["OutputPath"];

            var log = new Logger();

            // GitHub Client ID and Client Secret
            const string LicenseUtilityClientId = "LicenseUtility.ClientId";
            // Add user-secretes with command line:  dotnet user-secrets set LicenseUtility.ClientId XXXX-XX-ID...
            Console.WriteLine($"The Secret Id is {_configuration[LicenseUtilityClientId]}");

            LicenseUtility.ClientId = _configuration[LicenseUtilityClientId];
            const string LicenseUtilityClientSecret = "LicenseUtility.ClientSecret";
            // Add user-secretes with command line:  dotnet user-secrets set LicenseUtility.ClientSecret XYZ...
            LicenseUtility.ClientSecret = _configuration[LicenseUtilityClientSecret];
            Console.WriteLine($"The Client secret is {_configuration[LicenseUtilityClientId]}");

            var packages = PackageLicensesUtility.GetPackages(path, log);
            var list = new List<(LocalPackageInfo, License)>();
            var t = Task.Run(async () =>
            {
                foreach (var p in packages)
                {
                    Console.WriteLine($"{p.Nuspec.GetId()}.{p.Nuspec.GetVersion()}");
                    if (p.Nuspec.GetAuthors().StartsWith("Microsoft"))
                    {
                        Console.WriteLine("Ignore Microsoft");
                        continue;
                    }

                    if (p.Nuspec.GetAuthors().StartsWith("Jetbrain"))
                    {
                        Console.WriteLine("Ignore Jetbrain");
                        continue;
                    }

                    if (p.Nuspec.GetAuthors().StartsWith("xunit"))
                    {
                        Console.WriteLine("Ignore xUnit.net [Testing Framework]");
                        continue;
                    }

                    var license = await p.GetLicenseAsync(log);
                    list.Add((p, license));
                }
            });
            t.Wait();

            System.IO.DirectoryInfo di = new DirectoryInfo(_outputPath);
            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }

            try
            {
                CreateWorkbook(list);
                CreateTextLicensesFile(list);
            }
            catch (Exception)
            {
                throw;
            }

            Console.WriteLine("Completed.");
        }

        private static void CreateTextLicensesFile(List<(LocalPackageInfo, License)> list)
        {
            // Create a file to write to.
            using (StreamWriter sw = File.CreateText(_outputPath + "LICENSE-3RD-PARTY.txt"))
            {
                if (_configuration["ExportProductAndCompanyName"] == "show")
                {
                    sw.WriteLine("=================================");
                    sw.WriteLine(_configuration["ProductName"]);
                    sw.WriteLine(_configuration["Company"]);
                    sw.WriteLine("=================================");
                }
                sw.WriteLine(string.Format("Lists of {0} third-party dependencies.", list.Count));
                sw.WriteLine("");

                foreach (var (package, license) in list)
                {
                    var nuspec = package.Nuspec;
                    var title = nuspec.GetTitle() ?? "";
                    sw.WriteLine(string.Format("{0} - Author: {1} - Version: {2} - License Url: {3}", string.IsNullOrEmpty(title) ? nuspec.GetId() : nuspec.GetTitle(), nuspec.GetAuthors() ?? "", nuspec.GetVersion(), nuspec.GetLicenseUrl() ?? ""));
                }
            }
        }

        private static void CreateWorkbook(List<(LocalPackageInfo, License)> list)
        {
            var book = new XLWorkbook();
            var sheet = book.Worksheets.Add("Packages");

            // header
            var headers = new[] { "Id", "Version", "Authors", "Title", "ProjectUrl", "LicenseUrl", "RequireLicenseAcceptance", "Copyright", "Inferred License ID", "Inferred License Name" };
            for (var i = 0; i < headers.Length; i++)
            {
                sheet.Cell(1, 1 + i).SetValue(headers[i]).Style.Font.SetBold();
            }

            // values
            var row = 2;
            foreach (var (p, l) in list)
            {
                var nuspec = p.Nuspec;

                sheet.Cell(row, 1).SetValue(nuspec.GetId() ?? "");
                sheet.Cell(row, 2).SetValue($"{nuspec.GetVersion()}");
                sheet.Cell(row, 3).SetValue(nuspec.GetAuthors() ?? "");
                sheet.Cell(row, 4).SetValue(nuspec.GetTitle() ?? "");
                sheet.Cell(row, 5).SetValue(nuspec.GetProjectUrl() ?? "");
                sheet.Cell(row, 6).SetValue(nuspec.GetLicenseUrl() ?? "");
                sheet.Cell(row, 7).SetValue($"{nuspec.GetRequireLicenseAcceptance()}");
                sheet.Cell(row, 8).SetValue(nuspec.GetCopyright() ?? "");
                sheet.Cell(row, 9).SetValue(l?.Id ?? "");
                sheet.Cell(row, 10).SetValue(l?.Name ?? "");

                // save license text file
                if (!string.IsNullOrEmpty(l?.Text))
                {
                    string filename;
                    if (l.IsMaster)
                    {
                        filename = $"{l.Id}.txt";
                    }
                    else
                    {
                        if (l.DownloadUri != null)
                            filename = l.DownloadUri.PathAndQuery.Substring(1).Replace("/", "-").Replace("?", "-") + ".txt";
                        else
                            filename = $"{nuspec.GetId()}.{nuspec.GetVersion()}.txt";
                    }

                    File.WriteAllText(_outputPath + filename, l.Text, System.Text.Encoding.UTF8);

                    // set filename to cell
                    sheet.Cell(row, 11).SetValue(filename);
                    ++row;
                }
            }

            book.SaveAs(_outputPath + "Licenses.xlsx");
        }
    }
}