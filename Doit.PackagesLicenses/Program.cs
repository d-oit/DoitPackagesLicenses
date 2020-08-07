using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using NuGet.Common;
using NuGet.Protocol;
using PackageLicenses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
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
					if (p.Nuspec.GetAuthors().ToLower().StartsWith("microsoft"))
					{
						Console.WriteLine("Ignore Microsoft");
						continue;
					}

					if (p.Nuspec.GetAuthors().ToLower().StartsWith("jetbrain"))
					{
						Console.WriteLine("Ignore Jetbrain");
						continue;
					}

					if (p.Nuspec.GetAuthors().ToLower().StartsWith("xunit"))
					{
						Console.WriteLine("Ignore xUnit.net [Testing Framework]");
						continue;
					}

					var license = await p.GetLicenseAsync(log);
					list.Add((p, license));
				}
			});
			t.Wait();

			try
			{
				CreateWorkbook(list);
			}
			catch (Exception)
			{
				throw;
			}

			Console.WriteLine("Completed.");
		}

		private static string GetTitle(NuGet.Packaging.NuspecReader nuspec)
		{
			if (nuspec.GetTitle()?.Length > 0)
				return nuspec.GetTitle();

			return nuspec.GetId();
		}

		private static string GetLicence(License license)
		{
			if (license?.Name?.Length > 0)
				return license.Name;

			if (license?.Text?.Length > 0)
				return ParseLicence(license.Text);

			return null;
		}

		private static string GetLicence(NuGet.Packaging.NuspecReader nuspec)
		{
			using var webClient = new WebClient();
			var url = nuspec.GetLicenseUrl();

			if (string.IsNullOrWhiteSpace(url))
				return null;

			try
			{
				var text = webClient.DownloadString(url);
				return ParseLicence(text);
			}
			catch
			{
				return null;
			}
		}

		private static string ParseLicence(string text)
		{
			if (text == null)
				return null;

			text = text.ToLower();

			if (text.Contains("apache licence"))
			{
				if (text.Contains("version 2.0"))
					return "Apache License 2.0";

				return "Apache";
			}

			if (text.Contains("mit license"))
				return "MIT License";

			if (text.Contains("unlicence"))
				return "The Unlicense";

			if (text.Contains("new bsd license"))
				return "New BSD License";

			if (text.Contains("gpl licence"))
				return "GPL";

			return null;
		}

		private static void CreateWorkbook(List<(LocalPackageInfo, License)> list)
		{
			var book = new XLWorkbook();
			var sheet = book.Worksheets.Add("Packages");

			// header
			var headers = new[] { "Title", "Licence", "LicenceUrl", "ProjectUrl" };
			for (var i = 0; i < headers.Length; i++)
			{
				sheet.Cell(1, 1 + i).SetValue(headers[i]).Style.Font.SetBold();
			}

			// values
			var row = 2;
			foreach (var (package, licence) in list)
			{
				var nuspec = package.Nuspec;

				sheet.Cell(row, 1).SetValue(GetTitle(nuspec));
				sheet.Cell(row, 2).SetValue(GetLicence(licence) ?? GetLicence(nuspec) ?? "[???]");
				sheet.Cell(row, 3).SetValue(nuspec.GetLicenseUrl() ?? "");
				sheet.Cell(row, 4).SetValue(nuspec.GetProjectUrl() ?? "");

				row++;
			}

			var filePath = _outputPath + "Licenses.xlsx";

			if (File.Exists(filePath))
				File.Delete(filePath);

			book.SaveAs(filePath);
		}
	}
}