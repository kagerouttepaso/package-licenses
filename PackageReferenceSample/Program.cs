using NuGet.Common;
using NuGet.Protocol;
using PackageLicenses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Linq;
using NuGet.Versioning;
using NuGet.Packaging.Core;

namespace PackageReferenceSample
{
	class Logger : ILogger
	{
		public Action<string> Write { get; set; } = Console.WriteLine;

		public void Log(LogLevel level, string data) { }
		public void Log(ILogMessage message) { }
		public System.Threading.Tasks.Task LogAsync(LogLevel level, string data) => System.Threading.Tasks.Task.FromResult(0);
		public System.Threading.Tasks.Task LogAsync(ILogMessage message) => System.Threading.Tasks.Task.FromResult(0);
		public void LogDebug(string data) => Write?.Invoke($"DEBUG: {data}\n");
		public void LogError(string data) => Write?.Invoke($"ERROR: {data}\n");
		public void LogInformation(string data) => Write?.Invoke($"INFORMATION: {data}\n");
		public void LogInformationSummary(string data) => Write?.Invoke($"SUMMARY: {data}\n");
		public void LogMinimal(string data) => Write?.Invoke($"MINIMAL: {data}\n");
		public void LogVerbose(string data) => Write?.Invoke($"VERBOSE: {data}\n");
		public void LogWarning(string data) => Write?.Invoke($"WARNING: {data}\n");
	}

	class Program
	{
		static void Main(string[] args)
		{

			const string samplefile = "PackageReferenceSample.csproj.xml";
			var dir = System.AppDomain.CurrentDomain.BaseDirectory;
			var samplepath = Path.Combine(dir, samplefile);

			// Create Output folder
			var savefolder = Path.Combine(dir, $"{Path.GetFileName("Licenses")}-{DateTime.Now.ToFileTimeUtc()}");
			Directory.CreateDirectory(savefolder);

			var package = args.Length > 1 ? args[1] : samplepath;
			var command = args.Length > 0 ? args[0] : "proj";

			switch (command)
			{
				case "sln":
					SolutionPackageReferencesList(package, savefolder);
					break;

				case "proj":
					ProjectPackageReferences(package, savefolder);
					break;
			}

			Console.WriteLine("Completed.");
#if DEBUG
			Console.ReadKey();
#endif
		}

		private static void ProjectPackageReferences(string project, string savefolder)
		{
			var log = new Logger();
			var result = PackageLicensesUtility.TryProjectPackageReferencesListAsync(project, savefolder, log);

			if (result.Result == false)
				Directory.Delete(savefolder, true);
		}

		private static void SolutionPackageReferencesList(string solutionPath, string saveFolderPath)
		{
			var log = new Logger();
			var result = PackageLicensesUtility.TrySolutionPackageReferencesListAsync(solutionPath, saveFolderPath, log);
			result.Wait();
		}
	}
}
