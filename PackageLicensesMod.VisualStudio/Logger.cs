using Microsoft.VisualStudio.Shell;
using NuGet.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PackageLicensesMod.VisualStudio
{
	internal class Logger : ILogger
	{
		private static Logger _instance;

		public static Logger Instance
		{
			get
			{
				if (_instance == null)
					_instance = new Logger();

				return _instance;
			}
		}

		public Action<string> Write { get; set; }

		private Logger()
		{
		}

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

		System.Threading.Tasks.Task ILogger.LogAsync(LogLevel level, string data)
		{
			throw new NotImplementedException();
		}

		System.Threading.Tasks.Task ILogger.LogAsync(ILogMessage message)
		{
			throw new NotImplementedException();
		}
	}
}
