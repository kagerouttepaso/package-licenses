using EnvDTE;
using Microsoft.VisualStudio.Shell;
using NuGet.Common;
using PackageLicenses;
using System;
using System.ComponentModel.Design;
using System.IO;
using Task = System.Threading.Tasks.Task;

namespace PackageLicensesMod.VisualStudio
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class ListCommand
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 4129;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("08842038-328f-49f1-a02b-52c4478d100b");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		private readonly MenuCommand _menuCommand;

		/// <summary>
		/// Initializes a new instance of the <see cref="ListCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private ListCommand(AsyncPackage package, OleMenuCommandService commandService)
		{
			this.package = package ?? throw new ArgumentNullException(nameof(package));
			commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

			var menuCommandID = new CommandID(CommandSet, CommandId);
			_menuCommand = new MenuCommand(this.ExecuteAsync, menuCommandID);
			commandService.AddCommand(_menuCommand);
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static ListCommand Instance
		{
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static async Task InitializeAsync(AsyncPackage package)
		{
			// Verify the current thread is the UI thread - the call to AddCommand in ListCommand's constructor requires
			// the UI thread.
			ThreadHelper.ThrowIfNotOnUIThread();

			OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
			Instance = new ListCommand(package, commandService);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private async void ExecuteAsync(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			var dte = (DTE)ServiceProvider.GetServiceAsync(typeof(DTE));
			var solutionDir = Path.GetDirectoryName(dte.Solution.FullName);

			try
			{
				_menuCommand.Enabled = false;

				// Create temporary folder
				var tempPath = Path.Combine(Path.GetTempPath(), $"{Path.GetFileName(solutionDir)}-{DateTime.Now.ToFileTimeUtc()}");
				Directory.CreateDirectory(tempPath);

				// List
				var result = await PackageLicensesUtility.TryPackagesListAsync(solutionDir, tempPath, Logger.Instance);

				// Open folder
				if (result)
					System.Diagnostics.Process.Start(tempPath);
				else
					Directory.Delete(tempPath, true);
			}
			catch (Exception ex)
			{
				ServiceProvider.WriteOnOutputWindow($"ERROR: {ex.Message}\n");
			}
			finally
			{
				_menuCommand.Enabled = true;
			}
		}
	}

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
	}
}
