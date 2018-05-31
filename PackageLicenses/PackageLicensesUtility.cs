using ClosedXML.Excel;
using NuGet.Common;
using NuGet.Protocol;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using static PackageLicenses.SeparatedValuesWriter;

namespace PackageLicenses
{
	/// <summary>
	/// Nugetパッケージライセンスのユーティリティクラス
	/// </summary>
	public static class PackageLicensesUtility
	{
		/// <summary>
		/// Nugetパッケージ情報のリストを取得する
		/// </summary>
		/// <param name="packagesPath">Nugetパッケージのパス</param>
		/// <param name="log">logインターフェイス</param>
		/// <returns></returns>
		public static IEnumerable<LocalPackageInfo> GetPackages(string packagesPath, ILogger log = null)
		{
			var logger = log ?? NullLogger.Instance;
			var type = LocalFolderUtility.GetLocalFeedType(packagesPath, logger);
			switch (type)
			{
				case FeedType.FileSystemV2:
					return LocalFolderUtility.GetPackagesV2(packagesPath, logger);
				case FeedType.FileSystemV3:
					return LocalFolderUtility.GetPackagesV3(packagesPath, logger);
				default:
					break;
			}
			return new List<LocalPackageInfo>();
		}

		/// <summary>
		/// PackageReference形式のプロジェクトファイルからパッケージ情報のリストを取得
		/// </summary>
		/// <param name="fullpath">プロジェクトファイルのフルパス</param>
		/// <param name="log">log</param>
		/// <returns></returns>
		public static IEnumerable<LocalPackageInfo> GetGlobalPackages(string fullpath, ILogger log)
		{
			//Nugetグローバルパッケージのパス
			var pkg = GetGlobalPackagesFolder();

			XDocument xml = XDocument.Load(fullpath);
			var result = xml.Element("Project").Elements("ItemGroup")
				.Where(i => i.Element("PackageReference") != null).Select(i => i.Elements("PackageReference")).First();

			var packages = new List<LocalPackageInfo>();

			foreach (var item in result)
			{
				var inc = item.Attribute("Include");
				var ver = item.Attribute("Version");

				var nuspec = Path.Combine(pkg, inc.Value, ver.Value, inc.Value + "." + ver.Value + ".nupkg");

				packages.Add(LocalFolderUtility.GetPackage(new Uri(nuspec), log));
			}

			return packages;
		}

		/// <summary>
		/// Nugetグローバルパッケージのパスを取得する
		/// </summary>
		/// <returns></returns>
		private static string GetGlobalPackagesFolder()
		{
			/*全ユーザー 	%ProgramFiles(x86)%\NuGet\Config\***.config 	
			個別ユーザー 	%APPDATA%\NuGet\NuGet.Config*/
			//<add key="globalPackagesFolder" value="F:\.nuget\packages" />

			var pkg = @"%UserProfile%\.nuget\packages";
			return Environment.ExpandEnvironmentVariables(pkg);
		}

		/// <summary>
		/// Nugetのライセンス情報を取得する
		/// </summary>
		/// <param name="info">Nugetパッケージ情報</param>
		/// <param name="log">logインターフェイス</param>
		/// <returns></returns>
		public static async Task<License> GetLicenseAsync(this LocalPackageInfo info, ILogger log = null)
		{
			var licenseUrl = info.Nuspec.GetLicenseUrl();
			if (!string.IsNullOrWhiteSpace(licenseUrl) && Uri.IsWellFormedUriString(licenseUrl, UriKind.Absolute))
			{
				var license = await new Uri(licenseUrl).GetLicenseAsync(log);
				if (license != null) return license;
			}

			var projectUrl = info.Nuspec.GetProjectUrl();
			if (!string.IsNullOrWhiteSpace(projectUrl) && Uri.IsWellFormedUriString(projectUrl, UriKind.Absolute))
			{
				var license = await new Uri(projectUrl).GetLicenseAsync(log);
				if (license != null) return license;
			}
			return null;
		}

		/// <summary>
		/// ヘッダ文字列
		/// </summary>
		private static readonly List<string> _headers = new List<string>() { "Id", "Version", "Authors", "Title", "ProjectUrl", "LicenseUrl", "RequireLicenseAcceptance", "Copyright", "Inferred License ID", "Inferred License Name", "Downloaded license text file" };

		/// <summary>
		/// ライセンスファイルをダウンロードして各種情報をファイルに書き出す
		/// </summary>
		/// <param name="projectPath">プロジェクトファイルのパス</param>
		/// <param name="saveFolderPath">各種ファイルを出力するフォルダ</param>
		/// <param name="log">logインターフェイス</param>
		/// <returns></returns>
		public static async Task<bool> TryProjectPackageReferencesListAsync(string projectPath, string saveFolderPath, ILogger log)
		{
			if (System.IO.File.Exists(projectPath) == false)
			{
				log.LogError($"Not Found: '{projectPath}'\n");
				return false;
			}

			// Get GitHub Client ID and Client Secret
			SetGitHubClientInfo();

			// Get packages
			var packages = PackageLicensesUtility.GetGlobalPackages(projectPath, log);

			if (packages.Count() == 0)
			{
				log.LogWarning($"No Packages\n");
				return false;
			}

			// Output metadata and get licenses
			var list = await GetLicenses(packages, log);

			// Save to files
			await CreateFilesAsync(list, saveFolderPath, log);

			return true;
		}

		/// <summary>
		/// ライセンスファイルをダウンロードして各種情報をファイルに書き出す
		/// </summary>
		/// <param name="solutionPath">ソリューションのパス</param>
		/// <param name="saveFolderPath">各種ファイルを出力するフォルダ</param>
		/// <param name="log">logインターフェイス</param>
		/// <returns></returns>
		public static async Task<bool> TryPackagesListAsync(string solutionPath, string saveFolderPath, ILogger log)
		{
			var root = Path.Combine(solutionPath, "packages");
			if (!Directory.Exists(root))
			{
				log.LogError($"Not Found: '{root}'\n");
				return false;
			}
			log.LogInformation($"Packages Path: '{root}'\n");

			// Get GitHub Client ID and Client Secret
			SetGitHubClientInfo();

			// Get packages
			var packages = PackageLicensesUtility.GetPackages(root, log);

			if (packages.Count() == 0)
			{
				log.LogWarning($"No Packages\n");
				return false;
			}

			// Output metadata and get licenses
			var list = await GetLicenses(packages, log);

			// Save to files
			await CreateFilesAsync(list, saveFolderPath, log);

			return true;
		}

		private static async Task<List<(LocalPackageInfo, License)>> GetLicenses(IEnumerable<LocalPackageInfo> packages, ILogger log)
		{
			var headers = string.Join("\t", _headers.Take(_headers.Count - 1));
			var dividers = string.Join("\t", _headers.Take(_headers.Count - 1).Select(i => new string('-', i.Length)));
			log.LogInformation($"\n{headers}\n{dividers}\n");

			var list = new List<(LocalPackageInfo, License)>();
			foreach (var p in packages)
			{
				var nuspec = p.Nuspec;
				var license = await p.GetLicenseAsync(log);
				log.LogInformation($"{nuspec.GetId()}\t{nuspec.GetVersion()}\t{nuspec.GetAuthors()}\t{nuspec.GetTitle()}\t{nuspec.GetProjectUrl()}\t{nuspec.GetLicenseUrl()}\t{nuspec.GetRequireLicenseAcceptance()}\t{nuspec.GetCopyright()}\t{license?.Id}\t{license?.Name}\n");

				list.Add((p, license));
			}
			log.LogInformation($"\n");
			return list;
		}

		private static void SetGitHubClientInfo()
		{
			var query = Environment.GetEnvironmentVariable("PACKAGE-LICENSES-GITHUB-QUERY", EnvironmentVariableTarget.User);
			if (!string.IsNullOrWhiteSpace(query))
			{
				var m = Regex.Match(query, "client_id=(?<id>.*?)&client_secret=(?<secret>.*)");
				if (m.Success)
				{
					LicenseUtility.ClientId = m.Groups["id"].Value;
					LicenseUtility.ClientSecret = m.Groups["secret"].Value;
				}
			}
		}

		/// <summary>
		/// 各種情報のリストとライセンスファイルを書き出す
		/// </summary>
		/// <param name="list">Nugetパッケージ情報とライセンス情報のペア</param>
		/// <param name="saveFolderPath">各種ファイルを書き出すパス</param>
		private static async Task CreateFilesAsync(List<(LocalPackageInfo, License)> list, string saveFolderPath, ILogger log)
		{
			var book = new XLWorkbook();
			var sheet = book.Worksheets.Add("Packages");

			var path = Path.Combine(saveFolderPath, "Licenses.txt");
			using (var writer = new SeparatedValuesWriter(path, SeparatedValuesWriterSetting.Tsv))
			{
				// header
				for (var i = 0; i < _headers.Count; i++)
				{
					sheet.Cell(1, 1 + i).SetValue(_headers[i]).Style.Font.SetBold();
					await writer.WriteLineAsync(_headers);
				}
				await writer.FlushAsync();

				// values
				var row = 2;
				foreach (var (p, l) in list)
				{
					var nuspec = p.Nuspec;
					var filename = WriteLicenseFile(saveFolderPath, l, nuspec);

					var v = CreateValues(l, nuspec, filename);

					WriteValuesCell(sheet, row, v);
					await WriteValuesTxtAsync(writer, v);

					++row;
				}
				await writer.FlushAsync();
			}
			book.SaveAs(Path.Combine(saveFolderPath, "Licenses.xlsx"));
			log.LogInformation($"Saved to '{saveFolderPath}'\n");
		}

		private static IEnumerable<string> CreateValues(License l, NuGet.Packaging.NuspecReader nuspec, string filename)
		{
			var values = new[]
			{
				nuspec.GetId() ?? "",
				$"{nuspec.GetVersion()}",
				nuspec.GetAuthors() ?? "",
				nuspec.GetTitle() ?? "",
				nuspec.GetProjectUrl() ?? "",
				nuspec.GetLicenseUrl() ?? "",
				$"{nuspec.GetRequireLicenseAcceptance()}",
				nuspec.GetCopyright() ?? "",
				l?.Id ?? "",
				l?.Name ?? "",
				filename ?? "",
			};

			return values;
		}

		/// <summary>
		/// テキストファイルにライセンス情報を書き込む
		/// </summary>
		/// <param name="writer">writer</param>
		/// <param name="values">values</param>
		/// <returns></returns>
		private static async Task WriteValuesTxtAsync(SeparatedValuesWriter writer, IEnumerable<string> values)
		{
			await writer.WriteLineAsync(values);
		}

		/// <summary>
		/// セルにライセンス情報を書き込む
		/// </summary>
		/// <param name="sheet">sheet</param>
		/// <param name="row">行</param>
		/// <param name="values">values</param>
		private static void WriteValuesCell(IXLWorksheet sheet, int row, IEnumerable<string> values)
		{
			int c = 1;
			foreach (var v in values)
			{
				sheet.Cell(row, c).SetValue(values);
				c++;
			}
		}

		/// <summary>
		/// ライセンスファイルをダウンロード
		/// </summary>
		/// <param name="saveFolderPath">ダウンロードするフォルダパス</param>
		/// <param name="l">ライセンス情報</param>
		/// <param name="nuspec">nuspec</param>
		/// <returns>ファイル名</returns>
		private static string WriteLicenseFile(string saveFolderPath, License l, NuGet.Packaging.NuspecReader nuspec)
		{
			string filename = null;
			// save license text file
			if (!string.IsNullOrEmpty(l?.Text))
			{

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

				var path = Path.Combine(saveFolderPath, filename);
				if (!File.Exists(path))
					File.WriteAllText(path, l.Text, System.Text.Encoding.UTF8);
			}
			return filename;
		}

		/// <summary>
		/// プロジェクトファイルを検索
		/// </summary>
		private static IEnumerable<FileInfo> FindProjectFile(string FolderPass)
		{
			var di = new DirectoryInfo(FolderPass);

			return di.EnumerateFiles("*.csproj", SearchOption.AllDirectories);
		}

		/// <summary>
		/// ソリューションフォルダからプロジェクトを取得してプロジェクトごとのライセンスを取得する
		/// </summary>
		/// <param name="solutionPath">ソリューションフォルダ</param>
		/// <param name="saveFolderPath">各種ファイルを出力するフォルダ</param>
		/// <param name="log">log</param>
		/// <returns>
		/// プロジェクトごとの成否
		/// </returns>
		public static async Task<IEnumerable<bool>> TrySolutionPackageReferencesListAsync(string solutionPath, string saveFolderPath, ILogger log)
		{
			var files = FindProjectFile(solutionPath);
			var ret = new List<bool>();
			foreach (var f in files)
			{
				ret.Add(await TryProjectPackageReferencesListAsync(f.FullName, Path.Combine(saveFolderPath, f.Name), log));
			}
			return ret;
		}
	}
}
