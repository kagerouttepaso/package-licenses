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

    static class LogExtension
    {
        public static void Dump(this string value) => Console.WriteLine(value);
    }

    class Program
    {
        static void Main(string[] args)
        {
            var log = new Logger();
            const string samplefile = "PackageReferenceSample.csproj.xml";
            var dir = System.AppDomain.CurrentDomain.BaseDirectory;
            var fullpath = Path.Combine(dir, samplefile);

            var packages = GetPackages(fullpath, log);

            var list = new List<(LocalPackageInfo, License)>();
            var t = Task.Run(async () =>
            {
                foreach (var p in packages)
                {
                    var license = await p.GetLicenseAsync(log);
                    list.Add((p, license));
                }
            });
            t.Wait();

            DispData(list);
            Console.WriteLine("Completed.");
#if DEBUG
            Console.ReadKey();
#endif
        }

        private static List<LocalPackageInfo> GetPackages(string fullpath, Logger log)
        {


            var pkg = @"%UserProfile%\.nuget\packages";

            XDocument xml = XDocument.Load(fullpath);
            var result = xml.Element("Project").Elements("ItemGroup")
                .Where(i => i.Element("PackageReference") != null).Select(i => i.Elements("PackageReference")).First();

            var packages = new List<LocalPackageInfo>();

            foreach (var item in result)
            {
                var inc = item.Attribute("Include");
                var ver = item.Attribute("Version");

                var nuspec = Environment.ExpandEnvironmentVariables(Path.Combine(pkg, inc.Value, ver.Value, inc.Value + "." + ver.Value + ".nupkg"));

                packages.Add(LocalFolderUtility.GetPackage(new Uri(nuspec), log));
            }

            return packages;
        }

        private static void DispData(List<(LocalPackageInfo, License)> list)
        {
            // header
            var headers = new[] { "Id", "Version", "Authors", "Title", "ProjectUrl", "LicenseUrl", "RequireLicenseAcceptance", "Copyright", "Inferred License ID", "Inferred License Name" };
            for (var i = 0; i < headers.Length; i++)
            {
                Console.WriteLine(String.Join('\t',headers));
            }

            // values
            foreach (var (p, l) in list)
            {
                var nuspec = p.Nuspec;
                var text = new [] {
                    nuspec.GetId() ?? "",
                    $"{nuspec.GetVersion()}",
                    nuspec.GetAuthors() ?? "",
                    nuspec.GetTitle() ?? "",
                    nuspec.GetProjectUrl() ?? "",
                    nuspec.GetLicenseUrl() ?? "",
                    $"{nuspec.GetRequireLicenseAcceptance()}",
                    nuspec.GetCopyright() ?? "",
                    l?.Id ?? "",
                    l?.Name ?? ""};

                Console.WriteLine(String.Join('\t',text));
            }
        }
    }
}
