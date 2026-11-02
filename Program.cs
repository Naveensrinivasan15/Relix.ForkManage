using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace RelixQ.ForkManager;

public class ReportRecord
{
    public string Repository_Name { get; set; } = string.Empty;
    public string Forked_Repository { get; set; } = string.Empty;
    public string Upstream_URL { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public string Language { get; set; } = string.Empty;
    public string License { get; set; } = string.Empty;
    public int Stars { get; set; }
    public string Default_Branch { get; set; } = string.Empty;
    public string Upstream_Commit_SHA { get; set; } = string.Empty;
    public string Fork_Status { get; set; } = string.Empty;
    public string Fork_URL { get; set; } = string.Empty;
}

public class Program
{
    private const string TargetOrg = "relixq-oss-report";
    private const string CsvFilePath = "RelixQ_OSS_Final_Report.csv";
    private const string ExcelFilePath = "RelixQ_OSS_Final_Report.xlsx";
    private const int TargetTotal = 1000;

    private static readonly string GitHubToken = Environment.GetEnvironmentVariable("GITHUB_TOKEN") ?? "";
    private static readonly HttpClient HttpClient = new();

    public static async Task Main(string[] args)
    {
        Console.WriteLine("=================================================");
        Console.WriteLine(" RelixQ OSS Harvester & Batch Forking Engine ");
        Console.WriteLine("=================================================");

        if (string.IsNullOrWhiteSpace(GitHubToken))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("ERROR: GITHUB_TOKEN environment variable is not set.");
            Console.ResetColor();
            return;
        }

        HttpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("RelixQ-OSS-Harvester", "1.0"));
        HttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", GitHubToken);
        HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        HttpClient.DefaultRequestHeaders.Add("X-GitHub-Api-Version", "2022-11-28");

        // 1. Load existing records
        var records = LoadExistingRecords(CsvFilePath);
        var seenUpstream = new HashSet<string>(records.Select(r => r.Repository_Name.Trim()), StringComparer.OrdinalIgnoreCase);
        var seenOrgRepos = new HashSet<string>(records.Select(r => r.Forked_Repository.Trim()), StringComparer.OrdinalIgnoreCase);

        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine($"[Init] Loaded {records.Count} existing verified repositories from {CsvFilePath}.");
        Console.WriteLine($"[Init] Remaining repositories needed: {TargetTotal - records.Count}\n");
        Console.ResetColor();

        // 2. Category Quotas & Search Seeds
        var categoryConfigs = new (string Category, int TargetQuota, string[] Queries)[]
        {
            ("Cryptography and TLS", 120, new[] {
                "topic:cryptography stars:>500",
                "topic:tls stars:>300",
                "topic:post-quantum stars:>15",
                "topic:pqc stars:>10",
                "topic:encryption stars:>1000",
                "topic:cipher stars:>500",
                "topic:ssl stars:>500"
            }),
            ("Identity and Authentication", 80, new[] {
                "topic:authentication stars:>1200",
                "topic:oauth2 stars:>1000",
                "topic:identity stars:>800",
                "topic:iam stars:>500",
                "topic:sso stars:>500",
                "topic:jwt stars:>1500"
            }),
            ("Cloud and Kubernetes", 150, new[] {
                "topic:kubernetes stars:>3000",
                "topic:cloud-native stars:>1500",
                "topic:service-mesh stars:>800",
                "topic:container-runtime stars:>500",
                "topic:cncf stars:>1000"
            }),
            ("Databases", 120, new[] {
                "topic:database stars:>3000",
                "topic:distributed-database stars:>1000",
                "topic:key-value-store stars:>800",
                "topic:sql stars:>2000",
                "topic:nosql stars:>1500"
            }),
            ("Operating Systems, Kernels & Runtimes", 100, new[] {
                "topic:operating-system stars:>1500",
                "topic:kernel stars:>1000",
                "topic:runtime stars:>2000",
                "topic:compiler stars:>2000",
                "topic:embedded stars:>1000"
            }),
            ("Networking, Proxies & Web Servers", 100, new[] {
                "topic:reverse-proxy stars:>1000",
                "topic:http-server stars:>1500",
                "topic:networking stars:>2000",
                "topic:load-balancer stars:>1000",
                "topic:vpn stars:>1500"
            }),
            ("Web Frameworks", 130, new[] {
                "topic:web-framework stars:>4000",
                "topic:backend-framework stars:>2500",
                "topic:api-framework stars:>2000"
            }),
            ("Package Managers and Build Tools", 100, new[] {
                "topic:package-manager stars:>1500",
                "topic:build-tool stars:>1500",
                "topic:bundler stars:>1000",
                "topic:monorepo stars:>1000"
            }),
            ("Blockchain", 50, new[] {
                "topic:blockchain stars:>2500",
                "topic:smart-contracts stars:>1500",
                "topic:zero-knowledge stars:>500"
            }),
            ("Observability, Tracing & Infrastructure", 50, new[] {
                "topic:monitoring stars:>2000",
                "topic:observability stars:>1000",
                "topic:tracing stars:>800",
                "topic:metrics stars:>1500",
                "topic:telemetry stars:>500"
            })
        };

        foreach (var (category, targetQuota, queries) in categoryConfigs)
        {
            if (records.Count >= TargetTotal) break;

            int currentCount = records.Count(r => r.Category.Equals(category, StringComparison.OrdinalIgnoreCase));
            int neededForCategory = targetQuota - currentCount;

            Console.WriteLine($"--------------------------------------------------");
            Console.WriteLine($"Category: {category}");
            Console.WriteLine($"Current: {currentCount} | Quota: {targetQuota} | Needed: {Math.Max(0, neededForCategory)}");
            Console.WriteLine($"--------------------------------------------------");

            if (neededForCategory <= 0)
            {
                Console.WriteLine("Quota met. Moving to next category.\n");
                continue;
            }

            foreach (var query in queries)
            {
                if (records.Count >= TargetTotal) break;
                if (records.Count(r => r.Category.Equals(category, StringComparison.OrdinalIgnoreCase)) >= targetQuota) break;

                Console.WriteLine($"Querying: {query}...");
                var candidates = await SearchRepositoriesAsync(query);

                foreach (var candidate in candidates)
                {
                    if (records.Count >= TargetTotal) break;
                    if (records.Count(r => r.Category.Equals(category, StringComparison.OrdinalIgnoreCase)) >= targetQuota) break;

                    if (seenUpstream.Contains(candidate.FullName)) continue;

                    var targetForkName = $"{TargetOrg}/{candidate.Name}";
                    if (seenOrgRepos.Contains(targetForkName)) continue;

                    if (string.IsNullOrWhiteSpace(candidate.LicenseSpdx) || candidate.LicenseSpdx is "NOASSERTION" or "NONE")
                        continue;

                    Console.WriteLine($"[{records.Count + 1}/{TargetTotal}] Processing: {candidate.FullName} ({candidate.Stars} stars, {candidate.Language})...");

                    var commitSha = await GetHeadCommitShaAsync(candidate.FullName, candidate.DefaultBranch);

                    // Fork repository with retry loop on secondary rate limits
                    var (forkUrl, forkStatus) = await ForkRepositoryAsync(candidate.FullName);

                    if (forkStatus != "Verified")
                    {
                        Console.WriteLine($"Skipping {candidate.FullName}: {forkStatus}");
                        continue;
                    }

                    var newRecord = new ReportRecord
                    {
                        Repository_Name = candidate.FullName,
                        Forked_Repository = targetForkName,
                        Upstream_URL = candidate.HtmlUrl,
                        Category = category,
                        Language = candidate.Language ?? "Unknown",
                        License = candidate.LicenseSpdx,
                        Stars = candidate.Stars,
                        Default_Branch = candidate.DefaultBranch,
                        Upstream_Commit_SHA = commitSha,
                        Fork_Status = "Verified",
                        Fork_URL = forkUrl
                    };

                    records.Add(newRecord);
                    seenUpstream.Add(candidate.FullName);
                    seenOrgRepos.Add(targetForkName);

                    // Save checkpoint to both CSV and Excel every 5 items
                    if (records.Count % 5 == 0)
                    {
                        SaveRecords(CsvFilePath, records);
                        SaveToExcel(ExcelFilePath, records);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"--> [Checkpoint] {records.Count} records saved to CSV and Excel.");
                        Console.ResetColor();
                    }

                    // 12-second delay between forks to avoid triggering anti-abuse heuristics
                    await Task.Delay(12000);
                }

                await Task.Delay(2000);
            }
        }

        // Final save to both formats
        SaveRecords(CsvFilePath, records);
        SaveToExcel(ExcelFilePath, records);

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"\n=================================================");
        Console.WriteLine($"FINISHED! Master report contains {records.Count} records.");
        Console.WriteLine($"CSV:   {Path.GetFullPath(CsvFilePath)}");
        Console.WriteLine($"Excel: {Path.GetFullPath(ExcelFilePath)}");
        Console.WriteLine($"=================================================");
        Console.ResetColor();
    }

    private static async Task<List<GitHubRepoDto>> SearchRepositoriesAsync(string query)
    {
        var list = new List<GitHubRepoDto>();
        var url = $"https://api.github.com/search/repositories?q={Uri.EscapeDataString(query + " fork:false archived:false")}&sort=stars&order=desc&per_page=100";

        while (true)
        {
            try
            {
                var response = await HttpClient.GetAsync(url);

                if (response.StatusCode == HttpStatusCode.Forbidden || (int)response.StatusCode == 429)
                {
                    int waitSeconds = 45;
                    if (response.Headers.TryGetValues("Retry-After", out var retryValues) &&
                        int.TryParse(retryValues.FirstOrDefault(), out var parsedSeconds))
                    {
                        waitSeconds = parsedSeconds + 5;
                    }

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"[Search Rate Limit] Pausing {waitSeconds}s for query: '{query}'...");
                    Console.ResetColor();

                    await Task.Delay(waitSeconds * 1000);
                    continue; // Retry the search query!
                }

                if (!response.IsSuccessStatusCode)
                {
                    var err = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"[Search Error ({response.StatusCode})]: {err}. Retrying in 10s...");
                    await Task.Delay(10000);
                    continue;
                }

                var content = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(content);

                if (!doc.RootElement.TryGetProperty("items", out var items))
                {
                    Console.WriteLine("No items returned in payload.");
                    return list;
                }

                foreach (var item in items.EnumerateArray())
                {
                    // Discard inactive repos (> 24 months without a commit)
                    if (item.TryGetProperty("pushed_at", out var pushedElement))
                    {
                        if (DateTime.TryParse(pushedElement.GetString(), out var pushedDate))
                        {
                            if (pushedDate < DateTime.UtcNow.AddMonths(-24)) continue;
                        }
                    }

                    var licenseSpdx = item.TryGetProperty("license", out var lic) && lic.ValueKind == JsonValueKind.Object && lic.TryGetProperty("spdx_id", out var spdx)
                        ? spdx.GetString() ?? ""
                        : "";

                    list.Add(new GitHubRepoDto
                    {
                        FullName = item.GetProperty("full_name").GetString() ?? "",
                        Name = item.GetProperty("name").GetString() ?? "",
                        HtmlUrl = item.GetProperty("html_url").GetString() ?? "",
                        Language = item.TryGetProperty("language", out var lang) ? lang.GetString() : "Unknown",
                        Stars = item.GetProperty("stargazers_count").GetInt32(),
                        DefaultBranch = item.GetProperty("default_branch").GetString() ?? "main",
                        LicenseSpdx = licenseSpdx
                    });
                }

                Console.WriteLine($"-> Found {list.Count} active candidates.");
                return list;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Search Exception]: {ex.Message}. Retrying in 15s...");
                await Task.Delay(15000);
            }
        }
    }
    private static async Task<string> GetHeadCommitShaAsync(string fullName, string defaultBranch)
    {
        try
        {
            var response = await HttpClient.GetAsync($"https://api.github.com/repos/{fullName}/commits/{defaultBranch}");
            if (response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(content);
                return doc.RootElement.GetProperty("sha").GetString() ?? "UNKNOWN";
            }
        }
        catch { }
        return "UNKNOWN";
    }

    private static async Task<(string ForkUrl, string Status)> ForkRepositoryAsync(string fullName)
    {
        var repoName = fullName.Split('/')[1];
        var expectedForkUrl = $"https://github.com/{TargetOrg}/{repoName}";
        var url = $"https://api.github.com/repos/{fullName}/forks";

        var payload = new
        {
            organization = TargetOrg,
            default_branch_only = true
        };

        while (true)
        {
            try
            {
                var jsonContent = new StringContent(JsonSerializer.Serialize(payload), System.Text.Encoding.UTF8, "application/json");
                var response = await HttpClient.PostAsync(url, jsonContent);

                if (response.StatusCode is HttpStatusCode.Accepted or HttpStatusCode.OK)
                {
                    var content = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(content);
                    var htmlUrl = doc.RootElement.GetProperty("html_url").GetString() ?? expectedForkUrl;
                    return (htmlUrl, "Verified");
                }

                // 422: Repository already forked/present in the organization (recaptures previously forked repos)
                if (response.StatusCode == HttpStatusCode.UnprocessableEntity)
                {
                    return (expectedForkUrl, "Verified");
                }

                // 403 or 429: Secondary rate limit
                if (response.StatusCode is HttpStatusCode.Forbidden or (HttpStatusCode)429)
                {
                    int waitSeconds = 120;
                    if (response.Headers.TryGetValues("Retry-After", out var retryValues) &&
                        int.TryParse(retryValues.FirstOrDefault(), out var parsedSeconds))
                    {
                        waitSeconds = parsedSeconds + 5;
                    }

                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"\n[Secondary Rate Limit] Cooldown: waiting {waitSeconds}s before retrying {fullName}...");
                    Console.ResetColor();

                    await Task.Delay(waitSeconds * 1000);
                    continue;
                }

                var err = await response.Content.ReadAsStringAsync();
                return (expectedForkUrl, $"Failed: {response.StatusCode}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Network Error]: {ex.Message}. Retrying in 15s...");
                await Task.Delay(15000);
            }
        }
    }

    private static List<ReportRecord> LoadExistingRecords(string path)
    {
        if (!File.Exists(path)) return new List<ReportRecord>();

        var firstLine = File.ReadLines(path).FirstOrDefault() ?? "";
        var delimiter = firstLine.Contains('\t') ? "\t" : ",";

        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter,
            HeaderValidated = null,
            MissingFieldFound = null
        };

        using var reader = new StreamReader(path);
        using var csv = new CsvReader(reader, config);
        return csv.GetRecords<ReportRecord>().ToList();
    }

    private static void SaveRecords(string path, List<ReportRecord> records)
    {
        var tempPath = path + ".tmp";
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ",",
            ShouldQuote = args => true
        };

        using (var writer = new StreamWriter(tempPath, false, System.Text.Encoding.UTF8))
        using (var csv = new CsvWriter(writer, config))
        {
            csv.WriteRecords(records);
        }

        File.Move(tempPath, path, overwrite: true);
    }

    private static void SaveToExcel(string excelPath, List<ReportRecord> records)
    {
        try
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("RelixQ_OSS_Report");

            // Header Row
            string[] headers = {
                "Repository_Name", "Forked_Repository", "Upstream_URL", "Category",
                "Language", "License", "Stars", "Default_Branch",
                "Upstream_Commit_SHA", "Fork_Status", "Fork_URL"
            };

            for (int col = 0; col < headers.Length; col++)
            {
                var cell = worksheet.Cell(1, col + 1);
                cell.Value = headers[col];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(230, 230, 230);
            }

            // Data Rows
            for (int i = 0; i < records.Count; i++)
            {
                var r = records[i];
                int row = i + 2;

                worksheet.Cell(row, 1).Value = r.Repository_Name;
                worksheet.Cell(row, 2).Value = r.Forked_Repository;
                worksheet.Cell(row, 3).Value = r.Upstream_URL;
                worksheet.Cell(row, 4).Value = r.Category;
                worksheet.Cell(row, 5).Value = r.Language;
                worksheet.Cell(row, 6).Value = r.License;
                worksheet.Cell(row, 7).Value = r.Stars;
                worksheet.Cell(row, 8).Value = r.Default_Branch;
                worksheet.Cell(row, 9).Value = r.Upstream_Commit_SHA;
                worksheet.Cell(row, 10).Value = r.Fork_Status;
                worksheet.Cell(row, 11).Value = r.Fork_URL;
            }

            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(excelPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Excel Warning] Could not save Excel file (is it open in Excel?): {ex.Message}");
        }
    }
}

public class GitHubRepoDto
{
    public string FullName { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string HtmlUrl { get; set; } = string.Empty;
    public string? Language { get; set; }
    public int Stars { get; set; }
    public string DefaultBranch { get; set; } = string.Empty;
    public string LicenseSpdx { get; set; } = string.Empty;
}