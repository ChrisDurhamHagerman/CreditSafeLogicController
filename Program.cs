using Mono.Web;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Office.Interop.Access;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;


namespace CreditSafeController
{
    internal static class Program
    {
        // ======= Paths / files =======
        private static readonly string DbPath = @"C:\ADSK-Automation\Terms and Credit.accdb";
        private static readonly string LogFolder = @"C:\ADSK-Automation\Logs";
        private static readonly string LogFile = Path.Combine(LogFolder, "CreditSafeController.log");
        private static readonly string ImportRoot = @"C:\ADSK-Automation\CreditSafe Imports\NetSuiteImports";
        private static readonly string ArchiveRoot = Path.Combine(ImportRoot, "Old Imports");
        private static readonly string CreditSafeExportRoot = @"C:\ADSK-Automation\CreditSafe Imports";
        private static readonly string CreditSafeOldExports = Path.Combine(CreditSafeExportRoot, "OldExports");


        // ======= Access =======
        private static readonly string ConnStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DbPath};Persist Security Info=False;";
        private const string TargetTable = "[CreditSafe-Weekly Import]";

        // ======= NetSuite TBA (fill these with PROD values) =======
        // TIP: once verified, move to env vars or a config file
        private static readonly string NsAccount = "290783";  // PROD realm (no _SB1)
        private static readonly string NsConsumerKey = "7b10d2fd278c7d00e5b9c4f368ccfd2d1de8b242e938dd801c0467a4d003c863";
        private static readonly string NsConsumerSecret = "c3d533e2b638e5350283fa8f38a66d65787902e1ff9b0010574901163f04a954";
        private static readonly string NsTokenId = "8213746ee169d07db851a1d8fc23c4762c86b0c1d5acfd4a0718a7ac42873071";
        private static readonly string NsTokenSecret = "3983843dedb5eb237ab3eaf9efeca8c445934f7cdc58f6cd4cbc85efadbc73a6";
        private static readonly string NsCustomerRestletUrl = "https://290783.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=862&deploy=1";
        private static readonly string NsInvoiceRestletUrl = "https://290783.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=863&deploy=1";
        private static readonly string NsTriggerImportRestletUrl = "https://290783.restlets.api.netsuite.com/app/site/hosting/restlet.nl?script=864&deploy=1";



        private static int Main(string[] args)
        {
            try
            {
                SafeEnsureDir(LogFolder);
                Log("CreditSafeController start.");

                if (args == null || args.Length == 0)
                {
                    Log("Usage: CreditSafeController <ExcelFilePath>");
                    return 2;
                }

                string excelPath = args[0];
                if (!File.Exists(excelPath))
                {
                    Log($"Excel file not found: {excelPath}");
                    return 3;
                }

                Log($"Excel path: {excelPath}");

                // 1) Load Excel to DataTable
                var excelData = LoadExcelData(excelPath);
                if (excelData == null || excelData.Rows.Count == 0)
                {
                    Log("No rows loaded from Excel.");
                    return 4;
                }
                Log($"Loaded {excelData.Rows.Count} rows from Excel.");

                // 2) Clean strings
                CleanAllStrings(excelData);

                // 3) Clear target table
                int cleared = ClearTargetTable();
                Log($"Cleared {cleared} row(s) from table {TargetTable}.");

                // 4) Insert rows
                InsertIntoAccess(excelData);

                // ========================================================
                // 5) Fetch Customer and Invoice CSVs sequentially
                // ========================================================
                Log($"Calling paged RESTlet (customer): {NsCustomerRestletUrl} | realm={NsAccount}");
                var customerCsv = CallPagedCsvRestletAsync(NsCustomerRestletUrl, "customer").GetAwaiter().GetResult();

                if (string.IsNullOrEmpty(customerCsv))
                {
                    Log("Customer fetch failed (null/empty). Skipping CustomerSync.csv write.");
                }
                else
                {
                    SafeEnsureDir(ImportRoot);
                    SafeEnsureDir(ArchiveRoot);
                    SaveWithArchive("CustomerSync.csv", customerCsv);
                    Log("Saved CustomerSync.csv to NetSuiteImports.");
                }

                Log($"Calling paged RESTlet (invoice): {NsInvoiceRestletUrl} | realm={NsAccount}");
                var invoiceCsv = CallPagedCsvRestletAsync(NsInvoiceRestletUrl, "invoice").GetAwaiter().GetResult();

                if (string.IsNullOrEmpty(invoiceCsv))
                {
                    Log("Invoice fetch failed (null/empty). Skipping InvoiceDSOSync.csv write.");
                }
                else
                {
                    SafeEnsureDir(ImportRoot);
                    SafeEnsureDir(ArchiveRoot);
                    SaveWithArchive("InvoiceDSOSync.csv", invoiceCsv);
                    Log("Saved InvoiceDSOSync.csv to NetSuiteImports.");
                }

                // ========================================================
                // 6) Ensure both CSVs exist before running Access process
                // ========================================================
                var customerPath = Path.Combine(ImportRoot, "CustomerSync.csv");
                var invoicePath = Path.Combine(ImportRoot, "InvoiceDSOSync.csv");

                bool customerReady = File.Exists(customerPath) && new FileInfo(customerPath).Length > 0;
                bool invoiceReady = File.Exists(invoicePath) && new FileInfo(invoicePath).Length > 0;

                if (customerReady && invoiceReady)
                {
                    Log("Both Customer and Invoice CSVs are present.");

                    // ⭐⭐⭐ ADD THIS ⭐⭐⭐
                    // Import CSVs into Access tables dbo_tc_customer and dbo_tc_invoice_dso
                    try
                    {
                        CsvAccessImporter.ImportCreditSafeFiles(customerPath, invoicePath);
                        Log("Imported CreditSafe CSVs into Access successfully.");
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR during Access import: {ex.Message}");
                    }
                    // ⭐⭐⭐ END ADD ⭐⭐⭐

                    // After importing, run Access process + export for NetSuite
                    Log("Running Access process...");
                    RunProcessAndExport();
                }
                else
                {
                    Log($"Skipping Access process. CustomerReady={customerReady}, InvoiceReady={invoiceReady}");
                }

                // ========================================================
                // 7) Send CreditSafeImport.csv directly to NetSuite Saved Import
                // ========================================================

                string creditSafeCsv = Path.Combine(CreditSafeExportRoot, "CreditSafeImport.csv");

                if (File.Exists(creditSafeCsv) && new FileInfo(creditSafeCsv).Length > 0)
                {
                    Log("Cleaning timestamps from CreditSafeImport.csv...");
                    CleanCreditSafeExport(creditSafeCsv);

                    Log("Sending CreditSafeImport.csv to NetSuite saved import...");

                    bool importOk = TriggerNetSuiteCsvImport(creditSafeCsv).GetAwaiter().GetResult();

                    if (importOk)
                        Log("✔ NetSuite CSV import triggered successfully.");
                    else
                        Log("❌ NetSuite CSV import failed.");
                }
                else
                {
                    Log("❌ CreditSafeImport.csv not found or empty. Cannot send to NetSuite.");
                }






                Log("CreditSafeController complete.");
                return 0;

            }
            catch (Exception ex)
            {
                Log($"ERROR: {ex}");
                return 1;
            }
        }





        // ---------------- Excel Load ----------------

        private static DataTable LoadExcelData(string excelPath)
        {
            string ext = Path.GetExtension(excelPath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xls")
                Log($"Warning: file extension {ext} is not a typical Excel file.");

            string excelConn = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelPath};
Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"";";

            using (var conn = new OleDbConnection(excelConn))
            {
                conn.Open();

                var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schema == null || schema.Rows.Count == 0)
                    throw new InvalidOperationException("No worksheets or ranges found in the Excel file.");

                Log($"Schema rows: {schema.Rows.Count}");
                foreach (DataRow r in schema.Rows)
                {
                    string tblType = r["TABLE_TYPE"]?.ToString();
                    string tblName = r["TABLE_NAME"]?.ToString();
                    Log($"Schema -> TYPE: {tblType}, NAME: {tblName}");
                }

                string sheetName = null;

                foreach (DataRow r in schema.Rows)
                {
                    string type = r["TABLE_TYPE"]?.ToString();
                    string name = r["TABLE_NAME"]?.ToString();
                    if (!string.IsNullOrEmpty(name) &&
                        string.Equals(type, "TABLE", StringComparison.OrdinalIgnoreCase) &&
                        (name.EndsWith("$") || name.EndsWith("$'")))
                    {
                        sheetName = name;
                        break;
                    }
                }

                if (sheetName == null)
                {
                    foreach (DataRow r in schema.Rows)
                    {
                        string name = r["TABLE_NAME"]?.ToString();
                        if (!string.IsNullOrEmpty(name) && name.Contains("$"))
                        {
                            sheetName = name;
                            break;
                        }
                    }
                }

                if (sheetName == null)
                    sheetName = schema.Rows[0]["TABLE_NAME"]?.ToString();

                if (string.IsNullOrEmpty(sheetName))
                    throw new InvalidOperationException("Could not determine worksheet name from schema.");

                Log($"Using worksheet/table name: {sheetName}");

                var dt = new DataTable();
                using (var cmd = new OleDbCommand($"SELECT * FROM [{sheetName}]", conn))
                using (var da = new OleDbDataAdapter(cmd))
                {
                    da.Fill(dt);
                }
                return dt;
            }
        }

        private static void CleanAllStrings(DataTable dt)
        {
            int changed = 0;
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    if (row[col] == null) continue;

                    var obj = row[col];
                    if (obj is string s)
                    {
                        var cleaned = CleanString(s);
                        if (!string.Equals(s, cleaned, StringComparison.Ordinal))
                        {
                            row[col] = cleaned;
                            changed++;
                        }
                    }
                }
            }
            Log($"String cleaning complete. Fields adjusted: {changed}.");
        }

        private static string CleanString(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            try
            {
                var bytes = Encoding.Default.GetBytes(input);
                return Encoding.UTF8.GetString(bytes);
            }
            catch
            {
                return input;
            }
        }

        // ---------------- Access Insert ----------------

        private static int ClearTargetTable()
        {
            using (var conn = new OleDbConnection(ConnStr))
            using (var cmd = new OleDbCommand($"DELETE FROM {TargetTable}", conn))
            {
                conn.Open();
                var affected = cmd.ExecuteNonQuery();
                return affected < 0 ? 0 : affected;
            }
        }

        private static void InsertIntoAccess(DataTable excelData)
        {
            using (var conn = new OleDbConnection(ConnStr))
            {
                conn.Open();

                int inserted = 0;
                int failed = 0;
                int rowIndex = 1;

                string insertSql = @"
INSERT INTO [CreditSafe-Weekly Import]
([Safe number], [Company name], [Country Code], [Company Number], [Local score], [Score],
 [Score text], [Status], [Credit Limit], [Personal reference], [Personal limit],
 [Notes], [Address], [Number of Employees], [Date of last change], [Summary],
 [Previous Value], [New Value], [Company Type])
VALUES
(@SafeNumber, @CompanyName, @CountryCode, @CompanyNumber, @LocalScore, @Score,
 @ScoreText, @Status, @CreditLimit, @PersonalReference, @PersonalLimit,
 @Notes, @Address, @NumberOfEmployees, @DateOfLastChange, @Summary,
 @PreviousValue, @NewValue, @CompanyType)";

                foreach (DataRow row in excelData.Rows)
                {
                    try
                    {
                        using (var cmd = new OleDbCommand(insertSql, conn))
                        {
                            cmd.Parameters.Add("@SafeNumber", OleDbType.VarWChar).Value = row["Safe Number"] ?? DBNull.Value;
                            cmd.Parameters.Add("@CompanyName", OleDbType.VarWChar).Value = row["Company Name"] ?? DBNull.Value;
                            cmd.Parameters.Add("@CountryCode", OleDbType.VarWChar).Value = row["Country Code"] ?? DBNull.Value;
                            cmd.Parameters.Add("@CompanyNumber", OleDbType.VarWChar).Value = row["Company Number"] ?? DBNull.Value;

                            cmd.Parameters.Add("@LocalScore", OleDbType.Integer).Value = ParseInt(row["Local Score"], rowIndex, "Local Score");
                            cmd.Parameters.Add("@Score", OleDbType.VarWChar).Value = row["Score"] ?? DBNull.Value;

                            cmd.Parameters.Add("@ScoreText", OleDbType.VarWChar).Value = row["Score Text"] ?? DBNull.Value;
                            cmd.Parameters.Add("@Status", OleDbType.VarWChar).Value = row["Status"] ?? DBNull.Value;

                            cmd.Parameters.Add("@CreditLimit", OleDbType.Double).Value = ParseDoubleDefaultZero(row["Credit Limit"], rowIndex, "Credit Limit");

                            cmd.Parameters.Add("@PersonalReference", OleDbType.VarWChar).Value = row["Personal Reference"] ?? DBNull.Value;
                            cmd.Parameters.Add("@PersonalLimit", OleDbType.Double).Value = ParseDoubleNullable(row["Personal Limit"], rowIndex, "Personal Limit");

                            cmd.Parameters.Add("@Notes", OleDbType.VarWChar).Value = row["Notes"] ?? DBNull.Value;
                            cmd.Parameters.Add("@Address", OleDbType.VarWChar).Value = row["Address"] ?? DBNull.Value;

                            cmd.Parameters.Add("@NumberOfEmployees", OleDbType.Integer).Value = ParseInt(row["Number Of Employees"], rowIndex, "Number Of Employees");
                            cmd.Parameters.Add("@DateOfLastChange", OleDbType.Date).Value = ParseDate(row["Date Of Last Change"], rowIndex, "Date Of Last Change");

                            cmd.Parameters.Add("@Summary", OleDbType.VarWChar).Value = row["Summary"] ?? DBNull.Value;
                            cmd.Parameters.Add("@PreviousValue", OleDbType.VarWChar).Value = row["Previous Value"] ?? DBNull.Value;
                            cmd.Parameters.Add("@NewValue", OleDbType.VarWChar).Value = row["New Value"] ?? DBNull.Value;
                            cmd.Parameters.Add("@CompanyType", OleDbType.VarWChar).Value = row["Company Type"] ?? DBNull.Value;

                            cmd.ExecuteNonQuery();
                        }

                        Log($"Inserted row {rowIndex} (Safe Number: {row["Safe Number"]?.ToString() ?? "Unknown"})");
                        inserted++;
                    }
                    catch (Exception exRow)
                    {
                        var rawData = string.Join(" | ", row.ItemArray.Select(i => i?.ToString()));
                        Log($"ERROR inserting row {rowIndex}: {exRow.Message}");
                        Log($"    Row data: {rawData}");
                        failed++;
                    }

                    rowIndex++;
                }

                Log($"Insert complete. Success: {inserted}, Failed: {failed}, Total: {rowIndex - 1}");
            }
        }

        private static object ParseInt(object value, int rowIndex, string column)
        {
            if (value == null) return DBNull.Value;
            var s = value.ToString();
            if (int.TryParse(s, out int i)) return i;
            Log($"Row {rowIndex} - Invalid integer in '{column}': '{s}' -> NULL");
            return DBNull.Value;
        }

        private static object ParseDoubleNullable(object value, int rowIndex, string column)
        {
            if (value == null) return DBNull.Value;
            var s = value.ToString();
            if (double.TryParse(s, out double d)) return d;
            if (string.IsNullOrWhiteSpace(s)) return DBNull.Value;
            Log($"Row {rowIndex} - Invalid double in '{column}': '{s}' -> NULL");
            return DBNull.Value;
        }

        private static object ParseDoubleDefaultZero(object value, int rowIndex, string column)
        {
            if (value == null) return 0d;
            var s = value.ToString();
            if (double.TryParse(s, out double d)) return d;
            Log($"Row {rowIndex} - Missing or invalid '{column}', defaulting to 0.");
            return 0d;
        }

        private static object ParseDate(object value, int rowIndex, string column)
        {
            if (value == null) return DBNull.Value;

            var s = value.ToString().Trim();
            if (string.IsNullOrEmpty(s))
                return DBNull.Value;

            if (s.Length >= 10)
            {
                var datePart = s.Substring(0, 10);
                if (DateTime.TryParse(datePart, out DateTime dt))
                    return dt.Date;
            }

            if (DateTime.TryParse(s, out DateTime fullDt))
                return fullDt.Date;

            Log($"Row {rowIndex} - Invalid date in '{column}': '{s}' -> NULL");
            return DBNull.Value;
        }

        // ---------------- NetSuite RESTlet ----------------

        private static async Task<string> SignedGetAsync(string url, string labelForLogs)
        {
            // Mirrors your existing TBA signing approach, but returns raw body text.
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            string baseString, paramString, nonce, ts;
            var auth = BuildTbaAuthHeader_HS256(
                method: "GET",
                url: url,
                realm: NsAccount,
                consumerKey: NsConsumerKey,
                consumerSecret: NsConsumerSecret,
                tokenId: NsTokenId,
                tokenSecret: NsTokenSecret,
                baseString: out baseString,
                paramString: out paramString,
                nonce: out nonce,
                timestamp: out ts
            );

            Log($"RESTlet page request ({labelForLogs}): {url}");

            using (var http = new HttpClient())
            {
                var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.Add("Authorization", auth);
                req.Headers.Add("Accept", "text/plain"); // CSV text

                var res = await http.SendAsync(req);
                var body = await res.Content.ReadAsStringAsync();

                Log($"RESTlet HTTP {(int)res.StatusCode} {res.ReasonPhrase} (page request)");
                if (!res.IsSuccessStatusCode)
                {
                    Log($"RESTlet response body: {body}");
                    return null;
                }

                return body; // raw CSV text
            }
        }

        private static async Task<string> CallPagedCsvRestletAsync(string baseUrl, string label)
        {
            int page = 0;
            int pageSize = 1000;   // no cap; RESTlet honors 1000
            bool wroteHeader = false;
            var sb = new StringBuilder();

            while (true)
            {
                var url = AppendQuery(baseUrl, new Dictionary<string, string>
                {
                    ["page"] = page.ToString(),
                    ["pageSize"] = pageSize.ToString(),
                    ["header"] = wroteHeader ? "false" : "true"
                });

                var csv = await SignedGetAsync(url, label);
                if (string.IsNullOrEmpty(csv))
                {
                    Log($"{label} page {page} failed or returned null.");
                    break;
                }

                sb.Append(csv);
                if (!csv.EndsWith("\n")) sb.AppendLine();

                // Heuristic to know we're done:
                // Count lines in this page (subtract header on first page).
                int lineCount = csv.Split('\n').Length - (page == 0 ? 1 : 0);
                if (lineCount < pageSize) break;

                wroteHeader = true;
                page++;
            }

            return sb.ToString();
        }


        private static string AppendQuery(string baseUrl, IDictionary<string, string> queryParams)
        {
            var uri = new Uri(baseUrl, UriKind.Absolute);
            var qb = HttpUtility.ParseQueryString(uri.Query ?? string.Empty);
            foreach (var kvp in queryParams) qb[kvp.Key] = kvp.Value;
            var ub = new UriBuilder(uri) { Query = qb.ToString() ?? string.Empty };
            return ub.Uri.ToString();
        }

        private static async Task<string> CallRestletForCsvAsync(string url, string _propertyNameIgnored)
        {
            using (var http = new HttpClient())
            {
                string baseString, paramString, nonce, ts;
                var auth = BuildTbaAuthHeader_HS256(
                    method: "GET",
                    url: url,               // sign exact URL incl. query
                    realm: NsAccount,
                    consumerKey: NsConsumerKey,
                    consumerSecret: NsConsumerSecret,
                    tokenId: NsTokenId,
                    tokenSecret: NsTokenSecret,
                    baseString: out baseString,
                    paramString: out paramString,
                    nonce: out nonce,
                    timestamp: out ts
                );

                Log($"RESTlet page request: {url}");
                var req = new HttpRequestMessage(HttpMethod.Get, url);
                req.Headers.Add("Authorization", auth);
                // Accept anything; RESTlet returns plain text
                req.Headers.Add("Accept", "*/*");

                var res = await http.SendAsync(req);
                var body = await res.Content.ReadAsStringAsync();

                Log($"RESTlet HTTP {(int)res.StatusCode} {res.ReasonPhrase} (page request)");
                if (!res.IsSuccessStatusCode)
                {
                    Log($"RESTlet response body: {body}");
                    return null;
                }

                // RESTlet returns CSV TEXT directly
                return body ?? string.Empty;
            }
        }





        private static string BuildTbaAuthHeader_HS256(
    string method, string url, string realm,
    string consumerKey, string consumerSecret,
    string tokenId, string tokenSecret,
    out string baseString, out string paramString,
    out string nonce, out string timestamp)
        {
            nonce = Guid.NewGuid().ToString("N");
            timestamp = ((long)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds).ToString();

            var uri = new Uri(url);
            var baseUrl = $"{uri.Scheme}://{uri.Host}{uri.AbsolutePath}";
            var queryParams = ParseQueryString(uri.Query);

            var oauth = new[]
            {
        new KeyValuePair<string,string>("oauth_consumer_key", consumerKey),
        new KeyValuePair<string,string>("oauth_token", tokenId),
        new KeyValuePair<string,string>("oauth_nonce", nonce),
        new KeyValuePair<string,string>("oauth_timestamp", timestamp),
        new KeyValuePair<string,string>("oauth_signature_method", "HMAC-SHA256"),
        new KeyValuePair<string,string>("oauth_version", "1.0")
    };

            var allParams = queryParams.Concat(oauth)
                .Select(p => new KeyValuePair<string, string>(PercentEncode(p.Key), PercentEncode(p.Value)))
                .OrderBy(p => p.Key, StringComparer.Ordinal)
                .ThenBy(p => p.Value, StringComparer.Ordinal)
                .ToArray();

            paramString = string.Join("&", allParams.Select(p => $"{p.Key}={p.Value}"));
            baseString = $"{method.ToUpperInvariant()}&{PercentEncode(baseUrl)}&{PercentEncode(paramString)}";

            var signingKey = $"{PercentEncode(consumerSecret)}&{PercentEncode(tokenSecret)}";
            string signature;
            using (var hmac = new System.Security.Cryptography.HMACSHA256(Encoding.ASCII.GetBytes(signingKey)))
            {
                signature = Convert.ToBase64String(hmac.ComputeHash(Encoding.ASCII.GetBytes(baseString)));
            }

            var sb = new StringBuilder("OAuth ");
            sb.Append($"realm=\"{PercentEncode(realm)}\"");
            foreach (var kv in oauth)
                sb.Append($", {kv.Key}=\"{PercentEncode(kv.Value)}\"");
            sb.Append($", oauth_signature=\"{PercentEncode(signature)}\"");

            return sb.ToString();
        }

        private static string PercentEncode(string value)
        {
            return Uri.EscapeDataString(value ?? string.Empty).Replace("%7E", "~");
        }

        private static System.Collections.Generic.IEnumerable<KeyValuePair<string, string>> ParseQueryString(string query)
        {
            var list = new System.Collections.Generic.List<KeyValuePair<string, string>>();
            if (string.IsNullOrEmpty(query)) return list;
            var q = query[0] == '?' ? query.Substring(1) : query;
            foreach (var pair in q.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var idx = pair.IndexOf('=');
                if (idx < 0)
                    list.Add(new KeyValuePair<string, string>(Uri.UnescapeDataString(pair), ""));
                else
                    list.Add(new KeyValuePair<string, string>(
                        Uri.UnescapeDataString(pair.Substring(0, idx)),
                        Uri.UnescapeDataString(pair.Substring(idx + 1))
                    ));
            }
            return list;
        }





        private static string ExtractJsonProperty(string json, string prop)
        {
            // minimal string extractor for JSON string fields: "prop":"value"
            var marker = $"\"{prop}\":";
            var idx = json.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx < 0) return null;

            var start = json.IndexOf('"', idx + marker.Length);
            if (start < 0) return null;
            start++;

            var sb = new StringBuilder();
            for (int i = start; i < json.Length; i++)
            {
                var ch = json[i];
                if (ch == '"' && json[i - 1] != '\\') break;
                sb.Append(ch);
            }
            return sb.ToString().Replace("\\\"", "\"");
        }

        private static void SaveWithArchive(string fileName, string contents)
        {
            try
            {
                SafeEnsureDir(ImportRoot);
                SafeEnsureDir(ArchiveRoot);

                var path = Path.Combine(ImportRoot, fileName);
                if (File.Exists(path))
                {
                    var archived = Path.Combine(ArchiveRoot,
                        $"{Path.GetFileNameWithoutExtension(fileName)}_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
                    File.Move(path, archived);
                    Log($"Archived {fileName} to {archived}");
                }

                File.WriteAllText(path, contents, Encoding.UTF8);
                Log($"Wrote {fileName} ({contents.Length} chars) to {path}.");
            }
            catch (Exception ex)
            {
                Log($"Error saving {fileName}: {ex.Message}");
            }
        }

        private static void RunProcessAndExport()
        {
            Microsoft.Office.Interop.Access.Application accessApp = null;

            try
            {
                Log("Running Access macro: 02-Process Everything");

                accessApp = new Microsoft.Office.Interop.Access.Application();
                accessApp.OpenCurrentDatabase(DbPath, false);

                // Run macro
                accessApp.DoCmd.RunMacro("02-Process Everything");
                Log("Macro completed successfully.");

                // Export the output query using Access (NOT OleDb Text driver)
                ExportQueryToCsv_AccessCom(
                    accessApp: accessApp,
                    queryOrTableName: "09-Import to NetSuite Data",  // no brackets here
                    exportDir: @"C:\ADSK-Automation\CreditSafe Imports",
                    exportFileName: "CreditSafeImport.csv"
                );
            }
            catch (Exception ex)
            {
                Log($"ERROR running Access macro/export: {ex.Message}");
                if (ex.InnerException != null)
                    Log($"   ⤷ Inner: {ex.InnerException.Message}");
            }
            finally
            {
                try
                {
                    if (accessApp != null)
                    {
                        accessApp.CloseCurrentDatabase();
                        accessApp.Quit();
                        Marshal.ReleaseComObject(accessApp);
                        accessApp = null;
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        /// <summary>
        /// Exports a saved Access query (or table) to CSV using Access COM (DoCmd.TransferText),
        /// which avoids the fragile OleDb Text/ISAM SELECT INTO behavior.
        /// Archives old export first (same behavior as your prior method).
        /// </summary>
        private static void ExportQueryToCsv_AccessCom(
            Microsoft.Office.Interop.Access.Application accessApp,
            string queryOrTableName,
            string exportDir,
            string exportFileName)
        {
            try
            {
                SafeEnsureDir(exportDir);
                SafeEnsureDir(CreditSafeOldExports);

                var exportPath = Path.Combine(exportDir, exportFileName);

                // Archive old file if present
                if (File.Exists(exportPath))
                {
                    var archivedPath = Path.Combine(
                        CreditSafeOldExports,
                        $"{Path.GetFileNameWithoutExtension(exportFileName)}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                    );

                    try
                    {
                        File.Move(exportPath, archivedPath);
                        Log($"Archived previous export to {archivedPath}");
                    }
                    catch (Exception exMove)
                    {
                        Log($"WARN: Could not move old export (will try delete). {exMove.Message}");
                        File.Delete(exportPath);
                        Log("Deleted existing export to allow new write.");
                    }
                }

                // Export using Access engine
                Log($"Exporting '{queryOrTableName}' -> {exportPath}");

                accessApp.DoCmd.TransferText(
                    Microsoft.Office.Interop.Access.AcTextTransferType.acExportDelim,
                    Type.Missing,              // export spec (none)
                    queryOrTableName,          // saved query name OR table name
                    exportPath,                // output file
                    true                       // include headers
                );

                if (!File.Exists(exportPath))
                    throw new IOException($"Expected export file not found after export: {exportPath}");

                var len = new FileInfo(exportPath).Length;
                Log($"Exported {queryOrTableName} -> {exportPath} ({len:n0} bytes).");
            }
            catch (Exception ex)
            {
                Log($"ERROR exporting query to CSV (Access COM): {ex.Message}");
                if (ex.InnerException != null)
                    Log($"   ⤷ Inner: {ex.InnerException.Message}");
            }
        }



        private static async Task<bool> TriggerNetSuiteCsvImport(string csvFilePath)
        {
            try
            {
                if (!File.Exists(csvFilePath))
                {
                    Log("❌ CreditSafeImport.csv not found.");
                    return false;
                }

                string csvText = File.ReadAllText(csvFilePath, Encoding.UTF8);

                string url = NsTriggerImportRestletUrl;  // script=864

                string nonce, ts, baseString, paramString;
                string auth = BuildTbaAuthHeader_HS256(
                    "POST",
                    url,
                    NsAccount,
                    NsConsumerKey,
                    NsConsumerSecret,
                    NsTokenId,
                    NsTokenSecret,
                    out baseString,
                    out paramString,
                    out nonce,
                    out ts
                );

                Log($"Triggering saved import with RESTlet 864...");
                Log($"TBA: nonce={nonce}, ts={ts}");

                var json = "{ \"csv\": \"" + EscapeJson(csvText) + "\" }";

                using (var http = new HttpClient())
                {
                    var req = new HttpRequestMessage(HttpMethod.Post, url);
                    req.Headers.Add("Authorization", auth);
                    req.Content = new StringContent(json, Encoding.UTF8, "application/json");

                    var res = await http.SendAsync(req);
                    var body = await res.Content.ReadAsStringAsync();

                    Log($"Import trigger HTTP {(int)res.StatusCode}: {res.ReasonPhrase}");
                    Log($"Response: {body}");

                    return res.IsSuccessStatusCode;
                }
            }
            catch (Exception ex)
            {
                Log($"❌ ERROR TriggerNetSuiteCsvImport: {ex.Message}");
                return false;
            }
        }




        /// <summary>
        /// Exports a SELECTable table/query from Access to a CSV using the ACE Text driver.
        /// No COM automation required. Overwrites file; archives prior version.
        /// </summary>
        private static void ExportQueryToCsv(string queryName, string exportDir, string exportFileName)
        {
            try
            {
                // Ensure folders
                SafeEnsureDir(exportDir);
                SafeEnsureDir(CreditSafeOldExports);

                var exportPath = Path.Combine(exportDir, exportFileName);

                // If an old export exists, archive it first
                if (File.Exists(exportPath))
                {
                    var archivedPath = Path.Combine(
                        CreditSafeOldExports,
                        $"{Path.GetFileNameWithoutExtension(exportFileName)}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
                    );

                    try
                    {
                        File.Move(exportPath, archivedPath);
                        Log($"Archived previous export to {archivedPath}");
                    }
                    catch (Exception exMove)
                    {
                        Log($"WARN: Could not move old export (will try delete). {exMove.Message}");
                        try
                        {
                            File.Delete(exportPath);
                            Log("Deleted existing export to allow new write.");
                        }
                        catch (Exception exDel)
                        {
                            throw new IOException($"Existing export file is locked: {exportPath}. {exDel.Message}");
                        }
                    }
                }

                // Write new export via ACE Text driver (SELECT INTO)
                using (var conn = new OleDbConnection(ConnStr))
                {
                    conn.Open();

                    // Target: [Text;FMT=Delimited;HDR=Yes;Database=<folder>].[<file>]
                    string sql = $@"
SELECT *
INTO [Text;FMT=Delimited;HDR=Yes;Database={exportDir};].[{exportFileName}]
FROM {queryName}";
                    using (var cmd = new OleDbCommand(sql, conn))
                    {
                        // For SELECT INTO, ExecuteNonQuery usually returns -1 (expected)
                        cmd.ExecuteNonQuery();
                    }
                }

                if (!File.Exists(exportPath))
                    throw new IOException($"Expected export file not found after export: {exportPath}");

                var len = new FileInfo(exportPath).Length;
                Log($"Exported {queryName} -> {exportPath} ({len:n0} bytes).");
            }
            catch (Exception ex)
            {
                Log($"ERROR exporting query to CSV: {ex.Message}");
            }
        }



        private static void SafeEnsureDir(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                Log($"Failed to ensure directory '{path}': {ex.Message}");
            }
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";

            return s
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }

        private static void CleanCreditSafeExport(string csvPath)
        {
            try
            {
                if (!File.Exists(csvPath))
                {
                    Log($"CleanCreditSafeExport: file not found: {csvPath}");
                    return;
                }

                var originalText = File.ReadAllText(csvPath);

                // Match things like:
                //  11/17/2025 14:57:01
                //  1/7/2025 9:05:03
                // and replace with just the date part:
                //  11/17/2025
                var pattern = @"(\d{1,2}/\d{1,2}/\d{4})\s+\d{1,2}:\d{2}:\d{2}";
                var cleanedText = Regex.Replace(originalText, pattern, "$1");

                File.WriteAllText(csvPath, cleanedText, Encoding.UTF8);

                Log("CleanCreditSafeExport: stripped time component from date/time values in CSV.");
            }
            catch (Exception ex)
            {
                Log($"CleanCreditSafeExport ERROR: {ex.Message}");
            }
        }



        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var cur = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }

                if (c == ',' && !inQuotes)
                {
                    result.Add(cur.ToString());
                    cur.Clear();
                }
                else
                {
                    cur.Append(c);
                }
            }

            result.Add(cur.ToString());
            return result;
        }

        private static string RebuildCsvLine(List<string> parts)
        {
            return string.Join(",", parts.Select(v =>
            {
                if (v.Contains(",") || v.Contains("\""))
                    return "\"" + v.Replace("\"", "\"\"") + "\"";
                return v;
            }));
        }





        private static void Log(string message)
        {
            try
            {
                SafeEnsureDir(Path.GetDirectoryName(LogFile) ?? ".");
                File.AppendAllText(LogFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}");
            }
            catch
            {
                // swallow logging errors
            }
        }

        //private static System.Collections.Generic.IEnumerable<(string k, string v)> ParseQueryString(string query)
        //{
        //    var list = new System.Collections.Generic.List<(string, string)>();
        //    if (string.IsNullOrEmpty(query)) return list;
        //    var q = query[0] == '?' ? query.Substring(1) : query;
        //    foreach (var pair in q.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
        //    {
        //        var idx = pair.IndexOf('=');
        //        if (idx < 0)
        //        {
        //            list.Add((Uri.UnescapeDataString(pair), ""));
        //        }
        //        else
        //        {
        //            var key = Uri.UnescapeDataString(pair.Substring(0, idx));
        //            var val = Uri.UnescapeDataString(pair.Substring(idx + 1));
        //            list.Add((key, val));
        //        }
        //    }
        //    return list;
        //}

    }
}
