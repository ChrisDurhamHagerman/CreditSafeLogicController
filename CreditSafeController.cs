using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace CreditSafeController
{
    internal static class CsvAccessImporter
    {
        private static readonly string DbPath = @"C:\ADSK-Automation\Terms and Credit.accdb";
        private static readonly string LogFile = @"C:\ADSK-Automation\Logs\CreditSafeController.log";
        private static readonly string ConnStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DbPath};Persist Security Info=False;";

        // Main entry for both imports
        public static void ImportCreditSafeFiles(string customerCsv, string invoiceCsv)
        {
            try
            {
                Log("=== Starting CreditSafe CSV import process ===");

                if (!File.Exists(customerCsv))
                    throw new FileNotFoundException("Customer CSV was not found.", customerCsv);
                if (!File.Exists(invoiceCsv))
                    throw new FileNotFoundException("Invoice CSV was not found.", invoiceCsv);

                using (var connection = new OleDbConnection(ConnStr))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            Log($"Customer CSV detected: {customerCsv}");
                            ImportCustomerCsvToAccess(customerCsv, connection, transaction);

                            Log($"Invoice CSV detected: {invoiceCsv}");
                            ImportInvoiceCsvToAccess(invoiceCsv, connection, transaction);

                            transaction.Commit();
                        }
                        catch
                        {
                            try { transaction.Rollback(); } catch { }
                            throw;
                        }
                    }
                }

                Log("=== CreditSafe CSV import complete ===");
            }
            catch (Exception ex)
            {
                Log($"❌ ImportCreditSafeFiles ERROR: {ex}");
                throw;
            }
        }

        // --------------------------------------------------------------------
        // CUSTOMER CSV IMPORT
        // --------------------------------------------------------------------
        private static void ImportCustomerCsvToAccess(
            string csvPath,
            OleDbConnection connection,
            OleDbTransaction transaction)
        {
            try
            {
                Log("Importing Customer CSV into dbo_tc_customer...");

                var lines = File.ReadAllLines(csvPath);
                if (lines.Length <= 1)
                {
                    throw new InvalidDataException("Customer CSV contains no data rows.");
                }

                var headers = lines[0].Split(',').Select(h => h.Trim('"')).ToArray();

                // clear old rows within the shared customer/invoice transaction
                new OleDbCommand("DELETE FROM dbo_tc_customer", connection, transaction).ExecuteNonQuery();
                    Log("Cleared existing rows in dbo_tc_customer.");

                    for (int i = 1; i < lines.Length; i++)
                    {
                        var cols = ParseCsvLine(lines[i]);
                        if (cols.Length < headers.Length)
                            throw new InvalidDataException("Customer CSV row " + (i + 1) + " has fewer fields than its header.");

                        string sql = @"
INSERT INTO dbo_tc_customer
(
    InternalID,
    Customer,
    Credit_Limit,
    Credit_Safe_ID,
    Credit_Safe_Credit_Limit,
    [Credit Safe_Refresh_Required],
    Base_Deposit_Level,
    Effective_Deposit_Level,
    Weighted_DBT,
    Effective_DSO,
    HCO_Credit_Limit,
    Last_Effective_Deposit_Level_Change,
    Override_HCO_Credit_Level
)
VALUES
(
    @InternalID,
    @Customer,
    @Credit_Limit,
    @Credit_Safe_ID,
    @Credit_Safe_Credit_Limit,
    @Credit_Safe_Refresh_Required,
    @Base_Deposit_Level,
    @Effective_Deposit_Level,
    @Weighted_DBT,
    @Effective_DSO,
    @HCO_Credit_Limit,
    @Last_Effective_Deposit_Level_Change,
    @Override_HCO_Credit_Level
)";


                        using (var cmd = new OleDbCommand(sql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@InternalID", GetValue(cols, headers, "InternalID"));
                            cmd.Parameters.AddWithValue("@Customer", GetValue(cols, headers, "Customer"));
                            cmd.Parameters.AddWithValue("@Credit_Limit", ToDouble(cols, headers, "Credit Limit"));
                            cmd.Parameters.AddWithValue("@Credit_Safe_ID", GetValue(cols, headers, "Credit Safe ID"));
                            cmd.Parameters.AddWithValue("@Credit_Safe_Credit_Limit", ToDouble(cols, headers, "Credit Safe Credit Limit"));
                            cmd.Parameters.AddWithValue("@Credit_Safe_Refresh_Required", ToBool(cols, headers, "Credit Safe Refresh Required"));
                            cmd.Parameters.AddWithValue("@Base_Deposit_Level", GetValue(cols, headers, "Base Deposit Level"));
                            cmd.Parameters.AddWithValue("@Effective_Deposit_Level", GetValue(cols, headers, "Effective Deposit Level"));
                            cmd.Parameters.AddWithValue("@Weighted_DBT", ToDouble(cols, headers, "Weighted DBT"));
                            cmd.Parameters.AddWithValue("@Effective_DSO", ToDouble(cols, headers, "Effective DSO"));
                            cmd.Parameters.AddWithValue("@HCO_Credit_Limit", ToDouble(cols, headers, "HCO Credit Limit"));
                            cmd.Parameters.AddWithValue("@Last_Effective_Deposit_Level", ToDate(cols, headers, "Last Effective Deposit Level Change"));
                            cmd.Parameters.AddWithValue("@Override_HCO_Credit_Level", ToBool(cols, headers, "Override HCO Credit Level"));
                            cmd.ExecuteNonQuery();
                        }
                    }

                    Log($"✅ Imported {lines.Length - 1} rows into dbo_tc_customer.");
            }
            catch (Exception ex)
            {
                Log($"❌ ImportCustomerCsvToAccess ERROR: {ex}");
                throw;
            }
        }

        // --------------------------------------------------------------------
        // INVOICE CSV IMPORT
        // --------------------------------------------------------------------
        private static void ImportInvoiceCsvToAccess(
            string csvPath,
            OleDbConnection connection,
            OleDbTransaction transaction)
        {
            try
            {
                Log("Importing Invoice CSV into dbo_tc_invoice_dso...");

                var lines = File.ReadAllLines(csvPath);
                if (lines.Length <= 1)
                {
                    throw new InvalidDataException("Invoice CSV contains no data rows.");
                }

                var headers = lines[0].Split(',').Select(h => h.Trim('"')).ToArray();

                new OleDbCommand("DELETE FROM dbo_tc_invoice_dso", connection, transaction).ExecuteNonQuery();
                    Log("Cleared existing rows in dbo_tc_invoice_dso.");

                    for (int i = 1; i < lines.Length; i++)
                    {
                        var cols = ParseCsvLine(lines[i]);
                        if (cols.Length < headers.Length)
                            throw new InvalidDataException("Invoice CSV row " + (i + 1) + " has fewer fields than its header.");

                        string sql = @"
INSERT INTO dbo_tc_invoice_dso
(
    InternalID,
    Customer,
    Invoice_Nbr,
    Amount,
    Invoice_Date,
    Days_Open,
    Effective_Days_Open,
    Days_Overdue,
    Effective_Days_Overdue,
    Effective_Best_Possible_DSO,
    Effective_Final_DSO,
    Best_Possible_DSO_Amount,
    Current_Effective_DSO_Amount,
    SO,
    Status
)
VALUES
(
    @InternalID,
    @Customer,
    @Invoice_Nbr,
    @Amount,
    @Invoice_Date,
    @Days_Open,
    @Effective_Days_Open,
    @Days_Overdue,
    @Effective_Days_Overdue,
    @Effective_Best_Possible_DSO,
    @Effective_Final_DSO,
    @Best_Possible_DSO_Amount,
    @Current_Effective_DSO_Amount,
    @SO,
    @Status
)";


                        using (var cmd = new OleDbCommand(sql, connection, transaction))
                        {
                            cmd.Parameters.AddWithValue("@InternalID", GetValue(cols, headers, "InternalID"));
                            cmd.Parameters.AddWithValue("@Customer", GetValue(cols, headers, "Customer"));
                            cmd.Parameters.AddWithValue("@Invoice_Nbr", GetValue(cols, headers, "Invoice #"));
                            cmd.Parameters.AddWithValue("@Amount", ToDouble(cols, headers, "Amount"));
                            cmd.Parameters.AddWithValue("@Invoice_Date", ToDate(cols, headers, "Invoice Date"));
                            cmd.Parameters.AddWithValue("@Days_Open", ToInt(cols, headers, "Days Open"));
                            cmd.Parameters.AddWithValue("@Effective_Days_Open", ToInt(cols, headers, "Effective Days Open"));
                            cmd.Parameters.AddWithValue("@Days_Overdue", ToInt(cols, headers, "Days Overdue"));
                            cmd.Parameters.AddWithValue("@Effective_Days_Overdue", ToInt(cols, headers, "Effective Days Overdue"));
                            cmd.Parameters.AddWithValue("@Effective_Best_Possible_DSO", ToInt(cols, headers, "Effective Best Possible DSO"));
                            cmd.Parameters.AddWithValue("@Effective_Final_DSO", ToInt(cols, headers, "Effective Final DSO"));
                            cmd.Parameters.AddWithValue("@Best_Possible_DSO_Amount", ToDouble(cols, headers, "Best Possible DSO Amount"));
                            cmd.Parameters.AddWithValue("@Current_Effective_DSO_Amou", ToDouble(cols, headers, "Current Effective DSO Amount"));
                            cmd.Parameters.AddWithValue("@SO", GetValue(cols, headers, "SO#"));
                            cmd.Parameters.AddWithValue("@Status", GetValue(cols, headers, "Status"));
                            cmd.ExecuteNonQuery();
                        }
                    }

                    Log($"✅ Imported {lines.Length - 1} rows into dbo_tc_invoice_dso.");
            }
            catch (Exception ex)
            {
                Log($"❌ ImportInvoiceCsvToAccess ERROR: {ex}");
                throw;
            }
        }

        // --------------------------------------------------------------------
        // CSV HELPERS
        // --------------------------------------------------------------------
        private static string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            foreach (var c in line)
            {
                if (c == '"' && !inQuotes)
                    inQuotes = true;
                else if (c == '"' && inQuotes)
                    inQuotes = false;
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                    sb.Append(c);
            }

            result.Add(sb.ToString());
            return result.ToArray();
        }

        private static string GetValue(string[] cols, string[] headers, string name)
        {
            int i = Array.FindIndex(headers, h => h.Equals(name, StringComparison.OrdinalIgnoreCase));
            return (i >= 0 && i < cols.Length) ? cols[i].Trim('"') : "";
        }

        private static object ToInt(string[] cols, string[] headers, string name)
        {
            var s = GetValue(cols, headers, name);
            if (int.TryParse(s, out int val)) return val;
            return DBNull.Value;
        }

        private static object ToDouble(string[] cols, string[] headers, string name)
        {
            var s = GetValue(cols, headers, name).Replace(",", "");
            if (double.TryParse(s, out double val)) return val;
            return DBNull.Value;
        }

        private static object ToDate(string[] cols, string[] headers, string name)
        {
            var s = GetValue(cols, headers, name);
            if (DateTime.TryParse(s, out DateTime dt)) return dt;
            return DBNull.Value;
        }

        private static object ToBool(string[] cols, string[] headers, string name)
        {
            var s = GetValue(cols, headers, name).Trim().ToLower();
            if (s == "true" || s == "yes" || s == "1") return true;
            if (s == "false" || s == "no" || s == "0") return false;
            return DBNull.Value;
        }

        private static void Log(string message)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogFile) ?? ".");
                File.AppendAllText(LogFile, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}");
            }
            catch { }
        }
    }
}
