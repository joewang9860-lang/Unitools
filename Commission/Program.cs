using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using ClosedXML.Excel;

class Program
{
    // Input/output folders are placed next to the running executable so each
    // published tool has its own `input` and `output` directories alongside the exe.
    static string InputDir => Path.Combine(AppContext.BaseDirectory, "input");
    static string OutputDir => Path.Combine(AppContext.BaseDirectory, "output");
    static string salesExcelPath => Path.Combine(InputDir, "SalesData.xlsx");

    static void Main()
    {
        // Ensure input/output directories exist (create if missing)
        Directory.CreateDirectory(InputDir);
        Directory.CreateDirectory(OutputDir);

        // Staff-specific month offsets (how many months back from the input date to use)
        var staffOffsets = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["DMB"] = 1,
            ["DWF"] = 2,
            ["PPR"] = 2,
        };

        if (!File.Exists(salesExcelPath))
        {
            Console.WriteLine($"Reference sales file not found: {salesExcelPath}");
            Console.WriteLine("Place SalesData.xlsx into the input folder and re-run.");
            return;
        }

        // Show brief user manual before prompting
        ShowUserManual();

        Console.Write("Enter target year-month (yyyyMM) or press Enter for current: ");
        var inputYm = Console.ReadLine()?.Trim();
        string ym;
        if (string.IsNullOrWhiteSpace(inputYm) || inputYm.Length != 6)
            ym = DateTime.Now.ToString("yyyyMM");
        else
            ym = inputYm;

        // Discover available staff input files (files starting with staff code)
        var staffFiles = FindStaffFiles(InputDir, "SalesData.xlsx");
        Console.WriteLine($"Found {staffFiles.Count} staff file(s) in input folder.");
        foreach (var kv0 in staffFiles)
            Console.WriteLine($" - Staff: {kv0.Key}, File: {Path.GetFileName(kv0.Value)}");

        // For each configured staff, compute source year-month (input ym minus offset)
        foreach (var kv in staffOffsets)
        {
            var staffCode = kv.Key;
            var offset = kv.Value;
            // parse ym as yyyyMM
            if (!DateTime.TryParseExact(ym + "01", "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var baseDate))
            {
                Console.WriteLine($"Invalid year-month: {ym}");
                continue;
            }
            var srcDate = baseDate.AddMonths(-offset);
            var srcYm = srcDate.ToString("yyyyMM");

            // Determine staff input file by staff code (any file starting with the code)
            if (!staffFiles.TryGetValue(staffCode, out var staffPath))
            {
                Console.WriteLine($"No input file found for staff {staffCode} in input folder. Skipping.");
                continue;
            }

            var outName = $"{staffCode}{srcYm}.xlsx"; // output named by staff + source ym (e.g., DMB202602.xlsx)
            var outPath = Path.Combine(OutputDir, outName);
            Console.WriteLine($"Processing staff {staffCode} using {Path.GetFileName(staffPath)} -> {outName}");
            try
            {
                ProcessStaffFile(salesExcelPath, staffPath, outPath, staffCode);
                Console.WriteLine($"Written: {outPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {staffCode}: {ex.Message}");
            }
        }
    }

    // Process a staff file using SalesData as reference and produce an output workbook
    // with the same three-sheet structure: (1) all sales of the month, (2) unpaid sales of the month,
    // (3) history unpaid before the month. Output is saved to outputPath.
    static void ProcessStaffFile(string salesPath, string staffPath, string outputPath, string staffCode)
    {
        var sales = ReadSalesRecords(salesPath);

        // Determine output year-month from outputPath (e.g., DMB202602.xlsx -> 202602)
        var outYm = ExtractYearMonthFromFileName(Path.GetFileName(outputPath));
        if (!TryParseYearMonth(outYm, out var outYear, out var outMonth))
            throw new InvalidOperationException($"Cannot determine output year-month from filename: {outputPath}");

        var monthStart = new DateTime(outYear, outMonth, 1);
        var monthEnd = monthStart.AddMonths(1).AddDays(-1);

        // Filter records
        var staffAll = sales.Where(s => string.Equals(s.SalesRep, staffCode, StringComparison.OrdinalIgnoreCase)
                                        && s.TransactionDate >= monthStart && s.TransactionDate <= monthEnd).ToList();
        var staffUnpaid = staffAll.Where(s => !string.Equals(s.ARPaid, "Paid", StringComparison.OrdinalIgnoreCase)).ToList();
        var staffHistoryUnpaid = sales.Where(s => string.Equals(s.SalesRep, staffCode, StringComparison.OrdinalIgnoreCase)
                                                 && s.TransactionDate < monthStart
                                                 && !string.Equals(s.ARPaid, "Paid", StringComparison.OrdinalIgnoreCase)).ToList();

        using var staffWb = new XLWorkbook(staffPath);

        // Extract unpaid rows from staff input sheets 2 and 3 (if present). These will populate output sheet 3.
        var templateWs2 = staffWb.Worksheets.ElementAtOrDefault(1);
        var templateWs3 = staffWb.Worksheets.ElementAtOrDefault(2);
        var unpaidFromInput = new List<SaleRecord>();
        if (templateWs2 != null) unpaidFromInput.AddRange(ExtractUnpaidFromWorksheet(templateWs2));
        if (templateWs3 != null) unpaidFromInput.AddRange(ExtractUnpaidFromWorksheet(templateWs3));
        // Order sheet 3 entries by date ascending (early to late)
        unpaidFromInput = unpaidFromInput.OrderBy(s => s.TransactionDate).ToList();

        // For each unpaid row from input, check SalesData by invoice number. If SalesData marks it as Paid,
        // update the record and flag it so we can highlight the row in output.
        foreach (var rec in unpaidFromInput)
        {
            if (string.IsNullOrWhiteSpace(rec.InvoiceNumber)) continue;
            var match = sales.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s.InvoiceNumber) && string.Equals(s.InvoiceNumber.Trim(), rec.InvoiceNumber.Trim(), StringComparison.OrdinalIgnoreCase));
            if (match != null && string.Equals(match.ARPaid, "Paid", StringComparison.OrdinalIgnoreCase))
            {
                rec.ARPaid = "Paid";
                rec.IsPaidBySales = true;
            }
        }

        // Ensure three worksheets exist (preserve content if present)
        var wsAll = staffWb.Worksheets.ElementAtOrDefault(0) ?? staffWb.AddWorksheet("AllSales");
        var wsUnpaid = staffWb.Worksheets.ElementAtOrDefault(1) ?? staffWb.AddWorksheet("Unpaid");
        var wsHist = staffWb.Worksheets.ElementAtOrDefault(2) ?? staffWb.AddWorksheet("HistoryUnpaid");

        // Rename output worksheets to the requested names
        try { wsAll.Name = $"Sales {outYm}"; } catch { /* ignore rename errors */ }
        try { wsUnpaid.Name = $"Unpaid {outYm}"; } catch { /* ignore rename errors */ }
        try { wsHist.Name = "Outstanding"; } catch { /* ignore rename errors */ }

        void RewriteSheet(IXLWorksheet ws, List<SaleRecord> recs)
        {
            var headerRow = 1;
            var lastCol = Math.Max(12, ws.LastColumnUsed()?.ColumnNumber() ?? 12);
            var headers = new List<string>();
            for (int c = 1; c <= lastCol; c++) headers.Add(ws.Cell(headerRow, c).GetString());

            ws.Clear();
            for (int c = 1; c <= headers.Count; c++) ws.Cell(headerRow, c).Value = headers[c - 1];

            var r = headerRow + 1;
            foreach (var s in recs)
            {
                var maxCols = Math.Max(headers.Count, s.Cells?.Count ?? 0);
                for (int c = 1; c <= maxCols; c++)
                {
                    // Column 1: ensure Customer/Company name is present (use parsed Customer if available)
                    if (c == 1)
                    {
                        var custVal = !string.IsNullOrWhiteSpace(s.Customer)
                            ? s.Customer
                            : (s.Cells != null && s.Cells.Count >= 1 ? s.Cells[0] : string.Empty);
                        ws.Cell(r, c).Value = custVal;
                        continue;
                    }

                    // Column 2: prefer the parsed TransactionDate if available
                    if (c == 2 && s.TransactionDate != default)
                    {
                        ws.Cell(r, c).Value = s.TransactionDate;
                        continue;
                    }

                    var val = s.Cells != null && s.Cells.Count >= c ? s.Cells[c - 1] : string.Empty;
                    ws.Cell(r, c).Value = val;
                }

                // Highlight unpaid rows with light orange across columns A..L (1..12)
                try
                {
                    if (!string.Equals(s.ARPaid, "Paid", StringComparison.OrdinalIgnoreCase))
                    {
                        ws.Range(r, 1, r, 12).Style.Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.FromHtml("#FFDAB9"));
                    }
                    else if (s.IsPaidBySales)
                    {
                        // Paid by sales mapping - use light green
                        ws.Range(r, 1, r, 12).Style.Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.LightGreen);
                    }
                }
                catch { }

                r++;
            }
        }

        RewriteSheet(wsAll, staffAll);
        RewriteSheet(wsUnpaid, staffUnpaid);
        // Sheet 3 comes from staff input sheets 2 and 3 (rows where A/R paid is "Unpaid")
        // Write sheet 3 with highlights for rows flagged as paid by SalesData
        // Implemented inline here to access ClosedXML styling
        var headerRow = 1;
        var lastColHist = Math.Max(12, wsHist.LastColumnUsed()?.ColumnNumber() ?? 12);
        var headersHist = new List<string>();
        for (int c = 1; c <= lastColHist; c++) headersHist.Add(wsHist.Cell(headerRow, c).GetString());
        wsHist.Clear();
        for (int c = 1; c <= headersHist.Count; c++) wsHist.Cell(headerRow, c).Value = headersHist[c - 1];
        var rr = headerRow + 1;
        foreach (var s in unpaidFromInput)
        {
            var maxCols = Math.Max(headersHist.Count, s.Cells?.Count ?? 0);
            for (int c = 1; c <= maxCols; c++)
            {
                if (c == 1)
                {
                    var custVal = !string.IsNullOrWhiteSpace(s.Customer)
                        ? s.Customer
                        : (s.Cells != null && s.Cells.Count >= 1 ? s.Cells[0] : string.Empty);
                    wsHist.Cell(rr, c).Value = custVal;
                    continue;
                }
                if (c == 2 && s.TransactionDate != default)
                {
                    wsHist.Cell(rr, c).Value = s.TransactionDate;
                    continue;
                }
                var val = s.Cells != null && s.Cells.Count >= c ? s.Cells[c - 1] : string.Empty;
                wsHist.Cell(rr, c).Value = val;
            }

            if (s.IsPaidBySales)
            {
                try { wsHist.Cell(rr, 12).Value = "Paid"; } catch { }
                try { wsHist.Range(rr, 1, rr, 12).Style.Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.LightGreen); } catch { }
            }
            else
            {
                try { wsHist.Range(rr, 1, rr, 12).Style.Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.FromHtml("#FFDAB9")); } catch { }
            }

            rr++;
        }

        staffWb.SaveAs(outputPath);
    }

    class SaleRecord
    {
        public string Customer;
        public DateTime TransactionDate;
        public string InvoiceNumber;
        public double Amount;
        public string SalesRep;
        public string ARPaid;
        public List<string> Cells = new List<string>();
        public bool IsPaidBySales = false;
    }

    static List<SaleRecord> ReadSalesRecords(string salesPath)
    {
        var list = new List<SaleRecord>();
        using var wb = new XLWorkbook(salesPath);
        var ws = wb.Worksheets.First();
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;

        string currentCustomer = null;
        for (int r = 1; r <= lastRow; r++)
        {
            var a = ws.Cell(r, 1).GetString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(a) && !a.StartsWith("Total for", StringComparison.OrdinalIgnoreCase))
            {
                var d = ws.Cell(r, 4).GetString();
                var b = ws.Cell(r, 2).GetString();
                if (string.IsNullOrWhiteSpace(d) && string.IsNullOrWhiteSpace(b))
                {
                    currentCustomer = a;
                    continue;
                }
            }
            if (!string.IsNullOrEmpty(a) && a.StartsWith("Total for", StringComparison.OrdinalIgnoreCase))
            {
                currentCustomer = null;
                continue;
            }

            if (currentCustomer == null) continue;

            var dateStr = ws.Cell(r, 2).GetString();
            if (!DateTime.TryParseExact(dateStr, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                try { dt = ws.Cell(r, 2).GetDateTime(); }
                catch { continue; }
            }

            var inv = ws.Cell(r, 4).GetString();
            double amt = 0;
            var amtStr = ws.Cell(r, 9).GetString();
            if (!double.TryParse(amtStr, NumberStyles.Any, CultureInfo.InvariantCulture, out amt))
            {
                try { amt = ws.Cell(r, 9).GetDouble(); } catch { amt = 0; }
            }
            var rep = ws.Cell(r, 11).GetString();
            var ar = ws.Cell(r, 12).GetString();

            // capture all cell text values for this row
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 12;
            var cells = new List<string>();
            for (int c = 1; c <= lastCol; c++) cells.Add(ws.Cell(r, c).GetString());

            list.Add(new SaleRecord { Customer = currentCustomer, TransactionDate = dt, InvoiceNumber = inv, Amount = amt, SalesRep = rep, ARPaid = ar, Cells = cells });
        }

        return list;
    }

    static List<SaleRecord> ExtractUnpaidFromWorksheet(IXLWorksheet ws)
    {
        var list = new List<SaleRecord>();
        if (ws == null) return list;
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        for (int r = 2; r <= lastRow; r++)
        {
            var ar = ws.Cell(r, 12).GetString();
            if (string.Equals(ar, "Paid", StringComparison.OrdinalIgnoreCase)) continue;

            var cust = ws.Cell(r, 1).GetString();
            var dateStr = ws.Cell(r, 2).GetString();
            DateTime dt;
            if (!DateTime.TryParseExact(dateStr, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
            {
                try { dt = ws.Cell(r, 2).GetDateTime(); }
                catch { continue; }
            }
            var inv = ws.Cell(r, 4).GetString();
            double amt = 0;
            var amtStr = ws.Cell(r, 9).GetString();
            if (!double.TryParse(amtStr, NumberStyles.Any, CultureInfo.InvariantCulture, out amt))
            {
                try { amt = ws.Cell(r, 9).GetDouble(); } catch { amt = 0; }
            }
            var rep = ws.Cell(r, 11).GetString();

            // capture all cell text values for this row
            var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 12;
            var cells = new List<string>();
            for (int c = 1; c <= lastCol; c++) cells.Add(ws.Cell(r, c).GetString());

            list.Add(new SaleRecord { Customer = cust, TransactionDate = dt, InvoiceNumber = inv, Amount = amt, SalesRep = rep, ARPaid = ar, Cells = cells });
        }
        return list;
    }

    static string ExtractYearMonthFromFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return string.Empty;
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        if (string.IsNullOrEmpty(baseName)) return string.Empty;
        for (int i = 0; i + 6 <= baseName.Length; i++)
        {
            var part = baseName.Substring(i, 6);
            if (int.TryParse(part, out _)) return part;
        }
        return string.Empty;
    }

    static bool TryParseYearMonth(string ym, out int year, out int month)
    {
        year = 0; month = 0;
        if (string.IsNullOrWhiteSpace(ym) || ym.Length != 6) return false;
        if (!int.TryParse(ym.Substring(0, 4), out year)) return false;
        if (!int.TryParse(ym.Substring(4, 2), out month)) return false;
        return month >= 1 && month <= 12;
    }

    // Map of staff code -> full path to their Excel file
    static Dictionary<string, string> FindStaffFiles(string inputDir, string mainSalesFileName)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!Directory.Exists(inputDir)) return map;

        var files = Directory.GetFiles(inputDir, "*.xlsx");
        foreach (var f in files)
        {
            var name = Path.GetFileName(f);
            if (string.Equals(name, mainSalesFileName, StringComparison.OrdinalIgnoreCase)) continue;

            var code = ExtractStaffCodeFromFileName(name);
            if (string.IsNullOrWhiteSpace(code)) continue;
            if (!map.ContainsKey(code)) map[code] = f;
        }
        return map;
    }

    static string ExtractStaffCodeFromFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName)) return string.Empty;
        // Remove extension
        var baseName = Path.GetFileNameWithoutExtension(fileName);
        if (string.IsNullOrEmpty(baseName)) return string.Empty;

        // Staff code is the leading run of letters before digits or other separators
        var chars = baseName.ToCharArray();
        var idx = 0;
        while (idx < chars.Length && char.IsLetter(chars[idx])) idx++;
        var code = baseName.Substring(0, idx);
        return code;
    }

    static void ShowUserManual()
    {
        Console.WriteLine();
        var prev = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("=== COMMISSION TOOL ===");
        Console.ForegroundColor = prev;
        Console.WriteLine("User manual:");
        Console.WriteLine(" - Place the master sales file named 'SalesData.xlsx' into the 'input' folder next to the executable.");
        Console.WriteLine(" - Add staff input files (Excel .xlsx) into the same 'input' folder. Each staff file should start with the staff code, e.g. 'DMB202601.xlsx'.");
        Console.WriteLine(" - The program will use configured staff offsets to compute the source month and will produce outputs into the 'output' folder alongside the executable.");
        Console.WriteLine(" - Output files are named '<StaffCode><yyyyMM>.xlsx' and contain three sheets: 'Sales yyyyMM', 'Unpaid yyyyMM', and 'Outstanding'.");
        Console.WriteLine(" - Sheet 1 and 2 are taken from SalesData for the target month; sheet 3 is built from staff input unpaid rows (with updates from SalesData).");
        Console.WriteLine(" - Unpaid rows are highlighted with light orange; rows detected as paid in SalesData are set to 'Paid' and highlighted light green.");
        Console.WriteLine();
    }
}
