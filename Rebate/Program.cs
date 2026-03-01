using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

class Program
{
    // Paths are resolved relative to the running executable so the app works
    // whether run from IDE, dotnet run, or as a published executable.
    // Input/output folders are placed next to the running executable so each
    // published tool has its own `input` and `output` directories alongside the exe.
    static string InputDir => Path.Combine(AppContext.BaseDirectory, "input");
    static string OutputDir => Path.Combine(AppContext.BaseDirectory, "output");
    static string salesExcelPath => Path.Combine(InputDir, "SalesData.xlsx");
    // Output files will be created under an output subfolder per company and timestamp.

    static void Main()
    {
        // Ensure input/output directories exist (create if missing)
        Directory.CreateDirectory(InputDir);
        Directory.CreateDirectory(OutputDir);
        // Show a short user manual / usage notes before asking for company name
        ShowUserManual();

        // Ask for company name (used to create an output filename)
        Console.Write("Enter company name: ");
        var company = Console.ReadLine()?.Trim();
        if (string.IsNullOrWhiteSpace(company)) company = "Company";
        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var safeCompany = SanitizeCompanyName(company);
        // No per-company subfolder any more; output files go directly into the output folder
        // Ensure output folder exists (already created above)

        if (!File.Exists(salesExcelPath))
        {
            Console.WriteLine("Cannot find sales input file.");
            Console.WriteLine($"Expected location: {salesExcelPath}");
            // Help the user by listing files in the input directory
            var files = Directory.GetFiles(InputDir);
            if (files.Length == 0)
            {
                Console.WriteLine("The input directory is empty. Place your SalesData.xlsx (or other supported file) into the input folder.");
            }
            else
            {
                Console.WriteLine("Files found in input folder:");
                foreach (var f in files)
                    Console.WriteLine(" - " + Path.GetFileName(f));
            }
            return;
        }

        // Find XLSX files in the input folder; filenames are province codes (e.g., AB.xlsx, BC.xlsx).
        var xlsxFiles = Directory.GetFiles(InputDir, "*.xlsx").Where(p => !string.Equals(Path.GetFileName(p), Path.GetFileName(salesExcelPath), StringComparison.OrdinalIgnoreCase)).ToArray();
        var priceTablesByProvince = new Dictionary<string, Dictionary<string, double>>(StringComparer.OrdinalIgnoreCase);
        foreach (var pf in xlsxFiles)
        {
            var province = Path.GetFileNameWithoutExtension(pf);
            var table = ParsePriceTableFromExcel(pf);
            priceTablesByProvince[province] = table;
            Console.WriteLine($"Parsed {table.Count} entries from price file: {Path.GetFileName(pf)} (province: {province})");
        }

        // Merge all province tables into one fallback table (first occurrence wins)
        var mergedPriceTable = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in priceTablesByProvince)
        {
            foreach (var e in kv.Value)
            {
                if (!mergedPriceTable.ContainsKey(e.Key)) mergedPriceTable[e.Key] = e.Value;
            }
        }

        var salesRows = ReadSalesExcel(salesExcelPath, mergedPriceTable, priceTablesByProvince);

        // No rebate calculation. This program only maps sales rows to price table entries and exports results.
        // Output file placed directly in the output folder and named using the company/timestamp pattern.
        var outputPath = Path.Combine(OutputDir, $"{safeCompany}_{timestamp}.xlsx");
        ExportMappedSales(salesRows, outputPath);
        Console.WriteLine($"Mapped sales exported: {outputPath}");
    }

    // Keep input headers to reproduce the sales table in output
    static List<string> InputHeaders = new();

    // Simple model for a sales record
    class SalesRow
    {
        // Input layout: 1=City, 5=Product/Service full name, 6=Memo/Description, 8=Sales price, 9=Amount
        public string City;
        public string ProductFullName;
        public string Memo;
        public double UnitPrice; // as read from Excel (if present)
        public double SalesAmount; // Quantity * UnitPrice or from Excel
        public Dictionary<string, string> Extras = new();

        // Price table matching results
        public double? MatchedPremium;
        public string MatchedPriceKey;
        public bool MeetsPremium;
        public string MatchedProvince;
        public double? AmountIfMeetsPremium;
        // Whether this row represents a normal data row (not a marker or total row)
        public bool IsDataRow;
    }


    static List<SalesRow> ReadSalesExcel(string path, Dictionary<string, double> mergedPriceTable, Dictionary<string, Dictionary<string, double>> provincePriceTables = null)
    {
        var rows = new List<SalesRow>();
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();

        // Header is on the 5th row for the sales input file. The input layout uses
        // fixed columns: 1=City, 5=Product/Service full name, 6=Memo/Description,
        // 8=Sales price, 9=Amount. Read rows starting at row 6.
        const int headerRowNumber = 5;
        const int cityCol = 1;
        const int productCol = 5;
        const int memoCol = 6;
        const int priceCol = 8;
        const int amountCol = 9;

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? headerRowNumber;
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;

        // Read and store input headers (all columns)
        InputHeaders = new List<string>();
        for (int c = 1; c <= lastCol; c++)
        {
            var h = ws.Cell(headerRowNumber, c).GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(h)) h = "Column" + c;
            var orig = h; var idx = 1;
            while (InputHeaders.Contains(h)) { h = orig + "_" + idx; idx++; }
            InputHeaders.Add(h);
        }

        string currentCity = null; // tracked from group markers in first column
        for (int rn = headerRowNumber + 1; rn <= lastRow; rn++)
        {
            var dataRow = ws.Row(rn);
            if (dataRow == null || dataRow.IsEmpty()) continue;

            // Read raw first cell for marker detection
            var firstCellRaw = dataRow.Cell(cityCol).GetString()?.Trim() ?? string.Empty;
            var isStartMarker = !string.IsNullOrEmpty(firstCellRaw) && firstCellRaw.StartsWith("- ");
            var isEndMarker = !string.IsNullOrEmpty(firstCellRaw) && firstCellRaw.StartsWith("Total for -", StringComparison.OrdinalIgnoreCase);

            var sr = new SalesRow();

            // Store all original cell values keyed by header
            for (int c = 1; c <= lastCol; c++)
            {
                var hv = InputHeaders[c - 1];
                try { sr.Extras[hv] = dataRow.Cell(c).GetString(); } catch { sr.Extras[hv] = string.Empty; }
            }

            if (isStartMarker)
            {
                // start marker row - keep original row, update currentCity, but mark as non-data row
                sr.IsDataRow = false;
                sr.City = firstCellRaw;
                currentCity = firstCellRaw.Substring(2).Trim();
                rows.Add(sr);
                continue;
            }
            if (isEndMarker)
            {
                // total/summary row - keep original row, clear currentCity, mark as non-data
                sr.IsDataRow = false;
                sr.City = firstCellRaw;
                currentCity = null;
                rows.Add(sr);
                continue;
            }

            // Normal data row
            sr.IsDataRow = true;
            sr.City = currentCity ?? firstCellRaw;
            sr.ProductFullName = dataRow.Cell(productCol).GetString();
            sr.Memo = dataRow.Cell(memoCol).GetString();

            // Try parse numeric fields robustly
            var priceCell = dataRow.Cell(priceCol);
            double price = 0;
            if (!double.TryParse(priceCell.GetString(), out price))
            {
                try { price = priceCell.GetDouble(); } catch { price = 0; }
            }
            sr.UnitPrice = price;

            var amtCell = dataRow.Cell(amountCol);
            double amt = 0;
            if (!double.TryParse(amtCell.GetString(), out amt))
            {
                try { amt = amtCell.GetDouble(); } catch { amt = 0; }
            }
            sr.SalesAmount = amt;

            // If product present, decide whether to attempt mapping or skip based on keywords
            if (!string.IsNullOrWhiteSpace(sr.ProductFullName))
            {
                var lowerProd = sr.ProductFullName.ToLowerInvariant();
                // If product is in these categories, skip mapping and mark as not eligible
                var skipKeywords = new[] { "reducer", "moulding", "molding", "shoe", "stair" };
                if (skipKeywords.Any(k => lowerProd.Contains(k)))
                {
                    // explicit skip: no mapping, MeetsPremium remains false (default)
                    sr.MatchedPremium = null;
                    sr.MatchedPriceKey = null;
                    sr.MatchedProvince = null;
                    sr.AmountIfMeetsPremium = null;
                }
                else
                {
                    var province = ProvinceForCity(sr.City);
                    if (provincePriceTables != null && provincePriceTables.TryGetValue(province, out var provTable) && provTable.Count > 0)
                    {
                        if (TryFindPriceForProduct(sr.ProductFullName, provTable, out var matchedPrice, out var matchedKey, sr.Memo))
                        {
                            sr.MatchedPremium = matchedPrice;
                            sr.MatchedPriceKey = matchedKey;
                            sr.MatchedProvince = province;
                        }
                    }
                    // Fallback to merged table
                    if ((sr.MatchedPremium == null || sr.MatchedPremium == 0) && mergedPriceTable != null && mergedPriceTable.Count > 0)
                    {
                        if (TryFindPriceForProduct(sr.ProductFullName, mergedPriceTable, out var matchedPrice2, out var matchedKey2, sr.Memo))
                        {
                            sr.MatchedPremium = matchedPrice2;
                            sr.MatchedPriceKey = matchedKey2;
                            sr.MatchedProvince = province; // use row's province as reference when falling back
                        }
                    }
                }
            }

            // Determine MeetsPremium only for data rows with matched premium and a sales price
            if (sr.IsDataRow && sr.MatchedPremium.HasValue && sr.UnitPrice != 0)
            {
                sr.MeetsPremium = sr.UnitPrice >= sr.MatchedPremium.Value;
                if (sr.MeetsPremium)
                {
                    sr.AmountIfMeetsPremium = sr.SalesAmount;
                }
            }

            // Compute sales amount if missing — no quantity field available, so only rely on SalesAmount or UnitPrice
            // (If input contains a quantity column, map it into Extras or reintroduce a Quantity field.)

            rows.Add(sr);
        }
        return rows;
    }

    // Parse price table from an Excel file. Looks for header row containing "Product" and "Premium" and
    // reads subsequent rows to build a mapping of product name -> premium price.
    static Dictionary<string, double> ParsePriceTableFromExcel(string path)
    {
        var map = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return map;

        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();

        // Search for header row in the first 10 rows
        int headerRow = 1;
        int lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 1;
        int productCol = -1, premiumCol = -1;
        for (int r = 1; r <= Math.Min(10, lastRow); r++)
        {
            for (int c = 1; c <= lastCol; c++)
            {
                var v = ws.Cell(r, c).GetString();
                if (string.IsNullOrWhiteSpace(v)) continue;
                var low = v.ToLowerInvariant();
                if (productCol < 0 && low.Contains("product")) productCol = c;
                if (premiumCol < 0 && low.Contains("premium")) premiumCol = c;
            }
            if (productCol >= 0 && premiumCol >= 0) { headerRow = r; break; }
        }

        // Fallbacks
        if (productCol < 0) productCol = 1;
        if (premiumCol < 0) premiumCol = Math.Min(3, lastCol);

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            var prod = ws.Cell(r, productCol).GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(prod)) continue;
            var premCell = ws.Cell(r, premiumCol);
            double prem = 0;
            var s = premCell.GetString();
            if (!double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out prem))
            {
                try { prem = premCell.GetDouble(); } catch { prem = 0; }
            }
            if (prem == 0) continue;
            var name = Regex.Replace(prod, "\\s+", " ");
            if (!map.ContainsKey(name)) map[name] = prem;
        }

        return map;
    }

    

    static bool TryFindPriceForProduct(string productName, Dictionary<string, double> priceTable, out double price, out string matchedKey, string fallbackText = null)
    {
        price = 0; matchedKey = null;
        if (string.IsNullOrWhiteSpace(productName) || priceTable == null || priceTable.Count == 0) return false;

        // Exact match (case-insensitive)
        if (priceTable.TryGetValue(productName, out price)) { matchedKey = productName; return true; }
        var keyLower = productName.Trim();
        var exact = priceTable.Keys.FirstOrDefault(k => string.Equals(k.Trim(), keyLower, StringComparison.OrdinalIgnoreCase));
        if (exact != null) { price = priceTable[exact]; matchedKey = exact; return true; }

        // Tokenize source name
        var srcTokens = Tokenize(productName).Where(t => t.Length > 0).ToList();
        if (srcTokens.Count == 0) return false;

        // Prepare modifier/stopword lists
        // Tokens that indicate a style/variant (not the base product family)
        var modifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "evolved","elite","pro","plus","max","advanced","deluxe","standard","lite","mini",
            // style/shape tokens
            "chevron","wpc","chev","herringbone","plank","tile","mosaic","strip","rustic","handscraped","evo"
        };
        var stopwords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "the","and","for","of","a","an","by","with","-"
        };

        // classify tokens: strong = not modifier/stopword
        var srcStrong = srcTokens.Where(t => !modifiers.Contains(t) && !stopwords.Contains(t)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var srcHasVariant = srcTokens.Intersect(modifiers, StringComparer.OrdinalIgnoreCase).Any();

        double bestScore = double.NegativeInfinity;
        double secondBestScore = double.NegativeInfinity;
        string bestKeyStr = null;

        foreach (var kv in priceTable)
        {
            var cand = kv.Key;
            var candTokens = Tokenize(cand).Where(t => t.Length > 0).ToList();
            if (candTokens.Count == 0) continue;

            // quick prefilter: require at least one strong token overlap
            var overlap = srcStrong.Intersect(candTokens, StringComparer.OrdinalIgnoreCase).ToList();
            if (overlap.Count == 0) continue;

            // detect if candidate is a variant (chevron, evolved, etc.)
            var candHasVariant = candTokens.Intersect(modifiers, StringComparer.OrdinalIgnoreCase).Any();
            // If source does NOT mention variants but candidate is a variant, skip it.
            if (!srcHasVariant && candHasVariant)
            {
                // candidate is a style/variant (e.g., chevron) while source has no variant token -> skip
                continue;
            }

            // scoring signals
            var matchedStrong = overlap.Count;
            var matchedTokens = srcTokens.Intersect(candTokens, StringComparer.OrdinalIgnoreCase).Count();
            var srcTokenCount = srcTokens.Distinct(StringComparer.OrdinalIgnoreCase).Count();
            var candTokenCount = candTokens.Distinct(StringComparer.OrdinalIgnoreCase).Count();

            // penalty for candidate containing modifier tokens that are not in source
            var candMods = candTokens.Where(t => modifiers.Contains(t)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var extraMods = candMods.Where(m => !srcTokens.Contains(m, StringComparer.OrdinalIgnoreCase)).Count();

            // extra tokens count
            var extraTokens = candTokenCount - matchedTokens;

            // normalized levenshtein as tie-breaker (1-bestNormalizedDistance)
            var normSource = NormalizeForMatch(productName);
            var normCand = NormalizeForMatch(cand);
            var lev = LevenshteinDistance(normSource, normCand);
            var maxLen = Math.Max(normSource.Length, normCand.Length);
            double levScore = 0;
            if (maxLen > 0) levScore = 1.0 - (lev / (double)maxLen);

            // Compose score (weights chosen heuristically)
            double score = 0;
            score += matchedStrong * 40; // strong token matches are most important
            score += matchedTokens * 10; // other token matches
            score += (matchedTokens / (double)srcTokenCount) * 30; // coverage of source
            score += levScore * 20; // small boost from string similarity
            score -= extraMods * 60; // heavy penalty for variant modifiers (evolved/elite)
            score -= extraTokens * 2; // small penalty for extra words in candidate

            // Prefer shorter candidates if scores close
            score -= (candTokenCount - srcTokenCount) * 1;

            if (score > bestScore)
            {
                secondBestScore = bestScore;
                bestScore = score; bestKeyStr = kv.Key;
            }
            else if (score > secondBestScore)
            {
                secondBestScore = score;
            }
        }
        // Accept only if bestScore passes threshold
        if (bestKeyStr != null && bestScore > 40)
        {
            // If there's a tie (best and secondBest equal or very close), try fallbackText if provided
            if (Math.Abs(bestScore - secondBestScore) <= 0.0001 && !string.IsNullOrWhiteSpace(fallbackText))
            {
                // Try matching using fallback text (memo/description) once
                if (TryFindPriceForProduct(fallbackText, priceTable, out var fbPrice, out var fbKey, null))
                {
                    price = fbPrice; matchedKey = fbKey; return true;
                }
                // if fallback didn't help, fall through to pick bestKeyStr
            }

            price = priceTable[bestKeyStr]; matchedKey = bestKeyStr; return true;
        }

        return false;
    }

    static List<string> Tokenize(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return new List<string>();
        // split on non-word characters, keep alphanumerics
        var parts = Regex.Split(s.ToLowerInvariant(), "[^a-z0-9]+").Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
        return parts;
    }

    static string NormalizeForMatch(string s)
    {
        return Regex.Replace(s ?? string.Empty, "[^a-z0-9]", "", RegexOptions.IgnoreCase).ToLowerInvariant();
    }

    // Levenshtein distance implementation
    static int LevenshteinDistance(string a, string b)
    {
        if (a == null) a = string.Empty; if (b == null) b = string.Empty;
        var la = a.Length; var lb = b.Length;
        if (la == 0) return lb; if (lb == 0) return la;
        var d = new int[la + 1, lb + 1];
        for (int i = 0; i <= la; i++) d[i, 0] = i;
        for (int j = 0; j <= lb; j++) d[0, j] = j;
        for (int i = 1; i <= la; i++)
        {
            for (int j = 1; j <= lb; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        }
        return d[la, lb];
    }

    static string ProvinceForCity(string city)
    {
        if (string.IsNullOrWhiteSpace(city)) return "BC";
        var c = city.ToLowerInvariant();
        if (c.Contains("edmonton") || c.Contains("calgary")) return "AB";
        return "BC";
    }

    // Find the project directory by walking up parent folders looking for a .csproj file.
    static string FindProjectDirectory()
    {
        // Start from current directory then fallback to the runtime base directory
        var startDirs = new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory };
        foreach (var start in startDirs)
        {
            var found = SearchUpForCsproj(start);
            if (!string.IsNullOrEmpty(found)) return found;
        }
        // As a last resort, return current directory
        return Directory.GetCurrentDirectory();
    }

    static string SearchUpForCsproj(string startPath)
    {
        try
        {
            var dir = new DirectoryInfo(Path.GetFullPath(startPath));
            while (dir != null)
            {
                var files = dir.GetFiles("*.csproj");
                if (files != null && files.Length > 0) return dir.FullName;
                dir = dir.Parent;
            }
        }
        catch { }
        return null;
    }

    static string SanitizeCompanyName(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "Company";
        var invalid = Path.GetInvalidFileNameChars();
        var cleaned = new string(s.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
        cleaned = Regex.Replace(cleaned, "\\s+", "_");
        return cleaned;
    }

    static void ShowUserManual()
    {
        var prevColor = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("=== REBATE TOOL ===");
        Console.ForegroundColor = prevColor;
        Console.WriteLine();
        Console.WriteLine("User manual:");
        Console.WriteLine(" - Place your main sales file named 'SalesData.xlsx' into the 'input' folder next to the project.");
        Console.WriteLine(" - Optional province price files may be added to 'input' (e.g. 'AB.xlsx', 'BC.xlsx').");
        Console.WriteLine(" - Output will be created in the 'output' folder with filename pattern: <Company>_yyyyMMdd_HHmmss.xlsx");
        Console.WriteLine(" - The program maps products to prices and exports the mapped sales table.\n");
    }

    static string GetColumnLetter(int col)
    {
        // Convert 1-based column index to Excel column letter
        var dividend = col;
        var columnName = string.Empty;
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }

    static void ExportMappedSales(List<SalesRow> rows, string output)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("MappedSales");

        // Find the product/service header index so we can place MatchedPriceKey next to it
        var prodIdx = -1;
        for (int i = 0; i < InputHeaders.Count; i++)
        {
            var h = InputHeaders[i] ?? string.Empty;
            var low = h.ToLowerInvariant();
            if (low.Contains("product") || low.Contains("service"))
            {
                prodIdx = i;
                break;
            }
        }

        // Write headers, inserting MatchedPriceKey immediately after product/service when found
        var col = 1;
        for (int i = 0; i < InputHeaders.Count; i++)
        {
            ws.Cell(1, col).Value = InputHeaders[i];
            if (i == prodIdx)
            {
                ws.Cell(1, col + 1).Value = "MatchedPriceKey";
                col += 2;
            }
            else
            {
                col += 1;
            }
        }

        // Append MeetsPremium, MatchedPremium, MatchedProvince, and AmountIfMeetsPremium at the end
        ws.Cell(1, col++).Value = "MeetsPremium";
        ws.Cell(1, col++).Value = "MatchedPremium";
        ws.Cell(1, col++).Value = "MatchedProvince";
        ws.Cell(1, col++).Value = "AmountIfMeetsPremium";

        var r = 2;
        foreach (var row in rows)
        {
            // write original row cells in same header order, inserting MatchedPriceKey after product/service
            var ccol = 1;
            for (int i = 0; i < InputHeaders.Count; i++)
            {
                var hv = InputHeaders[i];
                row.Extras.TryGetValue(hv, out var val);
                ws.Cell(r, ccol).Value = val;
                if (i == prodIdx)
                {
                    ws.Cell(r, ccol + 1).Value = row.MatchedPriceKey ?? string.Empty;
                    ccol += 2;
                }
                else
                {
                    ccol += 1;
                }
            }

            // write MeetsPremium, MatchedPremium, MatchedProvince, AmountIfMeetsPremium at the end
            ws.Cell(r, ccol++).Value = row.IsDataRow ? (row.MeetsPremium ? "Yes" : "No") : string.Empty;
            if (row.MatchedPremium.HasValue)
                ws.Cell(r, ccol++).Value = row.MatchedPremium.Value;
            else
                ws.Cell(r, ccol++).Value = string.Empty;
            ws.Cell(r, ccol++).Value = row.MatchedProvince ?? string.Empty;
            if (row.AmountIfMeetsPremium.HasValue)
                ws.Cell(r, ccol++).Value = row.AmountIfMeetsPremium.Value;
            else
                ws.Cell(r, ccol++).Value = string.Empty;

            r++;
        }

        // Calculate total for AmountIfMeetsPremium (sum over rows 2..r-1)
        var totalRow = r;
        // AmountIfMeetsPremium column index is the last header column we added minus 1
        var amountColIndex = col - 1;
        // Write label and formula
        ws.Cell(totalRow, amountColIndex - 1).Value = "Total";
        var colLetter = GetColumnLetter(amountColIndex);
        ws.Cell(totalRow, amountColIndex).FormulaA1 = $"=SUM({colLetter}2:{colLetter}{r - 1})";

        wb.SaveAs(output);
    }
}