using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

class Program
{
    static string InputDir => Path.Combine(AppContext.BaseDirectory, "input");
    static string OutputDir => Path.Combine(AppContext.BaseDirectory, "output");

    static void Main()
    {
        Directory.CreateDirectory(InputDir);
        Directory.CreateDirectory(OutputDir);

        ShowUserManual();

        var subdirs = Directory.GetDirectories(InputDir);
        if (subdirs.Length == 0)
        {
            Console.WriteLine("No subfolders found in input folder. Place folders with files into 'input' and re-run.");
            return;
        }

        foreach (var dir in subdirs)
        {
            var folderName = Path.GetFileName(dir);
            var files = Directory.GetFiles(dir, "*", SearchOption.AllDirectories);

            // Collect unique CL codes (e.g., CL12345) and keep first relative path encountered
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var re = new Regex(@"\bCL(\d+)\b", RegexOptions.IgnoreCase);
            foreach (var f in files)
            {
                var fn = Path.GetFileName(f);
                var m = re.Match(fn);
                if (!m.Success) continue;
                var code = "CL" + m.Groups[1].Value;
                if (!map.ContainsKey(code))
                {
                    map[code] = Path.GetRelativePath(dir, f);
                }
            }

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Files");
            ws.Cell(1, 1).Value = "CLCode";
            ws.Cell(1, 2).Value = "RelativePath";

            var r = 2;
            foreach (var kv in map.OrderBy(k => k.Key))
            {
                ws.Cell(r, 1).Value = kv.Key;
                ws.Cell(r, 2).Value = kv.Value;
                r++;
            }

            var outPath = Path.Combine(OutputDir, $"{folderName}_CL_List.xlsx");
            wb.SaveAs(outPath);
            Console.WriteLine($"Created: {outPath} ({map.Count} unique CL codes)");
        }

        Console.WriteLine("Done.");
    }

    static void ShowUserManual()
    {
        Console.WriteLine();
        var prev = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("=== FILE NAME ORGANIZER ===");
        Console.ForegroundColor = prev;
        Console.WriteLine("User manual:");
        Console.WriteLine(" - Place subfolders of files you want to export under the 'input' folder next to the executable.");
        Console.WriteLine(" - For each subfolder, this tool will generate one Excel file named '<FolderName>_CL_List.xlsx' in the 'output' folder.");
        Console.WriteLine(" - The generated Excel contains columns: FileName, RelativePath, SizeBytes, LastModified.");
        Console.WriteLine();
    }
}
