using System;
using System.IO;

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

        Console.WriteLine("Hello, World!");
    }
}
