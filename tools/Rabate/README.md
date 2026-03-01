# Rebate Tool — User Manual

This small console tool maps sales rows to a price table and exports a mapped sales report. Use it to compare sales prices against premium prices and produce an Excel report.

Quick start

1. Place your main sales file named `SalesData.xlsx` in the `input` folder next to the project (create the folder if needed).
2. Optionally add province price files to `input` (example: `AB.xlsx`, `BC.xlsx`). These are used to match product premiums by province.
3. Run the tool from the project directory:
   - `dotnet run --project UniTools` (or run the built executable)
4. When prompted, enter a company name. Output will be created in the `output` folder using the pattern: `<Company>_yyyyMMdd_HHmmss.xlsx`.

Notes

- The program prints a short user manual on startup and highlights `=== REBATE TOOL ===` in the console.
- The tool does not perform rebate calculations — it maps products to prices and exports the mapped sales table.
- If the input file is missing, the tool lists files found in the `input` folder to help locate/correct the file.

If you want additional details (e.g., how to add a quantity column or change header row), tell me what you'd like and I can update this README.