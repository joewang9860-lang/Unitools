# Commission Tool — User Manual

This console tool generates per-staff commission workbooks by using a master `SalesData.xlsx` file as reference and staff input templates.

Quick start

1. Place `SalesData.xlsx` (master sales records) in the `input` folder next to the tool executable.
2. Add staff input files (Excel `.xlsx`) to the same `input` folder. Each staff file must start with the staff code, for example `DMB202601.xlsx` for staff `DMB`.
3. Run the tool (or run the built executable). The program will prompt for a target year-month (`yyyyMM`). Press Enter to use the current month.
4. The tool writes outputs to the `output` folder next to the executable. Output files are named `<StaffCode><yyyyMM>.xlsx` and contain three sheets:
   - `Sales yyyyMM` — all sales for the target month (from `SalesData.xlsx`)
   - `Unpaid yyyyMM` — unpaid sales for the target month (from `SalesData.xlsx`)
   - `Outstanding` — unpaid rows carried from the staff input (sheets 2 & 3), updated where SalesData shows payment

Behavior notes

- Staff-specific month offsets are configured in the program (DMB=1, DWF=2, PPR=2). The offset determines which source month is used for naming the output.
- The tool marks rows where SalesData indicates payment: A/R status is set to `Paid` and the row is highlighted light green.
- Unpaid rows are highlighted light orange. Highlighting is applied across columns A..L.

If you need the tool to perform additional mapping or commission calculations, describe the rules and I will implement them.