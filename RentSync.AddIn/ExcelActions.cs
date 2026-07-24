using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using RentSync.Core.Models;
using RentSync.Core.Services;

namespace RentSync.AddIn
{
    /// <summary>
    /// All Excel Object Model access lives here, and only ever runs on the
    /// UI thread (callers guarantee that via the ribbon pattern).
    ///
    /// Performance rule demonstrated: ONE Range.Value2 assignment with an
    /// object[,] array instead of cell-by-cell writes. For a 10,000-row rent
    /// roll that is the difference between ~0.1s and minutes — each cell
    /// write is a COM round-trip.
    /// </summary>
    internal static class ExcelActions
    {
        public static void WriteRentRoll(IReadOnlyList<RentRollRecord> records,
            string sheetName = "Rent Roll")
        {
            var sheet = GetOrCreateSheet(sheetName);
            sheet.Cells.Clear();

            var headers = new object[1, 8]
            {
                { "Property ID", "Property", "Unit", "Tenant",
                  "Monthly Rent", "Area (sqm)", "Lease Start", "Lease End" }
            };

            var data = new object[records.Count, 8];
            for (var i = 0; i < records.Count; i++)
            {
                var r = records[i];
                data[i, 0] = r.PropertyId;
                data[i, 1] = r.PropertyName;
                data[i, 2] = r.UnitId;
                data[i, 3] = r.TenantName;
                data[i, 4] = r.MonthlyRent;
                data[i, 5] = r.AreaSqm;
                data[i, 6] = r.LeaseStart.ToDateTime(TimeOnly.MinValue);
                data[i, 7] = r.LeaseEnd.ToDateTime(TimeOnly.MinValue);
            }

            WriteBlock(sheet, headers, data);
            FormatColumns(sheet, currency: new[] { 5 }, dates: new[] { 7, 8 });
            ((Excel.Worksheet)sheet).Activate();
        }

        public static void WriteSummaries(IReadOnlyList<PropertySummary> summaries)
        {
            var sheet = GetOrCreateSheet("Property Summary");
            sheet.Cells.Clear();

            var headers = new object[1, 6]
            {
                { "Property ID", "Property", "Units",
                  "Total Monthly Rent", "Total Area (sqm)", "Avg Rent / sqm" }
            };

            var data = new object[summaries.Count, 6];
            for (var i = 0; i < summaries.Count; i++)
            {
                var s = summaries[i];
                data[i, 0] = s.PropertyId;
                data[i, 1] = s.PropertyName;
                data[i, 2] = s.UnitCount;
                data[i, 3] = s.TotalMonthlyRent;
                data[i, 4] = s.TotalAreaSqm;
                data[i, 5] = s.AvgRentPerSqm;
            }

            WriteBlock(sheet, headers, data);
            FormatColumns(sheet, currency: new[] { 4, 6 }, dates: Array.Empty<int>());
            sheet.Activate();
        }

        // ---------------- helpers ----------------

        private static Excel.Worksheet GetOrCreateSheet(string name)
        {
            var app = Globals.ThisAddIn.Application;
            var book = app.ActiveWorkbook ?? app.Workbooks.Add();

            foreach (Excel.Worksheet ws in book.Worksheets)
                if (ws.Name == name) return ws;

            var created = (Excel.Worksheet)book.Worksheets.Add(
                After: book.Worksheets[book.Worksheets.Count]);
            created.Name = name;
            return created;
        }

        private static void WriteBlock(Excel.Worksheet sheet,
            object[,] headers, object[,] data)
        {
            var cols = headers.GetLength(1);

            var headerRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, cols]];
            headerRange.Value2 = headers;
            headerRange.Font.Bold = true;

            if (data.GetLength(0) > 0)
            {
                var dataRange = sheet.Range[
                    sheet.Cells[2, 1],
                    sheet.Cells[1 + data.GetLength(0), cols]];
                dataRange.Value2 = data; // single COM round-trip
            }

            sheet.Columns.AutoFit();
        }

        private static void FormatColumns(Excel.Worksheet sheet,
            int[] currency, int[] dates)
        {
            foreach (var c in currency)
                ((Excel.Range)sheet.Columns[c]).NumberFormat = "#,##0.00";
            foreach (var d in dates)
                ((Excel.Range)sheet.Columns[d]).NumberFormat = "yyyy-mm-dd";
        }
    }
}
