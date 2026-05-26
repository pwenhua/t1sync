# T1Sync C# — Visual Studio Setup

NuGet install ClosedXML

NuGet install Microsoft.Data.SqlClient

#
For Claude Code

I am thinking unify excel and csv, Excel Application can open online csv?

Yes — Excel.Application.Workbooks.Open on a SharePoint/OneDrive URL works for .csv too (same Office auth path as .xlsx), but Excel applies type inference on import (strips leading zeros, auto-converts dates), which is exactly what your existing ReadCsv avoids. To keep CSVs lossless you'd need Workbooks.OpenText with every column flagged as Text (xlTextFormat = 2), which is workable but more code per call.

My recommendation: for unifying online support, don't route CSVs through Excel — instead, when an online URL ends in .csv, download it to a temp file (HttpClient with the existing SharePoint auth cookie, or Excel.Workbooks.Open(...).SaveAs(localCsv); .Close(false) as a quick-and-dirty path) and feed it through the existing ReadCsv. That keeps the type-safety story uniform: xlsx → Excel COM (cells as values), csv → string parser (cells as strings), with the only branching being on extension. The tradeoff is one extra hop for online CSVs vs. opening them directly in Excel.

Want me to sketch that out?