# DnaHelper
A tiny tool for Excel-DNA and SharpDevelop

# Why
[SharpDevelop](https://sourceforge.net/projects/sharpdevelop/) is a very nice IDE for unprofessional programmer. It is only 13.8 MB with code hinting and Winform designer.

[Excel-DNA](https://excel-dna.net/) is also a very nice addin. With it, I can use C# instead of old VB6.0 in Excel.

To make the conversion be easy, I make this script.

# How
0. Change the `XLL_PATH` in the script to yours.
1. Create solution in SharpDevelop.
2. Copy this script to `*.sln` level and run it.
3. Modify the initial `*.xml` file.
4. Key your codes in SharpDevelop.
5. Run it.
6. Find `*.dna` and `*.xll` in `Distribution` -- They are what you need.

# Note
Due to my poor programming skill, you must do these:
1. Involve Excel namespace and rename `Excel.Application` to `XlApp`.
2. Delare `app`, `wb` and `ws` for `Excel.Application`, `Workbook` and `Worksheet`.
3. Init them in constructor by
``` csharp
    app = new XlApp();
    app.Visible = true;
    wb = (Workbook)app.Workbooks.Add();
    ws = (Worksheet)wb.Worksheets.Add();
```
