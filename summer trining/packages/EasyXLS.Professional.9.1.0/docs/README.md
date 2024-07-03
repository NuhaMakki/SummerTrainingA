# EasyXLS Excel Library for .NET
![EasyXLS](https://raw.githubusercontent.com/EasyXLS/EasyXLS.samples/main/images/easyxls-excel-library-component-small.jpg "EasyXLS")

## About
[EasyXLS](https://www.easyxls.com) is an API to import, export and convert MS Excel files without Microsoft Excel installed from C#, VB.NET, C++.NET, ASP.NET web pages, Windows Forms or Windows Services.   

[EasyXLS for .NET](https://www.easyxls.com/net-excel-library) is used for creating, reading and manipulating MS Excel files, including support for advanced features like formatting, formulas, charts, macros, images and pivot tables.  

EasyXLS library  is designed to be easy to use, with a straightforward API and comprehensive documentation.  It is also highly performant, with optimized algorithms for handling large Excel files.

## EasyXLS Features

* [Create Excel files](https://www.easyxls.com/manual/basics/create-excel-file.html), new files or from Excel templates
* [Import Excel data](https://www.easyxls.com/manual/FAQ/import-excel-in-dot-net.html), modify Excel file and resave the file
* [Convert Excel files](https://www.easyxls.com/manual/basics/convert-html-to-excel.html), between MS Excel file formats (XLSX, XLSM, XLSB, XLS and SpreadsheetML) and also text formats (HTML, XML, CSV and TXT).
* [Format cells](https://www.easyxls.com/manual/basics/format-excel-cells.html), rows, and columns with background, foreground, fonts, borders, alignments, number and date formats and other formatting elements. Conditional formatting is also supported.
* Multiple sheets 
* Complex [formulas](https://www.easyxls.com/manual/basics/import-export-excel-formulas.html) and functions, named ranges and formulas, [formula calculation engine](https://www.easyxls.com/manual/basics/excel-calculation-engine.html) included
* Hyperlinks, comments and [images](https://www.easyxls.com/manual/basics/excel-image-import-export.html)
* Data validation for cell values, including drop-down selection
* Print options and page breaks
* Group rows and columns, split and freeze panes, filter and auto-filter
* [Charts](https://www.easyxls.com/manual/basics/excel-chart-inside-sheet.html) with various supported types and formatting
* [Pivot tables](https://www.easyxls.com/manual/basics/excel-pivot-table.html) and pivot charts
* [Encryption and password protection](https://www.easyxls.com/manual/basics/password-protected-excel-file.html) to protect the Excel file from unauthorized access, [protect sheet data](https://www.easyxls.com/manual/basics/excel-protect-sheet.html) inside sheet from altering
* VB code and [macros](https://www.easyxls.com/manual/basics/excel-macros-vba-project.html) preservation
* File properties with details about the author and company that generated the Excel file or custom properties
* Import/export from data structures, SQL databases, lists of data, [export DataTable to Excel](https://www.easyxls.com/manual/FAQ/export-datatable-to-excel.html), [import Excel to DataTable](https://www.easyxls.com/manual/FAQ/import-excel-to-datatable.html), import/export from GridView or DataGridView

## Supported File Formats
**MS Excel Open XML:** XLSX, XLSM  
**MS Excel Binary:** XLSB, XLS  
**XML:** SpreadsheetML, XML specific schema  
**Web:** HTML, mso attributes, CSS  
**Text:** CSV, TXT  

## Getting Started

### **Step 1**: Download and install EasyXLS nuget package  

Download EasyXLS from upper link and execute below line in Package Manager Console from Visual Studio:  
```Install-Package easyxls.professional```  
or search for EasyXLS in NuGet Package Manager in Visual Studio and install.

### **Step 2**: License file setup   

Generate a trial license file from [EasyXLS trials](https://www.easyxls.com/trials#dotnet) page. The trial license is valid for 30-days.  
Setup the license file into your project using the [guidelines](https://www.easyxls.com/manual/licensing/license-setup.html).

### **Step 3**: Start coding

You can execute the code below in C# to create an Excel file having two sheets and a value set in "A1" cell.

```
// Create an instance of the class that creates Excel files, having two sheets
ExcelDocument workbook = new ExcelDocument(2);

// Get the table of data for the first sheet
ExcelTable xlsTable = ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

// Add data in A1 cell
xlsTable.easy_getCell("A1").setValue("Hello world!");

// Create Excel file
workbook.easy_WriteXLSXFile("C:\\Samples\\Excel.xlsx");
```

## Documentation
EasyXLS website provides detailed information on how to use the various features and functionalities of the EasyXLS library, including a complete [User Guide](https://www.easyxls.com/manual), [tutorials](https://www.easyxls.com/manual/tutorials/easyxls-tutorials.html), [demos](https://www.easyxls.com/net-excel-library#demo), and [API documentation](https://www.easyxls.com/manual/API_Documentation/index.html).

---
[Product Page](https://www.easyxls.com/net-excel-library) / [Trial License](https://www.easyxls.com/trials#dotnet) / [Getting Started](https://www.easyxls.com/manual/getting-started/easyxls-dot-net-excel-library.html) / [Tutorials](https://www.easyxls.com/tutorials) / [Documentation](https://www.easyxls.com/manual) / [Knowledge Base](https://www.easyxls.com/net-excel-library#faq) / [Support](https://www.easyxls.com/ask-a-question)