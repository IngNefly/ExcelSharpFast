# ExcelSharpFast
Lightweight library for reading and writing massive Excel (.xlsx) files using OpenXML with low memory consumption.

# ExcelSharpFast 📊⚡

**ExcelSharpFast** is a lightweight and high-performance C# library for reading and writing massive `.xlsx` files using [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK), without loading the entire file into memory. Ideal for large-scale data processing, validation, and export scenarios.

> 🚀 Built for speed and scalability  
> 💼 No Excel installation or commercial license required

---

## ✨ Features

- ✅ Stream `.xlsx` files with millions of rows using minimal memory
- ✅ Filter and validate records as you read them
- ✅ Write new `.xlsx` files from scratch (no templates required)
- ✅ Automatically detects and writes all column headers
- 📦 Perfect for ETLs, microservices, background jobs, and batch processors

---

## 📦 Installation

```bash
dotnet add package ExcelSharpFast
```
---

## 🛠️ Quick Usage

1. Read an Excel file with validation

```csharp
var validator = new RequiredFieldsValidator(new[] { "Column1", "Column1" });

var reader = new ExcelReaderService();
var rows = reader.ReadRows("input.xlsx", headerRowIndex: 1, validator: validator);

foreach (var row in rows)
{
    Console.WriteLine($"Row {row.RowIndex}: {string.Join(", ", row.Data)}");
}
```

2. Write a new Excel file from filtered rows

```csharp
var writer = new ExcelWriterService();
writer.WriteExcel("output.xlsx", rows.ToList());
```

---

## 📂 Data Structure

```csharp
public class ExcelRow
{
    public int RowIndex { get; set; }
    public Dictionary<string, string> Data { get; set; }
}
```

---

## 🧪 Custom Validation Logic

You can create your own validation rules by implementing the IValidator interface:

```csharp
public class MyCustomValidator : IValidator
{
    public bool IsValid(ExcelRow row)
    {
        return row.Data.TryGetValue("Status", out var status) && status == "Active";
    }
}
```

---
