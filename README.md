# Google Sheets API Integration in C# (.NET)

This C# class allows seamless interaction with **Google Sheets** using the **Google Sheets API v4** via the `Google.Apis.Sheets.v4` package. It's designed to help you perform common spreadsheet operations such as reading, writing, updating, and filtering rows — both online and offline.

---

## 🚀 Features

- ✅ Read full datasets from a Google Sheet
- 🔍 Search by ID or any value (returns row/column info)
- ✍️ Update specific cells or entire rows
- ➕ Append new rows
- ❌ Flag rows as deleted using a "delete" column
- 🧼 Filter and delete logic rows locally

---

## 🔧 Prerequisites

- [.NET 6+](https://dotnet.microsoft.com/en-us/download)
- NuGet package:  
  ```bash
  dotnet add package Google.Apis.Sheets.v4
