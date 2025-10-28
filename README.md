# 🧮 SheetUtils for Google Apps Script

A powerful, self-contained utility class to handle Google Sheets like a
lightweight database --- with filtering, updating, inserting, deleting,
highlighting, and logical conditions.

------------------------------------------------------------------------

## ⚙️ Features

✅ Read entire sheets into structured JSON-like arrays\
✅ Filter rows using **conditions**, including: - direct values
(`{ status: "Active" }`) - functions (`{ age: a => a > 18 }`) - regular
expressions (`{ name: /^A/ }`) - nested logic with `AND`, `OR`, `NOT` -
arrays of values (`{ status: ["Open", "Pending"] }`) - direct `rowID`
lookup (single or list)

✅ Safely handle **dates** (no timezone shift)\
✅ Convert numbers automatically\
✅ Insert, update, and delete rows\
✅ Highlight or clear rows directly in the sheet

------------------------------------------------------------------------

## 📦 Class Overview

``` js
class SheetUtils {
  static getTable(sheetObj)
  static getRowsFromTable(table, conditions, columnNames)
  static getValuesFromTable(table, conditions, columnName)
  static setValues(sheetObj, conditions, data)
  static insertRow(sheetObj, data)
  static deleteRows(sheetObj, conditions)
  static getHeaders(sheetObj)
  static stringifyObj(obj)
  static logRows(rows)
  static getTimeStamp()
  static highlightRows(sheetObj, conditions, color)
  static clearHighlight(sheetObj, highlightObj)
}
```

------------------------------------------------------------------------

## 🧱 Function Reference

### 1. `getTable(sheetObj)`

Reads the entire sheet (excluding headers) and returns array of
objects.\
Each row includes `rowID` (the actual sheet row number).

``` js
let table = SheetUtils.getTable(sheet);
```

**Return Example:**

``` js
[
  { rowID: 2, name: "Harish", age: 25, dob: "1986-07-01", status: "Active" },
  { rowID: 3, name: "Asha", age: 30, dob: "1992-05-02", status: "Pending" }
]
```

------------------------------------------------------------------------

### 2. `getRowsFromTable(table, conditions, columnNames)`

Filters rows by `conditions`.\
Supports nested logical objects.

**Examples:**

``` js
// Simple equality
SheetUtils.getRowsFromTable(table, { status: "Active" });

// Function condition
SheetUtils.getRowsFromTable(table, { age: a => a > 18 });

// Regex
SheetUtils.getRowsFromTable(table, { name: /^A/ });

// Multiple statuses
SheetUtils.getRowsFromTable(table, { status: ["Open", "Pending"] });

// Nested logic
SheetUtils.getRowsFromTable(table, {
  AND: {
    age: v => v > 18,
    OR: {
      status: ["Open", "Pending"],
      NOT: { country: "BannedLand" }
    }
  }
});

// Direct rowID match (skips all else)
SheetUtils.getRowsFromTable(table, { rowID: [2, 5, 7] });
```

------------------------------------------------------------------------

### 3. `getValuesFromTable(table, conditions, columnName)`

Returns an array of values from a specific column.

``` js
let emails = SheetUtils.getValuesFromTable(table, { status: "Active" }, "email");
```

------------------------------------------------------------------------

### 4. `setValues(sheetObj, conditions, data)`

Updates all rows matching `conditions`.

``` js
SheetUtils.setValues(sheet, { status: "Pending" }, { status: "Done", updated: SheetUtils.getTimeStamp() });
```

------------------------------------------------------------------------

### 5. `insertRow(sheetObj, data)`

Appends a new row at the bottom of the sheet.

``` js
SheetUtils.insertRow(sheet, { name: "New User", status: "Active" });
```

------------------------------------------------------------------------

### 6. `deleteRows(sheetObj, conditions)`

Deletes rows matching `conditions`.

``` js
SheetUtils.deleteRows(sheet, { rowID: [3, 4] });
```

------------------------------------------------------------------------

### 7. `getHeaders(sheetObj)`

Returns the list of header names.

``` js
let headers = SheetUtils.getHeaders(sheet);
```

------------------------------------------------------------------------

### 8. `stringifyObj(obj)`

Converts any nested object to a formatted JSON string.

``` js
Logger.log(SheetUtils.stringifyObj(userObj));
```

------------------------------------------------------------------------

### 9. `logRows(rows)`

Nicely logs the returned rows to the Apps Script console.

``` js
SheetUtils.logRows(filteredRows);
```

------------------------------------------------------------------------

### 10. `getTimeStamp()`

Returns current timestamp in **MM/dd/yy hh:mm** format.

``` js
let now = SheetUtils.getTimeStamp();
```

------------------------------------------------------------------------

### 11. `highlightRows(sheetObj, conditions, color = "#fff59d")`

Highlights rows matching the conditions.\
Returns an object you can later use to clear highlights.

``` js
let highlightSession = SheetUtils.highlightRows(sheet, { status: "Pending" }, "#fff59d");
```

**Returns:**

``` js
{
  id: "2a95ef45-bbb4-44e4-b7cf-20ab44c1b5e0",
  rowIDs: [3, 4, 9],
  color: "#fff59d",
  timestamp: "2025-10-05T12:10:22.789Z"
}
```

------------------------------------------------------------------------

### 12. `clearHighlight(sheetObj, highlightObj)`

Removes the highlight from rows using the previously returned object.

``` js
SheetUtils.clearHighlight(sheet, highlightSession);
```

------------------------------------------------------------------------

## ⚠️ Notes & Best Practices

-   **Dates** are read as plain strings to avoid timezone shifts.\

-   You can safely compare dates using:

    ``` js
    { dob: d => new Date(d) > new Date("2000-01-01") }
    ```

-   **rowID** corresponds to the actual row in the sheet (2 = first data
    row).\

-   All logical filtering (`AND`, `OR`, `NOT`) is recursive and can be
    deeply nested.\

-   For safety, test your filters with `logRows()` before modifying
    data.

------------------------------------------------------------------------

## 🧪 Example Workflow

``` js
function demo() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Users");

  // Load table
  const table = SheetUtils.getTable(sheet);

  // Get active users
  const active = SheetUtils.getRowsFromTable(table, { status: "Active" });
  SheetUtils.logRows(active);

  // Highlight active rows
  const hl = SheetUtils.highlightRows(sheet, { status: "Active" });

  // Later clear the highlight
  SheetUtils.clearHighlight(sheet, hl);

  // Update pending users
  SheetUtils.setValues(sheet, { status: "Pending" }, { status: "Approved" });
}
```

------------------------------------------------------------------------

## 🧩 License

This utility is open-source and free to use in any Google Apps Script
project.
You're welcome to extend it with:
- sorting functions
- range copy/move utilities
- formula-based column support
- event-trigger helpers
