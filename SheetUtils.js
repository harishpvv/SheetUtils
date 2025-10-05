class SheetUtils {
  /**
   * Get the table from a sheet object.
   * Reads all cells as strings (using getDisplayValues()) to avoid timezone shifts.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObj
   * @returns {Object[]} Array of row objects {colName: value, ...}
   */
  static getTable(sheetObj) {
    if (!sheetObj) throw new Error("Sheet object is required.");

    const values = sheetObj.getDataRange().getDisplayValues();
    if (values.length < 2) return [];

    const headers = values[0];
    return values.slice(1).map((row, idx) => {
    let obj = { rowID: idx + 2 }; // real row number in sheet
    headers.forEach((h, i) => {
      let val = row[i];
      if (/^\d{4}-\d{2}-\d{2}$/.test(val)) obj[h] = val;
      else if (!isNaN(val) && val.trim() !== "") obj[h] = Number(val);
      else obj[h] = val;
    });
    return obj;
  });
  }


  /**
   * Get rows matching conditions from a table
   * @param {Object[]} table - output of getTable()
   * @param {Object} conditions - see _matchConditions for supported formats
   * @param {string[]} [column_names] - optional, select subset of columns
   * @returns {Object[]} Filtered array of rows
   */
  static getRowsFromTable(table, conditions, column_names) {
    return table.filter(row => this._matchConditions(row, conditions))
                .map(r => {
                  if (!column_names) return r;
                  let sub = {};
                  column_names.forEach(c => sub[c] = r[c]);
                  return sub;
                });
  }

  /**
   * Get array of values from a table/rows
   * @param {Object[]} table - output of getTable() or getRowsFromTable()
   * @param {Object} conditions
   * @param {string} column_name
   * @returns {Array} Array of column values
   */
  static getValuesFromTable(table, conditions, column_name) {
    return this.getRowsFromTable(table, conditions, [column_name])
               .map(r => r[column_name]);
  }

  /**
   * Update values in a sheet where conditions are met
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObj
   * @param {Object} conditions
   * @param {Object} data - {colName: newValue, ...}
   */
  static setValues(sheetObj, conditions, data) {
    const values = sheetObj.getDataRange().getValues();
    const headers = values[0];
    const table = this.getTable(sheetObj);

    table.forEach((row, idx) => {
      if (this._matchConditions(row, conditions)) {
        Object.keys(data).forEach(key => {
          const colIndex = headers.indexOf(key);
          if (colIndex >= 0) values[idx + 1][colIndex] = data[key];
        });
      }
    });

    sheetObj.getRange(1, 1, values.length, headers.length).setValues(values);
  }

  /**
   * Insert a row at the end
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObj
   * @param {Object} data - {colName: value, ...}
   */
  static insertRow(sheetObj, data) {
    const headers = this.getHeaders(sheetObj);
    const newRow = headers.map(h => data[h] !== undefined ? data[h] : "");
    sheetObj.appendRow(newRow);
  }

  /**
   * Delete rows where conditions are met
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObj
   * @param {Object} conditions
   */
  static deleteRows(sheetObj, conditions) {
    const table = this.getTable(sheetObj);
    const headers = this.getHeaders(sheetObj);
    const rowsToDelete = [];

    table.forEach((row, idx) => {
      if (this._matchConditions(row, conditions)) rowsToDelete.push(idx + 2); // header + 1-based
    });

    rowsToDelete.reverse().forEach(r => sheetObj.deleteRow(r));
  }

  /**
   * Get sheet headers
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObj
   * @returns {string[]}
   */
  static getHeaders(sheetObj) {
    return sheetObj.getRange(1, 1, 1, sheetObj.getLastColumn()).getValues()[0];
  }

  /**
   * Convert nested object to string safely
   * @param {Object} obj
   * @returns {string}
   */
  static stringifyObj(obj) {
    return JSON.stringify(obj, (key, value) => {
      if (typeof value === "function") return value.toString();
      return value;
    }, 2);
  }

  /**
   * Log rows or table
   * @param {Object[]} rows
   */
  static logRows(rows) {
    rows.forEach(r => Logger.log(this.stringifyObj(r)));
  }

  /**
   * Get timestamp in MM/dd/yy HH:mm
   * @returns {string}
   */
  static getTimeStamp() {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yy HH:mm");
  }

  /**
   * Highlights rows that match given conditions.
   * @param {Sheet} sheetObj - Target sheet.
   * @param {Object} conditions - Filter conditions.
   * @param {String} [color="#fff59d"] - Highlight color.
   * @returns {Object} highlightObj - { id: <UUID>, rowIDs: [..], color: <color> }
   */
  static highlightRows(sheetObj, conditions, color = "#fff59d") {
    if (!sheetObj) throw new Error("Sheet object is required.");

    const table = this.getTable(sheetObj);
    const matches = table.filter(row => this._matchConditions(row, conditions));

    if (matches.length === 0) {
      Logger.log("No matching rows found to highlight.");
      return null;
    }

    // Collect all rowIDs
    const rowIDs = matches.map(r => r.rowID);

    // Apply color
    rowIDs.forEach(id => {
      sheetObj.getRange(id, 1, 1, sheetObj.getLastColumn()).setBackground(color);
    });

    // Create highlight session ID
    const highlightObj = {
      id: Utilities.getUuid(),
      rowIDs: rowIDs,
      color: color,
      timestamp: new Date().toISOString(),
    };

    Logger.log(`Highlighted ${rowIDs.length} row(s): ${JSON.stringify(highlightObj)}`);
    return highlightObj;
  }

  /**
   * Clears highlights created by highlightRows().
   * @param {Sheet} sheetObj - Target sheet.
   * @param {Object} highlightObj - The returned object from highlightRows().
   */
  static clearHighlight(sheetObj, highlightObj) {
    if (!sheetObj || !highlightObj || !highlightObj.rowIDs) {
      throw new Error("Invalid arguments passed to clearHighlight.");
    }

    highlightObj.rowIDs.forEach(id => {
      sheetObj.getRange(`A${id}:${sheetObj.getLastColumn()}`).setBackground("white");
    });

    Logger.log(`Cleared highlight for rows: ${highlightObj.rowIDs.join(", ")}`);
  }


  // ----------------- Internal helpers -----------------

  static _getSpreadsheetTimeZone() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss && ss.getSpreadsheetTimeZone) return ss.getSpreadsheetTimeZone();
    } catch (e) {}
    return Session.getScriptTimeZone() || 'UTC';
  }

  static _compareValues(a, b) {
    return a === b;
  }

  /**
   * Internal: match row against conditions
   * Supports:
   *  - Equality with string, number, or Date objects
   *  - Functions: col => boolean (row string converted to Date if looks like yyyy-MM-dd)
   *  - Regex
   *  - Arrays of allowed values
   *  - Nested AND, OR, NOT
   */
  static _matchConditions(row, conditions) {
    if (!conditions || Object.keys(conditions).length === 0) return true;

    // --- Direct rowID match (ignore all else)
    if ("rowID" in conditions) {
      const condVal = conditions.rowID;
      if (Array.isArray(condVal)) {
        return condVal.includes(row.rowID);
      }
      return row.rowID === Number(condVal);
    }

    for (let key in conditions) {
      const condVal = conditions[key];

      if (key === "AND") {
        if (!this._matchConditions(row, condVal)) return false;
      } 
      else if (key === "OR") {
        let ok = false;
        for (let subKey in condVal) {
          if (this._matchConditions(row, { [subKey]: condVal[subKey] })) {
            ok = true;
            break;
          }
        }
        if (!ok) return false;
      } 
      else if (key === "NOT") {
        if (this._matchConditions(row, condVal)) return false;
      } 
      else {
        let rowVal = row[key];

        if (condVal instanceof RegExp) {
          if (!condVal.test(rowVal)) return false;
        } 
        else if (typeof condVal === "function") {
          let fnVal = rowVal;
          // Convert yyyy-MM-dd strings to Date for function comparisons
          if (/^\d{2}-\d{2}-\d{4}$/.test(fnVal)) fnVal = new Date(fnVal);
          if (!condVal(fnVal)) return false;
        } 
        else if (Array.isArray(condVal)) {
          if (!condVal.some(v => v instanceof Date 
                                  ? new Date(rowVal).getTime() === v.getTime() 
                                  : rowVal === v)) return false;
        } 
        else if (condVal instanceof Date) {
          if (new Date(rowVal).getTime() !== condVal.getTime()) return false;
        } 
        else {
          if (rowVal !== condVal) return false;
        }
      }
    }

    return true;
  }
}
