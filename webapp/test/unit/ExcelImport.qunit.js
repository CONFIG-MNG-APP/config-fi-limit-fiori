sap.ui.define([
  "zgsp26/conf/mng/filimit/confmngfefilimit/ext/main/ExcelImport"
], function (ExcelImport) {
  "use strict";

  QUnit.module("ExcelImport – parseNumber");

  QUnit.test("parses valid numbers", function (assert) {
    assert.strictEqual(ExcelImport.parseNumber(5000), 5000, "integer 5000");
    assert.strictEqual(ExcelImport.parseNumber(99.99), 99.99, "float 99.99");
    assert.strictEqual(ExcelImport.parseNumber("5000"), 5000, "string '5000'");
    assert.strictEqual(ExcelImport.parseNumber("1234.56"), 1234.56, "string '1234.56'");
    assert.strictEqual(ExcelImport.parseNumber(0), 0, "zero");
  });

  QUnit.test("returns 0 for invalid values", function (assert) {
    assert.strictEqual(ExcelImport.parseNumber(""), 0, "empty string");
    assert.strictEqual(ExcelImport.parseNumber(null), 0, "null");
    assert.strictEqual(ExcelImport.parseNumber(undefined), 0, "undefined");
    assert.strictEqual(ExcelImport.parseNumber("abc"), 0, "non-numeric string");
  });

  QUnit.module("ExcelImport – mapHeaders");

  QUnit.test("maps known headers correctly", function (assert) {
    var result = ExcelImport.mapHeaders(["Expense Type", "GL Account", "Approval Limit", "Currency", "Change Note"]);
    assert.strictEqual(result.mapped["Expense Type"], "ExpenseType");
    assert.strictEqual(result.mapped["GL Account"], "GlAccount");
    assert.strictEqual(result.mapped["Approval Limit"], "AutoApprLim");
    assert.strictEqual(result.mapped["Currency"], "Currency");
    assert.strictEqual(result.mapped["Change Note"], "ChangeNote");
    assert.strictEqual(result.unmapped.length, 0);
  });

  QUnit.test("maps alternative header names", function (assert) {
    var result = ExcelImport.mapHeaders(["ExpenseType", "Account", "Limit", "Curr", "Note"]);
    assert.strictEqual(result.mapped["ExpenseType"], "ExpenseType");
    assert.strictEqual(result.mapped["Account"], "GlAccount");
    assert.strictEqual(result.mapped["Limit"], "AutoApprLim");
    assert.strictEqual(result.mapped["Curr"], "Currency");
    assert.strictEqual(result.mapped["Note"], "ChangeNote");
  });

  QUnit.test("reports unmapped headers", function (assert) {
    var result = ExcelImport.mapHeaders(["Expense Type", "UnknownCol", "FooBar"]);
    assert.strictEqual(Object.keys(result.mapped).length, 1);
    assert.deepEqual(result.unmapped, ["UnknownCol", "FooBar"]);
  });

  QUnit.module("ExcelImport – transformRow");

  QUnit.test("transforms a valid row", function (assert) {
    var mapping = { "Expense Type": "ExpenseType", "GL Account": "GlAccount", "Limit": "AutoApprLim", "Currency": "Currency" };
    var rawRow = { "Expense Type": "TRAVEL", "GL Account": "600100", "Limit": "5000", "Currency": "VND" };
    var row = ExcelImport.transformRow(rawRow, mapping, "DEV");

    assert.strictEqual(row.ExpenseType, "TRAVEL");
    assert.strictEqual(row.GlAccount, "600100");
    assert.strictEqual(row.AutoApprLim, 5000);
    assert.strictEqual(row.Currency, "VND");
    assert.strictEqual(row.EnvId, "DEV");
    assert.strictEqual(row.ActionType, "C");
    assert.strictEqual(row._state, "new");
    assert.strictEqual(row._reqItemId, null);
    assert.strictEqual(row.VersionNo, 0);
  });

  QUnit.test("parses AutoApprLim as number from string", function (assert) {
    var mapping = { "Limit": "AutoApprLim" };
    var rawRow = { "Limit": "12345.67" };
    var row = ExcelImport.transformRow(rawRow, mapping, "DEV");
    assert.strictEqual(row.AutoApprLim, 12345.67);
  });

  QUnit.test("AutoApprLim defaults to 0 for non-numeric", function (assert) {
    var mapping = { "Expense Type": "ExpenseType", "Limit": "AutoApprLim" };
    var rawRow = { "Expense Type": "TRAVEL", "Limit": "abc" };
    var row = ExcelImport.transformRow(rawRow, mapping, "DEV");
    assert.strictEqual(row.AutoApprLim, 0);
    assert.strictEqual(row.ExpenseType, "TRAVEL");
  });

  QUnit.test("returns null for empty row", function (assert) {
    var mapping = { "Expense Type": "ExpenseType", "GL Account": "GlAccount" };
    var rawRow = { "Expense Type": "", "GL Account": "" };
    var row = ExcelImport.transformRow(rawRow, mapping, "DEV");
    assert.strictEqual(row, null);
  });

  QUnit.test("uses provided EnvId", function (assert) {
    var mapping = { "Expense Type": "ExpenseType" };
    var rawRow = { "Expense Type": "OFFICE" };
    var row = ExcelImport.transformRow(rawRow, mapping, "QAS");
    assert.strictEqual(row.EnvId, "QAS");
  });

  QUnit.module("ExcelImport – parseWorkbook");

  QUnit.test("returns error for empty workbook", function (assert) {
    var result = ExcelImport.parseWorkbook(null, "DEV");
    assert.strictEqual(result.rows.length, 0);
    assert.ok(result.errors.length > 0);
  });

  QUnit.test("parses workbook with valid data (mock XLSX global)", function (assert) {
    var origXLSX = window.XLSX;
    window.XLSX = {
      utils: {
        sheet_to_json: function () {
          return [
            { "Expense Type": "TRAVEL", "GL Account": "600100", "Limit": 5000, "Currency": "VND" },
            { "Expense Type": "OFFICE", "GL Account": "600200", "Limit": 2000, "Currency": "USD" },
            { "Expense Type": "", "GL Account": "", "Limit": "", "Currency": "" }
          ];
        }
      }
    };

    var workbook = { SheetNames: ["Sheet1"], Sheets: { "Sheet1": {} } };
    var result = ExcelImport.parseWorkbook(workbook, "DEV");

    assert.strictEqual(result.rows.length, 2, "2 valid rows parsed");
    assert.strictEqual(result.skipped, 1, "1 empty row skipped");
    assert.strictEqual(result.rows[0].ExpenseType, "TRAVEL");
    assert.strictEqual(result.rows[0].AutoApprLim, 5000);
    assert.strictEqual(result.rows[0].Currency, "VND");
    assert.strictEqual(result.rows[1].ExpenseType, "OFFICE");
    assert.strictEqual(result.rows[1].AutoApprLim, 2000);

    window.XLSX = origXLSX;
  });

  QUnit.test("returns error when no headers match", function (assert) {
    var origXLSX = window.XLSX;
    window.XLSX = {
      utils: {
        sheet_to_json: function () {
          return [{ "Foo": "bar", "Baz": "qux" }];
        }
      }
    };

    var workbook = { SheetNames: ["Sheet1"], Sheets: { "Sheet1": {} } };
    var result = ExcelImport.parseWorkbook(workbook, "DEV");

    assert.strictEqual(result.rows.length, 0);
    assert.ok(result.errors[0].indexOf("No recognizable column headers") !== -1);

    window.XLSX = origXLSX;
  });
});
