// jonahar@GitHub
//
// Spreadsheet Service API:
// https://developers.google.com/apps-script/reference/spreadsheet/



// ------------ Utils ------------
// ===============================

// A special function that runs when the spreadsheet is open, used to add a
// custom menu to the spreadsheet.
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Aggregate', functionName: 'aggregate'},
    {name: 'Create Chart', functionName: 'create_chart'},
  ];
  spreadsheet.addMenu('Expenses Tracker', menuItems);
}


function assert(condition, message) {
  if (!condition) {
    throw message || "Assertion failed";
  }
}


// sort array and remove duplicates
function sort_unique(arr) {
  if (arr.length === 0) return arr;
  arr = arr.sort()
  var ret = [arr[0]];
  for (var i = 1; i < arr.length; i++) {
    if (arr[i - 1] !== arr[i])
      ret.push(arr[i]);
  }
  return ret;
}




// ------------ ExpensesBook object ------------
// =============================================

MONTHS = {
  '01': 'January',
  '02': 'February',
  '03': 'March',
  '04': 'April',
  '05': 'May',
  '06': 'June',
  '07': 'July',
  '08': 'August',
  '09': 'September',
  '10': 'October',
  '11': 'November',
  '12': 'December',
}


var ExpensesBook = function() {
  this.expenses = {};
  this.all_months_sorted = [];
  this.all_categories_sorted = [];

  // add amount to the given category in the iven month
  this.add_expense = function(month_id, category, amount) {
    if (!([month_id, category] in this.expenses)) {
      this.expenses[[month_id, category]] = 0;
    }
    this.expenses[[month_id, category]] += amount;
  }


  // get the amount spent in a category in a given month
  this.get_expenses = function(month_id, category) {
    if (!([month_id, category] in this.expenses))
      return 0;
    return this.expenses[[month_id, category]];
  }


  // initialize the arrays 'all_months_sorted' and 'all_categories_sorted' according to the
  // current data in this expenses book.
  this.set_months_and_categories = function() {
    var all_months = [];
    var all_categories = [];

    for (var key in this.expenses) {
      var [month_id, category] = key.split(',');
      all_months.push(month_id);
      all_categories.push(category);
    }
    this.all_months_sorted = sort_unique(all_months);
    this.all_categories_sorted = sort_unique(all_categories);
    this.all_months_sorted.pop(); // remove the last month. we assume it's incomplete
  }


  // return the sum of all expenses in the given month
  this.get_expenses_in_month = function(month_id) {
    var sum = 0;
    for (var i = 0; i < this.all_categories_sorted.length; i++) {
      var category = this.all_categories_sorted[i];
      sum += this.get_expenses(month_id, category);
    }
    return sum;
  }


  // return the rows of a summary sheet. statistics for the last month are not exported, as they are
  // considered incomplete
  this.export_summary = function() {
    this.set_months_and_categories();

    var rows = [];
    var header_row = ['Category'];
    for (var i = 0; i < this.all_months_sorted.length; i++) {
      var month_id = this.all_months_sorted[i];
      var [yyyy, mm] = month_id.split('-');
      header_row.push(MONTHS[mm] + ' ' + yyyy);
    }

    header_row = header_row.concat(['', 'Category average']);
    rows.push(header_row);

    // create row for every category
    for (var i = 0; i < this.all_categories_sorted.length; i++) {
      var category = this.all_categories_sorted[i];
      var row = [category];
      category = this.all_categories_sorted[i];
      var total_category_expenses = 0;
      for (var j = 0; j < this.all_months_sorted.length; j++) {
        var month_id = this.all_months_sorted[j];
        var amount = this.get_expenses(month_id, category);
        row.push(amount);
        total_category_expenses += amount;
      }
      var avg = total_category_expenses / this.all_months_sorted.length; // average monthly expense for this category
      row = row.concat(['', avg]);
      rows.push(row);
    }

    rows.push([' ']); // separating row. space is necessary: empty string will not be written as new row

    // monthly total row
    row = ['Monthly Total:'];
    for (var i = 0; i < this.all_months_sorted.length; i++) {
      var month_id = this.all_months_sorted[i];
      var month_expenses = this.get_expenses_in_month(month_id);
      row.push(month_expenses);
    }
    rows.push(row);

    return rows;
  }

} // end ExpensesBook




// ------------ Aggregator ------------
// ====================================

EXPENSES_SHEET = 'Expenses'
SUMMARY_SHEET = 'Summary'
BASE=10


// create an ExpensesBook object from the given rows. first row is header
function build_expenses_book(rows) {
  var expenses_book = new ExpensesBook();
  var header_row = rows[0];

  // make sure the header is as we expect
  assert(
    header_row[0] == 'Date' && header_row[1] == 'Description' &&
    header_row[2] == 'Amount' && header_row[3] == 'Category',
    'Unexpected header row'
  );

  for (var i = 1; i < rows.length; i++) {
    var [date, description, amount_str, category] = rows[i];
    amount = parseInt(amount_str, BASE); // parseInt handles commas (e.g. 3,600.00)
    var [yyyy, mm, dd] = date.split('-');
    month_id = yyyy + '-' + mm;
    expenses_book.add_expense(month_id, category, amount);
  }
  return expenses_book;
}


// remove all data and charts from the given sheet
function clear_sheet(sheet) {
  sheet.clear();
  var charts = sheet.getCharts();
  for (var i = 0; i < charts.length; i++) {
    sheet.removeChart(charts[i]);
  }
}


// Create a 'Monthly Expenses' chart in the summary sheet
function create_chart() {
  var ss = SpreadsheetApp.getActive();
  var summary_sheet = ss.getSheetByName(SUMMARY_SHEET);
  summary_sheet.activate();

  var range = summary_sheet.getDataRange();
  var last_col = range.getLastColumn();
  var last_row = range.getLastRow();
  // ignore the two last rows and two last columns in the summary sheet. they contain 'total' values.
  var chart_range = range.offset(0, 0, last_row - 2, last_col - 2);

  var chartBuilder = summary_sheet.newChart();
  chartBuilder.addRange(chart_range)
    .setChartType(Charts.ChartType.COLUMN)
    .setPosition(4, 2, 0, 0)
    .setOption('title', 'Monthly Expenses')
    .setTransposeRowsAndColumns(true)
    .setNumHeaders(1);

  summary_sheet.insertChart(chartBuilder.build());
}


// read data, aggregate (compute statistics) and write to summary sheet
function aggregate() {
  var ss = SpreadsheetApp.getActive(); // type: Spreadsheet

  var expenses_sheet = ss.getSheetByName(EXPENSES_SHEET); // type: Sheet
  var summary_sheet = ss.getSheetByName(SUMMARY_SHEET); // type: Sheet

  var expenses_range = expenses_sheet.getDataRange(); // type: Range
  var expenses_rows = expenses_range.getValues(); // list of lists

  var book = build_expenses_book(expenses_rows);
  var summary_rows = book.export_summary();

  summary_sheet.activate(); // jump to the summary sheet in the UI
  clear_sheet(summary_sheet);
  for (var i = 0; i < summary_rows.length; i++) {
    summary_sheet.appendRow(summary_rows[i]);
  }
}
