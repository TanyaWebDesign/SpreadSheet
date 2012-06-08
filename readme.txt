  /*
  * Public functions
  */
  
$.fn.spreadsheet( options )
  options:         
       { rows: ,          // initial number of rows.  default: 3
         cols: ,          // initial number of columns.  default: 3
         data: ,          // initial 2-dimensional array of strings to display. default: none
         rowheader: ,     // display row headers.  default: true
         colheader: ,     // display column headers.  default: true
         zebra_striped: , // display every odd row on a gray background.  default: false
         read_only: ,     // disallow user to enter or change data.  default: false
         context_menu:    // display custom context menu.  default: true
       }

  // is spreadsheet read-only
  $.fn.isReadOnly()
  
  // make spreadsheet read only ( bTrue - true or false )
  $.fn.setReadOnly ( bTrue )
  
  // display column headers ( bShow - true or false )
  $.fn.toggleColumnHeader( bShow )

  // display row headers ( bShow - true or false )
  $.fn.toggleRowHeader( bShow )

  // are row headers displayed
  $.fn.isRowHeader()

  // are column headers displayed
  $.fn.isColumnHeader()

  // how many rows spreadsheet has (column header row not included)
  $.fn.rowCount()

  // how many columns spreadsheet has (row header column not included)
  $.fn.colCount()

  // insert row at the specified location. If nIndex is missing or negative, a row will be added at the end
  $.fn.insertRow(nIndex)

  // remove row at the specified location. If nIndex is missing or negative, the last row will be removed
  $.fn.removeRow  = function(nIndex)

  // insert column at the specified location. If nIndex is missing or negative, a column will be added at the end.  Maximum number of columns is 26.
  $.fn.insertColumn(nIndex)

  // remove column at the specified location. If nIndex is missing or negative, the last column will be removed
  $.fn.removeColumn(nIndex)

  // display odd rows on gray background. bStriped - true or false
  $.fn.setZebraStriped(bStriped)

  // sort spreadsheet by column at the specified location. Sort will be numeric for numeric columns, 
  // and string for string columns.
  // sort will be in opposite order (ascending or descending) if the column is currently sorted, and it will be
  // descending if column is currently unsorted
  $.fn.sortColumn(iCol)

  // treat column as having numeric data ( iCol - column index, bTrue - true or false).
  $.fn.setNumeric(iCol, bTrue)

  // does column at the specified location have numeric data
  $.fn.isNumeric(iCol)
  
  // return 2-dimensional array of data in the spreadsheet
  $.fn.getData()

/*
 * Keyboard interaction
 */
Ctrl + C			Copy
Ctrl + X			Cut
Ctrl + V			Paste
Enter on editable cell		Saves the contents of a cell and removes the edit mode
Tab / Shift + Tab		Navigates between cells, rows, and columns
