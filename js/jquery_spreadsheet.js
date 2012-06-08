( function( $ ) {

var asCopiedElements = new Array();
var nTableCount = 0;

$.fn.spreadsheet = function ( options )
{
    var settings = $.extend( {
        rows: 3,
        cols: 3,
        data: null,
        rowheader: true,
        colheader: true,
        zebra_striped: false,
        read_only: false,
        context_menu: true
    }, options|| {} );
    
    var oDiv = $( this );
    if ( oDiv.size() == 0 ) {
      return false;
    }
    
    var oTable = $( "<table class = 'sp_spreadsheet' ></table>" ).appendTo( oDiv );
    nTableCount++;
    oTable.attr( 'id', "sp_" + nTableCount );
    
    var nRowInd, nColInd;
    var nRowCount = 3; 
    var nColCount = 3; 
    var oRow;
    var oCol;
    
    if ( settings.data != null ) {
      var nInd;
      nRowCount = settings.data.length;
      nColCount = 0;
      
      for ( nInd = 0; nInd < nRowCount; nInd++ ) {
        nColCount = ( nColCount < settings.data[ nInd ].length ) ? settings.data[ nInd ].length : nColCount;
      }
    }
    if ( nRowCount < settings.rows ) {
      nRowCount = settings.rows;
    }
    if ( nColCount < settings.cols ) {
      nColCount = settings.cols;
    }
    
    // check for 0 defaults
    nRowCount = ( nRowCount == 0 ) ? 5 : nRowCount+1;
    nColCount = ( nColCount == 0 ) ? 5 : nColCount+1;
    
    if ( nColCount > 25 ) {
      nColCount = 25;
    }
   
    for ( nRowInd = 0; nRowInd < nRowCount; nRowInd++ ) {
      oRow = $( "<tr></tr>" ).appendTo( oTable );
      oRow.height( 20 );
      
      if ( nRowInd == 0 ) {
        oRow.addClass( "sp_rowheader" );
      }
      for ( nColInd = 0; nColInd < nColCount; nColInd++ ) {
        oCol = $( "<td></td>" ).appendTo( oRow );
        if ( nColInd == 0 ) {
          oCol.addClass( "sp_colheader" );
          if ( nRowInd > 0 ) {
            oCol.text( nRowInd );
          }
        }
        else {
          if ( nRowInd == 0 ) {
            oCol.text( String.fromCharCode(65 + nColInd - 1 ) );
          }
          else {
            if ( settings.data != null ) {
              if ( ( settings.data[ nRowInd - 1 ] != undefined ) && ( settings.data[ nRowInd - 1 ][ nColInd - 1 ] != undefined ) ) {
                oCol.text( settings.data[ nRowInd - 1 ][ nColInd - 1 ] );
              }
            }
          }
          oCol.width( 100 );
        }
      }
    }  
    
    var timeout;
    
    oTable.attr( 'tabindex', nTableCount );
    oTable.blur( function () {
      timeout = setTimeout(function() {
        unselect(oTable);}, 300);
    } );
   
    var sTD = "table#sp_"  + nTableCount + " td";
    var td = $( sTD );
    
    // double-click
    td.live( "dblclick", function ( event ) { 
      destroyContextMenu();
      unSelectRow( oTable );
      unSelectColumn( oTable );
      unSelectCell( oTable );
      
      var oCell = $(this);
      
      if ( oCell.parent().hasClass( 'sp_rowheader' ) ) {
        return;
      }
      if ( oCell.hasClass( 'sp_colheader' ) ) {
        return;
      }
      positionTextCtrl( oCell );   
    } );
 
      // click
     td.live( "click", function ( event ) { 
        destroyContextMenu();
        unSelectRow( oTable );
        unSelectColumn( oTable );
        unSelectCell( oTable );
        
        var oCell = $(this);
        
        if ( oCell.parent().hasClass( 'sp_rowheader' ) ) {
          var iIndex = oCell.parent().find( 'td' ).index( oCell );
          if ( iIndex > 0 ) {
            selectColumn( oTable, iIndex );
          }
          return;
        }
           
        if ( oCell.hasClass( 'sp_colheader' ) ) {
          oCell.parent().addClass( "sp_selected" );
          oCell.parent().removeClass( "sp_oddrow" );
          return;
        }
        oCell.addClass( "sp_selected" );  
      } );
  
      // keypress - doesn't work on IE and Chrome
      /*
       oDiv.keypress( function ( event ) { 
        if ( event.ctrlKey ) {
          if ( ( event.which == 67 ) || ( event.which == 99 ) ) {   // Ctrl + C
            doMenuAction( oTable, "Copy" );
            oTable.focus();
          }
          if ( ( event.which == 86 ) || ( event.which == 118 ) ) {   // Ctrl + V
            doMenuAction( oTable, "Paste" );
            oTable.focus();
          }
        }
      } );
      */
      
      oDiv.keydown( function ( event ) {    // keypress doesn't capture non-character keys
        if ( event.which == 9 ) {
          nextCell( oTable, event.shiftKey );
          event.preventDefault();
        }
        if ( event.ctrlKey ) {
          if ( ( event.which == 67 ) || ( event.which == 99 ) ) {   // Ctrl + C
            doMenuAction( oTable, "Copy" );
            oTable.focus();
          }
          if ( ( event.which == 86 ) || ( event.which == 118 ) ) {   // Ctrl + V
            doMenuAction( oTable, "Paste" );
            oTable.focus();
          }
        }
          if ( ( event.which == 88 ) || ( event.which == 120 ) ) {   // Ctrl + X
            doMenuAction( oTable, "Cut" );
            oTable.focus();
          }
      } );
      
      // right mouse click
      td.live( "contextmenu", function ( event ) { 
         if ( !settings.context_menu ) {
          return true;
         }
         destroyContextMenu();   // destroy old

          var oCell = $( this );
          var oRow = oCell.parent();
          
          if ( oRow.hasClass( "sp_selected" ) ) {    // if clicked on a selected row
            displayContextMenu( event.pageX, event.pageY, oTable, true, false, false );
          }
          else if ( oCell.hasClass( "sp_colselected" ) ) {   // if clicked on a selected column
            displayContextMenu( event.pageX, event.pageY, oTable, false, true, false );
          }
          else if ( oRow.hasClass( 'sp_rowheader' ) && isColSelected( oCell ) ) {    // if clicked on a selected column
            displayContextMenu( event.pageX, event.pageY, oTable, false, true, false );
          }
          else {    // if clicked on row or column other than selected
               unSelectRow( oTable );   
               unSelectColumn( oTable );
               unSelectCell( oTable );
               
              if ( ( oRow.hasClass( 'sp_rowheader' ) == false ) && ( oCell.hasClass( 'sp_colheader' ) == false ) ) {
                 oCell.addClass( "sp_selected" );
                 displayContextMenu( event.pageX, event.pageY, oTable, false, false, true );
              }
              else {
                displayContextMenu( event.pageX, event.pageY, oTable, false, false, false );
              }
          }
          return false;
      } );
      
      var id;
      
      id = oTable.attr( 'id' ) + "_cell";
      $( "#" + id ).live( "blur", function () { 
        saveCell();
      } );
      
      // keypress
      $( "#" + id ).live( "keypress", function ( event ) {
        if ( event.which == 13 ) {
          saveCell();
        }
      } );
      
       id = oTable.attr( 'id' ) + "_menu";
       $( "ul#" + id + " li").live( "click", function ( event ) { 
        clearTimeout( timeout );
        doMenuAction( oTable, $( this ).text() );
        oTable.focus();
        destroyContextMenu();
      } );
      
      // resizing
      $( 'div.drag_handle' ).live( "mousedown", function ( event ) { 
        if ( event.which == 1 ) {
          var draghandle = $( this );
          draghandle.attr( "dragging", true );
          draghandle.css( 'border-left', '2px dotted #BAD0EF' );
          return false;
        }
      } );
      
       $(window).mouseup( function ( event ) {
         var draghandle = $( 'div.drag_handle[ dragging = true ]' );
         if ( ( event.which == 1 ) && ( draghandle.size() > 0 ) ) {
            draghandle.removeAttr( "dragging" );
            draghandle.css( 'border', 'none' );
            resizeColumn( draghandle, event.pageX );
          }
       } );
      
        oDiv.mousemove( function ( event ) { 
        var draghandle = $( 'div.drag_handle[ dragging = true ]' );
        if ( draghandle.size() > 0 ) {
          dragBorder( draghandle, event.pageX );
        }
      } );
      
      $('div.drag_handle').live( "hover", function () {
            $(this).css( 'cursor', 'crosshair' );
          },         
         function () {     
            $(this).css( 'cursor', 'default' );
         }); 
     
      $( window ).resize( function () {
        makeResizable( oTable );
      } );
    
      makeResizable( oTable );
      oTable.data( "zebra_striped", false )
      
      oTable.setZebraStriped( settings.zebra_striped );
      oTable.data( "read_only", settings.read_only );
      
      oTable.toggleRowHeader( settings.rowheader );
      oTable.toggleColumnHeader( settings.colheader );
      return oTable;
}      
/*
 * Private functions
 */
  function unselect(oTable)
  {
        destroyContextMenu();
        unSelectRow( oTable );
        unSelectColumn( oTable );
        unSelectCell( oTable );
  }

  
  function positionTextCtrl( oCell )
  {
    var oTable = oCell.closest( '.sp_spreadsheet' );
    if ( oTable.isReadOnly() ) {
      return;
    }
    var ctrl = $( "<input type='text' name='sp_cell'/>" ).insertAfter( oTable );
    ctrl.attr( 'id', oTable.attr( 'id' ) + "_cell" );
    ctrl.addClass( 'sp_cell' );
    
    var offset = oCell.offset();
    var nLeft = offset.left + 2;
    var nTop = offset.top + 2;
    var nWidth = oCell.width();
    var nHeight = oCell.height()
    
    ctrl.css( { position: 'fixed' } );
    ctrl.css( { left: nLeft + "px", top: nTop + "px" } );
    ctrl.width( nWidth - 4 );
    ctrl.height( nHeight - 4 );
    ctrl.focus();
    ctrl.attr( 'value', oCell.text() );
    ctrl.data( 'cell', oCell );
    
    asCopiedElements = [];
  }

  function saveCell()
  {
    var ctrl = $( "input[ name = 'sp_cell']" );
    var oTableCell = ctrl.data( 'cell' );
    var sText = ctrl.attr( 'value' );
    sText = $.trim( sText );
    oTableCell.text( sText );
    ctrl.remove();
    
    var oParent = oTableCell.parent();
    //var iSort = $( 'table#spreadsheet' ).data( 'sort' );
    
    var oTable = oTableCell.closest( '.sp_spreadsheet' );
    var iSort = oTable.data( 'sort' );
    var iIndex = oParent.find( 'td').index( oTableCell ) - 1;
    
    if ( iSort == iIndex ) {
      unSort( oTable );
    }
  }

  function displayContextMenu( left, top, oTable, bRow, bColumn, bCell )
  {
    var menu = $( "<ul></ul>" ).insertAfter( oTable );
    menu.attr( 'id', oTable.attr( 'id' ) + "_menu" );
    menu.attr( 'name', 'sp_context' );
    menu.addClass( 'sp_context' );
    menu.css( { 'top': top, 'left': left } );
   
    var bBorder = false;
    if ( !oTable.isReadOnly() ) {
          bBorder = true;
          
          if ( bRow || bCell ) {
                menu.append( "<li>Cut</li>" );
                menu.append( "<li>Copy</li>" );
                if ( asCopiedElements.length > 0 ) {
                    menu.append( "<li>Paste</li>" );
                }
          }
          var nRowCount = oTable.rowCount();
          var nColCount = oTable.colCount();
          
          if ( bRow ) {
            menu.append( "<li>Insert row</li>" );
          }
          else {
            menu.append( "<li>Add row</li>" );
          }
          
          if ( nRowCount > 1 ) {
            menu.append( "<li>Delete row</li>" );
          }
          
          if ( nColCount < 26 ) {
            if ( !bColumn ) {
              menu.append( "<li>Add column</li>" );
            }
            else {
              menu.append( "<li>Insert column</li>" );
            }
          }
          if ( nColCount > 1 ) {
            menu.append( "<li>Delete column</li>" );
          }
    }
    if ( bColumn ) {
      bBorder = true;
      menu.append( "<li>Sort column</li>" );
      menu.append( "<li>Toggle numeric</li>" );
    }
    
    if ( oTable.isRowHeader() ) {
      menu.append( "<li>Hide row headers</li>" );
    }
    else {
      menu.append( "<li>Show row headers</li>" );
    }
    
    if ( bBorder ) {
      menu.find( 'li' ).last().css( {'border-top': '1px inset #cccccc', 'margin-top': '5px' } );
    }
    
    if ( oTable.isColumnHeader() ) {
      menu.append( "<li>Hide column headers</li>" );
    }
    else {
      menu.append( "<li>Show column headers</li>" );
    }
    menu.append( "<li>Toggle zebra striped</li>" );
    menu.show();
  }

  function nextCell( oTable, bBackward ) 
  {
    var rows = oTable.find( 'tr' );
    
    var selectedRow = oTable.find('tr.sp_selected');
    if ( selectedRow.size() > 0 ) {       // row is selected
      var nSelectedRowInd = rows.index( selectedRow[0] );
      var oNextRow;
      unSelectRow( oTable );
      
      if ( bBackward ) {
          oNextRow = ( nSelectedRowInd ==  1 ) ? rows.eq( rows.size() - 1 ) : rows.eq( nSelectedRowInd - 1 );
      }
      else {
          oNextRow = ( nSelectedRowInd == rows.size() - 1 ) ? rows.eq( 1 ) : rows.eq( nSelectedRowInd + 1 );
      }
      oNextRow.removeClass( "sp_oddrow" );
      oNextRow.addClass( 'sp_selected');
      return;
    }
    
    var selectedCol = oTable.find('td.sp_colselected');
    if ( selectedCol.size() > 0 ) {    // column is selected
      var cols = rows.eq( 1 ).find( 'td' );
      var nCellInd = cols.index( selectedCol[ 0 ] );
      var nColInd;
      
      if ( bBackward ) {
        nColInd = ( nCellInd == 1 ) ? cols.size() - 1 : nCellInd - 1;
      }
      else {
        nColInd = ( nCellInd == cols.size() - 1 ) ?  1 : nCellInd + 1;
      }
      unSelectColumn( oTable );
      selectColumn( oTable, nColInd );
      return;
    }
    
    var oSelectedCell = oTable.find('td.sp_selected');
    if ( oSelectedCell.size() == 0 ) {    // select first cell
      oTable.find( 'tr' ).eq( 1 ).find( 'td' ).eq( 1 ).addClass( 'sp_selected');
      return;
    }
    var thisRow = oSelectedCell.parent();
    var nThisRow = rows.index( thisRow );
    var cells = thisRow.find( 'td' );
    var nColCount = cells.size();
    var nSelectedInd = cells.index( oSelectedCell );
   
    oSelectedCell.removeClass( 'sp_selected');
    if ( bBackward ) {
      if ( nSelectedInd == 1 ) {
        if ( nThisRow == 1 ) {
          rows.eq( rows.size() - 1 ).find( 'td' ).eq( nColCount - 1 ).addClass( 'sp_selected'); // select last column
        }
        else {
          rows.eq( nThisRow - 1 ).find( 'td' ).eq( nColCount - 1 ).addClass( 'sp_selected'); // select previous row, last column
        }
      }
      else {
        cells.eq( nSelectedInd - 1 ).addClass( 'sp_selected');
      }
    }
    else {
      if ( nSelectedInd == nColCount - 1 ) {
        if ( nThisRow == rows.size() - 1 ) {
          rows.eq( 1 ).find( 'td' ).eq( 1 ).addClass( 'sp_selected');   // select first row, first column
        }
        else {
          rows.eq( nThisRow + 1 ).find( 'td' ).eq( 1 ).addClass( 'sp_selected');  // select next row, first column
        }
      }
      else {
        cells.eq( nSelectedInd + 1 ).addClass( 'sp_selected');
      }
    }
  }
  
  function destroyContextMenu()
  {
      $("ul[name = 'sp_context']").remove();
  }

  function doMenuAction( oTable, action )
  {
    var selectedRow = oTable.find('tr.sp_selected');
    var selectedCol = oTable.find('td.sp_colselected');
    var oSelectedRow;
    var oSelectedCol;
    var oSelectedCell = null;
    var nSelectedRow = undefined;
    var nSelectedCol = undefined;
    var rows = oTable.find( 'tr' );
    
    if ( selectedCol.size() > 0 ) {    // column is selected
      oSelectedCol = selectedCol.first();   
      nSelectedCol = rows.eq( 1 ).find( 'td' ).index( oSelectedCol ) - 1; // first selected column cell is in 2nd row
    }
    else if ( selectedRow.size() > 0 ) {   // row is selected
      oSelectedRow = selectedRow.first();
      nSelectedRow = rows.index( oSelectedRow ) - 1;
    }
    else {
      oSelectedCell = oTable.find('td.sp_selected');
    }
    
    if ( action == "Cut" ) {
      if ( nSelectedRow != undefined ) {
        copyRow( oSelectedRow, true );
      }
      else {
        copyCell( oSelectedCell, true );
      }
    }
    else if ( action == "Copy" ) {
      if ( nSelectedRow != undefined ) {
        copyRow( oSelectedRow, false );
      }
      else {
        copyCell( oSelectedCell, false );
      }
    }
    else if ( action == "Paste" ) {
        if ( !oTable.isReadOnly() ) {
            var nInd;
            
            if ( nSelectedRow != undefined ) {
              var nCount = asCopiedElements.length;
              var oCells = oSelectedRow.find( 'td' );
              
              for ( nInd = 0; nInd < nCount; nInd++ ) {   // skip header
                oCells.eq( nInd + 1 ).text( asCopiedElements[ nInd ] );
              }
           }
           else {
            oSelectedCell.text( asCopiedElements[0] );
           }
        }
    }
    else if ( action == "Insert row" ) {
      oTable.insertRow( nSelectedRow );
    }
    else if ( action == "Add row" ) {
      oTable.insertRow();
    }
    else if ( action == "Delete row" ) {
      oTable.removeRow( nSelectedRow );
    }
    else if ( action == "Add column" ) {
      oTable.insertColumn();
    }
    else if ( action == "Insert column" ) {
      oTable.insertColumn( nSelectedCol );
    }
    else if ( action == "Delete column" ) {
      oTable.removeColumn( nSelectedCol );
    }
    else if ( action == "Sort column" ) {
      oTable.sortColumn( nSelectedCol );
    }
    else if ( action == "Toggle numeric" ) {
      var bTrue = oTable.isNumeric( nSelectedCol );
      oTable.setNumeric( nSelectedCol, !bTrue );
    }
    else if ( action.indexOf( "row header" ) != -1 ) {
      oTable.toggleRowHeader( !oTable.isRowHeader() );
    }
    else if ( action.indexOf( "column header" ) != -1 ) {
      oTable.toggleColumnHeader( !oTable.isColumnHeader() );
    }
    else if ( action.indexOf( "zebra striped" ) != -1 ) {
      oTable.setZebraStriped( !oTable.data( "zebra_striped" ) );
    }
  }

  function compareA( num1, num2 ) {
    return compareHelper(num1, num2, true);
  }

  function compareD( num1, num2 ) {
    return compareHelper(num1, num2, false);
  }

  function compareHelper(num1, num2, order){
      var nSeparator = num1.lastIndexOf( "^" );
      var num1Helper = parseFloat( num1.substring( 0, nSeparator ) );
      nSeparator = num2.lastIndexOf( "^" );
      var num2Helper = parseFloat( num2.substring( 0, nSeparator ) );
      var result;
      if (order)
      {
        result = num2Helper - num1Helper; 
      }
      else{
        result = num1Helper - num2Helper; 
      }
      return result;
  }

  function unSort( oTable )
  {
    var oFirstRow = oTable.find( 'tr' ).first();
    oFirstRow.find( 'td' ).each( function() { $(this).children().remove() } );
    oTable.removeData( 'sort' );
    oTable.removeData( 'sort_order' );
  }

  function unSelectRow( oTable )
  {
    oTable.find( 'tr' ).removeClass( "sp_selected" );
    if ( oTable.data( "zebra_striped" ) ) {
      oTable.find( 'tr:odd' ).addClass( "sp_oddrow" );
    }
  }

  function unSelectCell( oTable )
  {
    oTable.find( 'td').removeClass( "sp_selected" );
  }

  function copyRow( oSelectedRow, bEmptySelected )
  {
    var data = oSelectedRow.find( 'td' );
    var nCount = data.size();
    var nInd;
    
    asCopiedElements = [];
    for ( nInd = 1; nInd < nCount; nInd++ ) {   // skip header
      asCopiedElements[ nInd - 1 ] = data.eq( nInd ).text();
      if ( bEmptySelected ) {
        data.eq( nInd ).text("");
      }
    }
  }

  function copyCell( oSelectedCell, bEmptySelected )
  {
    asCopiedElements = [];
    asCopiedElements[0] = oSelectedCell.text();
    if ( bEmptySelected ) {
      oSelectedCell.text("");
    }
  }

  function selectColumn( oTable, iIndex )
  {
    var nInd;
    var nRowCount = oTable.rowCount();
    var rows = oTable.find( 'tr' );
    
    unSelectColumn(oTable);
    
    for ( nInd = 1; nInd <= nRowCount; nInd++ ) {
      rows.eq( nInd ).find( 'td' ).eq( iIndex ).addClass( "sp_colselected" );
    }
  }

  function unSelectColumn( oTable ) 
  {
    oNextRow = oTable.find( 'td' ).removeClass( "sp_colselected" );
  }

  function makeResizable( oTable )
  {
    $( "div.drag_handle[ name = '" + oTable.attr( 'id') + "']" ).remove();
    
    var nColCount = oTable.colCount();
    var nHeight = oTable.height();
    var nTop = oTable.css( 'top' );
    var nInd;
    var oFirstRow = oTable.find( 'tr' ).first();
    var oDiv;
    var oCell;
    
    oFirstRow.find( 'td' ).each( function( nIndex )  {    // all columns including header
      oDiv = $( "<div class = 'drag_handle' oncontextmenu = 'return false'></div>" ).insertBefore ( oTable );
      oDiv.attr( 'name', oTable.attr( 'id' ) );
      oCell = $(this);
      oDiv.data( 'column', nIndex );
      oDiv.data( 'table', oTable );
      oDiv.css( { 'position': 'absolute', 'top': nTop, 'left': oCell.offset().left + oCell.width() + 3 } );
      oDiv.height( nHeight );
      oDiv.width( 3 );
    } );
  }

  function dragBorder( oDiv, nLeft ) {
    oDiv.css( 'left', nLeft + 2 );
  }

  function resizeColumn( oDiv, nRight ) 
  {
    var nColumn = oDiv.data( 'column' );
    var oTable = oDiv.data( 'table' );
    var oFirstRow = oTable.find( 'tr' ).first();
    var oCell = oFirstRow.find( 'td' ).eq( nColumn );
    oCell.width( nRight - oCell.offset().left );
    resetSliders(oTable);
  }

  function resetSliders(oTable)  //put all sliders on the correct position
  {
      var dragHandles = $( "div.drag_handle[ name = '" + oTable.attr( 'id') + "']" );
      var cols = oTable.find( 'tr' ).first().find( 'td' );
      var oDiv;
      var oCell;
      var nPos;
      var nCellInd;
      
      dragHandles.each(function()
      {
          oDiv = $( this );
          nCellInd = oDiv.data( 'column' );
          oCell = cols.eq( nCellInd );
          nPos = oCell.offset().left + oCell.outerWidth() ;
          oDiv.css( 'left', nPos );
      });
  }

  function isColSelected( oHeaderCell ) 
  {
    var oTable = oHeaderCell.closest( '.sp_spreadsheet' );
    var selectedCol = oTable.find('td.sp_colselected');
    if ( selectedCol.size() == 0 ) {   // column not selected
      return false;
    }
    var rows = oTable.find( 'tr' );
    nHeaderColInd = rows.first().find( 'td' ).index( oHeaderCell );
    
    var oSelectedCol = selectedCol.first();   // first selected column cell is in 2nd row
    return ( rows.eq( 1 ).find( 'td' ).index( oSelectedCol ) == nHeaderColInd );
  }
  
  /*
  * Public functions
  */
  
  $.fn.isReadOnly = function()
  {
    return $( this ).data( "read_only" );
  }
  
  $.fn.setReadOnly = function( bTrue )
  {
    $( this ).data( "read_only", bTrue );
    return $( this );
  }
  
  $.fn.toggleColumnHeader = function( bShow )
  {
    var oRowHeader = $( this ).find( '.sp_rowheader' );
    if ( bShow ) {
      oRowHeader.show();
    }
    else {
      oRowHeader.hide();
    }
    return $( this );
 }

  $.fn.toggleRowHeader = function( bShow )
  {
    var oColHeader = $( this ).find( '.sp_colheader' );
    if ( bShow ) {
      oColHeader.show();
    }
    else {
      oColHeader.hide();
    }
    return $( this );
  }

  $.fn.isRowHeader = function()
  {
    var oColHeader = $( this ).find( '.sp_colheader' );
    return ( oColHeader.is( ':visible' ) );
  }

  $.fn.isColumnHeader = function()
  {
    var oRowHeader = $( this ).find( '.sp_rowheader' );
    return ( oRowHeader.is( ':visible' ) );
  }

  $.fn.rowCount = function()
  {
    return $( this ).find( 'tr' ).size() - 1;
  }

  $.fn.colCount = function()
  {
    return $( this ).find( 'tr' ).first().children().size() - 1;
  }

  $.fn.insertRow = function(nIndex)
  {
    var oTable = $( this );
    if ( oTable.isReadOnly() ) {
      return;
    }
    var oInsertAtRow;
    var oNewRow;
    var bInsertLast;
    var nNewIndex;
    
    unSelectRow( oTable );
    unSelectColumn( oTable );
    unSelectCell( oTable );
    
    var rows = oTable.find( 'tr' );
    var oLastRow = rows.last();
    var nLastRowIndex = rows.index(oLastRow);
    bInsertLast = ( nIndex == undefined ) || ( nIndex > nLastRowIndex - 1 ) ||
                    isNaN( nIndex ) || ( nIndex < 0 );
    
    if ( !bInsertLast ) {
      oInsertAtRow = rows.eq( nIndex + 1 );
      oNewRow = oInsertAtRow.clone( true ).insertBefore( oInsertAtRow );
      
      // recalculate row headers
      var nCount;
      var oNextRow;
      
      rows = oTable.find( 'tr' );
      for ( nCount = nIndex + 1; nCount < rows.size(); nCount++ ) {
        oNextRow = rows.eq( nCount );
        oNextRow.find( 'td.sp_colheader' ).text( nCount );
      }
    }
    else {
      oNewRow = oLastRow.clone( true ).insertAfter( oLastRow );
      oNewRow.find( 'td.sp_colheader' ).text( nLastRowIndex + 1 );
    }
    oNewRow.height( ( oLastRow.height() ) );    // in case there is no row header
    oNewRow.children().not( 'td.sp_colheader' ).text( "" );
    
    
    unSort( oTable );           //sorting
    makeResizable( oTable );            //resizing
    oTable.setZebraStriped( oTable.data( "zebra_striped" ) );
    return oTable;
  }

  $.fn.removeRow  = function(nIndex)
  {
    var oTable = $( this );
    if ( oTable.isReadOnly() ) {
      return;
    }
    var rows = oTable.find( 'tr' );
    var oLastRow = rows.last();
    var nLastRowIndex = $( 'table#spreadsheet tr' ).index(oLastRow);
    var bRemoveLast = ( nIndex == undefined ) || ( nIndex > nLastRowIndex - 1 ) || isNaN( nIndex ) || ( nIndex < 0 );
    
    if ( !bRemoveLast ) {
      var oRemoveRow = rows.eq( nIndex + 1 );
      oRemoveRow.remove();
      
      // recalculate row headers
      var nCount;
      var oNextRow;
      
      rows = oTable.find( 'tr' );
      for ( nCount = nIndex + 1; nCount < rows.size(); nCount++ ) {
        oNextRow = rows.eq( nCount );
        oNextRow.find( 'td.sp_colheader' ).text( nCount );
      }
    }
    else {
      oLastRow.remove();
    }
    makeResizable( oTable );    //resizing
    oTable.setZebraStriped( oTable.data( "zebra_striped" ) );
    return oTable;
  }

  $.fn.insertColumn = function(nIndex)
  {
    var oTable = $( this );
    if ( oTable.isReadOnly() ) {
      return;
    }
    unSelectRow(oTable);
    unSelectColumn(oTable);
    unSelectCell(oTable);

    var rows = oTable.find( 'tr' );
    var oHeaderRow = rows.first();
    var oLastColumn = oHeaderRow.find( 'td' ).last();
    var nLastColIndex = oHeaderRow.find( 'td' ).index(oLastColumn);
    
    if ( nLastColIndex == 25 ) {    // max number of columns
      return;
    }
    var bInsertLast = ( nIndex == undefined ) || ( nIndex > nLastColIndex - 1 ) || isNaN( nIndex ) || ( nIndex < 0 );
    
    var nRowCount = oTable.rowCount();
    var nInd;
    var oNextRow;
    var oCol;
    
     for ( nInd = 0; nInd <= nRowCount; nInd++ ) {
        oNextRow = rows.eq( nInd );
        if ( !bInsertLast ) {
          oCol = $( '<td></td>' ).insertBefore( oNextRow.find( 'td' ).eq( nIndex + 1 ) );
        }
        else {
          oCol = $( '<td></td>' ).appendTo( oNextRow );
        }
        oCol.width( 100 );
    }
    
    // header
    if ( !bInsertLast ) {
       for ( nInd = nIndex; nInd <= nLastColIndex; nInd++ ) {
        oHeaderRow.find( 'td' ).eq( nInd + 1 ).text( String.fromCharCode(65 + nInd ) );
       }
    }
    else {
       oHeaderRow.find( 'td' ).eq( nLastColIndex + 1 ).text( String.fromCharCode(65 + nLastColIndex )  );
    }
    makeResizable( oTable );    //resizing
    return oTable;
  }

  $.fn.removeColumn = function(nIndex)
  {
    var oTable = $( this );
    if ( oTable.isReadOnly() ) {
      return;
    }
    var rows = oTable.find( 'tr' );
    var oHeaderRow = rows.first();
    var oLastColumn = oHeaderRow.find( 'td' ).last();
    var nLastColIndex = oHeaderRow.find( 'td' ).index(oLastColumn);
    var bRemoveLast = ( nIndex == undefined ) || ( nIndex > nLastColIndex - 1 ) || isNaN( nIndex ) || ( nIndex < 0 );
    
    var nRowCount = oTable.rowCount();
    var nInd;
    var oNextRow;
    
     for ( nInd = 0; nInd <= nRowCount; nInd++ ) {
        oNextRow = rows.eq( nInd );
        if ( !bRemoveLast ) {
          oNextRow.find( 'td' ).eq( nIndex + 1 ).remove();
        }
        else {
          oNextRow.find( 'td' ).eq( nLastColIndex ).remove();
       }
    }
    
    // header
    if ( !bRemoveLast ) {
       for ( nInd = nIndex; nInd <= nLastColIndex; nInd++ ) {
        oHeaderRow.find( 'td' ).eq( nInd + 1 ).text( String.fromCharCode(65 + nInd ) );
       }
    }
    
    // sorting
    var nSort = oTable.data( 'sort' );
    if ( nSort == nIndex ) {
      unSort(oTable);
    }
    makeResizable( oTable );    //resizing
    return oTable;
  }

  $.fn.setZebraStriped = function(bStriped)
  {
    var oTable = $( this );
    unSelectRow(oTable);
    unSelectColumn(oTable);
    unSelectCell(oTable);
    
    var rows = oTable.find( 'tr' );
    if ( bStriped ) {
      rows.filter( "tr:odd").addClass( "sp_oddrow" );
      rows.filter( "tr:even").removeClass( "sp_oddrow" );
    }
    else {
      rows.filter( "tr:odd").removeClass( "sp_oddrow" );
    }
    oTable.data( "zebra_striped", bStriped );
    return oTable;
  }

  $.fn.sortColumn  = function(iCol)
  {
    var oTable = $( this );
    var aTDStrings = new Array();
    var nInd;
    var nRowCount = oTable.rowCount();
    var sText;
    var nSeparator;
    var oNextRow;
    
    if ( ( iCol == undefined ) || ( iCol > oTable.colCount() - 1 ) || isNaN( iCol ) || ( iCol < 0 ) ) {
      return;
    }
    // remove images
    var rows = oTable.find( 'tr' );
    rows.first().find( 'td' ).each( function() { $(this).children().remove() } );

    var iSort = oTable.data( 'sort' );
    
    if ( iSort == null ) {
      oTable.data( 'sort', iCol );
      oTable.data( 'sort_order', 'd' );
      bDescending = true;
    }
    else if ( iSort != iCol ) {
      unSort(oTable);
      oTable.data( 'sort', iCol );
      oTable.data( 'sort_order', 'd' );
      bDescending = true;
    }
    else  {   // ( iSort == iCol )
      bDescending = ( oTable.data( 'sort_order' ) == "d" ) ? false : true;
      if ( bDescending ) {
        oTable.data( 'sort_order', 'd' );
      }
      else {
        oTable.data( 'sort_order', 'a' );
      }
    }
    
    for ( nInd = 1; nInd <= nRowCount; nInd++ ) {
      oNextRow = rows.eq( nInd );
      sText = oNextRow.find( 'td' ).eq( iCol + 1 ).text();
      sText += "^" + (nInd - 1);
      aTDStrings[ nInd - 1 ] = sText;
    }
    
      
    if ( bDescending ) {
      if ( oTable.isNumeric( iCol ) ) {
        aTDStrings.sort( compareD );
      }
      else {
        aTDStrings.sort();
     }
    }
    else {
      if ( oTable.isNumeric( iCol ) ) {
        aTDStrings.sort( compareA );
      }
      else {
        aTDStrings.reverse();
      }
    }
   
    var anRows = new Array();
    for ( nInd = 0; nInd < aTDStrings.length; nInd++ ) {
      sText = aTDStrings[ nInd ];
      nSeparator = sText.lastIndexOf( "^" );
      anRows[ nInd ] = sText.substr( nSeparator + 1 );
    }
    
    var aRows = new Array();
    aRows[0] = rows.first();  // header
    for ( nInd = 0; nInd < anRows.length; nInd++ ) {
      oNextRow = rows.eq( parseInt( anRows[ nInd ] ) + 1 );
      aRows[ nInd + 1 ] = oNextRow;
      oNextRow.find( 'td:first-child' ).text( nInd + 1 );
    }
    
    oTable.empty();
    for ( nInd = 0; nInd < aRows.length; nInd++ ) {
      oTable.append( aRows[ nInd ] );
    }
    oTable.setZebraStriped( oTable.data( "zebra_striped" ) );
      
    var sImgPath = bDescending ? "img/arrow_down.png" : "img/arrow_up.png"
    var img = $( '<img></img>' ).attr( 'src', sImgPath ).css( 'float', 'right' );
    var oCell = oTable.find( 'tr' ).first().find( 'td' ).eq( iCol + 1 );
    img.appendTo( oCell );
    return oTable;
  }

  $.fn.setNumeric  = function(iCol, bTrue)
  {
    var nInd;
    var oCell;
    var oTable = $( this );
    
    if ( ( iCol == undefined ) || ( iCol > oTable.colCount() - 1 ) || isNaN( iCol ) || ( iCol < 0 ) ) {
      return;
    }
    var nRowCount = oTable.rowCount();
    var rows = oTable.find( 'tr' );
    
    for ( nInd = 1; nInd <= nRowCount; nInd++ ) {
      oNextRow = rows.eq( nInd );
      oCell = oNextRow.find( 'td' ).eq( iCol + 1 );
      oCell.attr( "numeric", bTrue );
      if ( bTrue ) {
        oCell.css( 'text-align', 'right' );
      }
      else {
        oCell.css( 'text-align', 'left' );
      }
    }
    return oTable;
  }

  $.fn.isNumeric  = function(iCol)
  {
    var oTable = $( this );
    if ( ( iCol == undefined ) || ( iCol > oTable.colCount() - 1 ) || isNaN( iCol ) || ( iCol < 0 ) ) {
      return false;
    }
    var oFirstRow = oTable.find( 'tr' ).eq( 1 );
    var bData = oFirstRow.find( 'td' ).eq( iCol + 1 ).attr( "numeric" );
    return (bData == "true");
  }
  
  $.fn.getData  = function()
  {
    var data = new Array();
    var nRow, nCol;
    
    var oTable = $( this );
    var rows = oTable.find( 'tr' );
    var nRowCount = rows.size();
    
    var cols = rows.eq( 0 ).find( 'td' );
    var nColCount = cols.size();
    
    for ( nRow = 1; nRow < nRowCount; nRow++ ) {
      data[ nRow - 1 ] = new Array();
      cols = rows.eq( nRow ).find( 'td' );
      for ( nCol = 1; nCol < nColCount; nCol++ ) {
        data[ nRow-1 ][nCol-1] = cols.eq( nCol ).text();
      }
    }
    return data;
  }
})(jQuery);