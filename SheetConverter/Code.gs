/**
 * Converter
 *
 * Library of spreadsheet range conversion functions to return cell contents as strings, applying the spreadsheet formats defined for the cell(s).
 *
 * <pre>
 * Usage:
 *         var converter = Converter.init(ss.getSpreadsheetTimeZone(),
 *                                        ss.getSpreadsheetLocale());
 *         // Get formats from range
 *         for (row=0;row<data.length;row++) {
 *           for (col=0;col<data[row].length;col++) {
 *             // Get formatted data
 *             var cellText = converter.convertCell(data[row][col],);
 *             Logger.log(cellText);
 *           }
 *         }    
 * </pre>
 *
 * Copyright (c) 2014, David Bingham, All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 *
 * Redistributions of source code must retain the above copyright notice, this list of
 * conditions and the following disclaimer.
 *
 * Redistributions in binary form must reproduce the above copyright notice, this list
 * of conditions and the following disclaimer in the documentation and/or other
 * materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
 * OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT
 * SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 * INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
 * TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR
 * BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
 * ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH
 * DAMAGE.
 */

var thisInstance_ = this;

/**
 * Utility function to test object class.
 * Note: does not work on Google Apps Script service classes.
 *
 * Example: Logger.log(toClass_.call(someday));  logs [object Date]
 * if (objIsClass_(someday,"Date")) ...
 * 
 * from http://javascript.info/tutorial/type-detection
 */
var toClass_ = {}.toString;
function objIsClass_(object,className) {
  return (toClass_.call(object).indexOf(className) !== -1);
}

/**
 * Initialize Converter to use a specific time zone and locale. This is optional, as
 * the library will use the script Session settings by default. Those defaults may
 * be different than the Spreadsheet settings, which could result in incorrect
 * dates and times.
 *
 * Caveat: localization stubbed for future development.
 *
<pre>
Usage:
&nbsp;  var values = Converter.init(ss.getSpreadsheetTimeZone(),
&nbsp;                              ss.getSpreadsheetLocale())
&nbsp;                        .convertRange(range);
</pre>
 *
 * @param  {String}  tzone  Timezone to use for date/time interpretation
 * @param  {String}  locale Locale to use for default date/time/currency.
 *
 * @return {Converter}      this object, for chaining
 */
function init( tzone, locale ) { // Constructor
  thisInstance_.tzone = tzone || thisInstance_.tzone || Session.getScriptTimeZone();
  thisInstance_.locale = locale || thisInstance_.locale || Session.getActiveUserLocale();  // For future localization of numbers, times, dates.
  return thisInstance_;
}


/**
 * Performs conversion of range values according to range formats.
 * 
<pre>
Usage:
&nbsp;  var values = Converter.convertRange(range);
</pre>
 * 
 *
 * @param  {Object}  range     Spreadsheet range to convert
 *
 * @return {Array}             2-D array of Strings. e.g.
<pre>
&nbsp; [["Stock", "Company", "Shares", "Value", "% Port", "% Goal", "C. Price"],
&nbsp;  ["AAPL", "Apple Inc.", "200", "$19596.00", "15.73%", "15.00%", "$97.98"],
&nbsp;  ["AMZN", "Amazon.com, Inc.", "25", "$8340.75", "6.70%", "15.00%", "$333.63"],
&nbsp;  ["Totals", "", "", "$124571.28", "100.00%", "100.00%", ""]]
</pre>
 */
function convertRange(range){
  // Must have 1 parameter that is a range object
  if (arguments.length !== 1 || Object.keys(range).indexOf("getValues") === -1)
    throw new Error( 'Invalid parameter(s)' );
  
  // Read range contents
  var data = range.getDisplayValues();
  
  // Build output array
  var result = [];
  
  // Populate rows
  for (row=0;row<data.length;row++) {
    result[row] = [];
    for (col=0;col<data[row].length;col++) {
      // Get formatted data
      result[row][col] = convertCell(data[row][col]);
    }
  }

  //debugger;

  return result;
}

function test_convertRange_() {
  var ss = SpreadsheetApp.openById('1ujI8lmwj1SBiNA1Sg99RXYOECEIbDGAFCnR4ISRCjuE');
  var range = ss.getActiveSheet().getDataRange();

  // Get ready to convert data
  var converter = init(ss.getSpreadsheetTimeZone(),
                       ss.getSpreadsheetLocale());
  var vals = converter.convertRange(range);
  debugger;
}

/**
 * Return a string containing an HTML table representation
 * of the given range, preserving style settings.
 * 
<pre>
Usage:
&nbsp;  var values = Converter.convertRange2html(range);
</pre>
 * 
 * @param {Range} range Spreadsheet range to render as HTML
 * 
 * @returns {String}    HTML version of range, e.g.
<pre>
&nbsp; &lt;table cellspacing...>
&nbsp;  &lt;colgroup>
&nbsp;   &lt;col width="62">
&nbsp;   &lt;col width="150">
&nbsp;  &lt;/colgroup>
&nbsp;  &lt;tbody>
&nbsp;   &lt;tr style="height: 21px;">
&nbsp;    &lt;td style="...>Stock&lt;/td>
&nbsp;    &lt;td style="...>Company&lt;/td>
&nbsp;   &lt;/tr>
&nbsp;   &lt;tr style=3D"height: 21px;">
&nbsp;    &lt;td style="...>AAPL&lt;/td>
&nbsp;    &lt;td style="...>Apple Inc.&lt;/td>
&nbsp;   &lt;/tr>
&nbsp;  &lt;/tbody>
&nbsp; &lt;/table>
</pre>
 */
function convertRange2html(range){
  var ss = range.getSheet().getParent();
  var sheet = range.getSheet();
  startRow = range.getRow();
  startCol = range.getColumn();
  lastRow = range.getLastRow();
  lastCol = range.getLastColumn();
  
  // Get ready to convert data
  var converter = thisInstance_.init();

  // Read table contents
  var data = range.getDisplayValues();

  // Get css style attributes from range
  var fontColors = range.getFontColors();
  var backgrounds = range.getBackgrounds();
  var fontFamilies = range.getFontFamilies();
  var fontSizes = range.getFontSizes();
  // getFontLines() ignores strike-through if cell also uses underline
  // https://code.google.com/p/google-apps-script-issues/issues/detail?id=4200
  var fontLines = range.getFontLines();
  var fontStyles = range.getFontStyles();
  var fontWeights = range.getFontWeights();
  var horizontalAlignments = range.getHorizontalAlignments();
  var verticalAlignments = range.getVerticalAlignments();
  var mergedRanges = range.getMergedRanges();
  
  // https://code.google.com/p/google-apps-script-issues/issues/detail?id=4187
  // Reported widths and heights can be incorrect, with "default" values.
  // If we read GAS defaults (120 wide, 17 high) replace with sheets defaults
  // (100 wide, 21 high). Will be wrong sometimes, but not with defaults!
  // Get column widths in pixels
  var colWidths = [];
  var tableWidth = 0;
  for (var col=startCol; col<=lastCol; col++) { 
    colWidths.push(120==sheet.getColumnWidth(col)?100:sheet.getColumnWidth(col));
    tableWidth += colWidths[colWidths.length-1];
  }
  // Get Row heights in pixels
  var rowHeights = [];
  for (var row=startRow; row<=lastRow; row++) { 
    rowHeights.push(17==sheet.getRowHeight(row)?21:sheet.getRowHeight(row));
  }

  // Get the UrlLinks for range
    var urlLinks = [];
    for (var row=startRow; row<=lastRow; row++) {
       var w = [];
       for (var col=startCol; col<=lastCol; col++) {
          var url = null;
          var url = sheet.getRange(row,col).getRichTextValue();
          var urlString = null;
          if (url){
            urlString = sheet.getRange(row,col).getRichTextValue().getLinkUrl();
          }
          if ((url == null)|(urlString ==null)) { urlString = 0;}
          w.push(urlString);
          }
       urlLinks.push(w);
       }
    
  // Future consideration...
  //var wraps = range.getWraps();
  
  // Build HTML Table, with inline styling for each cell
  // Default cell styling appears in table or row, so only minimal overrides need to be given for each cell
  var tableFormat = 'cellspacing="0" cellpadding="0" dir="ltr" border="1" style="width:TABLEWIDTHpx;table-layout:fixed;font-size:9pt;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc;font-weight:normal;color:black;background-color:white;text-align:right;text-decoration:none;font-style:normal;"';
   
  var html = ['<table '+tableFormat+'>'];
  // Column widths appear outside of table rows
  html.push('<colgroup>');
  for (col=0;col<colWidths.length;col++) {
    html.push('<col width=XXX>'.replace('XXX',colWidths[col]))
  }
  html.push('</colgroup>');
  html.push('<tbody>');
  
  // Populate rows
  for (row=0;row<data.length;row++) {
    html.push('<tr style="height:XXXpx;vertical-align:bottom;">'.replace('XXX',rowHeights[row]));
    for (col=0;col<data[row].length;col++) {
      // Get formatted data
      var cellText = converter.convertCell(data[row][col],true);

      var style = 'style="' 
                + 'padding:1px 2px; '
                + 'color:XXX;'.replace('XXX',fontColors[row][col].replace('general-','')).replace('color:black;','')
                + 'font-family:XXX;'.replace('XXX',fontFamilies[row][col]).replace('font-family:arial,sans,sans-serif;','')
                + 'font-size:XXXpt;'.replace('XXX',fontSizes[row][col]).replace('font-size:10pt;','')
                + 'font-weight:XXX;'.replace('XXX',fontWeights[row][col]).replace('font-weight:normal;','')
                + 'background-color:XXX;'.replace('XXX',backgrounds[row][col]).replace('background-color:white;','')
                + 'text-align:XXX;'.replace('XXX', horizontalAlignments[row][col]
                                             .replace('general-','')
                                             .replace('general','center')).replace('text-align:right;','')
                + 'vertical-align:XXX;'.replace('XXX',verticalAlignments[row][col]).replace('vertical-align:bottom;','')
                + 'text-decoration:XXX;'.replace('XXX',fontLines[row][col]).replace('text-decoration:none;','')
                + 'font-style:XXX;'.replace('XXX',fontStyles[row][col]).replace('font-style:normal;','')
                + 'border:1px solid black;'  // Need this, to override caja-guest td border-bottom
                + 'overflow:hidden;'
                +'"';

      var thisRow = range.getRow() + row;
      var thisCol = range.getColumn() + col;
      var rcSpan = "";
      var isTdPartOfRange = false;
      
      for (var i = 0; i < mergedRanges.length; i++) {
        var currentMergedRange = mergedRanges[i];
        
        var currentMergedRangeBoundaries = {
          top : currentMergedRange.getRow(),
          bottom : currentMergedRange.getRow() + currentMergedRange.getNumRows() - 1,
          left : currentMergedRange.getColumn(),
          right : currentMergedRange.getColumn() + currentMergedRange.getNumColumns() - 1
        };
        
        if ((thisRow == currentMergedRangeBoundaries.top && thisCol == currentMergedRangeBoundaries.left)) {
          // top left cell of range
          if (currentMergedRange.getNumRows() > 0) {
            rcSpan = " rowspan='"+currentMergedRange.getNumRows()+"'";
          }
          if (currentMergedRange.getNumColumns() > 0) {
            rcSpan += " colspan='"+currentMergedRange.getNumColumns()+"'";
          }
        } else if ((thisRow >= currentMergedRangeBoundaries.top && thisRow <= currentMergedRangeBoundaries.bottom) && (thisCol >= currentMergedRangeBoundaries.left && thisCol <= currentMergedRangeBoundaries.right)) {
          // falls in range
          isTdPartOfRange = true;
          break;
        }
      }
      
      if (!isTdPartOfRange) {
        var link = (urlLinks[row][col]);
        if (link !== 0) {
              var link = (urlLinks[row][col]);
              cellText = '<a href="' + link + '">' + cellText + '</a>';
         }
        html.push('<td XXX SPAN>'.replace('SPAN', rcSpan).replace('XXX',style)
                +String(cellText)
                +'</td>');
                }
    }
    html.push('</tr>');
  }
  html.push('</tbody>');
  html.push('</table>');
  
  //debugger;
  //return '<!--StartFragment--><meta name="generator" content="Sheets"><style type="text/css"><!--td {border: 1px solid #ccc;}br {mso-data-placement:same-cell;}--></style><table cellspacing="0" cellpadding="0" dir="ltr" border="1" style="table-layout:fixed;font-size:13px;font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc"><colgroup><col width="100"><col width="155"><col width="100"></colgroup><tbody><tr style="height:21px;"><td style="padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#ffff00;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;" data-sheets-value="[null,2,&quot;100px wide&quot;]">100px wide</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;font-family:georgia;font-size:140%;font-style:italic;color:#00ffff;" data-sheets-value="[null,2,&quot;Blue/Black&quot;]">Blue/Black</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;text-align:right;" data-sheets-value="[null,3,null,2]">2</td></tr><tr style="height:21px;"><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;text-align:right;" data-sheets-value="[null,3,null,2.3]">2.3</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;font-family:courier new,monospace;text-decoration:underline line-through;" data-sheets-value="[null,2,&quot;155 px wide&quot;]">155 px wide</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;text-align:right;" data-sheets-value="[null,3,null,41839]" data-sheets-numberformat="[null,5]" data-sheets-formula="=today()">7/19/2014</td></tr><tr style="height:30px;"><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;font-weight:bold;color:#ff0000;text-align:center;" data-sheets-value="[null,3,null,41822]" data-sheets-numberformat="[null,5,&quot;M/d/yyyy&quot;,1]">7/2/2014</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;"></td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;font-family:georgia;font-weight:bold;" data-sheets-value="[null,2,&quot;sadf&quot;]">sadf</td></tr><tr style="height:50px;"><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;border-left:1px solid #000000;text-decoration:underline line-through;vertical-align:top;text-align:right;" data-sheets-value="[null,2,&quot;asr&quot;]">asr</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;font-size:360%;vertical-align:bottom;" data-sheets-value="[null,2,&quot;asdf&quot;]">asdf</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;border-right:1px solid #000000;border-bottom:1px solid #000000;vertical-align:middle;text-align:center;" data-sheets-value="[null,2,&quot;100 px wide&quot;]">100 px wide</td></tr></tbody></table><!--EndFragment-->'

  return html.join('').replace('TABLEWIDTH',tableWidth);
}


/**
 * Performs conversion of cell contents according to their spreadsheet format.
 *
 * @param  {String}  cellDisplayValue
 * @param  {Boolean} htmlReady (optional) Set true if strings should be html-friendly.
 *                             Default is false, output is plain text.
 *
 * @return {String}            Formatted string. May contain HTML, depending on
 *                             htmlReady.
 */
function convertCell(cellDisplayValue,htmlReady) {
  // Must have 2 or 3 parameters, format must be string
  if (arguments.length < 1)
    throw new Error( 'Invalid parameter(s)' );
  htmlReady = htmlReady || false;
  thisInstance_.init(); // Ensure instance variables are set
  
  if (cellDisplayValue == "") return '';  // Not much to do with blank cells - just return an empty string

  // Sanitize string if output is for html
  if (htmlReady) cellDisplayValue = cellDisplayValue.replace(/ /g,"&nbsp;").replace(/</g,"&lt;").replace(/\n/g,"<br>");
  return cellDisplayValue;
}
