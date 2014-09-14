SheetConverter library (formerly "Spreadsheet to HTML")
==============

Google Apps Script library - interprets Google Sheets Formats, converts to formatted text or html

This is the source repository for the [SheetConverter Google Apps Script library](https://sites.google.com/site/scriptsexamples/custom-methods/sheetconverter).

Libary documentation is [available here](https://script.google.com/macros/library/d/Mo6Ljr7ZKrMeYO9mHqOdo9oxc-w7mEonb/5).

## Caveats ##

This script is incomplete, ignoring some types of formatting. (Feel free to fork and enhance it, if you wish! Broadly applicable enhancements can be merged and the library updated) There are also some known issues:

 - It produces <del>WAY too much</del> *a minimal amount* of inline style for each cell. It would be nicer if it could optionally generate and reuse css class styles where possible. (On the flip side - this is just what is needed for embedding into gmail, so if it's changed, it should support both.)
 - The general table style, including borders, is set in the `tableFormat` variable. Since there is no way to determine what borders are in place on a spreadsheet, it isn't possible to transfer them. (Possible enhancement: support a parameter for borders.)
 - <del>Numeric formatting can be read from a spreadsheet, but is not directly adaptable in Javascript, so numbers aren't rendered as they appear in the spreadsheet</del>. Most formats are now supported! Fractions are very slow. Custom formats can only be supported best-effort. Further extension of special negative numbering is needed (e.g. support for brackets & colours). Locale-based radix & thousand separators are also a future extension.
 - No special treatment for table headers - that would be a nice addition, allowing adoption of existing css.
 - No handling of Merged cells.
 - Overflow is not supported; wrap is applied to all cells.
 - Images and charts aren't supported.

## License ##

A BSD 2-clause license applies.

Copyright (c) 2014, David Bingham
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

 1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

 2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

## To use ##

    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var sheet = ss.getActiveSheet();
    var range = sheet.getDataRange();
    var htmlTable = getHtmlTable(range);  // Produce HTML table of entire spreadsheet
