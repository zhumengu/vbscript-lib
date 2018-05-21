dim objExcel, objWorkbook, objWorksheet
set objExcel = createobject("excel.application")

Const xlToLeft = -4159
Const xlUp = -4162
Const xlCenter = -4108
Const xlContext = -5002
Const xlLandscape = 2
include "excel\openworkbook.vbs"
include "excel\copysheet.vbs"
include "excel\opensheetfile.vbs"
