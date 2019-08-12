#INCLUDE ..\src\foxyxls.h
LOCAL m.lcFile, m.loExcel
m.lcPath=SYS(16)
m.lcPath=IIF(RAT("\", m.lcPath)>0, LEFT(m.lcPath, RAT("\", m.lcPath)), m.lcPath)
m.lcFile = m.lcPath+"..\out\Test1.xls"


SET PROCEDURE TO (m.lcPath+"..\src\FoxyXLS.prg") ADDITIVE
m.loExcel = CREATEOBJECT("FoxyXLS")

m.loExcel.cAuthor = "VFPIMAGING"
m.loExcel.nCodePage = 1252

m.loExcel.nDefaultRowHeight   = 30 && Points
m.loExcel.nDefaultColumnWidth = 14 && Characters

*!* m.loExcel.SetColumnWidth(3, 180)
*!* m.loExcel.SetColumnWidth(1, 180)

m.loExcel.AddCell( 1, 1, "White"    , "Segoe UI,10,B,White")
m.loExcel.AddCell( 2, 1, "Red"      , "Segoe UI,10,B,Red")
m.loExcel.AddCell( 3, 1, "Green"    , "Segoe UI,10,B,Green")
m.loExcel.AddCell( 4, 1, "Blue"     , "Segoe UI,10,B,Blue")
m.loExcel.AddCell( 5, 1, "Yellow"   , "Segoe UI,10,B,Yellow")
m.loExcel.AddCell( 6, 1, "Magenta"  , "Segoe UI,10,B,Magenta")
m.loExcel.AddCell( 7, 1, "Cyan"     , "Segoe UI,10,B,Cyan")
m.loExcel.AddCell( 8, 1, "DarkRed"  , "Segoe UI,10,B,DarkRed")
m.loExcel.AddCell( 9, 1, "DarkGreen", "Segoe UI,10,B,DarkGreen")
m.loExcel.AddCell(10, 1, "DarkBlue" , "Segoe UI,10,B,DarkBlue")
m.loExcel.AddCell(11, 1, "Olive"    , "Segoe UI,10,B,Olive")
m.loExcel.AddCell(12, 1, "Purple"   , "Segoe UI,10,B,Purple")
m.loExcel.AddCell(13, 1, "Teal"     , "Segoe UI,10,B,Teal")
m.loExcel.AddCell(14, 1, "Silver"   , "Segoe UI,10,B,Silver")
m.loExcel.AddCell(15, 1, "Gray"     , "Segoe UI,10,B,Gray")
m.loExcel.AddCell(16, 1, "Black"    , "Segoe UI,10,B,Black")
m.loExcel.AddCell(17, 1, "Automatic", "Segoe UI,10,B,Automatic")

m.loExcel.AddCell(20, 1, "Date in BRITISH format", "SEGOE UI,12,I")
m.loExcel.AddCell(20, 3, DATE(), "SEGOE UI,12,I", "dd/mm/yyyy", XLSALIGN_CENTER)

m.loExcel.AddCell(21, 1, "Date in AMERICAN format", "SEGOE UI,12,I")
m.loExcel.AddCell(21, 3, DATE(), "SEGOE UI,12,I", "m/d/yy"    , XLSALIGN_CENTER)

m.loExcel.AddCell(23, 1, "Values", "SEGOE UI,12,I")
m.loExcel.AddCell(23, 3, 1500)

m.loExcel.AddCell(24, 1, "Formatted Values")
m.loExcel.AddCell(24, 3, 1500, , "#,##0.00")

m.loExcel.AddCell(25, 1, "Currency formatted Values")
m.loExcel.AddCell(25, 3, 1500, , "$#,##0.00")
m.loExcel.AddCell(26, 3, -1500, , "$#,##0.00")

m.loExcel.AddCell(27, 1, "Percentage")
m.loExcel.AddCell(27, 3, 0.252, , "0.00%")

m.loExcel.WriteFile(m.lcFile)

RELEASE m.loExcel

CLEAR CLASS FoxyXLS
RELEASE PROCEDURE (m.lcPath+"..\src\FoxyXLS")
