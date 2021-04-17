VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()


Dim OpenFileName As String
Dim wb As Workbook
'Select and Open workbook
OpenFileName = Application.GetOpenFilename("hand receipt file,*.xlsx")
If OpenFileName = "False" Then Exit Sub
Set wb = Workbooks.Open(OpenFileName)
'notify start
    Application.StatusBar = "Processing Worksheets... Please do not disturb..."
    DoEvents

    Application.ScreenUpdating = False
'notify end

Dim n As Integer
n = wb.Sheets.Count
p = 0
    Do
        ThisWorkbook.Sheets("template").Copy Before:=ThisWorkbook.Sheets("START")
            p = p + 1
        Loop Until p = n


On Error Resume Next
'Get data EXAMPLE
ThisWorkbook.Sheets(1).Name = wb.Sheets(1).Name
ThisWorkbook.Sheets(1).Range("a7:a107").Value = wb.Sheets(1).Range("a13:a113").Value
ThisWorkbook.Sheets(1).Range("f7:f107").Value = wb.Sheets(1).Range("c13:c113").Value
ThisWorkbook.Sheets(1).Range("s7:s107").Value = wb.Sheets(1).Range("d13:d113").Value
ThisWorkbook.Sheets(1).Range("t7:t107").Value = wb.Sheets(1).Range("e13:e113").Value
ThisWorkbook.Sheets(1).Range("u7:u107").Value = wb.Sheets(1).Range("f13:f113").Value
ThisWorkbook.Sheets(1).Range("v7:v107").Value = wb.Sheets(1).Range("h13:h113").Value
ThisWorkbook.Sheets(1).Range("c4").Value = wb.Sheets(1).Range("a6").Value
ThisWorkbook.Sheets(1).Range("i3").Value = wb.Sheets(1).Range("c6").Value
ThisWorkbook.Sheets(1).Range("r3").Value = wb.Sheets(1).Range("c7").Value
ThisWorkbook.Sheets(1).Range("x4").Value = wb.Sheets(1).Range("g7").Value
ThisWorkbook.Sheets(1).Range("y1").Value = wb.Sheets(1).Range("g6").Value
ThisWorkbook.Sheets(1).Range("y2").Value = wb.Sheets(1).Range("a7").Value
ThisWorkbook.Sheets(1).Range("aa2").Value = wb.Sheets(1).Range("g5").Value

ThisWorkbook.Sheets(2).Name = wb.Sheets(2).Name
ThisWorkbook.Sheets(2).Range("a7:a107").Value = wb.Sheets(2).Range("a13:a113").Value
ThisWorkbook.Sheets(2).Range("f7:f107").Value = wb.Sheets(2).Range("c13:c113").Value
ThisWorkbook.Sheets(2).Range("s7:s107").Value = wb.Sheets(2).Range("d13:d113").Value
ThisWorkbook.Sheets(2).Range("t7:t107").Value = wb.Sheets(2).Range("e13:e113").Value
ThisWorkbook.Sheets(2).Range("u7:u107").Value = wb.Sheets(2).Range("f13:f113").Value
ThisWorkbook.Sheets(2).Range("v7:v107").Value = wb.Sheets(2).Range("h13:h113").Value
ThisWorkbook.Sheets(2).Range("c4").Value = wb.Sheets(2).Range("a6").Value
ThisWorkbook.Sheets(2).Range("i3").Value = wb.Sheets(2).Range("c6").Value
ThisWorkbook.Sheets(2).Range("r3").Value = wb.Sheets(2).Range("c7").Value
ThisWorkbook.Sheets(2).Range("x4").Value = wb.Sheets(2).Range("g7").Value
ThisWorkbook.Sheets(2).Range("y1").Value = wb.Sheets(2).Range("g6").Value
ThisWorkbook.Sheets(2).Range("y2").Value = wb.Sheets(2).Range("a7").Value
ThisWorkbook.Sheets(2).Range("aa2").Value = wb.Sheets(2).Range("g5").Value

ThisWorkbook.Sheets(3).Name = wb.Sheets(3).Name
ThisWorkbook.Sheets(3).Range("a7:a107").Value = wb.Sheets(3).Range("a13:a113").Value
ThisWorkbook.Sheets(3).Range("f7:f107").Value = wb.Sheets(3).Range("c13:c113").Value
ThisWorkbook.Sheets(3).Range("s7:s107").Value = wb.Sheets(3).Range("d13:d113").Value
ThisWorkbook.Sheets(3).Range("t7:t107").Value = wb.Sheets(3).Range("e13:e113").Value
ThisWorkbook.Sheets(3).Range("u7:u107").Value = wb.Sheets(3).Range("f13:f113").Value
ThisWorkbook.Sheets(3).Range("v7:v107").Value = wb.Sheets(3).Range("h13:h113").Value
ThisWorkbook.Sheets(3).Range("c4").Value = wb.Sheets(3).Range("a6").Value
ThisWorkbook.Sheets(3).Range("i3").Value = wb.Sheets(3).Range("c6").Value
ThisWorkbook.Sheets(3).Range("r3").Value = wb.Sheets(3).Range("c7").Value
ThisWorkbook.Sheets(3).Range("x4").Value = wb.Sheets(3).Range("g7").Value
ThisWorkbook.Sheets(3).Range("y1").Value = wb.Sheets(3).Range("g6").Value
ThisWorkbook.Sheets(3).Range("y2").Value = wb.Sheets(3).Range("a7").Value
ThisWorkbook.Sheets(3).Range("aa2").Value = wb.Sheets(3).Range("g5").Value

ThisWorkbook.Sheets(4).Name = wb.Sheets(4).Name
ThisWorkbook.Sheets(4).Range("a7:a107").Value = wb.Sheets(4).Range("a13:a113").Value
ThisWorkbook.Sheets(4).Range("f7:f107").Value = wb.Sheets(4).Range("c13:c113").Value
ThisWorkbook.Sheets(4).Range("s7:s107").Value = wb.Sheets(4).Range("d13:d113").Value
ThisWorkbook.Sheets(4).Range("t7:t107").Value = wb.Sheets(4).Range("e13:e113").Value
ThisWorkbook.Sheets(4).Range("u7:u107").Value = wb.Sheets(4).Range("f13:f113").Value
ThisWorkbook.Sheets(4).Range("v7:v107").Value = wb.Sheets(4).Range("h13:h113").Value
ThisWorkbook.Sheets(4).Range("c4").Value = wb.Sheets(4).Range("a6").Value
ThisWorkbook.Sheets(4).Range("i3").Value = wb.Sheets(4).Range("c6").Value
ThisWorkbook.Sheets(4).Range("r3").Value = wb.Sheets(4).Range("c7").Value
ThisWorkbook.Sheets(4).Range("x4").Value = wb.Sheets(4).Range("g7").Value
ThisWorkbook.Sheets(4).Range("y1").Value = wb.Sheets(4).Range("g6").Value
ThisWorkbook.Sheets(4).Range("y2").Value = wb.Sheets(4).Range("a7").Value
ThisWorkbook.Sheets(4).Range("aa2").Value = wb.Sheets(4).Range("g5").Value

ThisWorkbook.Sheets(5).Name = wb.Sheets(5).Name
ThisWorkbook.Sheets(5).Range("a7:a107").Value = wb.Sheets(5).Range("a13:a113").Value
ThisWorkbook.Sheets(5).Range("f7:f107").Value = wb.Sheets(5).Range("c13:c113").Value
ThisWorkbook.Sheets(5).Range("s7:s107").Value = wb.Sheets(5).Range("d13:d113").Value
ThisWorkbook.Sheets(5).Range("t7:t107").Value = wb.Sheets(5).Range("e13:e113").Value
ThisWorkbook.Sheets(5).Range("u7:u107").Value = wb.Sheets(5).Range("f13:f113").Value
ThisWorkbook.Sheets(5).Range("v7:v107").Value = wb.Sheets(5).Range("h13:h113").Value
ThisWorkbook.Sheets(5).Range("c4").Value = wb.Sheets(5).Range("a6").Value
ThisWorkbook.Sheets(5).Range("i3").Value = wb.Sheets(5).Range("c6").Value
ThisWorkbook.Sheets(5).Range("r3").Value = wb.Sheets(5).Range("c7").Value
ThisWorkbook.Sheets(5).Range("x4").Value = wb.Sheets(5).Range("g7").Value
ThisWorkbook.Sheets(5).Range("y1").Value = wb.Sheets(5).Range("g6").Value
ThisWorkbook.Sheets(5).Range("y2").Value = wb.Sheets(5).Range("a7").Value
ThisWorkbook.Sheets(5).Range("aa2").Value = wb.Sheets(5).Range("g5").Value

ThisWorkbook.Sheets(6).Name = wb.Sheets(6).Name
ThisWorkbook.Sheets(6).Range("a7:a107").Value = wb.Sheets(6).Range("a13:a113").Value
ThisWorkbook.Sheets(6).Range("f7:f107").Value = wb.Sheets(6).Range("c13:c113").Value
ThisWorkbook.Sheets(6).Range("s7:s107").Value = wb.Sheets(6).Range("d13:d113").Value
ThisWorkbook.Sheets(6).Range("t7:t107").Value = wb.Sheets(6).Range("e13:e113").Value
ThisWorkbook.Sheets(6).Range("u7:u107").Value = wb.Sheets(6).Range("f13:f113").Value
ThisWorkbook.Sheets(6).Range("v7:v107").Value = wb.Sheets(6).Range("h13:h113").Value
ThisWorkbook.Sheets(6).Range("c4").Value = wb.Sheets(6).Range("a6").Value
ThisWorkbook.Sheets(6).Range("i3").Value = wb.Sheets(6).Range("c6").Value
ThisWorkbook.Sheets(6).Range("r3").Value = wb.Sheets(6).Range("c7").Value
ThisWorkbook.Sheets(6).Range("x4").Value = wb.Sheets(6).Range("g7").Value
ThisWorkbook.Sheets(6).Range("y1").Value = wb.Sheets(6).Range("g6").Value
ThisWorkbook.Sheets(6).Range("y2").Value = wb.Sheets(6).Range("a7").Value
ThisWorkbook.Sheets(6).Range("aa2").Value = wb.Sheets(6).Range("g5").Value

ThisWorkbook.Sheets(7).Name = wb.Sheets(7).Name
ThisWorkbook.Sheets(7).Range("a7:a107").Value = wb.Sheets(7).Range("a13:a113").Value
ThisWorkbook.Sheets(7).Range("f7:f107").Value = wb.Sheets(7).Range("c13:c113").Value
ThisWorkbook.Sheets(7).Range("s7:s107").Value = wb.Sheets(7).Range("d13:d113").Value
ThisWorkbook.Sheets(7).Range("t7:t107").Value = wb.Sheets(7).Range("e13:e113").Value
ThisWorkbook.Sheets(7).Range("u7:u107").Value = wb.Sheets(7).Range("f13:f113").Value
ThisWorkbook.Sheets(7).Range("v7:v107").Value = wb.Sheets(7).Range("h13:h113").Value
ThisWorkbook.Sheets(7).Range("c4").Value = wb.Sheets(7).Range("a6").Value
ThisWorkbook.Sheets(7).Range("i3").Value = wb.Sheets(7).Range("c6").Value
ThisWorkbook.Sheets(7).Range("r3").Value = wb.Sheets(7).Range("c7").Value
ThisWorkbook.Sheets(7).Range("x4").Value = wb.Sheets(7).Range("g7").Value
ThisWorkbook.Sheets(7).Range("y1").Value = wb.Sheets(7).Range("g6").Value
ThisWorkbook.Sheets(7).Range("y2").Value = wb.Sheets(7).Range("a7").Value
ThisWorkbook.Sheets(7).Range("aa2").Value = wb.Sheets(7).Range("g5").Value

ThisWorkbook.Sheets(8).Name = wb.Sheets(8).Name
ThisWorkbook.Sheets(8).Range("a7:a107").Value = wb.Sheets(8).Range("a13:a113").Value
ThisWorkbook.Sheets(8).Range("f7:f107").Value = wb.Sheets(8).Range("c13:c113").Value
ThisWorkbook.Sheets(8).Range("s7:s107").Value = wb.Sheets(8).Range("d13:d113").Value
ThisWorkbook.Sheets(8).Range("t7:t107").Value = wb.Sheets(8).Range("e13:e113").Value
ThisWorkbook.Sheets(8).Range("u7:u107").Value = wb.Sheets(8).Range("f13:f113").Value
ThisWorkbook.Sheets(8).Range("v7:v107").Value = wb.Sheets(8).Range("h13:h113").Value
ThisWorkbook.Sheets(8).Range("c4").Value = wb.Sheets(8).Range("a6").Value
ThisWorkbook.Sheets(8).Range("i3").Value = wb.Sheets(8).Range("c6").Value
ThisWorkbook.Sheets(8).Range("r3").Value = wb.Sheets(8).Range("c7").Value
ThisWorkbook.Sheets(8).Range("x4").Value = wb.Sheets(8).Range("g7").Value
ThisWorkbook.Sheets(8).Range("y1").Value = wb.Sheets(8).Range("g6").Value
ThisWorkbook.Sheets(8).Range("y2").Value = wb.Sheets(8).Range("a7").Value
ThisWorkbook.Sheets(8).Range("aa2").Value = wb.Sheets(8).Range("g5").Value


'PURPOSE: Find & Replace text/values throughout entire workbook
'SOURCE: www.TheSpreadsheetGuru.com



Set wb = Workbooks.Open(OpenFileName, UpdateLinks:=0)
wb.Close

    
'clean up start


'cleanup end

'THIS IS THE KEY

Dim sht As Worksheet
'what is dim?? it works without below.
'Dim fnd1 As Variant
'Dim fnd2 As Variant
'Dim fnd3 As Variant
'Dim rplc As Variant

fnd1 = "          PUB NUM: "
fnd2 = "          ITEM DESC: "
fnd3 = "END ITEM NSN: "
fnd4 = "PUB DATE: "
fnd5 = "SHR NUMBER: "
fnd6 = "SIGNATURE"
fnd7 = "RECEIVED BY/GRADE"
fnd8 = "DATE (DD/MM/YYYY)"
fnd9 = "CL NUMBER: "
fnd10 = "LIN: "
fnd11 = " -- NOT ON CATALOG"
rplc = ""


For Each sht In ActiveWorkbook.Worksheets
  sht.Cells.Replace what:=fnd1, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
  sht.Cells.Replace what:=fnd2, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
  sht.Cells.Replace what:=fnd3, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd4, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd5, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd6, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd7, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd8, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd9, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd10, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
sht.Cells.Replace what:=fnd11, Replacement:=rplc, _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
        
Next sht

    Application.ScreenUpdating = True
    Application.StatusBar = "Done"
MsgBox ("where muh conops")

End Sub






