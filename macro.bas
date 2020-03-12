Option Compare Text
Public Function GetWorkbookByNamePattern(Pattern As String) As Workbook
Dim wb As Workbook
 For Each wb In Application.Workbooks
   If wb.Name Like Pattern Then
     Set GetWorkbookByNamePattern = wb
     Exit Function
   End If
 Next wb
 Set GetWorkbookByNamePattern = Nothing
End Function

Sub RAC_CVI_Consumer_Check_v2_Report_Macro()

'turns screen updates off -> speeds up code
Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim wb As Workbook
Dim ws As Worksheet

Dim FilePath As String
Dim FileName As String

Dim LastRowCount As Long

FilePath = "C:\Users\" & Environ$("Username") & "\Desktop\RAC_CVI_Consumer_Check_v2_Exports_Macro\"
FileName = "Copy of RAC CVI Consumer Check v2 " & Format(Now(), "MM-DD-YYYY")

'checks for FilePath directory, creates if not found
If Len(Dir(FilePath, vbDirectory)) = 0 Then
   MkDir (FilePath)
End If

'--------------- FILE SELECT ---------------'

'selects workbook based on name pattern
Set wb = GetWorkbookByNamePattern("**RAC CVI Consumer Check v2**")

'ends process if workbook not open
If wb Is Nothing Then
    MsgBox "'RAC CVI Consumer Check v2 Report' is not open", vbCritical
        Exit Sub
Else
    wb.Activate
End If

'--------------- SETUP ---------------'

'rename raw data sheet name
With wb
    .Sheets("Sheet1").Name = "Data"
End With

'sets variable for raw data sheet
Set ws = wb.Worksheets("Data")

'finds limits of data to select
With ws
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
End With




End Sub
