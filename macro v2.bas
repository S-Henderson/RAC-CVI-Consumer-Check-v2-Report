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

Sub RAC_CVI()

Dim wb As Workbook
Dim ws As Worksheet

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

'turns screen updates off -> speeds up code
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'finds limits of data to select
With ws
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
End With

'insert Raction column at start
With ws
    .Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("A1").Value = "Raction"
End With

'insert Patient Names Match column at after customer names
With ws
    .Columns("AN:AN").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("AN1").Value = "Patient Names Match"
End With

With ws
    With .Range("AN2")
            .Formula = "=IF(AJ2=AL2,""TRUE"",""FALSE"")"
    End With
    With .Range("AN2")
            .AutoFill Destination:=Range("AN2:AN" & LastRow)
    End With
End With

'set yellow conditional formatting
With ws.Range("=$A$1:$AZ$99999")
    With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""TAG""")
            .Interior.Color = RGB(255, 199, 206)
    End With
End With

'set green conditional formatting
With ws.Range("=$A$1:$AZ$99999")
    With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""Diff Patient""")
            .Interior.Color = RGB(198, 239, 206)
    End With
End With







End Sub
