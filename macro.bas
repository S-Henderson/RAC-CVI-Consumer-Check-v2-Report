'------------------------------------------------------------
' Name         : RAC CVI Consumer v2 Report Tool
' Author       : Scott Henderson
' Created      : 10/14/2020
' Last Updated : 05/28/2021
' Purpose      : Identify existing wearers and track them, in case the number of existing wearers claiming for the new wearers bonus ever need to be quantified for the client
' Input        : RAC CVI Consumer Check v2 Report from daily RAC email inbox
' Output       : Completed file of tagged and non-tagged transactions & CVI Exceptions of tagged transactions reports saved to the RAC_Reports_Exports folder on user Desktop
' Workflow     : 2 separate macro's. 1st to prep the report with Raction and patient name match checks.
'                2nd to create an exceptions file based on tagged transactions after a manual patient name check
'------------------------------------------------------------

Sub ASCII_Art()

'__________    _____  _________   _____________   ____.___  _________                                                   _________ .__                   __           ________
'\______   \  /  _  \ \_   ___ \  \_   ___ \   \ /   /|   | \_   ___ \  ____   ____   ________ __  _____   ___________  \_   ___ \|  |__   ____   ____ |  | __ ___  _\_____  \
' |       _/ /  /_\  \/    \  \/  /    \  \/\   Y   / |   | /    \  \/ /  _ \ /    \ /  ___/  |  \/     \_/ __ \_  __ \ /    \  \/|  |  \_/ __ \_/ ___\|  |/ / \  \/ //  ____/
' |    |   \/    |    \     \____ \     \____\     /  |   | \     \___(  <_> )   |  \\___ \|  |  /  Y Y  \  ___/|  | \/ \     \___|   Y  \  ___/\  \___|    <   \   //       \
' |____|_  /\____|__  /\______  /  \______  / \___/   |___|  \______  /\____/|___|  /____  >____/|__|_|  /\___  >__|     \______  /___|  /\___  >\___  >__|_ \   \_/ \_______ \
'        \/         \/        \/          \/                        \/            \/     \/            \/     \/                \/     \/     \/     \/     \/               \/

End Sub

Option Explicit

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

Sub Prep_Report()

Application.ScreenUpdating = False

'---------- DECLARE VARIABLES ----------'

Dim wb                As Workbook

Dim ws                As Worksheet
Dim ws_tracking       As Worksheet

Dim pre_last_row      As Long
Dim last_row          As Long

Dim patient_name_cell As Range

Dim save_file_path    As String
Dim file_name         As String

'---------- SELECT WORKBOOK ----------'

'Selects workbook based on name pattern
Set wb = GetWorkbookByNamePattern("**RAC CVI Consumer Check v2**")

'Ends process if workbook not open
If wb Is Nothing Then
    MsgBox "'RAC CVI Consumer Check v2 Report' is not open", vbCritical
        Exit Sub
Else
    wb.Activate
End If

Set ws = wb.Worksheets("Sheet1")

'---------- REMOVE FILTER ----------'

'Best practice to remove any filters at start

'Check for filter, turn off if exists
If ws.AutoFilterMode = True Then
    ws.AutoFilterMode = False
End If

'---------- REMOVE DUPLICATES ----------'

'Get the last row value BEFORE removing duplicates
With ws
    pre_last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

'Remove duplicates by Transaction (column 5) AND Previous Claim Number (column 29)
With ws
    .Range("A1:AS" & pre_last_row).RemoveDuplicates Columns:=Array(5, 29), Header:=xlYes
End With

'---------- FIND WS LIMITS ----------'

'Get the last row value AFTER removing duplicates
With ws
    last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

'---------- TRIM PATIENT NAMES ----------'

'Get rid of leading/trailing white space in patient names
For Each patient_name_cell In ws.Range("AK2:AN" & last_row)
    patient_name_cell = WorksheetFunction.Trim(patient_name_cell)
Next patient_name_cell

'---------- CREATE RACTION COLUMN ----------'

'Create Raction column
With ws
    .Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("A1").Value = "Raction"
End With

'---------- DATA VALIDATION LIST OPTIONS ----------'

'Data validation drop down list for Raction reasons
With ws
    With .Range("A2:A" & last_row)
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="TAG, PREV TAG, DIFF PATIENT, IS, BH TAG"
    End With
End With

'---------- PATIENT NAMES CHECK COLUMN ----------'

'Create Patient First Name Match column
With ws
    .Columns("AP:AP").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    .Range("AP1").Value = "Patient First Name Match"
End With

'Formula to check Patient First Name Match
With ws
    .Range("AP2:AP" & last_row).Formula = "=$AL2=$AN2"
End With

'---------- CREATE RACTION FORMULA ----------'

'put formula rules at end so ending formula not weird/offset since adding columns at start of data

'Formula for Raction -> lots of double escape quotes for string checks
With ws
    .Range("A2:A" & last_row).Formula = "=IFS($AB2=""True"",""BH TAG"",$AJ2=""Invalid Submission"",""IS"",$AU2<>"""",""PREV TAG"",$AP2=TRUE,""TAG"",$AP2=FALSE,""DIFF PATIENT"")"
End With

'---------- CONDITIONAL FORMATTING RULES ----------'

'Visually compare different Raction reasons
With ws

    'TAG -> red highlight rows
    With .Range("A:AU")
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""TAG""")
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
        End With
    End With
    
    'PREV TAG -> yellow highlight rows
    With .Range("A:AU")
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""PREV TAG""")
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 87, 0)
        End With
    End With
    
    'DIFF PATIENT -> green highlight rows
    With .Range("A:AU")
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""DIFF PATIENT""")
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
        End With
    End With
    
    'BH TAG -> red highlight rows
    With .Range("A:AU")
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""DIFF PATIENT""")
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
        End With
    End With
    
    'IS -> green highlight rows
    With .Range("A:AU")
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1=""DIFF PATIENT""")
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
        End With
    End With
    
End With

'---------- TRACKING INFO WORKSHEET ----------'

'Create new sheet for recording tracking info
wb.Sheets.Add(After:=Sheets("Sheet1")).Name = "Tracking_Info"

'Set worksheet object
Set ws_tracking = wb.Worksheets("Tracking_Info")

'Tracking headers
With ws_tracking
    .Range("A2").Value = "Channel"
    .Range("A3").Value = "Consumer"
    .Range("B1").Value = "Hits"
    .Range("C1").Value = "Assessed"
    .Range("D1").Value = "Worked"
    .Range("E1").Value = "Worked Value ($)"
End With

'Tracking Values -> Channel (Channel claims not applicable for this report/client so just 0)
With ws_tracking
    .Range("B2").Value = "0" 'Hits
    .Range("C2").Value = "0" 'Assessed
    .Range("D2").Value = "0" 'Worked
    .Range("E2").Value = "0" 'Worked Value
End With

'Tracking Values -> Consumer
'-1 values account for column headers
With ws_tracking
    .Range("B3").Value = pre_last_row - 1                          'Hits
    .Range("C3").Value = "=COUNTA(Sheet1!F:F)-1"                   'Assessed
    .Range("D3").Formula = "=COUNTIF(Sheet1!A:A,""TAG"")"          'Worked
    .Range("E3").Formula = "=SUMIF(Sheet1!A:A,""TAG"",Sheet1!H:H)" 'Worked Value
End With

'---------- VIEW MAIN DATA SHEET ----------'

ws.Activate

'---------- TURN ON FILTER ----------'

'Check for filter, turn on if none exists
If ws.AutoFilterMode = False Then
    ws.Range("A1").AutoFilter
End If

'---------- SAVE DIRECTORY CHECK ----------'

'This is path where report(s) are saved
save_file_path = "C:\Users\" & Environ$("Username") & "\Desktop\RAC_Reports_Exports\"

'Checks for save_file_path directory, creates if not found
If Len(Dir(save_file_path, vbDirectory)) = 0 Then
   MkDir (save_file_path)
End If

'---------- SAVE MAIN WORKBOOK ----------'

'Save file name
file_name = "Copy of RAC CVI Consumer Check v2 " & Format(Now(), "MM-DD-YYYY") & ".xlsx"

'Save file
wb.SaveAs Filename:=save_file_path & file_name

'---------- SCRIPT COMPLETED ----------'

Application.ScreenUpdating = True

MsgBox ("Prep Macro Completed Sucessfully")

End Sub

Sub Create_Exceptions()

Application.ScreenUpdating = False

'---------- DECLARE VARIABLES ----------'

Dim wb                   As Workbook
Dim wb_exceptions        As Workbook

Dim ws                   As Worksheet
Dim ws_exceptions        As Worksheet

Dim last_row             As Long
Dim exceptions_last_row  As Long

Dim save_file_path       As String
Dim exceptions_file_name As String

'---------- SELECT WORKBOOK ----------'

'Selects workbook based on name pattern
Set wb = GetWorkbookByNamePattern("**RAC CVI Consumer Check v2**")

'Ends process if workbook not open
If wb Is Nothing Then
    MsgBox "'RAC CVI Consumer Check v2 Report' is not open", vbCritical
        Exit Sub
Else
    wb.Activate
End If

Set ws = wb.Worksheets("Sheet1")

'---------- REMOVE FILTER ----------'

'Best practice to remove any filters at start

'Check for filter, turn off if exists
If ws.AutoFilterMode = True Then
    ws.AutoFilterMode = False
End If

'---------- FIND WS LIMITS ----------'

'Get the last row value
With ws
    last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

'---------- CREATE EXCEPTIONS SHEET ----------'

'Create new sheet for exceptions transactions (TAG)
wb.Sheets.Add(After:=Sheets("Tracking_Info")).Name = "Exceptions"

'Set worksheet object
Set ws_exceptions = wb.Worksheets("Exceptions")

'---------- EXCEPTIONS SHEET SETUP ----------'
    
'Exceptions headers
With ws_exceptions
    .Range("A1").Value = "Transaction"
    .Range("B1").Value = "Exception Reason"
End With

'---------- FILTER & COPY TAG TRANSACTIONS ----------'

'Filter and copy TAG transactions
With ws

    'Filter on Raction (column 1)
    With .Range("A1:AU" & last_row)
            .AutoFilter Field:=1, Criteria1:="TAG"
        
        'Copy/paste
        With .Range("F2:F" & last_row)
                .SpecialCells(xlCellTypeVisible).Copy Destination:=ws_exceptions.Range("A2")
        End With
    
    End With
    
End With

'Unfilter
On Error Resume Next
    ws.ShowAllData
On Error GoTo 0

'---------- FIND WS EXCEPTIONS LIMITS ----------'

'Get the last row value for the exceptions sheet
With ws_exceptions
    exceptions_last_row = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

'---------- FILL EXCEPTIONS SHEET ----------'

'Fill in exceptions data
With ws_exceptions
    .Range("B2:B" & exceptions_last_row).Value = "Existing Wearer"
End With

'---------- CREATE EXCEPTIONS WORKBOOK ----------'

'Copies exceptions to a new blank workbook

'Create new workbook
Set wb_exceptions = Workbooks.Add

'Copy exceptions sheet from main report to new workbook
ws_exceptions.Copy Before:=wb_exceptions.Sheets(1)

'Switch OFF the alert button for saving pop-up for deleting sheet
Application.DisplayAlerts = False

'Delete extra worksheet
With wb_exceptions
    .Sheets("Sheet1").Delete
End With

'Switch back ON the alert button
Application.DisplayAlerts = True

'---------- VIEW MAIN DATA SHEET ----------'

ws.Activate

'---------- TURN ON FILTER ----------'

'Check for filter, turn on if none exists
If ws.AutoFilterMode = False Then
    Range("A1").AutoFilter
End If

'---------- SAVE DIRECTORY CHECK ----------'

'This is path where report(s) are saved
save_file_path = "C:\Users\" & Environ$("Username") & "\Desktop\RAC_Reports_Exports\"

'Checks for save_file_path directory, creates if not found
If Len(Dir(save_file_path, vbDirectory)) = 0 Then
   MkDir (save_file_path)
End If

'---------- SAVE EXCEPTIONS FILE ----------'

'Save exceptions file name
exceptions_file_name = "CVI Exceptions " & Format(Now(), "MM-DD-YYYY") & ".xlsx"

'Save exceptions file
wb_exceptions.SaveAs Filename:=save_file_path & exceptions_file_name

'---------- SAVE MAIN WORKBOOK ----------'

'Save Workbook with exceptions sheet now included
wb.Save

'---------- SCRIPT COMPLETED ----------'

Application.ScreenUpdating = True

MsgBox ("Exceptions Macro Completed Sucessfully")

End Sub
