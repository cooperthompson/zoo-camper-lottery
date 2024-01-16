Option Explicit
Option Private Module

Public RegistrationWorksheet As Worksheet
Public RegistrationTable As ListObject

Public ConfigWorksheet As Worksheet
Public ConfigTable As ListObject

Public Sub Initialize()
    Dim RegistrationRange As Range
    
    On Error Resume Next
    ThisWorkbook.Sheets("Lottery Results").Delete
    ThisWorkbook.Sheets("Camp Config").Delete
    On Error GoTo 0
    
    ThisWorkbook.Sheets(1).Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Lottery Results"
    
    Set RegistrationWorksheet = ThisWorkbook.Sheets("Lottery Results")
    
    Dim EventCell As Range
    Set EventCell = Range("A:A").Find(What:="Event")
    ' MsgBox (EventCell.Address(False, False, xlA1, xlExternal))
    
    Dim LastRow As Long
    Dim LastColumn As Long
    Dim LastCell As Range
    
    LastRow = EventCell.CurrentRegion.Rows(EventCell.CurrentRegion.Rows.Count).Row
    LastColumn = EventCell.CurrentRegion.Columns(EventCell.CurrentRegion.Columns.Count).Column
    Set LastCell = Cells(LastRow, LastColumn)
    
    Set RegistrationTable = RegistrationWorksheet.ListObjects.Add(xlSrcRange, Range(EventCell, LastCell), , xlYes)
    RegistrationTable.Name = "LotteryResults"
    
    RegistrationTable.ListColumns.Add(RegistrationTable.ListColumns.Count + 1).Name = "Applicants"
    RegistrationTable.ListColumns("Applicants").Range.EntireColumn.AutoFit
    
    RegistrationTable.ListColumns.Add(RegistrationTable.ListColumns.Count + 1).Name = "Lottery Selection Status"
    RegistrationTable.ListColumns("Lottery Selection Status").Range.EntireColumn.AutoFit
    
    Call GenConfig
    
    RegistrationTable.ListColumns("Applicants").DataBodyRange.NumberFormat = "General"
    RegistrationTable.ListColumns("Applicants").DataBodyRange.Formula = "=VLOOKUP([@Event],ConfigTable[#All],2,FALSE)"
    
End Sub

Public Sub GenConfig()
    
    On Error Resume Next
    ThisWorkbook.Sheets("Camp Config").Delete
    On Error GoTo 0
    
    
    Dim ConfigSheetName As String
    ConfigSheetName = "Camp Config"
    
    
    Dim Sheet As Object
    Dim ConfigSheetExists As Boolean
    
    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Name = ConfigSheetName Then
            MsgBox ("The " & ConfigSheetName & "worksheet already exists.  Delete it, and re-run the config generator if you want to regen the config table.")
            ConfigSheetExists = True
        End If
    Next Sheet
                
    If Not ConfigSheetExists Then
        Set ConfigWorksheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ConfigWorksheet.Name = ConfigSheetName
    End If
    
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, RegistrationTable.Range.Address(False, False, xlA1, xlExternal))
    Set pt = pc.CreatePivotTable(ConfigWorksheet.Range("A1"))
    With pt.PivotFields("Event")
        .Orientation = xlRowField
        .Position = 1
    End With
    
    pt.AddDataField pt.PivotFields("Registration #"), "Count of Registrations", xlCount
    pt.PivotFields("Event").AutoSort xlDescending, "Count of Registrations"
    
    pc.Refresh
    
    pt.TableRange2.Copy
    ConfigWorksheet.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ConfigWorksheet.ListObjects.Add(xlSrcRange, Range("A1", Range("A1").End(xlToRight).End(xlDown)), , xlYes).Name = "ConfigTable"
    
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns.Add(3).Name = "Limit"
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns("Limit").DataBodyRange.Value = 10
    
    
    Dim TotalRow As Range
    ConfigWorksheet.ListObjects("ConfigTable").Range.Find(What:="Grand Total").EntireRow.Delete
    
    Range("C2").Select
    
End Sub
