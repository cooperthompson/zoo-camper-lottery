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
    
    LastRow = EventCell.CurrentRegion.Rows(EventCell.CurrentRegion.Rows.Count).row
    LastColumn = EventCell.CurrentRegion.Columns(EventCell.CurrentRegion.Columns.Count).Column
    Set LastCell = Cells(LastRow, LastColumn)
    
    Set RegistrationTable = RegistrationWorksheet.ListObjects.Add(xlSrcRange, Range(EventCell, LastCell), , xlYes)
    RegistrationTable.Name = "LotteryResults"
    
   
    Call GenConfig
    
    RegistrationTable.ListColumns.Add(4).Name = "Applicants"
    RegistrationTable.ListColumns("Applicants").DataBodyRange.NumberFormat = "General"
    RegistrationTable.ListColumns("Applicants").DataBodyRange.Formula = "=VLOOKUP([@Event],ConfigTable[#All],2,FALSE)"
    
    RegistrationTable.ListColumns.Add(5).Name = "Camp Limit"
    RegistrationTable.ListColumns("Camp Limit").DataBodyRange.NumberFormat = "General"
    RegistrationTable.ListColumns("Camp Limit").DataBodyRange.Formula = "=VLOOKUP([@Event],ConfigTable[#All],3,FALSE)"
    
    RegistrationTable.ListColumns.Add(6).Name = "Lottery Selection Status"
    
    Call FixColumnWidths(RegistrationTable)
    
End Sub

Public Sub FixColumnWidths(tbl As ListObject)
    tbl.Range.ColumnWidth = 200
    
    Dim col As ListColumn
    Dim row As ListRow
    
    For Each col In tbl.ListColumns
        col.Range.EntireColumn.AutoFit
    Next col
    
    For Each row In tbl.ListRows
        row.Range.EntireRow.AutoFit
    Next row
    
    

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
    pt.PivotFields("Event").AutoSort xlAscending, "Count of Registrations"
    
    pc.Refresh
    
    pt.TableRange2.Copy
    ConfigWorksheet.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    ConfigWorksheet.ListObjects.Add(xlSrcRange, Range("A1", Range("A1").End(xlToRight).End(xlDown)), , xlYes).Name = "ConfigTable"
    
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns.Add(3).Name = "Limit"
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns("Limit").DataBodyRange.Value = 10
    
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns.Add(4).Name = "Filled Spots"
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns("Filled Spots").DataBodyRange.Value = 0
        
    Dim TotalRow As Range
    ConfigWorksheet.ListObjects("ConfigTable").Range.Find(What:="Grand Total").EntireRow.Delete
    
    Range("C2").Select
   
End Sub

Public Sub GenRandomPermutation(tbl As ListObject)
    On Error Resume Next
    ThisWorkbook.Sheets("Random Draw").Delete
    On Error GoTo 0

    Dim RandomSheet As Worksheet
    Set RandomSheet = ThisWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    RandomSheet.Name = "Random Draw"
    
    Dim RandomTable As ListObject
    Dim TableHeader As ListObject
    Dim i As Long
    Dim row As ListRow
    Set RandomTable = RandomSheet.ListObjects.Add(xlSrcRange, Range("A1"), , xlYes)
    RandomTable.HeaderRowRange.Value2 = "Random Draw"
    For i = 1 To tbl.ListRows.Count
        Set row = RandomTable.ListRows.Add
        row.Range(1, 1) = i + 10000
    Next i
    
    RandomTable.DataBodyRange.Select
    Call Random
    
    tbl.ListColumns.Add(6).Name = "Random Draw"
    
    RandomTable.Range.Copy
    tbl.ListColumns("Random Draw").Range.PasteSpecial Paste:=xlPasteValues
End Sub

Sub Random()
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim r As Long
    
    For x = 1 To Selection.Rows.Count
       Randomize Timer
       r = Int(Rnd(1) * (Selection.Rows.Count) + 1)
       For z = 1 To Selection.Columns.Count
    
           y = Selection.Cells(x, z).Formula
           Selection.Cells(x, z).Formula = Selection.Cells(r, z).Formula
           Selection.Cells(r, z).Formula = y
       Next z
    Next x
End Sub

Public Sub RunLottery()
    Set RegistrationWorksheet = ThisWorkbook.Sheets("Lottery Results")
    Set RegistrationTable = RegistrationWorksheet.ListObjects("LotteryResults")
    Set ConfigWorksheet = ThisWorkbook.Sheets("Camp Config")
    Set ConfigTable = ConfigWorksheet.ListObjects("ConfigTable")

    Call GenRandomPermutation(RegistrationTable)
    Call FixColumnWidths(RegistrationTable)

    With RegistrationTable.Sort
        .SortFields.Clear
        .SortFields.Add Key:=RegistrationTable.ListColumns("Start Date").Range, Order:=xlAscending
        .SortFields.Add Key:=RegistrationTable.ListColumns("Applicants").Range, Order:=xlAscending
        .SortFields.Add Key:=RegistrationTable.ListColumns("Random Draw").Range, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    Dim SelectionStatusColumn As ListColumn
    Dim ApplicantNameColumn As ListColumn
    Dim SiblingNameColumn As ListColumn
    
    Set SelectionStatusColumn = RegistrationTable.ListColumns("Lottery Selection Status")
    Set ApplicantNameColumn = RegistrationTable.ListColumns("Camper Name")
    Set SiblingNameColumn = RegistrationTable.ListColumns("Please enter the full name of the friend or sibling.")
    
    Dim Application As ListRow
    Dim SiblingApplication As ListRow
    Dim SiblingApplicationRow As Range
    
    For Each Application In RegistrationTable.ListRows
        Dim CampName As String
        CampName = Application.Range(1).Value2
        
        Dim LimitsColumn As Range
        Dim FilledSpotsColumn As Range
        Dim Limit As Range
        Dim FilledSpots As Range
        
        Dim SelectedStatus As Range
        Dim SiblingName As Range
                
        Set LimitsColumn = ConfigTable.ListColumns("Limit").DataBodyRange
        Set FilledSpotsColumn = ConfigTable.ListColumns("Filled Spots").DataBodyRange
        
        Dim CampDataRow As Range
        Set CampDataRow = ConfigTable.ListColumns("Row Labels").DataBodyRange.Find(CampName).EntireRow
        
        Set Limit = Intersect(CampDataRow, LimitsColumn)
        Set FilledSpots = Intersect(CampDataRow, FilledSpotsColumn)
        
        If FilledSpots.Value2 < Limit.Value2 Then
            Set SelectedStatus = Intersect(Application.Range, SelectionStatusColumn.Range)
            Set SiblingName = Intersect(Application.Range, SiblingNameColumn.Range)
            
            SelectedStatus.Value2 = "Picked via Lottery"
            FilledSpots.Value2 = FilledSpots.Value2 + 1
            
            If SiblingName.Value2 <> vbNullString Then
                
                ' Set SiblingApplicationRow = ApplicantNameColumn.DataBodyRange.Find(SiblingName.Value2).EntireRow
                SiblingApplicationRow = WorksheetFunction.Match(SiblingName.Value2, ApplicantNameColumn.Range, 0)
                SiblingApplication = RegistrationTable.DataBodyRange.Rows(SiblingApplicationRow)
                                
                Set SelectedStatus = Intersect(SiblingApplication, SelectionStatusColumn)
                If SelectedStatus.Value2 = vbNullString Then
                    SelectedStatus.Value2 = "Picked via Sibling"
                    FilledSpots.Value2 = FilledSpots.Value2 + 1
                End If
            End If
            
        Else
            Application.Range(7).Value2 = "Not Picked"
        End If
        
        
       

    Next Application
           
End Sub
