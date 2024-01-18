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
    ActiveSheet.name = "Lottery Results"
    
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
    RegistrationTable.name = "LotteryResults"
    
   
    Call GenConfig
    
    RegistrationTable.ListColumns.Add(4).name = "Applicants"
    RegistrationTable.ListColumns("Applicants").DataBodyRange.NumberFormat = "General"
    RegistrationTable.ListColumns("Applicants").DataBodyRange.Formula = "=VLOOKUP([@Event],ConfigTable[#All],2,FALSE)"
    
    RegistrationTable.ListColumns.Add(5).name = "Camp Limit"
    RegistrationTable.ListColumns("Camp Limit").DataBodyRange.NumberFormat = "General"
    RegistrationTable.ListColumns("Camp Limit").DataBodyRange.Formula = "=VLOOKUP([@Event],ConfigTable[#All],3,FALSE)"
    
    RegistrationTable.ListColumns.Add(6).name = "Random Draw"
    
    RegistrationTable.ListColumns.Add(7).name = "Lottery Selection Status"
    
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
        If Sheet.name = ConfigSheetName Then
            MsgBox ("The " & ConfigSheetName & "worksheet already exists.  Delete it, and re-run the config generator if you want to regen the config table.")
            ConfigSheetExists = True
        End If
    Next Sheet
                
    If Not ConfigSheetExists Then
        Set ConfigWorksheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ConfigWorksheet.name = ConfigSheetName
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
    ConfigWorksheet.ListObjects.Add(xlSrcRange, Range("A1", Range("A1").End(xlToRight).End(xlDown)), , xlYes).name = "ConfigTable"
    
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns.Add(3).name = "Limit"
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns("Limit").DataBodyRange.Value = 15
    
    ConfigWorksheet.ListObjects("ConfigTable").ListColumns.Add(4).name = "Filled Spots"
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
    RandomSheet.name = "Random Draw"
    
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

Public Sub RemoveDuplicates()
    Set RegistrationWorksheet = ThisWorkbook.Sheets("Lottery Results")
    Set RegistrationTable = RegistrationWorksheet.ListObjects("LotteryResults")

    With RegistrationTable.Sort
        .SortFields.Clear
        .SortFields.Add Key:=RegistrationTable.ListColumns("Start Date").Range, Order:=xlAscending
        .SortFields.Add Key:=RegistrationTable.ListColumns("Applicants").Range, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    RegistrationTable.Range.RemoveDuplicates Columns:=Array(1, 15), Header:=xlYes

End Sub

Public Sub RunLottery()
    Set RegistrationWorksheet = ThisWorkbook.Sheets("Lottery Results")
    Set RegistrationTable = RegistrationWorksheet.ListObjects("LotteryResults")
    Set ConfigWorksheet = ThisWorkbook.Sheets("Camp Config")
    Set ConfigTable = ConfigWorksheet.ListObjects("ConfigTable")

    ' Call GenRandomPermutation(RegistrationTable)
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
    Set SelectionStatusColumn = RegistrationTable.ListColumns("Lottery Selection Status")
    
    Dim ApplicantNameColumn As ListColumn
    Set ApplicantNameColumn = RegistrationTable.ListColumns("Camper Name")
    
    Dim SiblingNameColumn As ListColumn
    Set SiblingNameColumn = RegistrationTable.ListColumns("Please enter the full name of the friend or sibling.")
    
    Dim CampStartDateColumn As ListColumn
    Set CampStartDateColumn = RegistrationTable.ListColumns("Start Date")
    
    Dim PreRegisteredColumn As ListColumn
    Set PreRegisteredColumn = RegistrationTable.ListColumns("Registered")
    
    Dim Application As ListRow
    Dim SiblingApplication As ListRow
    
    Dim ApplicationAccepted As Boolean
        
    SelectionStatusColumn.DataBodyRange.ClearContents
    ConfigTable.ListColumns("Filled Spots").DataBodyRange.ClearContents
    
    ' Automatically accept applications for anyone who is pre-registered
    For Each Application In RegistrationTable.ListRows
        Dim PreRegistrationStatus As Range
         Set PreRegistrationStatus = Intersect(Application.Range, PreRegisteredColumn.Range)
        If PreRegistrationStatus = 1 Then
            ApplicationAccepted = AcceptApplication(Application, ConfigTable, RegistrationTable, "Picked via Pre-registration")
        End If
    Next Application
    
    
    For Each Application In RegistrationTable.ListRows
        Dim CampName As String
        CampName = Application.Range(1).Value2
        
        Dim SelectedStatus As Range
        Dim SiblingName As Range
        Dim CampStartDate As Range
        
        Dim CampDataRow As Range
        Set CampDataRow = ConfigTable.ListColumns("Row Labels").DataBodyRange.Find(CampName).EntireRow
        
        Dim FilledSpotsColumn As Range
        Set FilledSpotsColumn = ConfigTable.ListColumns("Filled Spots").DataBodyRange
        Dim FilledSpots As Range
        Set FilledSpots = Intersect(CampDataRow, FilledSpotsColumn)
                
        Dim LimitsColumn As Range
        Set LimitsColumn = ConfigTable.ListColumns("Limit").DataBodyRange
        Dim Limit As Range
        Set Limit = Intersect(CampDataRow, LimitsColumn)
        
        If FilledSpots.Value2 < Limit.Value2 Then
            ApplicationAccepted = AcceptApplication(Application, ConfigTable, RegistrationTable, "Picked via Lottery")
            
            Set SiblingName = Intersect(Application.Range, SiblingNameColumn.Range)
            Set CampStartDate = Intersect(Application.Range, CampStartDateColumn.Range)
            
            If SiblingName.Value2 <> vbNullString Then
            
                Set SiblingApplication = GetSiblingApplication(RegistrationTable, SiblingName.Value2, CampStartDate.Text)
                If Not SiblingApplication Is Nothing Then
                    ApplicationAccepted = AcceptApplication(SiblingApplication, ConfigTable, RegistrationTable, "Picked via Sibling")
                End If
            End If
        Else
            Set SelectedStatus = Intersect(Application.Range, SelectionStatusColumn.Range)

            If SelectedStatus.Value2 = vbNullString Then
                SelectedStatus.Value2 = "Not Picked"
            End If
        End If
    Next Application
           
End Sub

Public Function AcceptApplication(Application As ListRow, ConfigTable As ListObject, RegistrationTable As ListObject, AcceptReason As String) As Boolean
    
    Dim CampName As String
    CampName = Application.Range(1).Value2
    
    Dim CampDataRow As Range
    Set CampDataRow = ConfigTable.ListColumns("Row Labels").DataBodyRange.Find(CampName).EntireRow
    
    Dim FilledSpotsColumn As Range
    Set FilledSpotsColumn = ConfigTable.ListColumns("Filled Spots").DataBodyRange
    Dim FilledSpots As Range
    Set FilledSpots = Intersect(CampDataRow, FilledSpotsColumn)
    
    Dim SelectionStatusColumn As ListColumn
    Set SelectionStatusColumn = RegistrationTable.ListColumns("Lottery Selection Status")
    Dim SelectedStatus As Range
    Set SelectedStatus = Intersect(Application.Range, SelectionStatusColumn.Range)
    
    If SelectedStatus.Value2 = vbNullString Then
        SelectedStatus.Value2 = AcceptReason
        FilledSpots.Value2 = FilledSpots.Value2 + 1
    End If

End Function


Public Sub Test()
    Set RegistrationWorksheet = ThisWorkbook.Sheets("Lottery Results")
    Set RegistrationTable = RegistrationWorksheet.ListObjects("LotteryResults")

    Dim row As ListRow
    Set row = GetSiblingApplication(RegistrationTable, "Luna Wahle", "8/28/2023  8:30:00 AM")
End Sub

Public Function GetSiblingApplication(tbl As ListObject, ApplicantNameCriteria As String, CampStartDateCriteria As String) As ListRow
    Dim Application As ListRow
    
    Dim ApplicantNameColumn As Range
    Dim CampStartDateColumn As Range

    Dim ApplicantName As Range
    Dim CampStartDate As Range

    Set ApplicantNameColumn = tbl.ListColumns("Camper Name").DataBodyRange
    Set CampStartDateColumn = tbl.ListColumns("Start Date").DataBodyRange

    For Each Application In tbl.ListRows
        Set ApplicantName = Intersect(Application.Range, ApplicantNameColumn)
        Set CampStartDate = Intersect(Application.Range, CampStartDateColumn)
        
        If InStr(1, ApplicantNameCriteria, ApplicantName.Value2) > 0 Then
            If CampStartDate.Text = CampStartDateCriteria Then
                Set GetSiblingApplication = Application
            End If
        End If
    Next Application
End Function
