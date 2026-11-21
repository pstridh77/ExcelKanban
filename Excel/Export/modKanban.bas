Attribute VB_Name = "modKanban"
Option Explicit

Private Const FORM_SHEET_NAME As String = "Formular"

Public Sub RefreshKanban()

    Call RensaKort
    Call RitaKort

End Sub

Public Sub RensaKort()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = Worksheets("Kanban")

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "Task_" Then
            ws.Shapes(i).Delete
        End If
    Next i

End Sub

Public Sub RitaKort()

    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim PosBacklog As Long
    Dim PosTodo As Long
    Dim PosDoing As Long
    Dim PosDone As Long
    Dim PosBlocked As Long

    Set wsData = Worksheets("Uppgifter")
    Set tbl = wsData.ListObjects("tblUppgifter")

    For i = 1 To tbl.ListRows.Count
        Call RitaKortRad(tbl, i, PosBacklog, PosTodo, PosDoing, PosDone, PosBlocked)
    Next i

End Sub

Private Function StatusKolumn(ByVal Status As String) As Long

    Select Case Status
        Case "Backlog"
            StatusKolumn = 2
        Case "Att g" & ChrW$(&HF6) & "ra"
            StatusKolumn = 5
        Case "P" & ChrW$(&HE5) & "g" & ChrW$(&HE5) & "r"
            StatusKolumn = 8
        Case "Klar"
            StatusKolumn = 11
        Case "Blockerad"
            StatusKolumn = 14
    End Select

End Function

Private Sub RitaKortRad(ByVal tbl As ListObject, ByVal Rad As Long, _
                        ByRef PosBacklog As Long, ByRef PosTodo As Long, _
                        ByRef PosDoing As Long, ByRef PosDone As Long, _
                        ByRef PosBlocked As Long)

    Dim ws As Worksheet
    Dim shp As Shape
    Dim Status As String
    Dim Uppgift As String
    Dim Ansvarig As String
    Dim ID As Long
    Dim Kolumn As Long
    Dim KortTop As Double
    Dim KortLeft As Double

    Set ws = Worksheets("Kanban")

    ID = CLng(tbl.DataBodyRange(Rad, 1).Value)
    Uppgift = CStr(tbl.DataBodyRange(Rad, 2).Value)
    Ansvarig = CStr(tbl.DataBodyRange(Rad, 3).Value)
    Status = CStr(tbl.DataBodyRange(Rad, 4).Value)
    Kolumn = StatusKolumn(Status)

    If Kolumn = 0 Then Exit Sub

    KortLeft = ws.Cells(3, Kolumn).Left

    Select Case Kolumn
        Case 2
            PosBacklog = PosBacklog + 1
            KortTop = 50 + PosBacklog * 80
        Case 5
            PosTodo = PosTodo + 1
            KortTop = 50 + PosTodo * 80
        Case 8
            PosDoing = PosDoing + 1
            KortTop = 50 + PosDoing * 80
        Case 11
            PosDone = PosDone + 1
            KortTop = 50 + PosDone * 80
        Case 14
            PosBlocked = PosBlocked + 1
            KortTop = 50 + PosBlocked * 80
    End Select

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, KortLeft, KortTop, 150, 60)
    shp.Name = "Task_" & ID
    shp.OnAction = "KortKlick"
    shp.TextFrame2.TextRange.Text = _
        "#" & ID & vbCrLf & _
        Uppgift & vbCrLf & _
        Ansvarig

End Sub

Public Sub NyUppgift()

    VisaUppgift 0

End Sub

Public Sub KortKlick()

    Dim ShapeName As String
    Dim TaskIDText As String

    ShapeName = CStr(Application.Caller)

    If Left$(ShapeName, 5) <> "Task_" Then
        MsgBox "Vald figur ar inte ett uppgiftskort.", vbExclamation
        Exit Sub
    End If

    TaskIDText = Mid$(ShapeName, 6)
    If Not IsNumeric(TaskIDText) Then
        MsgBox "Uppgiftskortet har ett ogiltigt ID.", vbExclamation
        Exit Sub
    End If

    VisaUppgift CLng(TaskIDText)

End Sub

Private Sub VisaUppgift(ByVal ID As Long)

    Dim wsForm As Worksheet
    Dim tbl As ListObject
    Dim taskRow As ListRow
    Dim statuses As Variant
    Dim priorities As Variant

    Set wsForm = GetTaskFormSheet()
    Set tbl = Worksheets("Uppgifter").ListObjects("tblUppgifter")

    If ID > 0 Then
        Set taskRow = FindTaskRow(tbl, ID)
        If taskRow Is Nothing Then
            MsgBox "Uppgift " & ID & " hittades inte i uppgiftstabellen.", vbExclamation
            Exit Sub
        End If
    End If

    If ID = 0 Then
        wsForm.Range("B4").Value = "Ny"
        wsForm.Range("Z1").Value = 0
        wsForm.Range("B5").Value = vbNullString
        wsForm.Range("B6").Value = vbNullString
        statuses = StatusOptions()
        priorities = PriorityOptions()
        wsForm.Range("B7").Value = statuses(0)
        wsForm.Range("B8").Value = priorities(1)
        wsForm.Range("B9").Value = vbNullString
        wsForm.Range("B10").Value = vbNullString
        wsForm.Range("B11").Value = vbNullString
        wsForm.Shapes("btnTaskDelete").Visible = False
    Else
        wsForm.Range("Z1").Value = ID
        wsForm.Range("B4").Value = ID
        wsForm.Range("B5").Value = taskRow.Range.Cells(1, tbl.ListColumns("Uppgift").Index).Value
        wsForm.Range("B6").Value = taskRow.Range.Cells(1, tbl.ListColumns("Ansvarig").Index).Value
        wsForm.Range("B7").Value = taskRow.Range.Cells(1, tbl.ListColumns("Status").Index).Value
        wsForm.Range("B8").Value = taskRow.Range.Cells(1, tbl.ListColumns("Prioritet").Index).Value
        If Len(Trim$(CStr(wsForm.Range("B8").Value))) = 0 Then
            priorities = PriorityOptions()
            wsForm.Range("B8").Value = priorities(1)
        End If
        wsForm.Range("B9").Value = taskRow.Range.Cells(1, tbl.ListColumns("Startdatum").Index).Value
        wsForm.Range("B10").Value = taskRow.Range.Cells(1, tbl.ListColumns("Deadline").Index).Value
        wsForm.Range("B11").Value = taskRow.Range.Cells(1, tbl.ListColumns("Kommentar").Index).Value
        wsForm.Shapes("btnTaskDelete").Visible = True
    End If

    wsForm.Visible = xlSheetVisible
    wsForm.Activate
    wsForm.Range("B5").Select

End Sub

Private Function GetTaskFormSheet() As Worksheet

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = FORM_SHEET_NAME Then
            Set GetTaskFormSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = FORM_SHEET_NAME
    InitializeTaskForm ws
    Worksheets("Kanban").Activate
    ws.Visible = xlSheetHidden
    Set GetTaskFormSheet = ws

End Function

Private Sub InitializeTaskForm(ByVal ws As Worksheet)

    Dim saveButton As Shape
    Dim cancelButton As Shape
    Dim deleteButton As Shape
    Dim statuses As Variant
    Dim priorities As Variant
    Dim i As Long

    ws.Cells.Clear
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 5
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 5
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("Z").Hidden = True

    ws.Range("A1:D1").Merge
    ws.Range("A1").Value = "Uppgiftsformul" & ChrW$(&HE4) & "r"
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = "Fyll i uppgiften och klicka p" & ChrW$(&HE5) & " Spara."
    ws.Range("B4:D4").Merge

    For i = 5 To 11
        ws.Range("B" & i & ":D" & i).Merge
    Next i

    ws.Range("A4:A11").Value = Application.Transpose(Array( _
        "ID", "Uppgift", "Ansvarig", "Status", "Prioritet", _
        "Startdatum", "Deadline", "Kommentar"))

    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Font.Size = 18
    ws.Range("A1:D1").Interior.Color = RGB(21, 96, 130)
    ws.Range("A1:D1").Font.Color = RGB(255, 255, 255)
    ws.Range("A2:D2").Font.Italic = True
    ws.Range("A4:A11").Font.Bold = True
    ws.Range("B4:D11").Interior.Color = RGB(242, 246, 248)
    ws.Range("A4:D11").VerticalAlignment = xlVAlignCenter
    ws.Rows("4:10").RowHeight = 24
    ws.Rows("11:11").RowHeight = 72
    ws.Range("B11").WrapText = True
    ws.Range("B9").NumberFormat = "yyyy-mm-dd"
    ws.Range("B10").NumberFormat = "yyyy-mm-dd"

    statuses = StatusOptions()
    priorities = PriorityOptions()
    With ws.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=Join(statuses, Application.International(xlListSeparator))
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With ws.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=Join(priorities, Application.International(xlListSeparator))
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    Set saveButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("B13").Left, ws.Range("B13").Top, 90, 30)
    saveButton.Name = "btnTaskSave"
    saveButton.TextFrame2.TextRange.Text = "Spara"
    saveButton.OnAction = "modKanban.SparaUppgift"
    saveButton.Fill.ForeColor.RGB = RGB(21, 96, 130)
    saveButton.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

    Set cancelButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("D13").Left, ws.Range("D13").Top, 90, 30)
    cancelButton.Name = "btnTaskCancel"
    cancelButton.TextFrame2.TextRange.Text = "Avbryt"
    cancelButton.OnAction = "modKanban.AvslutaUppgift"

    Set deleteButton = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("F13").Left, ws.Range("F13").Top, 90, 30)
    deleteButton.Name = "btnTaskDelete"
    deleteButton.TextFrame2.TextRange.Text = "Ta bort"
    deleteButton.OnAction = "modKanban.TaBortUppgift"

End Sub

Private Function StatusOptions() As Variant

    StatusOptions = Array( _
        "Backlog", _
        "Att g" & ChrW$(&HF6) & "ra", _
        "P" & ChrW$(&HE5) & "g" & ChrW$(&HE5) & "r", _
        "Blockerad", _
        "Klar")

End Function

Private Function PriorityOptions() As Variant

    PriorityOptions = Array( _
        "L" & ChrW$(&HE5) & "g", _
        "Medel", _
        "H" & ChrW$(&HF6) & "g", _
        "Kritisk")

End Function

Private Function IsAllowedValue(ByVal Value As String, ByVal Options As Variant) As Boolean

    Dim item As Variant

    For Each item In Options
        If StrComp(Value, CStr(item), vbTextCompare) = 0 Then
            IsAllowedValue = True
            Exit Function
        End If
    Next item

End Function

Private Function FindTaskRow(ByVal tbl As ListObject, ByVal ID As Long) As ListRow

    Dim item As ListRow
    Dim cellValue As Variant
    Dim idColumn As Long

    idColumn = tbl.ListColumns("ID").Index

    For Each item In tbl.ListRows
        cellValue = item.Range.Cells(1, idColumn).Value
        If IsNumeric(cellValue) Then
            If CLng(cellValue) = ID Then
                Set FindTaskRow = item
                Exit Function
            End If
        End If
    Next item

End Function

Private Function NextTaskID(ByVal tbl As ListObject) As Long

    Dim item As ListRow
    Dim cellValue As Variant
    Dim currentID As Long
    Dim maxID As Long
    Dim idColumn As Long

    idColumn = tbl.ListColumns("ID").Index

    For Each item In tbl.ListRows
        cellValue = item.Range.Cells(1, idColumn).Value
        If IsNumeric(cellValue) Then
            currentID = CLng(cellValue)
            If currentID > maxID Then maxID = currentID
        End If
    Next item

    NextTaskID = maxID + 1

End Function

Public Sub SparaUppgift()

    Dim wsForm As Worksheet
    Dim tbl As ListObject
    Dim taskRow As ListRow
    Dim taskID As Long
    Dim taskTitle As String
    Dim taskStatus As String
    Dim taskPriority As String
    Dim startValue As Variant
    Dim deadlineValue As Variant

    Set wsForm = GetTaskFormSheet()
    Set tbl = Worksheets("Uppgifter").ListObjects("tblUppgifter")

    taskTitle = Trim$(CStr(wsForm.Range("B5").Value))
    taskStatus = Trim$(CStr(wsForm.Range("B7").Value))
    taskPriority = Trim$(CStr(wsForm.Range("B8").Value))
    startValue = wsForm.Range("B9").Value
    deadlineValue = wsForm.Range("B10").Value

    If Len(taskTitle) = 0 Then
        MsgBox "Ange en titel f" & ChrW$(&HF6) & "r uppgiften.", vbExclamation
        Exit Sub
    End If
    If Not IsAllowedValue(taskStatus, StatusOptions()) Then
        MsgBox "Ange en giltig status.", vbExclamation
        Exit Sub
    End If
    If Not IsAllowedValue(taskPriority, PriorityOptions()) Then
        MsgBox "Ange en giltig prioritet.", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(CStr(startValue))) > 0 And Not IsDate(startValue) Then
        MsgBox "Ange ett giltigt startdatum.", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(CStr(deadlineValue))) > 0 And Not IsDate(deadlineValue) Then
        MsgBox "Ange ett giltigt slutdatum.", vbExclamation
        Exit Sub
    End If
    If IsDate(startValue) And IsDate(deadlineValue) Then
        If CDate(deadlineValue) < CDate(startValue) Then
            MsgBox "Slutdatum kan inte vara f" & ChrW$(&HF6) & "re startdatum.", vbExclamation
            Exit Sub
        End If
    End If

    taskID = CLng(wsForm.Range("Z1").Value)
    If taskID > 0 Then
        Set taskRow = FindTaskRow(tbl, taskID)
        If taskRow Is Nothing Then
            MsgBox "Uppgift " & taskID & " hittades inte i uppgiftstabellen.", vbExclamation
            Exit Sub
        End If
    Else
        taskID = NextTaskID(tbl)
        Set taskRow = tbl.ListRows.Add
    End If

    With taskRow.Range
        .Cells(1, tbl.ListColumns("ID").Index).Value = taskID
        .Cells(1, tbl.ListColumns("Uppgift").Index).Value = taskTitle
        .Cells(1, tbl.ListColumns("Ansvarig").Index).Value = Trim$(CStr(wsForm.Range("B6").Value))
        .Cells(1, tbl.ListColumns("Status").Index).Value = taskStatus
        .Cells(1, tbl.ListColumns("Prioritet").Index).Value = taskPriority
        If IsDate(startValue) Then
            .Cells(1, tbl.ListColumns("Startdatum").Index).Value = CDate(startValue)
        Else
            .Cells(1, tbl.ListColumns("Startdatum").Index).ClearContents
        End If
        If IsDate(deadlineValue) Then
            .Cells(1, tbl.ListColumns("Deadline").Index).Value = CDate(deadlineValue)
        Else
            .Cells(1, tbl.ListColumns("Deadline").Index).ClearContents
        End If
        .Cells(1, tbl.ListColumns("Kommentar").Index).Value = CStr(wsForm.Range("B11").Value)
    End With

    RefreshKanban
    AvslutaUppgift

End Sub

Public Sub TaBortUppgift()

    Dim wsForm As Worksheet
    Dim tbl As ListObject
    Dim taskRow As ListRow
    Dim taskID As Long

    Set wsForm = GetTaskFormSheet()
    taskID = CLng(wsForm.Range("Z1").Value)

    If taskID <= 0 Then
        MsgBox "Spara uppgiften innan du tar bort den.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Ta bort uppgift " & taskID & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    Set tbl = Worksheets("Uppgifter").ListObjects("tblUppgifter")
    Set taskRow = FindTaskRow(tbl, taskID)
    If taskRow Is Nothing Then
        MsgBox "Uppgift " & taskID & " hittades inte i uppgiftstabellen.", vbExclamation
        Exit Sub
    End If

    taskRow.Delete
    RefreshKanban
    AvslutaUppgift

End Sub

Public Sub AvslutaUppgift()

    Dim wsForm As Worksheet

    Set wsForm = GetTaskFormSheet()
    Worksheets("Kanban").Activate
    wsForm.Visible = xlSheetHidden

End Sub
