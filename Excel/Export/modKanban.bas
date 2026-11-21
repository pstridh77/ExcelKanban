Attribute VB_Name = "modKanban"
Option Explicit

Public Sub RefreshKanban()

    Call RensaKort
    Call RitaKort

End Sub

Public Sub RensaKort()

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = Worksheets("Kanban")

    For Each shp In ws.Shapes

        If Left(shp.Name, 5) = "Task_" Then
            shp.Delete
        End If

    Next shp

End Sub

Public Sub RitaKort()

    Dim wsData As Worksheet
    Dim wsKanban As Worksheet
    Dim tbl As ListObject

    Dim i As Long

    Set wsData = Worksheets("Uppgifter")
    Set wsKanban = Worksheets("Kanban")

    Set tbl = wsData.ListObjects("tblUppgifter")

    For i = 1 To tbl.ListRows.Count

        Call RitaKortRad(tbl, i)

    Next i

End Sub

Private Function StatusKolumn(Status As String) As Long

    Select Case Status

        Case "Backlog"
            StatusKolumn = 2

        Case "Att Göra"
            StatusKolumn = 5

        Case "Pågår"
            StatusKolumn = 8

        Case "Klar"
            StatusKolumn = 11

        Case "Blockerad"
            StatusKolumn = 14

    End Select

End Function

Private Sub RitaKortRad(tbl As ListObject, Rad As Long)

    Dim ws As Worksheet
    Dim shp As Shape

    Dim Status As String
    Dim Uppgift As String
    Dim Ansvarig As String
    Dim ID As Long

    Static PosBacklog As Long
    Static PosTodo As Long
    Static PosDoing As Long
    Static PosDone As Long
    Static PosBlocked As Long

    Set ws = Worksheets("Kanban")

    ID = tbl.DataBodyRange(Rad, 1).Value
    Uppgift = tbl.DataBodyRange(Rad, 2).Value
    Ansvarig = tbl.DataBodyRange(Rad, 3).Value
    Status = tbl.DataBodyRange(Rad, 4).Value

    Dim KortTop As Double
    Dim KortLeft As Double

    Select Case Status

        Case "Backlog"

            PosBacklog = PosBacklog + 1

            KortLeft = ws.Cells(3, 2).Left
            KortTop = 50 + PosBacklog * 80

        Case "Att Göra"

            PosTodo = PosTodo + 1

            KortLeft = ws.Cells(3, 5).Left
            KortTop = 50 + PosTodo * 80

        Case "Pågår"

            PosDoing = PosDoing + 1

            KortLeft = ws.Cells(3, 8).Left
            KortTop = 50 + PosDoing * 80

        Case "Klar"

            PosDone = PosDone + 1

            KortLeft = ws.Cells(3, 11).Left
            KortTop = 50 + PosDone * 80

        Case "Blockerad"

            PosBlocked = PosBlocked + 1

            KortLeft = ws.Cells(3, 14).Left
            KortTop = 50 + PosBlocked * 80

    End Select

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        KortLeft, _
        KortTop, _
        150, _
        60)

    shp.Name = "Task_" & ID
    
    shp.OnAction = "KortKlick"

    shp.TextFrame2.TextRange.Text = _
        "#" & ID & vbCrLf & _
        Uppgift & vbCrLf & _
        Ansvarig

End Sub

Public Sub KortKlick()

    Dim ShapeName As String
    Dim ID As Long

    ShapeName = Application.Caller

    ID = CLng(Replace(ShapeName, "Task_", ""))

    MsgBox "Du klickade på uppgift " & ID

End Sub



