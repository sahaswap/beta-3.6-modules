Attribute VB_Name = "modLayoutDump"
'==================================================================
' modLayoutDump - run DumpSheet1Layout once; it writes Sheet1's layout
' to "Sheet1 layout.txt" on the Desktop and opens it. Only positions and
' formatting settings - no cell values, so no case data. Delete this
' module afterwards.
'==================================================================
Option Explicit

Sub DumpSheet1Layout()
    Dim ws As Worksheet, shp As Shape, s As String, c As Long, r As Long, path As String, fh As Integer
    Dim txt As String, tl As String, br As String, theme As String

    Set ws = ThisWorkbook.Sheets("Sheet1")
    On Error Resume Next
    theme = CStr(Application.Run("ActiveTheme"))
    ws.Activate
    s = "Theme: " & theme & "   Zoom: " & ActiveWindow.Zoom & "   Excel " & Application.Version & vbCrLf & vbCrLf

    s = s & "SHAPES: name | text | type/autoshape | left, top, width, height | rot | flipH/flipV | " & _
        "lockAspect | placement | autosize/wrap | topleft-bottomright | macro" & vbCrLf
    For Each shp In ws.Shapes
        txt = "": tl = "": br = ""
        txt = Replace(Left$(shp.TextFrame2.TextRange.Text, 40), vbCr, " ")
        tl = shp.TopLeftCell.Address(0, 0)
        br = shp.BottomRightCell.Address(0, 0)
        s = s & shp.Name & " | " & txt & " | " & shp.Type & "/" & shp.AutoShapeType & " | " & _
            Format$(shp.Left, "0.0") & ", " & Format$(shp.Top, "0.0") & ", " & _
            Format$(shp.Width, "0.0") & ", " & Format$(shp.Height, "0.0") & " | " & shp.Rotation & " | " & _
            shp.HorizontalFlip & "/" & shp.VerticalFlip & " | " & shp.LockAspectRatio & " | " & _
            shp.Placement & " | " & shp.TextFrame2.AutoSize & "/" & shp.TextFrame2.WordWrap & " | " & _
            tl & "-" & br & " | " & shp.OnAction & vbCrLf
    Next shp

    s = s & vbCrLf & "COLUMNS: col | left | width (points) | char width | hidden" & vbCrLf
    For c = 1 To 29
        With ws.Columns(c)
            s = s & Split(ws.Cells(1, c).Address(1, 0), "$")(0) & " | " & Format$(.Left, "0.0") & " | " & _
                Format$(.Width, "0.0") & " | " & .ColumnWidth & " | " & .Hidden & vbCrLf
        End With
    Next c

    s = s & vbCrLf & "ROWS: row | top | height | hidden" & vbCrLf
    For r = 1 To 30
        With ws.Rows(r)
            s = s & r & " | " & Format$(.Top, "0.0") & " | " & Format$(.Height, "0.0") & " | " & .Hidden & vbCrLf
        End With
    Next r

    s = s & vbCrLf & "MERGED AREAS (rows 1-30, columns A-AC):" & vbCrLf
    Dim cell As Range, seen As String
    For Each cell In ws.Range("A1:AC30")
        If cell.MergeCells Then
            If InStr(seen, "|" & cell.MergeArea.Address(0, 0) & "|") = 0 Then
                seen = seen & "|" & cell.MergeArea.Address(0, 0) & "|"
                s = s & cell.MergeArea.Address(0, 0) & "  "
            End If
        End If
    Next cell
    s = s & vbCrLf

    path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Sheet1 layout.txt"
    fh = FreeFile
    Open path For Output As #fh
    Print #fh, s
    Close #fh
    Shell "notepad.exe """ & path & """", vbNormalFocus
End Sub
