Attribute VB_Name = "modSheet1Setup"
'==================================================================
' modSheet1Setup - ONE-TIME setup. Run SetupSheet1 once, then delete
' this module.
'
' Lays Sheet1 out for 5 counterparties and the OSDD inputs:
'   - the OSDD Search Inputs panel in V3:AB10: a column per entity
'     (customer, counterparties 1-5), a row per input (Negative News
'     name, 3 additional addresses, telephone, email)
'   - row 23 becomes Other Parties (it was Counterparty 6)
'   - the theme picker moves from U28 to V12:V13
' It first clears what an earlier SetupOsddInputs put in V:AB and shows
' row 23 again if that hid it. Sheet7's [Counterparties] tag is pointed
' at J18:J22 and its [CP 6] tag emptied; Search Matrix's Counterparty 6
' rows are hidden. Nothing else on Sheet1 moves.
'
' Then, with "Trust access to the VBA project object model" on, it makes
' the matching code changes in the other modules - the theme picker's
' new cell, the theme manager calling StyleOsddInputs, and the
' counterparty menus and narrative going to 5 - and lists what it
' changed. Without that setting it lists the changes to make by hand.
'==================================================================
Option Explicit

Private Const PICKER_OLD As String = "U28"
Private Const PICKER_NEW As String = "V13"

Sub SetupSheet1()
    Dim ws As Worksheet, ws7 As Worksheet, sm As Worksheet
    Dim labels As Variant, c As Long, f As String, themeList As String, themeVal As String
    Dim evts As Boolean, smProt As Boolean, stepName As String, errNum As Long, errDesc As String

    Set ws = ThisWorkbook.Sheets("Sheet1")
    evts = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    stepName = "unprotecting Sheet1"
    ws.Unprotect Password:="p7ss"

    ' The theme picker's list and value, from wherever it is now.
    On Error Resume Next
    themeVal = CStr(ws.Range(PICKER_OLD).Value)
    themeList = ws.Range(PICKER_OLD).Validation.Formula1
    If themeVal = "" Then themeVal = CStr(ws.Range(PICKER_NEW).Value)
    If themeList = "" Then themeList = ws.Range(PICKER_NEW).Validation.Formula1
    On Error GoTo Fail
    If themeList = "" Then themeList = "Default (no theme),Dark Blue,Navy & Gold"

    ' Clear the old panel and the old picker; give them the plain sheet
    ' background (U26's) until the panel is styled.
    stepName = "clearing the old panel (V3:AB28) and picker (U27:U28)"
    With ws.Range("V3:AB28")
        .UnMerge
        .Validation.Delete
        .Clear
    End With
    ws.Range("U27:U28").Validation.Delete
    ws.Range("U27:U28").Clear
    PlainLook ws.Range("V3:AB28,U27:U28"), ws.Range("U26")

    ' Row 23: Counterparty 6 becomes Other Parties.
    stepName = "turning row 23 into Other Parties"
    ws.Rows(23).Hidden = False
    ws.Range("G23").Value = "Other Parties"
    ws.Range("J23").MergeArea.ClearContents
    ws.Range("T23").ClearContents
    InputTip ws.Range("J23").MergeArea, "Other Parties", _
        "Others named in the narrative and relied upon (beneficial owners, directors...). " & _
        "Separate names with ;. Each gets a Name and a Negative News search."

    ' The panel.
    stepName = "building the panel (V3:AB10)"
    ws.Columns("V:AB").Hidden = False
    ws.Columns("V").ColumnWidth = 18
    ws.Columns("W:AB").ColumnWidth = 20
    ws.Range("V3:AB3").Merge
    ws.Range("V3").Value = "OSDD Search Inputs"
    ws.Range("W4").Formula = "=""Customer""&IF($J$13="""","""",CHAR(10)&$J$13)"
    For c = 1 To 5
        ws.Cells(4, 23 + c).Formula = "=""CP " & c & """&IF($J$" & (17 + c) & "="""","""",CHAR(10)&$J$" & (17 + c) & ")"
    Next c
    labels = Array("Negative News Name", "Additional Address 1", "Additional Address 2", _
        "Additional Address 3", "Telephone", "Email")
    For c = 0 To 5
        ws.Cells(5 + c, 22).Value = labels(c)
    Next c
    InputTip ws.Range("W5:AB5"), "Negative News Name", _
        "Leave blank to search the name without legal extensions (LLC, Inc...), middle names " & _
        "or suffixes, as the procedure asks - or type the name to search."
    InputTip ws.Range("W6:AB8"), "Additional Address", _
        "Any other address identified for this entity. Each gets an Address and a Name + Address search."
    InputTip ws.Range("W9:AB10"), "Telephone / Email", _
        "Each gets its own search. Several? Separate them with ;"

    ' The theme picker.
    stepName = "moving the theme picker to " & PICKER_NEW
    ws.Range("V12").Value = "Theme"
    With ws.Range(PICKER_NEW).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=themeList
        .InCellDropdown = True
    End With
    ws.Range(PICKER_NEW).Value = themeVal

    stepName = "colouring the panel (Module2 StyleOsddInputs)"
    StyleOsddInputs
    stepName = "protecting Sheet1"
    ws.Protect Password:="p7ss"
    stepName = "fitting Sheet1 to the window"
    FitSheet1ToWindow

    ' Sheet7: [Counterparties] joins CP1-5 only; [CP 6] is no more.
    Set ws7 = ThisWorkbook.Sheets("Sheet7")
    On Error Resume Next
    ws7.Unprotect Password:="p7ss"
    f = ws7.Range("K8").Formula
    If InStr(f, "J18:J23") > 0 Then ws7.Range("K8").Formula = Replace(f, "J18:J23", "J18:J22")
    If InStr(ws7.Range("K22").Formula, "J23") > 0 Then ws7.Range("K22").Formula = "="""""

    ' Search Matrix: hide Counterparty 6's block (rows 32-36).
    Set sm = ThisWorkbook.Sheets("Search Matrix")
    If InStr(sm.Range("A32").Formula, "$J$23") > 0 Then
        smProt = sm.ProtectContents
        If smProt Then sm.Unprotect Password:="p7ss"
        sm.Rows("32:36").Hidden = True
        If smProt Then sm.Protect Password:="p7ss"
    End If
    On Error GoTo 0
    Application.EnableEvents = evts

    MsgBox "Sheet1 is laid out: OSDD Search Inputs panel in V3:AB10, Other Parties in row 23, " & _
        "theme picker in V13." & vbCrLf & vbCrLf & PatchCode(), vbInformation, "Sheet1 Setup"
    Exit Sub

Fail:
    ' Keep the error before On Error Resume Next clears it.
    errNum = Err.Number: errDesc = Err.Description
    Application.EnableEvents = evts
    On Error Resume Next
    ws.Protect Password:="p7ss"
    MsgBox "Sheet1 setup stopped while " & stepName & ":" & vbCrLf & vbCrLf & _
        "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
        "It's safe to run SetupSheet1 again once that's sorted.", vbExclamation, "Sheet1 Setup"
End Sub

' The code changes that go with the new layout, made in whichever module
' holds each line. Returns what was changed, and what to change by hand.
Private Function PatchCode() As String
    Dim proj As Object, comp As Object, cm As Object
    Dim i As Long, ln As String, t As String, inApply As Boolean, hookThere As Boolean, hookAlready As Boolean
    Dim nPicker As Long, nHook As Long, nNarr As Long, nRenLoop As Long, nRenPick As Long
    Dim nMergeMenu As Long, nMergeCase As Long, nMergeMsg As Long, out As String

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    i = proj.VBComponents.Count
    If Err.Number <> 0 Then
        PatchCode = "Couldn't change the code: turn on File > Options > Trust Center > Trust Center " & _
            "Settings > Macro Settings > 'Trust access to the VBA project object model' and run " & _
            "SetupSheet1 again (it's safe to re-run)."
        Exit Function
    End If
    On Error GoTo 0

    For Each comp In proj.VBComponents
        If comp.Name <> "modSheet1Setup" Then
            Set cm = comp.CodeModule
            hookThere = False
            If cm.CountOfLines > 0 Then hookThere = (InStr(cm.Lines(1, cm.CountOfLines), "StyleOsddInputs") > 0)
            inApply = False
            i = 1
            Do While i <= cm.CountOfLines
                ln = cm.Lines(i, 1)
                t = Trim$(ln)

                If Left$(t, 1) <> "'" And InStr(ln, """" & PICKER_OLD & """") > 0 Then
                    cm.ReplaceLine i, Replace(ln, """" & PICKER_OLD & """", """" & PICKER_NEW & """")
                    nPicker = nPicker + 1
                ElseIf t = "For rowIdx = 18 To 23" Then
                    cm.ReplaceLine i, Replace(ln, "18 To 23", "18 To 22")
                    nNarr = nNarr + 1
                ElseIf t = "For i = 1 To 6" And i < cm.CountOfLines Then
                    If InStr(cm.Lines(i + 1, 1), "Range(""J18"").Offset(i - 1") > 0 Then
                        cm.ReplaceLine i, Replace(ln, "To 6", "To 5")
                        nRenLoop = nRenLoop + 1
                    End If
                ElseIf t = "ElseIf Val(entityChoice) >= 1 And Val(entityChoice) <= 6 Then" Then
                    cm.ReplaceLine i, Replace(ln, "<= 6", "<= 5")
                    nRenPick = nRenPick + 1
                ElseIf t = """5 = Counterparty 5"" & vbCrLf & _" And i < cm.CountOfLines Then
                    If Trim$(cm.Lines(i + 1, 1)) = """6 = Counterparty 6"", ""Select Entity"", ""0"")" Then
                        cm.ReplaceLine i, Replace(ln, " & vbCrLf & _", ", ""Select Entity"", ""0"")")
                        cm.DeleteLines i + 1
                        nMergeMenu = nMergeMenu + 1
                    End If
                ElseIf t Like "Case ""6"": entity = ""CP6_""*" Then
                    cm.DeleteLines i
                    i = i - 1
                    nMergeCase = nMergeCase + 1
                ElseIf InStr(ln, "enter a number from 0 to 6.") > 0 Then
                    cm.ReplaceLine i, Replace(ln, "0 to 6.", "0 to 5.")
                    nMergeMsg = nMergeMsg + 1
                ElseIf t Like "*Sub ApplyThemeByKey(*" Then
                    inApply = True
                    If hookThere Then hookAlready = True
                ElseIf inApply And t = "End Sub" Then
                    inApply = False
                ElseIf inApply And t = "NormaliseSheetConventions" And Not hookThere Then
                    cm.InsertLines i + 1, Space$(Len(ln) - Len(LTrim$(ln))) & "RunList ""StyleOsddInputs"""
                    hookThere = True
                    nHook = nHook + 1
                    i = i + 1
                End If
                i = i + 1
            Loop
        End If
    Next comp

    out = "Code changes made:" & vbCrLf
    out = out & Line1(nPicker, 4, "Theme picker cell U28 -> V13 (theme manager, Dark Blue theme, Sheet1 code)")
    out = out & Line1(nHook, 1, "Theme manager re-styles the panel after a theme change", _
        IIf(hookAlready, "already there", ""))
    out = out & Line1(nNarr, 1, "Narrative counterparty review: 5 counterparties")
    out = out & Line1(nRenLoop, 2, "Rename: counterparty list 1-5")
    out = out & Line1(nRenPick, 1, "Rename: counterparty choice 1-5")
    out = out & Line1(nMergeMenu, 1, "PDF Merge: menu without Counterparty 6")
    out = out & Line1(nMergeCase, 1, "PDF Merge: Counterparty 6 choice removed")
    out = out & Line1(nMergeMsg, 1, "PDF Merge: 'enter a number from 0 to 5'")
    PatchCode = out & vbCrLf & "Now: Debug > Compile VBAProject, save, and delete modSheet1Setup."
End Function

Private Function Line1(ByVal got As Long, ByVal expected As Long, ByVal what As String, _
    Optional ByVal note As String = "") As String
    If note <> "" Then
        Line1 = "   OK  " & what & " (" & note & ")" & vbCrLf
    ElseIf got >= expected Then
        Line1 = "   OK  " & what & vbCrLf
    Else
        Line1 = "   NOT FOUND (" & got & " of " & expected & ")  " & what & " - change by hand" & vbCrLf
    End If
End Function

' Help text shown when a cell is selected.
Private Sub InputTip(ByVal rng As Range, ByVal title As String, ByVal msg As String)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputTitle = title
        .InputMessage = msg
        .ShowInput = True
    End With
End Sub

Private Sub PlainLook(ByVal dst As Range, ByVal src As Range)
    On Error Resume Next
    If src.Interior.Pattern = xlNone Then
        dst.Interior.Pattern = xlNone
    Else
        dst.Interior.Color = src.Interior.Color
    End If
    dst.Font.Name = src.Font.Name
    dst.Font.Size = src.Font.Size
    dst.Font.Color = src.Font.Color
End Sub
