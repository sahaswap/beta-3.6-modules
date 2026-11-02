Attribute VB_Name = "modInstallUpdate"
'==================================================================
' modInstallUpdate - ONE-TIME installer for this update. Import it, run
' InstallUpdate once, then delete this module.
'
' Needs "Trust access to the VBA project object model" (File > Options >
' Trust Center > Trust Center Settings > Macro Settings), and one folder
' holding the new module files: Module2.bas, modTronLegacy.bas,
' modThemeManager.bas and modDesert.bas (one left out keeps the
' workbook's own copy, if it has one).
'
' Step 1, the code:
'   - replaces OSDD Search (Module2), the Dark Blue theme (Module13), the
'     theme manager (Module15) and the Desert theme, found by what they
'     contain, after exporting each one to a backup folder beside the new
'     files; removes the earlier one-off helpers (frmOsddInputs,
'     modBuildFrmOsddInputs, modSheet1Setup)
'   - edits single lines in the other modules: Reset clears the OSDD
'     panel; Sheet1's code watches the theme picker in V13; the
'     narrative, Rename and PDF Merge take 5 counterparties; ThisWorkbook
'     fits Sheet1 to the window on opening
' Step 2, Sheet1 - runs by itself once step 1 has finished, so the
' replaced modules are really gone:
'   - the OSDD Search Inputs panel in V3:AB10 and the theme picker in V13
'     (the theme manager puts copies of the form's banner shapes over
'     them, and the sidebar back in its 3.6.1 places, on every theme
'     change); gaps, column T and the panel narrowed and the text set to
'     one readable size, so Sheet1 fits at a larger zoom; row 23 (the old
'     Counterparty 6) emptied, locked and hidden
'   - Sheet7's [Counterparties] tag on CP1-5 and [CP 6] emptied; Search
'     Matrix's Counterparty 6 rows hidden
'   - the current theme applied again, Sheet1 fitted to the window
'==================================================================
Option Explicit

Private Const PWD As String = "p7ss"
Private Const PICKER_OLD As String = "U28"
Private Const PICKER_NEW As String = "V13"
Private Const SELF As String = "modInstallUpdate"
Private Const RESET_LINE As String = _
    "ws.Range(""J5:J11, J13:J15, J18:J23, T18:T23, G27, L27:N27, O27, W5:AB10"").Value = """""

'------------------------------------------------------------------
' STEP 1 - the code
'------------------------------------------------------------------
Sub InstallUpdate()
    Dim proj As Object, comp As Object, folder As String, bakDir As String
    Dim files As Variant, names As Variant, marks As Variant, copies(0 To 3) As String
    Dim f As Variant, i As Long, n As Long
    Dim problem As String, removeList As New Collection, log As String, code As String

    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    n = proj.VBComponents.Count
    If Err.Number <> 0 Or n = 0 Then
        MsgBox "Turn on File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
            "'Trust access to the VBA project object model', then run InstallUpdate again.", _
            vbExclamation, "Install Update"
        Exit Sub
    End If
    On Error GoTo Fail

    ' The folder with the new files. Each one is checked to be the right
    ' module and imported from a clean copy (see PrepareFile), so it does
    ' not matter where it was downloaded from.
    With Application.FileDialog(4)                    ' msoFileDialogFolderPicker
        .Title = "Folder with Module2.bas, modTronLegacy.bas, modThemeManager.bas, modDesert.bas"
        .InitialFileName = Environ$("USERPROFILE") & "\Downloads\"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    files = Array("Module2.bas", "modTronLegacy.bas", "modThemeManager.bas", "modDesert.bas")
    names = Array("Module2", "modTronLegacy", "modThemeManager", "modDesert")
    marks = Array("Sub SearchAndSavePDF_Direct(", "Sub ApplyTronLegacy(", "Sub ApplyThemeByKey(", "Sub ApplyDesert(")
    ' A file left out of the folder keeps the workbook's own copy of that
    ' module - fine when it has not changed since the last update.
    For i = 0 To UBound(files)
        If Dir$(folder & files(i)) = "" And HasModule(proj, CStr(marks(i))) Then
            log = log & "   kept the workbook's " & names(i) & " (not in the folder)" & vbCrLf
        Else
            copies(i) = PrepareFile(folder & files(i), CStr(names(i)), CStr(marks(i)), problem)
        End If
    Next i
    If problem <> "" Then
        MsgBox "Nothing was changed - fix these first:" & vbCrLf & problem, vbExclamation, "Install Update"
        Exit Sub
    End If

    ' What gets replaced or removed.
    For Each comp In proj.VBComponents
        If comp.Name <> SELF Then
            Select Case comp.Name
                Case "frmOsddInputs", "modBuildFrmOsddInputs", "modSheet1Setup"
                    removeList.Add comp
                Case Else
                    If comp.Type = 1 And comp.CodeModule.CountOfLines > 0 Then    ' standard module
                        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                        For i = 0 To UBound(marks)
                            If copies(i) <> "" And InStr(code, marks(i)) > 0 Then
                                removeList.Add comp
                                Exit For
                            End If
                        Next i
                    End If
            End Select
        End If
    Next comp

    bakDir = folder & "Backup before update " & Format$(Now, "yyyy-mm-dd hhnnss") & "\"
    MkDir bakDir
    n = 0
    For Each comp In removeList
        comp.Export bakDir & comp.Name & IIf(comp.Type = 3, ".frm", ".bas")
        log = log & "   replaced / removed " & comp.Name & vbCrLf
        ' A removed module only goes when this macro ends; renaming it
        ' first lets the new one take its name now.
        n = n + 1
        comp.Name = "zOldModule" & n & Format$(Now, "nnss")
        proj.VBComponents.Remove comp
    Next comp

    For i = 0 To UBound(files)
        If copies(i) <> "" Then
            Set comp = proj.VBComponents.Import(copies(i))
            log = log & "   imported " & comp.Name & vbCrLf
        End If
    Next i

    log = log & PatchCode(proj)

    ' Step 2 needs the replaced modules gone, which only happens once this
    ' macro has ended.
    Application.OnTime Now + TimeSerial(0, 0, 2), "'" & ThisWorkbook.Name & "'!InstallUpdate_Step2"
    MsgBox "Step 1 of 2 done - the code:" & vbCrLf & log & vbCrLf & _
        "Backups of what was replaced: " & bakDir & vbCrLf & vbCrLf & _
        "Click OK and step 2 lays out Sheet1.", vbInformation, "Install Update"
    Exit Sub

Fail:
    MsgBox "Install stopped: error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
        "Done so far:" & vbCrLf & log, vbExclamation, "Install Update"
End Sub

' A clean copy of a module file to import, in the TEMP folder: Windows
' (CRLF) line endings - the VBA editor needs them - with its
' Attribute VB_Name line first and no byte-order mark. mark is a line
' only that module has. Returns the copy's path, or "" with the reason
' added to problem.
Private Function PrepareFile(ByVal path As String, ByVal modName As String, ByVal mark As String, _
    ByRef problem As String) As String
    Dim fh As Integer, s As String, outPath As String

    If Dir$(path) = "" Then
        problem = problem & "   " & modName & ".bas is not in the folder" & vbCrLf
        Exit Function
    End If
    fh = FreeFile
    Open path For Binary Access Read As #fh
    s = Space$(LOF(fh))
    Get #fh, , s
    Close #fh

    If Left$(s, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then s = Mid$(s, 4)
    If InStr(s, mark) = 0 Then
        problem = problem & "   " & modName & ".bas is not the " & modName & " module" & vbCrLf
        Exit Function
    End If
    s = Replace(Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf), vbLf, vbCrLf)
    If Left$(s, 17) = "Attribute VB_Name" Then s = Mid$(s, InStr(s, vbCrLf) + 2)
    s = "Attribute VB_Name = """ & modName & """" & vbCrLf & s

    outPath = Environ$("TEMP") & "\" & modName & ".bas"
    If Dir$(outPath) <> "" Then Kill outPath
    fh = FreeFile
    Open outPath For Binary Access Write As #fh
    Put #fh, , s
    Close #fh
    PrepareFile = outPath
End Function

' Whether a standard module (other than this one) holds the line mark.
Private Function HasModule(ByVal proj As Object, ByVal mark As String) As Boolean
    Dim comp As Object
    For Each comp In proj.VBComponents
        If comp.Name <> SELF And comp.Type = 1 Then
            If comp.CodeModule.CountOfLines > 0 Then
                If InStr(comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines), mark) > 0 Then
                    HasModule = True
                    Exit Function
                End If
            End If
        End If
    Next comp
End Function

' The single-line changes, in whichever module holds each line.
Private Function PatchCode(ByVal proj As Object) As String
    Dim comp As Object, cm As Object, i As Long, ln As String, t As String, ind As String
    Dim nPicker As Long, nReset As Long, nResetDone As Long, nNarr As Long, nRenLoop As Long
    Dim nRenPick As Long, nMergeMenu As Long, nMergeCase As Long, nMergeMsg As Long, nNavy As Long
    Dim nNavyBtn As Long, nNavyTitle As Long, hasFindButton As Boolean
    Dim openNote As String, out As String

    For Each comp In proj.VBComponents
        If comp.Name <> SELF And Left$(comp.Name, 10) <> "zOldModule" Then
            Set cm = comp.CodeModule
            hasFindButton = False
            If cm.CountOfLines > 0 Then hasFindButton = (InStr(cm.Lines(1, cm.CountOfLines), "Private Function FindButton(") > 0)
            i = 1
            Do While i <= cm.CountOfLines
                ln = cm.Lines(i, 1)
                t = Trim$(ln)
                ind = Space$(Len(ln) - Len(LTrim$(ln)))

                If Left$(t, 1) <> "'" And InStr(ln, """" & PICKER_OLD & """") > 0 Then
                    cm.ReplaceLine i, Replace(ln, """" & PICKER_OLD & """", """" & PICKER_NEW & """")
                    nPicker = nPicker + 1
                ElseIf Left$(t, 62) = Left$(RESET_LINE, 62) And Right$(t, 11) = ".Value = """"" Then
                    If t = RESET_LINE Then
                        nResetDone = nResetDone + 1
                    Else
                        cm.ReplaceLine i, ind & RESET_LINE
                        nReset = nReset + 1
                    End If
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
                ElseIf t = "If InStr(1, txt, t, vbTextCompare) > 0 Then" And hasFindButton Then
                    ' Navy & Gold's FindButton: only a shape with a macro is a button.
                    cm.ReplaceLine i, ind & "If InStr(1, txt, t, vbTextCompare) > 0 And Len(shp.OnAction) > 0 Then"
                    nNavyBtn = nNavyBtn + 1
                ElseIf InStr(ln, "Characters.text = ""Beta 3.5""") > 0 Then
                    cm.ReplaceLine i, Replace(ln, """Beta 3.5""", """Beta 3.6.1""")
                    nNavyTitle = nNavyTitle + 1
                ElseIf t = "InStr(1, txt, ""OSDD"", vbTextCompare) > 0 Or _" Then
                    ' Navy & Gold takes any shape saying OSDD for a button; the
                    ' OSDD Search Inputs banner is not one.
                    cm.ReplaceLine i, ind & "(InStr(1, txt, ""OSDD"", vbTextCompare) > 0 And " & _
                        "InStr(1, txt, ""Inputs"", vbTextCompare) = 0) Or _"
                    nNavy = nNavy + 1
                End If
                i = i + 1
            Loop
        End If
    Next comp

    ' ThisWorkbook: fit Sheet1 to the window on opening.
    Set cm = proj.VBComponents("ThisWorkbook").CodeModule
    t = ""
    If cm.CountOfLines > 0 Then t = cm.Lines(1, cm.CountOfLines)
    If InStr(t, "FitSheet1ToWindow") > 0 Then
        openNote = "already there"
    ElseIf InStr(t, "Sub Workbook_Open(") > 0 Then
        openNote = "NOT DONE - ThisWorkbook has its own Workbook_Open: add  Application.Run ""FitSheet1ToWindow""  to it by hand"
    Else
        cm.AddFromString vbCrLf & "' Fit Sheet1 to this screen (Module2 FitSheet1ToWindow)." & vbCrLf & _
            "Private Sub Workbook_Open()" & vbCrLf & "    On Error Resume Next" & vbCrLf & _
            "    Application.Run ""FitSheet1ToWindow""" & vbCrLf & "End Sub"
        openNote = "added"
    End If

    out = Item(nPicker, 1, "Sheet1's code: theme picker U28 -> V13")
    out = out & Item(nReset + nResetDone, 1, "Reset clears the OSDD panel" & IIf(nReset = 0 And nResetDone > 0, " (already)", ""))
    out = out & Item(nNarr, 1, "Narrative counterparty review: 5 counterparties")
    out = out & Item(nRenLoop, 2, "Rename: counterparty list 1-5")
    out = out & Item(nRenPick, 1, "Rename: counterparty choice 1-5")
    out = out & Item(nMergeMenu, 1, "PDF Merge: menu without Counterparty 6")
    out = out & Item(nMergeCase, 1, "PDF Merge: Counterparty 6 choice removed")
    out = out & Item(nMergeMsg, 1, "PDF Merge: 'enter a number from 0 to 5'")
    out = out & Item(nNavy, 1, "Navy & Gold styles the OSDD Search Inputs banner as a banner")
    out = out & Item(nNavyBtn, 1, "Navy & Gold moves only shapes with a macro as buttons")
    out = out & Item(nNavyTitle, 1, "Navy & Gold title: Beta 3.6.1")
    out = out & "   ThisWorkbook fits Sheet1 on opening: " & openNote & vbCrLf
    PatchCode = out
End Function

Private Function Item(ByVal got As Long, ByVal expected As Long, ByVal what As String) As String
    If got >= expected Then
        Item = "   OK  " & what & vbCrLf
    Else
        Item = "   NOT FOUND (" & got & " of " & expected & ")  " & what & " - already changed, or change by hand" & vbCrLf
    End If
End Function

'------------------------------------------------------------------
' STEP 2 - Sheet1 (started by step 1)
'------------------------------------------------------------------
Public Sub InstallUpdate_Step2()
    Dim ws As Worksheet, ws7 As Worksheet, sm As Worksheet
    Dim labels As Variant, c As Long, f As String, themeList As String, themeVal As String, key As String
    Dim evts As Boolean, wbProt As Boolean, smProt As Boolean, stepName As String
    Dim pickerErr As String, errNum As Long, errDesc As String

    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' The theme and the picker's list, from wherever the picker is now.
    On Error Resume Next
    key = CStr(Application.Run("ActiveTheme"))
    themeVal = CStr(ws.Range(PICKER_OLD).Value)
    themeList = ws.Range(PICKER_OLD).Validation.Formula1
    If themeVal = "" Then themeVal = CStr(ws.Range(PICKER_NEW).Value)
    If themeList = "" Then themeList = ws.Range(PICKER_NEW).Validation.Formula1
    On Error GoTo 0
    If key = "" Then key = "NONE"
    If InStr(themeList, "Desert") = 0 Then themeList = "Default (no theme),Dark Blue,Desert,Navy & Gold"

    evts = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    stepName = "unprotecting Sheet1"
    wbProt = ThisWorkbook.ProtectStructure
    If wbProt Then ThisWorkbook.Unprotect Password:=PWD
    ws.Unprotect Password:=PWD

    stepName = "clearing the old panel (V3:AB28) and picker (U27:U28)"
    With ws.Range("V3:AB28")
        .UnMerge
        .Validation.Delete
        .Clear
    End With
    ws.Range("U27:U28").Validation.Delete
    ws.Range("U27:U28").Clear
    PlainLook ws.Range("V3:AB28,U27:U28"), ws.Range("U26")

    ' Row 23 (the old Counterparty 6) emptied, locked and hidden. The theme
    ' manager gives the sidebar buttons fixed sizes, so hiding it no longer
    ' squashes PDF Merge.
    stepName = "hiding row 23 (the old Counterparty 6)"
    ws.Range("G23").MergeArea.ClearContents
    ws.Range("J23").MergeArea.ClearContents
    ws.Range("J23").MergeArea.Validation.Delete
    ws.Range("T23").ClearContents
    ws.Range("J23").MergeArea.Locked = True
    ws.Range("T23").Locked = True
    ws.Rows(23).Hidden = True

    ' The panel: a column-header row like the counterparty table's (row 17)
    ' and five input rows, rows 5-10 beside the Alert Related Information
    ' rows.
    stepName = "building the panel (V5:AB10)"
    ' Narrow gaps either side of the form and a narrower address column and
    ' panel, so the whole sheet fits the window at a larger zoom.
    ws.Columns("V:AB").Hidden = False
    ws.Columns("F").ColumnWidth = 1.55
    ws.Columns("T").ColumnWidth = 36
    ws.Columns("U").ColumnWidth = 1.55
    ws.Columns("V").ColumnWidth = 15
    ws.Columns("W:AB").ColumnWidth = 12.5
    ' Short headers and labels, so none of them has to shrink to fit; the
    ' names are beside them on the form.
    ws.Range("V5").Value = "Input"
    ws.Range("W5").Value = "Customer"
    For c = 1 To 5
        ws.Cells(5, 23 + c).Value = "CP " & c
    Next c
    labels = Array("Negative News", "Other Address 1", "Other Address 2", "Other Address 3", "Phone / Email")
    For c = 0 To 4
        ws.Cells(6 + c, 22).Value = labels(c)
    Next c
    InputTip ws.Range("W6:AB6"), "Negative News", _
        "Leave blank to search the name without legal extensions (LLC, Inc...), middle names " & _
        "or suffixes, as the procedure asks - or type the name to search."
    InputTip ws.Range("W7:AB9"), "Other Address", _
        "Any other address identified for this entity. Each gets an Address and a Name + Address search."
    InputTip ws.Range("W10:AB10"), "Phone / Email", _
        "Each gets its own search. Several? Separate them with ;"
    ' One readable text size for the whole form and panel (it was 8 to 9).
    ws.Range("G5:T10,G13:T14,G17:T22,G26:T27,V5:AB10,V13").Font.Size = 10

    stepName = "moving the theme picker to " & PICKER_NEW
    ws.Range(PICKER_NEW).Value = themeVal
    pickerErr = PickerList(ws.Range(PICKER_NEW), themeList)

    stepName = "updating Sheet7 and Search Matrix"
    Set ws7 = ThisWorkbook.Sheets("Sheet7")
    On Error Resume Next
    ws7.Unprotect Password:=PWD
    f = ws7.Range("K8").Formula
    If InStr(f, "J18:J23") > 0 Then ws7.Range("K8").Formula = Replace(f, "J18:J23", "J18:J22")
    If InStr(ws7.Range("K22").Formula, "J23") > 0 Then ws7.Range("K22").Formula = "="""""
    Set sm = ThisWorkbook.Sheets("Search Matrix")
    If InStr(sm.Range("A32").Formula, "$J$23") > 0 Then
        smProt = sm.ProtectContents
        If smProt Then sm.Unprotect Password:=PWD
        sm.Rows("32:36").Hidden = True
        If smProt Then sm.Protect Password:=PWD
    End If

    ws.Protect Password:=PWD
    If wbProt Then ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
    Application.EnableEvents = evts

    ' The current theme again. The theme manager first puts the sidebar
    ' back in its 3.6.1 places and copies the form's banner shapes over the
    ' panel and the picker, so the theme styles them like the rest; then
    ' the panel's cells take its colours and Sheet1 is fitted to the window.
    Application.Run "ApplyThemeByKey", key
    Application.Run "StyleOsddInputs"
    Application.Run "FitSheet1ToWindow"
    On Error GoTo 0

    MsgBox "Step 2 of 2 done - Sheet1: OSDD Search Inputs panel in V3:AB10 and the theme picker " & _
        "in V13, both under banners like the form's; row 23 hidden; text at one size; Sheet7 and " & _
        "Search Matrix on 5 counterparties." & vbCrLf & vbCrLf & _
        IIf(pickerErr = "", "", "The theme dropdown couldn't be added (" & pickerErr & "). Run " & _
            "BuildThemePicker once to add it." & vbCrLf & vbCrLf) & _
        "Now: Debug > Compile VBAProject, save the workbook, and delete modInstallUpdate.", _
        vbInformation, "Install Update"
    Exit Sub

Fail:
    ' Keep the error before On Error Resume Next clears it.
    errNum = Err.Number: errDesc = Err.Description
    Application.EnableEvents = evts
    On Error Resume Next
    ws.Protect Password:=PWD
    If wbProt Then ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
    MsgBox "Sheet1 layout stopped while " & stepName & ":" & vbCrLf & vbCrLf & _
        "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
        "Run InstallUpdate_Step2 again once that's sorted (it's safe to re-run).", _
        vbExclamation, "Install Update"
End Sub

' Puts the theme list on the picker cell; "" if it worked, otherwise why
' not. Tried a second time with Sheet1 active and the cell selected.
Private Function PickerList(ByVal rng As Range, ByVal list As String) As String
    Dim attempt As Long
    On Error Resume Next
    For attempt = 1 To 2
        If attempt = 2 Then
            rng.Parent.Activate
            rng.Select
        End If
        Err.Clear
        rng.Validation.Delete
        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=list
        If Err.Number = 0 Then
            rng.Validation.IgnoreBlank = True
            rng.Validation.InCellDropdown = True
            rng.Validation.ShowInput = False
            rng.Validation.ShowError = False
            PickerList = ""
            Exit Function
        End If
        PickerList = "error " & Err.Number & ": " & Err.Description
    Next attempt
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
