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
'     panel; Sheet1's code watches the theme picker's hidden cell; the
'     narrative, Rename and PDF Merge take 5 counterparties (and the Rename
'     form offers CTS Report for counterparties); ThisWorkbook
'     fits Sheet1 to the window on opening
' Step 2, Sheet1 - runs by itself once step 1 has finished, so the
' replaced modules are really gone:
'   - the OSDD inputs beside the customer (row 14) and each counterparty
'     (rows 18-22) in V:AB, with a Listed Entity tick box; the Alert
'     Related Information and country risk rows widened to AB; the theme
'     picker hidden
'     (Ctrl+Shift+T chooses themes)
'     (the theme manager puts copies of the form's banner shapes over
'     them, and the sidebar back in its 3.6.1 places, on every theme
'     change); gaps, column T and the panel narrowed and the text set to
'     one readable size, so Sheet1 fits at a larger zoom; row 23 (the old
'     Counterparty 6) emptied, locked and hidden
'   - Sheet7's [Counterparties] tag on CP1-5 and [CP 6] emptied; Search
'     Matrix rebuilt with every search OSDD Search runs (BuildOsddSearchMatrix)
'   - the current theme applied again, Sheet1 fitted to the window
'==================================================================
Option Explicit

Private Const PWD As String = "p7ss"
Private Const PICKER_OLD As String = "U28"
Private Const PICKER_NEW As String = "AD1"     ' hidden - themes are chosen with Ctrl+Shift+T
Private Const SELF As String = "modInstallUpdate"
' The procedure's Negative News string, in two parts: a formula's text
' can only be 255 characters at a time.
Private Const NEG1 As String = """ AND (arrest OR corruption OR sentencing OR money laundering OR AML OR launder OR embezzle OR evas OR evad OR crime OR corrupt OR bribe OR theft OR extort OR drug OR traffic OR trafficking OR felony OR sanction OR counterfeit OR terror OR """
Private Const NEG2 As String = """investigat*, indict*, charge?, court, lawsuit, judgment, fines, penalties, regulat*, fraud, scam, incarcerat*)"""
Private Const SHEET_HOOK_MARK As String = "' (Inputs V14:AB14, V18:AB22. Search Matrix updates its links when opened.)"
Private Const RESET_LINE As String = _
    "ws.Range(""J5:J11, J13:J15, J18:J23, T18:T23, G27, L27:N27, O27, V14:AB14, V18:AB22"").Value = """""

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
    Dim nBlock As Long, nSheetHook As Long, needSheetHook As Boolean, nAlertRows As Long
    Dim nRenRow As Long, nRenCts As Long
    Dim openNote As String, out As String, smNote As String

    For Each comp In proj.VBComponents
        If comp.Name <> SELF And Left$(comp.Name, 10) <> "zOldModule" Then
            Set cm = comp.CodeModule
            hasFindButton = False
            needSheetHook = False
            If cm.CountOfLines > 0 And comp.Type = 100 Then
                ' Sheet1's change handler: add the panel hook, or swap the
                ' earlier, slower one (it refreshed Search Matrix on every edit).
                needSheetHook = InStr(cm.Lines(1, cm.CountOfLines), "Sub Worksheet_Change(") > 0 And _
                    InStr(cm.Lines(1, cm.CountOfLines), SHEET_HOOK_MARK) = 0
            End If
            If cm.CountOfLines > 0 Then hasFindButton = (InStr(cm.Lines(1, cm.CountOfLines), "Private Function FindButton(") > 0)
            i = 1
            Do While i <= cm.CountOfLines
                ln = cm.Lines(i, 1)
                t = Trim$(ln)
                ind = Space$(Len(ln) - Len(LTrim$(ln)))

                If Left$(t, 1) <> "'" And (InStr(ln, """" & PICKER_OLD & """") > 0 Or InStr(ln, """V13""") > 0) Then
                    cm.ReplaceLine i, Replace(Replace(ln, """" & PICKER_OLD & """", """" & PICKER_NEW & """"), _
                        """V13""", """" & PICKER_NEW & """")
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
                ElseIf t = "Set ProgNameRange = Range(""J7:T7"")" Or t = "Set EscScenarioRange = Range(""J8:T8"")" Then
                    ' Sheet1's greying of Program Name / Escalation Scenario: the
                    ' rows now run to AB.
                    cm.ReplaceLine i, Replace(Replace(ln, ":T7""", ":AB7"""), ":T8""", ":AB8""")
                    nAlertRows = nAlertRows + 1
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
                ElseIf comp.Name = "clsRenameRow" And t = "For i = 1 To 6" Then
                    ' Rename form: counterparties 1-5.
                    cm.ReplaceLine i, Replace(ln, "To 6", "To 5")
                    nRenRow = nRenRow + 1
                ElseIf comp.Name = "clsRenameRow" And Left$(t, 25) = "Private Const CP_CATS As " And _
                    InStr(t, "CTS Report") = 0 Then
                    ' Rename form: a CTS Report category for counterparties too.
                    cm.ReplaceLine i, Left$(ln, Len(RTrim$(ln)) - 1) & "|CTS Report"""
                    cm.InsertLines i + 1, _
                        "' A counterparty's CTS report: ECM_AlertID_CTS Report_<Counterparty n_name>." & vbCrLf & _
                        "Private Const CAT_CTS_CP As String = ""CTS Report"""
                    i = i + 2
                    nRenCts = nRenCts + 1
                ElseIf comp.Name = "clsRenameRow" And t = "newName = ECM & ""_"" & AlertID & ""_"" & CboCategory.Value" And _
                    InStr(cm.Lines(1, cm.CountOfLines), "= CAT_CTS_CP Then") = 0 Then
                    cm.InsertLines i + 1, _
                        Mid$(ind, 5) & "ElseIf CboCategory.Value = CAT_CTS_CP Then" & vbCrLf & _
                        ind & "' Counterparty CTS report - its identifier after the label." & vbCrLf & _
                        ind & "newName = ECM & ""_"" & AlertID & ""_CTS Report_"" & EntityIdentifier()"
                    i = i + 3
                    nRenCts = nRenCts + 1
                ElseIf Left$(t, 30) = "Private Const BLOCK As Long = " And InStr(t, "= 5 ") > 0 Then
                    ' Navy & Gold's Search Matrix bands one entity block at a time.
                    cm.ReplaceLine i, Replace(ln, "= 5 ", "= 13")
                    nBlock = nBlock + 1
                ElseIf needSheetHook And (t = "' OSDD panel (Module2): an email typed there turns into a link and takes" Or _
                    t = "' OSDD inputs (Module2): an email typed there turns into a link and takes") Then
                    ' An earlier 9-line hook (an older range): replace it.
                    cm.DeleteLines i, 9
                    cm.InsertLines i, SheetHook(ind)
                    needSheetHook = False
                    nSheetHook = nSheetHook + 1
                    i = i + 7
                ElseIf needSheetHook And t = "' OSDD panel (Module2): an email typed there turns into a link and" Then
                    ' The earlier hook (10 lines): replace it.
                    cm.DeleteLines i, 10
                    cm.InsertLines i, SheetHook(ind)
                    needSheetHook = False
                    nSheetHook = nSheetHook + 1
                    i = i + 7
                ElseIf needSheetHook And t = "Me.Unprotect Password:=""p7ss""" Then
                    If InStr(cm.Lines(1, cm.CountOfLines), "(Module2): an email typed there turns into a link") = 0 Then
                        cm.InsertLines i + 1, SheetHook(ind)
                        needSheetHook = False
                        nSheetHook = nSheetHook + 1
                        i = i + 8
                    End If
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

    ' Search Matrix: links brought up to date whenever it is opened.
    smNote = ""
    Set cm = Nothing
    On Error Resume Next
    Set cm = proj.VBComponents(ThisWorkbook.Sheets("Search Matrix").CodeName).CodeModule
    On Error GoTo 0
    If cm Is Nothing Then
        smNote = "NOT DONE - Search Matrix's code module not found"
    Else
        t = ""
        If cm.CountOfLines > 0 Then t = cm.Lines(1, cm.CountOfLines)
        If InStr(t, "RefreshOsddMatrix") > 0 Then
            smNote = "already there"
        ElseIf InStr(t, "Sub Worksheet_Activate(") > 0 Then
            smNote = "NOT DONE - Search Matrix has its own Worksheet_Activate: add  Application.Run ""RefreshOsddMatrix""  to it"
        Else
            cm.AddFromString vbCrLf & "' Search Matrix up to date whenever it is opened (Module2 RefreshOsddMatrix)." & vbCrLf & _
                "Private Sub Worksheet_Activate()" & vbCrLf & "    On Error Resume Next" & vbCrLf & _
                "    Application.Run ""RefreshOsddMatrix""" & vbCrLf & "End Sub"
            smNote = "added"
        End If
    End If

    ' ThisWorkbook: fit Sheet1 to the window on opening, and Ctrl+Shift+T for
    ' the theme chooser (let go again on closing, so the shortcut does not
    ' reopen the workbook).
    Set cm = proj.VBComponents("ThisWorkbook").CodeModule
    t = ""
    If cm.CountOfLines > 0 Then t = cm.Lines(1, cm.CountOfLines)
    If InStr(t, "ChooseTheme") > 0 Then
        openNote = "already there"
    ElseIf InStr(t, "    Application.Run ""FitSheet1ToWindow""") > 0 Then
        ' Our earlier Workbook_Open: add the shortcut after the fit.
        For i = 1 To cm.CountOfLines
            If cm.Lines(i, 1) = "    Application.Run ""FitSheet1ToWindow""" Then
                cm.InsertLines i + 1, "    Application.OnKey ""^+t"", ""ChooseTheme""   ' Ctrl+Shift+T: themes"
                Exit For
            End If
        Next i
        openNote = "shortcut added"
    ElseIf InStr(t, "Sub Workbook_Open(") > 0 Then
        openNote = "NOT DONE - ThisWorkbook has its own Workbook_Open: add  Application.Run ""FitSheet1ToWindow""" & _
            "  and  Application.OnKey ""^+t"", ""ChooseTheme""  to it by hand"
    Else
        cm.AddFromString vbCrLf & "' Fit Sheet1 to this screen (Module2 FitSheet1ToWindow)." & vbCrLf & _
            "Private Sub Workbook_Open()" & vbCrLf & "    On Error Resume Next" & vbCrLf & _
            "    Application.Run ""FitSheet1ToWindow""" & vbCrLf & _
            "    Application.OnKey ""^+t"", ""ChooseTheme""   ' Ctrl+Shift+T: themes" & vbCrLf & "End Sub"
        openNote = "added"
    End If
    t = cm.Lines(1, cm.CountOfLines)
    If InStr(t, "Sub Workbook_BeforeClose(") = 0 Then
        cm.AddFromString vbCrLf & "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbCrLf & _
            "    Application.OnKey ""^+t""" & vbCrLf & "End Sub"
    End If

    out = Item(nPicker, 1, "Sheet1's code: theme picker moved to the hidden cell " & PICKER_NEW)
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
    out = out & Item(nBlock, 1, "Navy & Gold Search Matrix: one band per entity")
    out = out & Item(nSheetHook, 1, "Sheet1's code: a link Excel makes of a typed email is taken off")
    out = out & Item(nAlertRows, 2, "Sheet1's code: Program Name / Escalation Scenario greying covers J:AB")
    out = out & Item(nRenRow, 2, "Rename form: counterparties 1-5")
    out = out & Item(nRenCts, 2, "Rename form: CTS Report for counterparties")
    out = out & "   ThisWorkbook fits Sheet1 and sets Ctrl+Shift+T on opening: " & openNote & vbCrLf
    out = out & "   Search Matrix refreshes its links when opened: " & smNote & vbCrLf
    PatchCode = out
End Function

' The lines Sheet1's Worksheet_Change gets, indented by ind. Only an edit
' that left a link behind (an email Excel turned into one) costs anything;
' Search Matrix brings its own links up to date when it is opened.
Private Function SheetHook(ByVal ind As String) As String
    SheetHook = _
        ind & "' OSDD inputs (Module2): a typed email turns into a link and takes the" & vbCrLf & _
        ind & "' link style - take the link off and put the inputs' look back." & vbCrLf & _
        ind & SHEET_HOOK_MARK & vbCrLf & _
        ind & "If Not Intersect(target, Me.Range(""V14:AB14,V18:AB22"")) Is Nothing Then" & vbCrLf & _
        ind & "    If Intersect(target, Me.Range(""V14:AB14,V18:AB22"")).Hyperlinks.Count > 0 Then" & vbCrLf & _
        ind & "        Intersect(target, Me.Range(""V14:AB14,V18:AB22"")).Hyperlinks.Delete" & vbCrLf & _
        ind & "        Application.Run ""StyleOsddInputs""" & vbCrLf & _
        ind & "    End If" & vbCrLf & _
        ind & "End If"
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
    Dim ws As Worksheet, ws7 As Worksheet, cb As Object
    Dim heads As Variant, c As Long, r As Variant, f As String, themeList As String, themeVal As String
    Dim key As String, evts As Boolean, wbProt As Boolean, stepName As String
    Dim pickerErr As String, errNum As Long, errDesc As String

    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' The theme, and the picker's value and list from wherever it was.
    On Error Resume Next
    key = CStr(Application.Run("ActiveTheme"))
    For Each r In Array("U28", "V13", PICKER_NEW)
        If themeVal = "" Then themeVal = CStr(ws.Range(r).Value)
        If themeList = "" Then themeList = ws.Range(r).Validation.Formula1
    Next r
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

    ' Whatever earlier layouts put to the right of the form, and the old
    ' visible theme picker, cleared.
    ' Unmerge only the merged blocks that reach into V:AB, one block at a
    ' time. (Unmerging V3:AB28 as a whole makes Excel widen it to every
    ' block it touches - the Alert rows, then G4:T4 - and unmerge the
    ' whole form. That is what an earlier version of this did.)
    stepName = "clearing the old layout (V3:AB28, U27:U28)"
    UnmergeBlocksIn ws.Range("U3:AB28")
    With ws.Range("V3:AB28")
        .Validation.Delete
        .Clear
    End With
    ws.Range("U27:U28").Validation.Delete
    ws.Range("U27:U28").Clear
    PlainLook ws.Range("V3:AB28,U27:U28"), ws.Range("U26")
    For Each r In Array(12, 13, 14, 18, 19, 20, 21, 22)
        On Error Resume Next
        ws.Shapes("OSDD Listed " & r).Delete
        On Error GoTo Fail
    Next r

    ' The Alert Related Information rows run across to AB, like the banner
    ' over them (the theme manager widens that). Each row's value cell is
    ' still J5-J10, so Sheet7's tags and everything else reading them is
    ' unaffected.
    ' The form's own merged cells, as Beta 3.6.1 has them (they were lost
    ' to the old clean-up described above).
    stepName = "restoring the form's merged cells"
    Application.DisplayAlerts = False
    For Each r In Array("G4:T4", "G5:I5", "G6:I6", "G7:I7", "G8:I8", "G9:I9", "G10:I10", _
        "G12:T12", "G13:I13", "J13:T13", "G14:I14", "J14:T14", _
        "G16:T16", "G17:I17", "J17:S17", "G18:I18", "J18:S18", "G19:I19", "J19:S19", "G20:I20", "J20:S20", _
        "G21:I21", "J21:S21", "G22:I22", "J22:S22", "G23:I23", "J23:S23", _
        "G26:K26", "L26:N26", "P26:Q26", "R26:S26", "G27:K27", "L27:N27", "P27:Q27", "R27:S27")
        If ws.Range(r).Cells(1, 1).MergeArea.Address <> ws.Range(r).Address Then ws.Range(r).Merge
    Next r
    Application.DisplayAlerts = True

    stepName = "widening the Alert Related Information and country risk rows to AB"
    Application.DisplayAlerts = False
    For c = 5 To 10
        ws.Range("J" & c & ":AB" & c).Merge
    Next c
    ' Country Risk Rating too: Excel Risk Class (T26 / T27, still the cells
    ' its formula and Sheet7's [Country Risk] tag use) runs across to AB.
    For c = 26 To 27
        ws.Range("T" & c & ":AB" & c).Merge
    Next c
    Application.DisplayAlerts = True

    ' Row 23 (the old Counterparty 6) emptied, locked and hidden. The theme
    ' manager gives the sidebar buttons fixed sizes, so hiding it does not
    ' squash PDF Merge.
    stepName = "hiding row 23 (the old Counterparty 6)"
    ws.Range("G23").MergeArea.ClearContents
    ws.Range("J23").MergeArea.ClearContents
    ws.Range("J23").MergeArea.Validation.Delete
    ws.Range("T23").ClearContents
    ws.Range("J23").MergeArea.Locked = True
    ws.Range("T23").Locked = True
    ws.Rows(23).Hidden = True

    ' The OSDD inputs, beside the entities they are for, under banners the
    ' theme manager copies from the tables' own (rows 12 and 16): headers
    ' on row 13 and the customer's inputs on row 14; headers on row 17 and
    ' each counterparty's inputs on its own row, 18-22.
    stepName = "laying out the OSDD inputs (V13:AB22)"
    ws.Columns("V:AB").Hidden = False
    ws.Columns("F").ColumnWidth = 1.55
    ws.Columns("T").ColumnWidth = 36
    ws.Columns("U").ColumnWidth = 1.55
    ' Seven equal input columns, as wide together as the form (G:T), so the
    ' paired banners on rows 12 and 16 mirror each other.
    ws.Columns("V:AB").ColumnWidth = 16.5
    heads = Array("Listed Entity", "Negative News", "Other Address 1", "Other Address 2", "Other Address 3", _
        "Phone", "Email")
    For c = 0 To 6
        ws.Cells(13, 22 + c).Value = heads(c)
        ws.Cells(17, 22 + c).Value = heads(c)
    Next c
    InputTip ws.Range("W14,W18:W22"), "Negative News", _
        "Leave blank to search the name without legal extensions (LLC, Inc...), middle names " & _
        "or suffixes, as the procedure asks - or type the name to search."
    InputTip ws.Range("X14:Z14,X18:Z22"), "Other Address", _
        "Any other address identified for this entity. Each gets an Address and a Name + Address search."
    InputTip ws.Range("AA14,AA18:AA22"), "Phone", _
        "Each number gets its own search. Several? Separate them with ;"
    InputTip ws.Range("AB14,AB18:AB22"), "Email", _
        "Each email address gets its own search. Several? Separate them with ;"
    ' One readable text size for the whole form and the inputs (it was 8-9).
    ws.Range("G5:AB10,G13:T14,G17:T22,G26:AB27,V13:AB14,V17:AB22").Font.Size = 10

    ' Listed Entity: a tick box per entity. Ticked, OSDD Search and Search
    ' Matrix leave out that entity's Negative News search.
    stepName = "adding the Listed Entity tick boxes"
    For Each r In Array(14, 18, 19, 20, 21, 22)
        ws.Range("V" & r).Locked = False
        With ws.Range("V" & r)
            Set cb = ws.CheckBoxes.Add(.Left + (.Width - 14) / 2, .Top + (.Height - 14) / 2, 14, 14)
        End With
        cb.Name = "OSDD Listed " & r
        cb.Caption = ""
        cb.LinkedCell = "$V$" & r
        cb.Value = xlOff
        cb.Locked = False
        cb.Placement = xlMove
    Next r

    ' Even rows: every data row and gap 18pt, every banner row 30pt (the
    ' Alert banner's row 4 like the others, with row 3 a gap above it).
    stepName = "evening out the rows"
    ws.Rows(3).RowHeight = 12
    For Each r In Array(4, 12, 16, 25)
        ws.Rows(r).RowHeight = 30
    Next r
    For Each r In Array(5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 17, 18, 19, 20, 21, 22, 24, 26, 27, 28)
        ws.Rows(r).RowHeight = 18
    Next r

    ' Everything centred, across and down.
    stepName = "centring the form"
    With ws.Range("G5:AB10,G13:T14,G17:T22,G26:AB27,V13:AB14,V17:AB22")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Phones as text, so a number keeps its digits (not 5.41169E+11).
    For Each r In Array(14, 18, 19, 20, 21, 22)
        With ws.Range("AA" & r)
            .NumberFormat = "@"
            If Not IsEmpty(.Value) And IsNumeric(.Value) Then .Value = CStr(.Value)
        End With
    Next r

    ' Risk colours on the country risk cells (PDF Risk Class R27, Excel Risk
    ' Class T27): solid fills with their own text colour, so they read the
    ' same in every theme. Every older rule on those rows goes first - the
    ' original dotted ones covered R27:S28 and T27:T28.
    stepName = "colouring the country risk cells"
    ws.Range("R27:AB28").FormatConditions.Delete
    RiskColours ws.Range("R27:S27")
    RiskColours ws.Range("T27:AB27")

    ' The theme picker, out of sight: themes are chosen with Ctrl+Shift+T.
    stepName = "hiding the theme picker (" & PICKER_NEW & ")"
    ws.Range(PICKER_NEW).Value = themeVal
    pickerErr = PickerList(ws.Range(PICKER_NEW), themeList)
    Application.OnKey "^+t", "ChooseTheme"

    stepName = "updating Sheet7 and Search Matrix"
    Set ws7 = ThisWorkbook.Sheets("Sheet7")
    On Error Resume Next
    ws7.Unprotect Password:=PWD
    f = ws7.Range("K8").Formula
    If InStr(f, "J18:J23") > 0 Then ws7.Range("K8").Formula = Replace(f, "J18:J23", "J18:J22")
    If InStr(ws7.Range("K22").Formula, "J23") > 0 Then ws7.Range("K22").Formula = "="""""
    On Error GoTo Fail
    stepName = "rebuilding Search Matrix"
    BuildOsddSearchMatrix
    On Error Resume Next

    ws.Protect Password:=PWD
    If wbProt Then ThisWorkbook.Protect Password:=PWD, Structure:=True, Windows:=False
    Application.EnableEvents = evts

    ' The current theme again. The theme manager first puts the sidebar
    ' back in its 3.6.1 places and copies the tables' banner shapes over the
    ' inputs, so the theme styles them like the rest; then the inputs take
    ' its colours and Sheet1 is fitted to the window.
    Application.Run "ApplyThemeByKey", key
    Application.Run "StyleOsddInputs"
    Application.Run "RefreshOsddMatrix"
    Application.Run "FitSheet1ToWindow"
    On Error GoTo 0

    MsgBox "Step 2 of 2 done - Sheet1: OSDD inputs beside the customer (row 14) and each counterparty " & _
        "(rows 18-22) in columns V:AB, with a Listed Entity tick box; the Alert Related Information " & _
        "and Country Risk Rating blocks run across to AB; row 23 hidden; text at one size; Sheet7 on " & _
        "5 counterparties." & _
        vbCrLf & vbCrLf & _
        "Themes: Ctrl+Shift+T (nothing on the sheet any more)." & vbCrLf & vbCrLf & _
        "Search Matrix: " & IIf(CStr(ThisWorkbook.Sheets("Search Matrix").Range("B5").Value) = "Other Address 1", _
            "rebuilt - 13 searches per entity (customer and CP 1-5), rows 2 to 79.", _
            "NOT rebuilt - send a screenshot of this message.") & vbCrLf & vbCrLf & _
        IIf(pickerErr = "", "", "(The hidden theme list couldn't be set: " & pickerErr & ")" & vbCrLf & vbCrLf) & _
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

' Search Matrix rebuilt as one 13-row block per entity - the customer,
' then counterparties 1-5 - with every search OSDD Search runs: name,
' address, name + address, the panel's three other addresses (each on its
' own and with the name), phone, email, and Negative News pages 1 and 2,
' which use the panel's Negative News name when it has one. Page 2 now
' starts at result 11 (&start=10), as OSDD Search's does.
Private Sub BuildOsddSearchMatrix()
    Dim sm As Worksheet, smProt As Boolean, b As Long, k As Long, r As Long, s As Long
    Dim nm As String, ad As String, pc As String, ent As String, lbl As String
    Dim kinds As Variant, fileKinds As Variant, q As String, oa As String, er As Long

    Set sm = ThisWorkbook.Sheets("Search Matrix")
    smProt = sm.ProtectContents
    If smProt Then sm.Unprotect Password:=PWD

    sm.Rows("2:200").Hidden = False
    On Error Resume Next
    sm.Range("E2:E200").Hyperlinks.Delete
    On Error GoTo 0
    sm.Range("A2:N200").ClearContents
    sm.Range("A2:N2").Copy
    sm.Range("A3:N79").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    kinds = Array("Google name", "Address", "Name + Address", "Other Address 1", "Name + Other Address 1", _
        "Other Address 2", "Name + Other Address 2", "Other Address 3", "Name + Other Address 3", _
        "Phone", "Email", "Negative News, page 1", "Negative News, page 2")
    fileKinds = Array("Google", "Address", "Google+Address", "Address 2", "Google+Address 2", _
        "Address 3", "Google+Address 3", "Address 4", "Google+Address 4", _
        "Telephone", "Email", "Negative News 1", "Negative News 2")

    For b = 0 To 5
        s = 2 + b * 13
        ' This entity's OSDD inputs on Sheet1: row 14 for the customer, its
        ' own row for a counterparty; Listed V, Negative News W, other
        ' addresses X:Z, phone AA, email AB.
        er = IIf(b = 0, 14, 17 + b)
        If b = 0 Then
            nm = "Sheet1!$J$13": ad = "Sheet1!$J$14": ent = "Customer"
            lbl = "=IF(" & nm & "="""",""Customer""," & nm & "&"" (Customer)"")"
        Else
            nm = "Sheet1!$J$" & (17 + b): ad = "Sheet1!$T$" & (17 + b): ent = "CP" & b
            lbl = "=IF(" & nm & "="""",""Counterparty " & b & """," & nm & "&"" (CP" & b & ")"")"
        End If
        For k = 0 To 12
            r = s + k
            Select Case k
                Case 0: q = "=IF(" & nm & "="""",""""," & nm & ")"
                Case 1: q = "=IF(" & ad & "="""",""""," & ad & ")"
                Case 2: q = "=IF(OR(" & nm & "=""""," & ad & "=""""),""""," & nm & "&"" + ""&" & ad & ")"
                Case 3, 5, 7
                    oa = "Sheet1!$" & Choose((k - 1) \ 2, "X", "Y", "Z") & "$" & er
                    q = "=IF(OR(" & nm & "=""""," & oa & "=""""),""""," & oa & ")"
                Case 4, 6, 8
                    oa = "Sheet1!$" & Choose((k - 2) \ 2, "X", "Y", "Z") & "$" & er
                    q = "=IF(OR(" & nm & "=""""," & oa & "=""""),""""," & nm & "&"" + ""&" & oa & ")"
                Case 9
                    oa = "Sheet1!$AA$" & er
                    q = "=IF(OR(" & nm & "=""""," & oa & "=""""),"""",SUBSTITUTE(SUBSTITUTE(TRIM(" & oa & _
                        "),""; "","";""),"";"","" OR ""))"
                Case 10
                    oa = "Sheet1!$AB$" & er
                    q = "=IF(OR(" & nm & "=""""," & oa & "=""""),"""",CHAR(34)&SUBSTITUTE(SUBSTITUTE(TRIM(" & _
                        oa & "),""; "","";""),"";"",CHAR(34)&"" OR ""&CHAR(34))&CHAR(34))"
                Case 11, 12
                    ' None for a listed entity (its tick box on Sheet1, column V).
                    oa = "Sheet1!$W$" & er
                    q = "=IF(OR(" & nm & "="""",Sheet1!$V$" & er & "=TRUE),"""",CHAR(34)&IF(" & oa & "=""""," & _
                        nm & "," & oa & ")&CHAR(34)&" & NEG1 & "&" & NEG2 & ")"
            End Select
            sm.Range("A" & r).Formula = lbl
            sm.Range("B" & r).Value = kinds(k)
            sm.Range("C" & r).Formula = "=IF(L" & r & "="""","""",L" & r & ")"
            sm.Range("D" & r).Formula = "=IF(F" & r & "="""","""",Sheet1!$J$9&""_""&Sheet1!$J$10&""_" & ent & _
                "_""&N" & r & "&""_" & fileKinds(k) & ".pdf"")"
            sm.Range("F" & r).Formula = q
            sm.Range("G" & r).Formula = "=IF(F" & r & "="""","""",SUBSTITUTE(F" & r & ",CHAR(34),""%22""))"
            sm.Range("H" & r).Formula = "=IF(G" & r & "="""","""",SUBSTITUTE(G" & r & ",""&"",""%26""))"
            sm.Range("I" & r).Formula = "=IF(H" & r & "="""","""",SUBSTITUTE(H" & r & ",""("",""%28""))"
            sm.Range("J" & r).Formula = "=IF(I" & r & "="""","""",SUBSTITUTE(I" & r & ","")"",""%29""))"
            sm.Range("K" & r).Formula = "=IF(J" & r & "="""","""",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(J" & r & _
                ",""+"",""%2B""),""#"",""%23""),"" "",""+""))"
            sm.Range("L" & r).Formula = "=IF(K" & r & "="""","""",""https://www.google.com/search?q=""&K" & r & _
                "&""&num=100&hl=en" & IIf(k = 12, "&start=10", "") & """)"
            sm.Range("M" & r).Formula = "=IF(F" & r & "="""",0,1)"
            sm.Range("N" & r).Formula = "=SUM($M$" & s & ":M" & r & ")"
        Next k
    Next b
    sm.Calculate
    If smProt Then sm.Protect Password:=PWD
End Sub

' Prohibited, High, Medium, Low and N/A as solid fills with readable text.
Private Sub RiskColours(ByVal rng As Range)
    Dim levels As Variant, lv As Variant, fc As Object
    rng.FormatConditions.Delete
    levels = Array( _
        Array("Prohibited", RGB(112, 0, 0), RGB(255, 255, 255)), _
        Array("High", RGB(192, 0, 0), RGB(255, 255, 255)), _
        Array("Medium", RGB(255, 192, 0), RGB(0, 0, 0)), _
        Array("Low", RGB(0, 128, 64), RGB(255, 255, 255)), _
        Array("N/A", RGB(128, 128, 128), RGB(255, 255, 255)))
    For Each lv In levels
        Set fc = rng.FormatConditions.Add(Type:=xlTextString, String:=lv(0), TextOperator:=xlContains)
        fc.Interior.Pattern = xlSolid
        fc.Interior.Color = lv(1)
        fc.Font.Color = lv(2)
        fc.Font.Bold = True
        fc.StopIfTrue = True
    Next lv
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

' Unmerges each merged block that has a cell in rng - that block only,
' however far it reaches outside rng, and nothing else.
Private Sub UnmergeBlocksIn(ByVal rng As Range)
    Dim cell As Range, blocks As New Collection, addr As Variant, seen As String
    For Each cell In rng.Cells
        If cell.MergeCells Then
            If InStr(seen, "|" & cell.MergeArea.Address & "|") = 0 Then
                seen = seen & "|" & cell.MergeArea.Address & "|"
                blocks.Add cell.MergeArea.Address
            End If
        End If
    Next cell
    For Each addr In blocks
        rng.Worksheet.Range(addr).UnMerge
    Next addr
End Sub

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
