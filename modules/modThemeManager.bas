Attribute VB_Name = "modThemeManager"
Option Explicit
'=====================================================================
' modThemeManager - one active theme at a time, chosen from a dropdown.
'
' WHY THIS EXISTS
' Each theme module saves the "original" look of every shape it touches
' (into that shape's AltText) and of every cell (into a hidden backup
' sheet), so it can put things back. That only works if the theme is
' applied to an UNTHEMED sheet.
'
' Apply Tron on top of a live Navy & Gold sheet and Tron records Navy &
' Gold's navy fills and gold fonts as "the original". Removing Tron then
' returns the sheet to Navy & Gold - and Navy & Gold's own stash has been
' overwritten, so it can never be removed either. The two themes are now
' permanently fused, which is the state that produced gold corner folds
' sitting on a black Grid canvas.
'
' So every apply here removes whatever is active FIRST, then applies the
' requested theme to a clean baseline. Themes coexist in the workbook;
' they never overlap on the sheet.
'
' ---------------------------------------------------------------------
' ADDING A THEME LATER: add ONE row to ThemeRegistry below. Nothing else
' in this module, or anywhere else, needs to change - the dropdown, the
' clear-everything sweep and the state marker are all driven from that
' table. Then re-run BuildThemePicker once to refresh the list.
' ---------------------------------------------------------------------
'
' SETUP (once):  run BuildThemePicker
'                then paste the Worksheet_Change stub from
'                Sheet1_ThemePicker_snippet.txt into the Sheet1 code
'                module, so picking from the dropdown actually fires.
'
' EVERY cross-module call goes through Application.Run on purpose. Do not
' "tidy" them into qualified calls like modTronLegacy.ApplyTronLegacy: a
' qualified call is resolved at COMPILE time and needs a module of that
' exact name to exist, so if a .bas was pasted into the editor rather than
' imported (File > Import File) the module is called Module15 or similar,
' the qualifier resolves to nothing, and Option Explicit reports it as
' "Variable not defined" - a compile error that stops the whole project.
' Application.Run resolves by PROCEDURE name at run time, so it works
' whatever the module is called, and it lets this manager tolerate a theme
' module that is not installed at all.
'=====================================================================
Public Const THEME_NONE As String = "NONE"
Public Const THEME_TRON As String = "TRON"
Public Const THEME_DESERT As String = "DESERT"
Public Const THEME_NAVY As String = "NAVYGOLD"

Private Const SHEET_NAME As String = "Sheet1"
' The picker's cell is kept out of sight (row 1 and column AD are hidden):
' themes are chosen with Ctrl+Shift+T (ChooseTheme), not on the sheet.
Private Const PICKER_CELL As String = "AD1"
Private Const STATE_NAME As String = "_ActiveTheme"
Private Const PWD As String = "p7ss"

' Removers for themes that are NOT in the registry - older ones this
' workbook has carried at some point. They are swept on every clear so a
' leftover from one of them cannot sit under a new theme, but they are
' deliberately not offered in the dropdown. modDarkTheme is still in the
' project and this workbook has a _DarkBak sheet, so its remover is worth
' calling; RemoveGlassStyle is listed only in case that module is ever
' brought back. A missing procedure is skipped silently.
Private Const LEGACY_REMOVERS As String = "RemoveDarkCells,RemoveGlassStyle,RemoveDune"

' Guards against the picker's own Worksheet_Change firing again while a
' theme is being applied.
Private busy As Boolean

' SEAMLESS SWITCHING.
'
' One theme change used to raise up to SEVEN dialogs: ClearTheme sweeps
' every remover in the registry and each announced itself, then the apply
' announced itself, then Navy & Gold's two companion stylers announced
' themselves. Every one of them needed an OK before the next could run.
'
' Success messages are noise here anyway - the picker cell already shows
' which theme is active, which is better confirmation than a dialog you
' dismiss and forget. So while the manager is driving, theme modules stay
' silent and progress goes to the status bar, which does not block.
'
' FAILURES still raise a dialog. Silence on success, never on error.
Private quietMode As Boolean

Public Function ThemeIsQuiet() As Boolean
    ThemeIsQuiet = quietMode
End Function

'=====================================================================
' THE REGISTRY - the single place that knows what themes exist.
'
' Columns:
'   0 key
'   1 name shown in the dropdown
'   2 apply proc
'   3 remove proc
'   4 EXTRA apply procs, comma separated - companion sheets
'   5 EXTRA remove procs, comma separated
'
' Columns 4 and 5 exist because a theme is not a Sheet1 theme. Tron
' styles Search Matrix and Backend_Settings inside its own apply, but
' Navy & Gold does not: StyleSearchMatrix and StyleBackendSettings are
' separate MANUAL macros its apply never calls. Without these columns,
' switching Tron -> Navy & Gold left the dashboard themed and the other
' two tabs bare, because Tron's restore had put them back and nothing
' re-styled them.
'
' A proc that is not installed is harmless: Application.Run simply fails
' and is swallowed, so a row can sit here indefinitely.
'=====================================================================
Private Function ThemeRegistry() As Variant
    ThemeRegistry = Array( _
        Array(THEME_NONE, "Default (no theme)", "", "", "", ""), _
        Array(THEME_TRON, "Dark Blue", "ApplyTronLegacy", "RemoveTronLegacy", "", ""), _
        Array(THEME_DESERT, "Desert", "ApplyDesert", "RemoveDesert", "", ""), _
        Array(THEME_NAVY, "Navy & Gold", "ApplyNavyGold", "RemoveNavyGold", _
              "StyleSearchMatrix,StyleBackendSettings", "RemoveSearchMatrixTheme") _
    )
End Function

' Runs a comma-separated list of procedure names, skipping any that are
' missing. Used for the companion-sheet stylers.
Private Sub RunList(ByVal procs As String)
    If Len(Trim$(procs)) = 0 Then Exit Sub
    Dim parts() As String, i As Long
    parts = Split(procs, ",")
    On Error Resume Next
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then Application.Run Trim$(parts(i))
        Err.Clear
    Next i
    On Error GoTo 0
End Sub

'---- setup ----------------------------------------------------------
' Run once. Also safe to re-run - refreshes the list after a theme is
' added to the registry.
Sub BuildThemePicker()
    Dim ws As Worksheet, reg As Variant, i As Long, list As String
    Dim wasProt As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox SHEET_NAME & " not found.", vbCritical: Exit Sub

    ' EVENTS OFF FIRST - this is not optional.
    '
    ' Writing to U28 below (ClearContents, then Value) raises
    ' Worksheet_Change on Sheet1. That handler unprotects, does its work,
    ' and RE-PROTECTS the sheet before returning - so without this guard
    ' the sheet is locked again by the time this routine reaches the label
    ' cell, and setting HorizontalAlignment fails with "Unable to set the
    ' HorizontalAlignment property of the Range class". Protection blocks
    ' formatting even on unlocked cells.
    '
    ' Once the picker stub is in Sheet1, the same write would also schedule
    ' a theme apply - so building the picker would silently re-theme the
    ' workbook. Off for the whole routine.
    Application.EnableEvents = False
    On Error GoTo CleanUp

    wasProt = ws.ProtectContents
    ws.Unprotect Password:=PWD
    ThisWorkbook.Unprotect Password:=PWD

    reg = ThemeRegistry()
    For i = LBound(reg) To UBound(reg)
        list = list & IIf(Len(list) > 0, ",", "") & reg(i)(1)
    Next i

    With ws.Range(PICKER_CELL)
        .ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=list
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
        ' Must stay unlocked or the dropdown is dead once Sheet1 is
        ' protected again - which every theme does on the way out.
        .Locked = False
        .HorizontalAlignment = xlCenter
        .Value = DisplayNameFor(ActiveTheme())
    End With

    If wasProt Then ws.Protect Password:=PWD
    Application.EnableEvents = True

    MsgBox "Theme picker ready in " & PICKER_CELL & "." & vbCrLf & vbCrLf & _
           "If picking a theme does nothing, the Worksheet_Change block is " & _
           "not in the Sheet1 code module yet - see " & _
           "Sheet1_CodeModule_UPDATED.txt.", vbInformation, "Theme Manager"
    Exit Sub

CleanUp:
    Dim d As String: d = Err.Description
    On Error Resume Next
    If wasProt Then ws.Protect Password:=PWD
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "BuildThemePicker failed: " & d, vbCritical, "Theme Manager"
End Sub

'---- the dropdown's entry point -------------------------------------
' Called by Worksheet_Change on Sheet1 when U28 is edited.
Public Sub OnThemePicked()
    If busy Then Exit Sub

    Dim ws As Worksheet, picked As String
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    picked = Trim$(CStr(ws.Range(PICKER_CELL).Value))
    If Len(picked) = 0 Then Exit Sub

    ApplyThemeByKey KeyForDisplayName(picked)
End Sub

'---- apply ----------------------------------------------------------
' The one road in. Clears whatever is active, then runs the requested
' theme's apply proc.
Public Sub ApplyThemeByKey(ByVal key As String)
    If busy Then Exit Sub
    busy = True
    quietMode = True
    Application.ScreenUpdating = False
    Application.StatusBar = "Applying theme..."

    Dim reg As Variant, i As Long, applyProc As String, extraApply As String
    reg = ThemeRegistry()

    ClearTheme True
    RestoreSheet1Baseline

    For i = LBound(reg) To UBound(reg)
        If reg(i)(0) = key Then
            applyProc = reg(i)(2)
            extraApply = reg(i)(4)
            Exit For
        End If
    Next i

    On Error Resume Next
    If Len(applyProc) > 0 Then
        Err.Clear
        Application.Run applyProc
        If Err.Number <> 0 Then
            busy = False
            quietMode = False
            Application.ScreenUpdating = True
            Application.StatusBar = False
            MsgBox "Could not apply that theme: " & Err.Description & vbCrLf & vbCrLf & _
                   "Its module may not be installed in this workbook.", _
                   vbExclamation, "Theme Manager"
            SetActiveTheme THEME_NONE
            SyncPicker THEME_NONE
            Exit Sub
        End If
    End If
    On Error GoTo 0

    ' Companion sheets. Runs after the main apply so a theme that styles
    ' them itself (Tron) is untouched, while one that does not (Navy &
    ' Gold) gets its manual stylers called for it.
    RunList extraApply

    ' Conventions that are NOT a theme's business, re-asserted last so
    ' they survive whatever the theme just did.
    NormaliseSheetConventions

    ' Sheet1's OSDD input cells (Module2) take the look of the table
    ' next to them in the theme just applied.
    RunList "StyleOsddInputs"

    SetActiveTheme key
    SyncPicker key

    quietMode = False
    Application.ScreenUpdating = True
    Application.StatusBar = False        ' hand the status bar back to Excel
    busy = False
End Sub

' Kept so existing buttons wired to these names keep working.
Sub ApplyTron():          ApplyThemeByKey THEME_TRON:  End Sub
Sub ApplyNavyGold_Safe(): ApplyThemeByKey THEME_NAVY:  End Sub

Sub RemoveActiveTheme()
    ApplyThemeByKey THEME_NONE
End Sub

'---- clear ----------------------------------------------------------
' Sweeps EVERY remove proc in the registry, not just the recorded one.
' A sweep is cheap - each Remove is a no-op when its own stash tag is
' absent - and it recovers a workbook whose state marker was lost, or one
' themed before this manager existed.
Public Function ClearTheme(ByVal quiet As Boolean) As Boolean
    Dim reg As Variant, i As Long
    ClearTheme = True
    reg = ThemeRegistry()

    On Error Resume Next
    ' The recorded theme first, so its own restore runs against the state
    ' it actually saved.
    Dim active As String
    active = ActiveTheme()
    For i = LBound(reg) To UBound(reg)
        If reg(i)(0) = active And Len(reg(i)(3)) > 0 Then
            Application.Run reg(i)(3)
            RunList reg(i)(5)
        End If
    Next i
    ' Then everything else.
    For i = LBound(reg) To UBound(reg)
        If reg(i)(0) <> active And Len(reg(i)(3)) > 0 Then
            Application.Run reg(i)(3)
            RunList reg(i)(5)
        End If
    Next i
    ' ...and themes that predate this manager and are not in the dropdown.
    RunList LEGACY_REMOVERS
    Err.Clear
    On Error GoTo 0

    UnhideDashboardShapes
    NormaliseSheetConventions         ' also true with no theme applied
    SetActiveTheme THEME_NONE
End Function

'=====================================================================
' SHEET CONVENTIONS - things that are true regardless of theme.
'
' Alignment on the companion sheets is a property of the SHEET, not a
' styling choice. Every theme sets its own - Navy & Gold's
' StyleSearchMatrix left-aligns the entity column and indents it, and
' Tron did the same - so fixing it inside one theme would only hold until
' you switched. Re-asserting it here, after the theme and its companion
' stylers have finished, makes it true under every theme and under no
' theme, and means a future theme cannot quietly undo it either.
'
' Deliberately narrow: alignment only. No colours, no fills, nothing that
' belongs to whichever theme is active.
'=====================================================================
Private Sub NormaliseSheetConventions()
    NormaliseBackendSettings
    NormaliseSearchMatrix
End Sub

Private Sub NormaliseBackendSettings()
    On Error Resume Next
    Dim ws As Worksheet, wasProt As Boolean
    Set ws = ThisWorkbook.Sheets("Backend_Settings")
    If ws Is Nothing Then Exit Sub

    wasProt = ws.ProtectContents
    ws.Unprotect Password:=PWD
    CentreMiddle ws.Range("A1:B4")
    If wasProt Then ws.Protect Password:=PWD
    On Error GoTo 0
End Sub

Private Sub NormaliseSearchMatrix()
    On Error Resume Next
    Dim ws As Worksheet, wasProt As Boolean, lastRow As Long, lastA As Long, lastB As Long
    Set ws = ThisWorkbook.Sheets("Search Matrix")
    If ws Is Nothing Then Exit Sub

    wasProt = ws.ProtectContents
    ws.Unprotect Password:=PWD

    ' Take the deeper of A and B. The stylers key off B alone, which
    ' misses any block whose Type-of-Search cell happens to be blank.
    lastA = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    lastB = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    lastRow = IIf(lastA > lastB, lastA, lastB)
    If lastRow < 1 Then lastRow = 1

    CentreMiddle ws.Range("A1:E" & lastRow)
    If wasProt Then ws.Protect Password:=PWD
    On Error GoTo 0
End Sub

' IndentLevel is cleared FIRST. Navy & Gold indents the entity column,
' and Excel only honours an indent with left/right alignment - leaving a
' stale indent set while switching to centre can make the write fail
' outright rather than simply being ignored.
Private Sub CentreMiddle(ByVal rng As Range)
    On Error Resume Next
    With rng
        .IndentLevel = 0
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo 0
End Sub

' Repairs a shape that a theme hid and never un-hid.
'
' modNavyGold does not restyle the Beta title - RepositionBeta hides the
' workbook's real title shape and AddTitleBeta draws a gold replacement as
' NGADD_* shapes. RemoveNavyGold then deletes its replacement without ever
' setting the original back to visible, so removing Navy & Gold leaves the
' nav bar with no title at all. That is a bug in the theme module, but the
' manager is what promises a clean baseline, so the repair belongs here
' too rather than only inside one theme.
Private Sub UnhideDashboardShapes()
    On Error Resume Next
    Dim ws As Worksheet, shp As Shape
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    For Each shp In ws.Shapes
        If Left$(shp.Name, 8) <> "TRONADD_" And Left$(shp.Name, 6) <> "NGADD_" Then
            If shp.Visible = msoFalse Then shp.Visible = msoTrue
        End If
    Next shp
    On Error GoTo 0
End Sub

'---- registry lookups -----------------------------------------------
Private Function DisplayNameFor(ByVal key As String) As String
    Dim reg As Variant, i As Long
    reg = ThemeRegistry()
    DisplayNameFor = reg(LBound(reg))(1)          ' default to the NONE row
    For i = LBound(reg) To UBound(reg)
        If reg(i)(0) = key Then DisplayNameFor = reg(i)(1): Exit For
    Next i
End Function

Private Function KeyForDisplayName(ByVal nm As String) As String
    Dim reg As Variant, i As Long
    reg = ThemeRegistry()
    KeyForDisplayName = THEME_NONE
    For i = LBound(reg) To UBound(reg)
        If StrComp(reg(i)(1), nm, vbTextCompare) = 0 Then
            KeyForDisplayName = reg(i)(0): Exit For
        End If
    Next i
End Function

' Writes the picker back without re-triggering Worksheet_Change.
Private Sub SyncPicker(ByVal key As String)
    On Error Resume Next
    Dim ws As Worksheet, wasProt As Boolean
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    Application.EnableEvents = False
    wasProt = ws.ProtectContents
    ws.Unprotect Password:=PWD
    ws.Range(PICKER_CELL).Value = DisplayNameFor(key)
    ws.Range(PICKER_CELL).Locked = False
    If wasProt Then ws.Protect Password:=PWD
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

'---- state ----------------------------------------------------------
' Stored as a hidden workbook-level defined name: it survives save/close,
' never appears on a sheet, and needs no extra backup tab.
Public Function ActiveTheme() As String
    Dim v As String
    ActiveTheme = THEME_NONE
    On Error Resume Next
    v = ThisWorkbook.Names(STATE_NAME).RefersTo      ' arrives as ="TRON"
    On Error GoTo 0
    v = Replace(Replace(Replace(v, "=", ""), Chr$(34), ""), " ", "")
    If Len(v) > 0 Then ActiveTheme = UCase$(v)
End Function

Public Sub SetActiveTheme(ByVal v As String)
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PWD
    ThisWorkbook.Names(STATE_NAME).Delete
    ThisWorkbook.Names.Add Name:=STATE_NAME, RefersTo:="=""" & v & """", Visible:=False
    On Error GoTo 0
End Sub

'---- theme chooser --------------------------------------------------
' Ctrl+Shift+T (set up in ThisWorkbook's Workbook_Open). There is no theme
' control on the sheet: the picker cell is hidden, so only someone who
' knows the shortcut changes themes.
Public Sub ChooseTheme()
    Dim reg As Variant, i As Long, msg As String, pick As String, n As Long
    reg = ThemeRegistry()
    msg = "Type the number of the theme:" & vbCrLf
    For i = LBound(reg) To UBound(reg)
        msg = msg & vbCrLf & (i - LBound(reg)) & "  =  " & reg(i)(1) & _
            IIf(reg(i)(0) = ActiveTheme(), "   (current)", "")
    Next i
    pick = InputBox(msg, "Theme")
    If StrPtr(pick) = 0 Or Not IsNumeric(Trim$(pick)) Then Exit Sub
    n = CLng(Trim$(pick)) + LBound(reg)
    If n < LBound(reg) Or n > UBound(reg) Then Exit Sub
    ApplyThemeByKey CStr(reg(n)(0))
End Sub

'---- Sheet1 baseline ------------------------------------------------
' Every theme starts from the same Sheet1: the sidebar's buttons and
' panel labels back where Beta 3.6.1 has them, and the OSDD panel's
' banners copied fresh from the plain originals. Run between ClearTheme
' and the next theme's apply. Without it the sidebar drifted from theme
' to theme until the utility buttons sat stacked on one spot: Navy &
' Gold puts positions back by shape name, and Excel renames a shape each
' time a theme changes its shape type, so its restore found nothing.
Private Sub RestoreSheet1Baseline()
    Dim ws As Worksheet, wasProt As Boolean, shp As Shape, spec As Variant, s As Variant
    Dim key As String, k As Long, stack As Variant, top0 As Single, foot As Single, gap As Single
    Const BTN_H As Single = 28.6

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    wasProt = ws.ProtectContents
    ws.Unprotect Password:=PWD

    ' Sheet1 active and all of it on screen first. Clearing a theme visits
    ' the other sheets, and until Sheet1 is laid out again Excel holds a
    ' shape short of where it is put - which is how the utility buttons
    ' ended up stacked on one spot.
    Application.ScreenUpdating = True
    Application.Run "FitSheet1ToWindow"
    DoEvents

    ' The panel labels and the action buttons where Beta 3.6.1 has them:
    ' macro (or caption), column, offset, row, offset, width, height.
    spec = Array( _
        Array("ACTION PANEL", 1, 18.9, 4, 23.7, 161.6, 24.9), _
        Array("Start_Button_Create_Folders", 2, 3.9, 7, 12.3, 135.5, 28.8), _
        Array("Consolidated_AML_Workflow", 2, 3.8, 9, 16.9, 135.5, 28.8), _
        Array("SearchAndSavePDF_Direct", 2, 2.6, 12, 8.5, 135.5, 28.8), _
        Array("ExportToWord", 2, 2.4, 14, 1.9, 135.5, 28.9), _
        Array("UTILITY PANEL", 1, 18.3, 17, 10.6, 161.6, 25.2))
    ' The utility buttons, top to bottom, share out the space between the
    ' UTILITY label and the foot of the sheet evenly - so a hidden row
    ' (23) narrows the gaps rather than squashing a button.
    stack = Array("Run_Mass_Rename_v4", "Trigger_PAD_Merge_Flow", "WarmUpWorkerProfiles", "ClearForm")
    top0 = ws.Rows(17).Top + 10.6 + 25.2
    foot = ws.Rows(28).Top + ws.Rows(28).Height
    gap = (foot - top0 - 4 * BTN_H) / 5

    For Each shp In ws.Shapes
        key = shp.OnAction
        If InStr(key, "!") > 0 Then key = Mid$(key, InStrRev(key, "!") + 1)
        If key = "" Then key = UCase$(Trim$(Replace(Replace(shp.TextFrame2.TextRange.Text, vbCr, ""), vbLf, "")))
        For Each s In spec
            If key = s(0) Then
                PutShape shp, ws.Columns(s(1)).Left + s(2), ws.Rows(s(3)).Top + s(4), s(5), s(6)
                Exit For
            End If
        Next s
        For k = 0 To UBound(stack)
            If key = stack(k) Then
                PutShape shp, ws.Columns(2).Left + 2, top0 + gap + k * (BTN_H + gap), 135.4, BTN_H
                Exit For
            End If
        Next k
    Next shp

    ' With the OSDD inputs beside the tables (columns V:AB): Alert Related
    ' Information and Country Risk Rating run across to AB over their
    ' widened rows; Customer Information and Counterparty Information keep
    ' to G:U, each with a matching copy beside it over the inputs -
    ' "Customer OSDD Inputs" and "Counterparty OSDD Inputs". Copies from
    ' earlier layouts go.
    ws.Shapes("Rectangle: Single Corner Snipped OSDD").Delete
    ws.Shapes("Rectangle: Single Corner Snipped Theme").Delete
    If Left$(CStr(ws.Range("W13").Value), 13) = "Negative News" Then
        StretchBanner ws, "Alert Related Information", "AB"
        StretchBanner ws, "Country Risk Rating", "AB"
        MatchBanner ws, "Counterparty Information", "Customer Information"
        PlaceBanner ws, "Customer Information", "OSDD Banner Customer", "Customer OSDD Inputs", "V", "AB"
        PlaceBanner ws, "Counterparty Information", "OSDD Banner Counterparty", "Counterparty OSDD Inputs", "V", "AB"
    Else
        ws.Shapes("OSDD Banner Customer").Delete
        ws.Shapes("OSDD Banner Counterparty").Delete
    End If
    Application.ScreenUpdating = False

    If wasProt Then ws.Protect Password:=PWD
    On Error GoTo 0
End Sub

' Puts a sidebar shape at x, y in w x h and checks it landed there. It
' moves with its cells but keeps its size, so hiding a row cannot squash
' it.
Private Sub PutShape(ByVal shp As Shape, ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single)
    Dim k As Long
    On Error Resume Next
    shp.Placement = xlMove
    For k = 1 To 3
        shp.Width = w
        shp.Height = h
        shp.Left = x
        shp.Top = y
        If Abs(shp.Left - x) < 0.5 And Abs(shp.Top - y) < 0.5 Then Exit For
        DoEvents
    Next k
End Sub

' Widens the sheet's banner saying txt so its right end sits at the right
' of column lastCol, as far in from the edge as its left end is from G.
Private Sub StretchBanner(ByVal ws As Worksheet, ByVal txt As String, ByVal lastCol As String)
    Dim shp As Shape, inset As Single, W As Single, k As Long
    On Error Resume Next
    Set shp = PlainBanner(ws, txt)
    If shp Is Nothing Then Exit Sub
    inset = shp.Left - ws.Range("G1").Left
    If inset < 0 Then inset = 0
    W = ws.Range(lastCol & "1").Left + ws.Range(lastCol & "1").Width - inset - shp.Left
    For k = 1 To 3
        shp.Width = W
        If Abs(shp.Width - W) < 1 Then Exit For
        DoEvents
    Next k
End Sub

' Gives the banner saying txt the same left and width as the one saying
' refTxt (Counterparty Information back to Customer Information's G:U,
' after an earlier layout stretched it).
Private Sub MatchBanner(ByVal ws As Worksheet, ByVal txt As String, ByVal refTxt As String)
    Dim shp As Shape, ref As Shape
    On Error Resume Next
    Set shp = PlainBanner(ws, txt)
    Set ref = PlainBanner(ws, refTxt)
    If shp Is Nothing Or ref Is Nothing Then Exit Sub
    shp.Left = ref.Left
    shp.Width = ref.Width
End Sub

' A copy of the banner shape saying srcText, named nm and saying txt,
' over columns firstCol:lastCol with the original's top, height and side
' inset. Narrowed before it moves and checked after, in case Excel holds
' it short of where it was put.
Private Sub PlaceBanner(ByVal ws As Worksheet, ByVal srcText As String, ByVal nm As String, _
    ByVal txt As String, ByVal firstCol As String, ByVal lastCol As String)
    Dim src As Shape, shp As Shape, inset As Single, L As Single, W As Single, k As Long

    On Error Resume Next
    ws.Shapes(nm).Delete
    Set src = PlainBanner(ws, srcText)
    If src Is Nothing Then Exit Sub
    inset = src.Left - ws.Range("G1").Left
    If inset < 0 Then inset = 0
    L = ws.Range(firstCol & "1").Left + inset
    W = ws.Range(lastCol & "1").Left + ws.Range(lastCol & "1").Width - inset - L

    Set shp = src.Duplicate
    shp.Name = nm
    shp.AlternativeText = ""
    shp.TextFrame2.TextRange.Text = txt
    shp.Placement = xlMove
    For k = 1 To 3
        shp.Width = 10
        shp.Left = L
        shp.Top = src.Top
        shp.Height = src.Height
        shp.Width = W
        If Abs(shp.Left - L) < 1 And Abs(shp.Width - W) < 1 Then Exit For
        DoEvents
    Next k
End Sub

' The sheet's own banner saying txt - not a copy, not a theme's decoration.
Private Function PlainBanner(ByVal ws As Worksheet, ByVal txt As String) As Shape
    Dim shp As Shape, nm As String
    On Error Resume Next
    For Each shp In ws.Shapes
        nm = shp.Name
        If Left$(nm, 12) <> "OSDD Banner " And nm <> "Rectangle: Single Corner Snipped OSDD" _
           And nm <> "Rectangle: Single Corner Snipped Theme" _
           And Left$(nm, 8) <> "TRONADD_" And Left$(nm, 6) <> "NGADD_" _
           And Left$(nm, 10) <> "DESERTADD_" And Left$(nm, 8) <> "DUNEADD_" Then
            If InStr(1, shp.TextFrame2.TextRange.Text, txt, vbTextCompare) > 0 Then
                Set PlainBanner = shp
                Exit Function
            End If
        End If
    Next shp
End Function
