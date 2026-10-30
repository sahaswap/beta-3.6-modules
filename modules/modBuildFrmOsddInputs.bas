Attribute VB_Name = "modBuildFrmOsddInputs"
'==================================================================
' modBuildFrmOsddInputs - ONE-TIME setup macro. Run BuildFrmOsddInputs
' once to build the frmOsddInputs UserForm: OSDD Search's questions on
' one screen (Negative News name, additional addresses, telephone /
' email per entity, and other parties), instead of a prompt each.
' Module2 uses the form when it is there and falls back to the prompts
' when it isn't.
'
' REQUIRES: File > Options > Trust Center > Trust Center Settings >
' Macro Settings > "Trust access to the VBA project object model"
' (the same setting BuildFrmExportMode needed).
'
' Safe to re-run: clears and rebuilds frmOsddInputs from scratch.
' Delete this module once frmOsddInputs is built.
'==================================================================
Option Explicit

Sub BuildFrmOsddInputs()
    Dim vbProj As Object, vbComp As Object, frm As Object, c As Object, i As Long

    On Error GoTo TrustFail
    Set vbProj = ThisWorkbook.VBProject
    i = vbProj.VBComponents.count
    On Error GoTo 0

    On Error Resume Next
    Set vbComp = vbProj.VBComponents("frmOsddInputs")
    On Error GoTo 0
    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(3)      ' vbext_ct_MSForm
        vbComp.Name = "frmOsddInputs"
    End If

    ' Touch .Designer first - a new component's Properties aren't there
    ' until its Designer has been loaded once.
    Set frm = vbComp.Designer
    vbComp.Properties("Width").Value = 870
    vbComp.Properties("Height").Value = 376
    vbComp.Properties("Caption").Value = "OSDD Search - Inputs"

    For i = frm.Controls.count - 1 To 0 Step -1
        frm.Controls.Remove frm.Controls(i).Name
    Next i

    Set c = frm.Controls.Add("Forms.Label.1", "lblIntro")
    c.Left = 8: c.Top = 6: c.Width = 846: c.Height = 28
    c.WordWrap = True
    c.Caption = "Check each row, then click Run Searches. To put several entries in one box, separate them " & _
        "with a semicolon ( ; ). Negative News names are pre-filled without legal extensions, middle names " & _
        "or suffixes, as the procedure asks - edit them if needed, or clear one to skip that Negative News search."

    AddHeader frm, "lblHdrEntity", 14, 170, "Entity"
    AddHeader frm, "lblHdrNeg", 188, 196, "Negative News name"
    AddHeader frm, "lblHdrAddr", 388, 266, "Additional address identified? (tick) - address(es)"
    AddHeader frm, "lblHdrContact", 664, 180, "Telephone / email"

    Set c = frm.Controls.Add("Forms.Frame.1", "FrameRows")
    c.Left = 8: c.Top = 56: c.Width = 852: c.Height = 196
    c.Caption = ""
    c.ScrollBars = 2                                 ' fmScrollBarsVertical
    c.Font.Name = "Segoe UI": c.Font.Size = 9

    Set c = frm.Controls.Add("Forms.Label.1", "lblOther")
    c.Left = 8: c.Top = 260: c.Width = 846: c.Height = 16
    c.Caption = "Other individuals or entities named in the narrative and relied upon (e.g. beneficial owners, " & _
        "directors) - each gets a Name and a Negative News search:"

    Set c = frm.Controls.Add("Forms.TextBox.1", "txtOther")
    c.Left = 8: c.Top = 278: c.Width = 852: c.Height = 18
    c.Font.Name = "Segoe UI": c.Font.Size = 9

    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnRun")
    c.Caption = "Run Searches": c.Left = 650: c.Top = 308: c.Width = 110: c.Height = 28
    ' Not the Default button: Enter while typing in a box shouldn't start the run.

    Set c = frm.Controls.Add("Forms.CommandButton.1", "btnCancel")
    c.Caption = "Cancel": c.Left = 768: c.Top = 308: c.Width = 92: c.Height = 28
    c.Cancel = True

    With vbComp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString FormCode()
    End With

    MsgBox "frmOsddInputs built." & vbCrLf & vbCrLf & _
        "OSDD Search now asks its questions on one screen. You can delete this " & _
        "modBuildFrmOsddInputs module - it's done its job.", vbInformation, "Build Complete"
    Exit Sub

TrustFail:
    MsgBox "Could not access the VBA project object model." & vbCrLf & vbCrLf & _
        "Enable 'Trust access to the VBA project object model' in:" & vbCrLf & _
        "File > Options > Trust Center > Trust Center Settings > Macro Settings," & vbCrLf & _
        "then run BuildFrmOsddInputs again." & vbCrLf & vbCrLf & _
        "Until then, OSDD Search keeps asking with its prompts.", vbCritical, "Trust Setting Required"
End Sub

Private Sub AddHeader(ByVal frm As Object, ByVal nm As String, ByVal x As Single, ByVal w As Single, _
    ByVal cap As String)
    Dim c As Object
    Set c = frm.Controls.Add("Forms.Label.1", nm)
    c.Left = x: c.Top = 40: c.Width = w: c.Height = 14
    c.Caption = cap
    c.Font.Bold = True
End Sub

' The form's code-behind.
Private Function FormCode() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "' frmOsddInputs - OSDD Search's inputs on one screen (Module2)" & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' One row per entity (customer, counterparties): the Negative News name" & vbCrLf
    s = s & "' (pre-filled without legal extension / middle names / suffixes, per" & vbCrLf
    s = s & "' the Operating Manual), whether an additional address was identified" & vbCrLf
    s = s & "' and which (pre-filled with addresses found in the transaction data)," & vbCrLf
    s = s & "' and telephone numbers / email addresses. Below them, the other parties" & vbCrLf
    s = s & "' named in the narrative. Rows are added at run time by AddRow, the same" & vbCrLf
    s = s & "' way the Rename form (UserForm2) builds its rows." & vbCrLf
    s = s & "'" & vbCrLf
    s = s & "' Built by BuildFrmOsddInputs (modBuildFrmOsddInputs)." & vbCrLf
    s = s & "' ============================================================" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Cancelled As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private mRows As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const ROW_H As Long = 24" & vbCrLf
    s = s & "Private Const X_ENTITY As Long = 6" & vbCrLf
    s = s & "Private Const W_ENTITY As Long = 170" & vbCrLf
    s = s & "Private Const X_NEG As Long = 180" & vbCrLf
    s = s & "Private Const W_NEG As Long = 196" & vbCrLf
    s = s & "Private Const X_CHK As Long = 382" & vbCrLf
    s = s & "Private Const X_ADDR As Long = 400" & vbCrLf
    s = s & "Private Const W_ADDR As Long = 250" & vbCrLf
    s = s & "Private Const X_CONTACT As Long = 656" & vbCrLf
    s = s & "Private Const W_CONTACT As Long = 176" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' One entity row. key names its controls (""C"" = customer, ""18""..""23"" =" & vbCrLf
    s = s & "' counterparty rows of Sheet1); the address box keeps what it was" & vbCrLf
    s = s & "' pre-filled with in its Tag (see RowAddresses)." & vbCrLf
    s = s & "Public Sub AddRow(ByVal key As String, ByVal entityLabel As String, ByVal negName As String, _" & vbCrLf
    s = s & "    ByVal txAddrs As String)" & vbCrLf
    s = s & "    Dim y As Long, c As Object" & vbCrLf
    s = s & "    y = 4 + mRows * ROW_H" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = FrameRows.Controls.Add(""Forms.Label.1"", ""lbl_"" & key)" & vbCrLf
    s = s & "    c.Caption = entityLabel" & vbCrLf
    s = s & "    c.Left = X_ENTITY: c.Top = y + 3: c.Width = W_ENTITY: c.Height = 16" & vbCrLf
    s = s & "    c.WordWrap = False" & vbCrLf
    s = s & "    c.ControlTipText = entityLabel" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = FrameRows.Controls.Add(""Forms.TextBox.1"", ""txtNeg_"" & key)" & vbCrLf
    s = s & "    c.Text = negName" & vbCrLf
    s = s & "    c.Left = X_NEG: c.Top = y: c.Width = W_NEG: c.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = FrameRows.Controls.Add(""Forms.CheckBox.1"", ""chkAddr_"" & key)" & vbCrLf
    s = s & "    c.Caption = """"" & vbCrLf
    s = s & "    c.Left = X_CHK: c.Top = y + 1: c.Width = 16: c.Height = 16" & vbCrLf
    s = s & "    c.ControlTipText = ""Tick if an additional address was identified""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = FrameRows.Controls.Add(""Forms.TextBox.1"", ""txtAddr_"" & key)" & vbCrLf
    s = s & "    c.Text = txAddrs" & vbCrLf
    s = s & "    c.Tag = txAddrs" & vbCrLf
    s = s & "    c.Left = X_ADDR: c.Top = y: c.Width = W_ADDR: c.Height = 18" & vbCrLf
    s = s & "    If txAddrs <> """" Then c.ControlTipText = ""Found in the transaction data: "" & txAddrs" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Set c = FrameRows.Controls.Add(""Forms.TextBox.1"", ""txtContact_"" & key)" & vbCrLf
    s = s & "    c.Left = X_CONTACT: c.Top = y: c.Width = W_CONTACT: c.Height = 18" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    mRows = mRows + 1" & vbCrLf
    s = s & "    FrameRows.ScrollHeight = 8 + mRows * ROW_H" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' A row's text box (""Neg"" or ""Contact""); """" if there is no such row." & vbCrLf
    s = s & "Public Function RowText(ByVal key As String, ByVal what As String) As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    RowText = FrameRows.Controls(""txt"" & what & ""_"" & key).Text" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "' A row's additional addresses: the box's text when the tick box is" & vbCrLf
    s = s & "' ticked, or when the analyst typed in the box (it differs from what it" & vbCrLf
    s = s & "' was pre-filled with); otherwise none." & vbCrLf
    s = s & "Public Function RowAddresses(ByVal key As String) As String" & vbCrLf
    s = s & "    Dim t As Object, chk As Object" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Set t = FrameRows.Controls(""txtAddr_"" & key)" & vbCrLf
    s = s & "    Set chk = FrameRows.Controls(""chkAddr_"" & key)" & vbCrLf
    s = s & "    If t Is Nothing Then Exit Function" & vbCrLf
    s = s & "    If chk.Value = True Or t.Text <> t.Tag Then RowAddresses = t.Text" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Property Get OtherPartiesText() As String" & vbCrLf
    s = s & "    OtherPartiesText = txtOther.Text" & vbCrLf
    s = s & "End Property" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnRun_Click()" & vbCrLf
    s = s & "    Cancelled = False" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    Cancelled = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        Cancelled = True" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCode = s
End Function
