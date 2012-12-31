VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAPL006 
   Caption         =   "AR Update"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmAPL006.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2160
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   560
      Width           =   4665
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   8760
      OleObjectBlob   =   "frmAPL006.frx":030A
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   4980
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   915
      Width           =   1815
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   915
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   2040
      Width           =   6660
      Begin VB.OptionButton optBy 
         Caption         =   "Key In Exchange Rate"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton optBy 
         Caption         =   "Default"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL006.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   3255
      Left            =   840
      OleObjectBlob   =   "frmAPL006.frx":66AD
      TabIndex        =   6
      Top             =   2760
      Width           =   7455
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   840
      TabIndex        =   13
      Top             =   615
      Width           =   1260
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   4620
      TabIndex        =   11
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "Customer"
      Height          =   225
      Left            =   840
      TabIndex        =   10
      Top             =   960
      Width           =   1290
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   840
      TabIndex        =   7
      Top             =   1365
      Width           =   1290
   End
End
Attribute VB_Name = "frmAPL006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Dim wcCombo As Control
Dim wsMsg As String

Private waScrToolTip As New XArrayDB
Private waResult As New XArrayDB
Dim wgsTitle As String
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private Const Curr = 0
Private Const CurrDesc = 1
Private Const Excr = 2
Private Const Dummy = 3
'Maximum Exchange Rate
Private Const wdExchangeRate = 9999.999999

Private wsBaseCurCd As String
Private wsCtlPrd As String

Private Sub cmdCancel()
    Ini_Scr
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsDate As String
    Dim wsSQL As String
    Dim wsMapPrd As String
    Dim adcmdSave As New ADODB.Command
    Dim wlCtr As Long
    
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    

On Error GoTo cmdSave_Err

    wsDteTim = gsSystemDate
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'If MsgBox("Are you sure to update exchange rate Now?", vbYesNo, gsTitle) = vbNo Then
    '        cmdCancel
    '        MousePointer = vbDefault
    '        Exit Sub
    'End If
    
    wsDate = medPrdFr.Text
    wsMapPrd = Get_FiscalPeriod(wsDate)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_CRTTMPEXCR"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    For wlCtr = 0 To waResult.UpperBound(1)
        If Trim(waResult(wlCtr, Curr)) <> "" Then
            Call SetSPPara(adcmdSave, 1, gsUserID)
            Call SetSPPara(adcmdSave, 2, Change_SQLDate(wsDteTim))
            Call SetSPPara(adcmdSave, 3, wsFormID)
            Call SetSPPara(adcmdSave, 4, Set_Quote(waResult(wlCtr, Curr)))
            Call SetSPPara(adcmdSave, 5, To_Value(waResult(wlCtr, Excr)))
            adcmdSave.Execute
        End If
    Next wlCtr
    
    cnCon.CommitTrans 'Create Stored Procedure String
    Set adcmdSave = Nothing
    
    'Create Selection Criteria
    ReDim wsSelection(2)
    wsSelection(1) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo.Text)
    wsSelection(2) = lblPrdFr.Caption & " " & Set_Quote(medPrdFr)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTAPL006 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(15, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSQL = wsSQL & "'', "
    wsSQL = wsSQL & "'" & Set_Quote(medPrdFr) & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTAPL006"
    Else
    wsRptName = "RPTAPL006"
    End If
    
    NewfrmPrint.ReportID = "APL006"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "APL006"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    
    Me.MousePointer = vbDefault
    
    'gsMsg = "Update Process is completed!"
    'MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        
    Call cmdCancel
    
    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
End Sub

Private Sub cboCusNoFr_LostFocus()
    FocusMe cboCusNoFr, True
End Sub

Private Sub cboCusNoTo_LostFocus()
    FocusMe cboCusNoTo, True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF9
            Call cmdOK
            
        Case vbKeyF11
            Call cmdCancel
        
        Case vbKeyF12
            Unload Me
    End Select
End Sub

Private Sub Form_Load()

    MousePointer = vbHourglass
    
    Call Ini_Form
    Call Ini_Caption
    Call Ini_Grid
    Call Ini_Scr

    MousePointer = vbDefault

End Sub

Private Sub Ini_Form()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsFormID = "APL006"
    
End Sub

Private Sub Ini_Scr()
    Dim wsFromDate As String

    waResult.ReDim 0, -1, Curr, Excr
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    Me.Caption = wsFormCaption
   
    'wsCtlPrd = getCtrlMth("AP")
    wsCtlPrd = getCtrlMth("AP")
    wsFromDate = Left(wsCtlPrd, 4) & "/" & Right(wsCtlPrd, 2) & "/01"
    
    tblCommon.Visible = False
    cboCusNoFr.Text = ""
    cboCusNoTo.Text = ""
   
    Call SetDateMask(medPrdFr)
   
    medPrdFr.Text = Format(DateAdd("D", -1, DateAdd("M", 1, CDate(wsFromDate))), "YYYY/MM/DD")
   
    optBy(0).Value = True
    tblDetail.Enabled = False
   
    Call LoadRecord
    
    'cboCusNoFr.SetFocus
End Sub

Private Function InputValidation() As Boolean
    InputValidation = False
    
    With tblDetail
        If .EditActive = True Then Exit Function
        
        .Update
        If Chk_GrdRow(To_Value(.FirstRow) + .Row) = False Then
            .SetFocus
            Exit Function
        End If
    End With
    
    If chk_medPrdFr = False Then
        Exit Function
    End If

    InputValidation = True
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 6600
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmAPL006 = Nothing
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    optBy(0).Caption = Get_Caption(waScrItm, "OPT1")
    optBy(1).Caption = Get_Caption(waScrItm, "OPT2")
    
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    
    With tblDetail
        .Columns(Curr).Caption = Get_Caption(waScrItm, "CURR")
        .Columns(CurrDesc).Caption = Get_Caption(waScrItm, "CURRDESC")
        .Columns(Excr).Caption = Get_Caption(waScrItm, "EXCR")
    End With
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
End Sub

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr.Text) = "/  /" Then
       gsMsg = "Must Input Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr.SetFocus
       Exit Function
    End If
    
    If Chk_Date(medPrdFr) = False Then
       gsMsg = "Invalid Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr.SetFocus
       Exit Function
    End If
    
    chk_medPrdFr = True
End Function

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        Call Opt_Setfocus(optBy, 2, 0)
        
    End If
End Sub

Private Sub optBy_Click(Index As Integer)
    If Index = 0 Then
        Call LoadRecord
        tblDetail.Enabled = False
    Else
        tblDetail.Enabled = True
    End If
End Sub

Private Sub optBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Index = 1 Then
        tblDetail.SetFocus
        End If
    End If
End Sub

Private Sub optBy_LostFocus(Index As Integer)
    tblDetail.Enabled = IIf(Opt_Getfocus(optBy, 2, 0) = 0, False, True)
End Sub

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo tblDetail_AfterColUpdate_Err
        With tblDetail
            .Update
        End With
    Exit Sub
    
tblDetail_AfterColUpdate_Err:
    MsgBox "tblDetail_AfterColUpdate_Err!"
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    With tblDetail
        Select Case ColIndex
            Case Excr
                If chk_grdExchRate(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(ColIndex).Text = NBRnd(.Columns(ColIndex).Text, giExrDp)
        End Select
    End With
    
    Exit Sub
    
Tbl_BeforeColUpdate_Err:
    tblDetail.Columns(ColIndex).Text = OldValue
    Cancel = True
    Exit Sub
    
tblDetail_BeforeColUpdate_Err:
    MsgBox "tblDetail_BeforeColUpdate_Err!"
    tblDetail.Columns(ColIndex).Text = OldValue
    Cancel = True

End Sub

Private Sub tblDetail_BeforeRowColChange(Cancel As Integer)

    On Error GoTo tblDetail_BeforeRowColChange_Err
    
    With tblDetail
        If .Bookmark <> .DestinationRow Then
            If Chk_GrdRow(To_Value(.Bookmark)) = False Then
                Cancel = True
                Exit Sub
            End If
        End If
    End With
    
    Exit Sub

tblDetail_BeforeRowColChange_Err:
    MsgBox "tblDetail_BeforeRowColChange_Err!"
    Cancel = True

End Sub

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo tblDetail_KeyDown_Err
    
        With tblDetail
            Select Case KeyCode
                Case vbKeyReturn
                    Select Case .Col
                        Case Excr
                            KeyCode = vbKeyDown
                            .Col = Excr
                    End Select
            End Select
        End With
    
    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "tblDetail_KeyDown_Err!"
End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)

    On Error GoTo tblDetail_KeyPress_Err

        With tblDetail
            Select Case .Col
                Case Excr
                    Call Chk_InpNum(KeyAscii, .Text, False, True)
                    Call chk_InpLen(tblDetail, 11, KeyAscii)
            End Select
        End With
        
    Exit Sub
    
tblDetail_KeyPress_Err:
    MsgBox "tblDetail_KeyPress_Err!"

End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case tcGo
            Call cmdOK
        Case tcCancel
                Call cmdCancel
        Case tcExit
            Unload Me
    End Select
End Sub

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = Curr To Excr
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case Curr
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).DataWidth = 3
                Case CurrDesc
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Locked = True
                Case Excr
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 11
                    .Columns(wiCtr).NumberFormat = gsExrFmt
                
                
            End Select
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wsCtlYr As String
    Dim wsCtlMon As String
          
    LoadRecord = False
    
    wsCtlYr = Left(wsCtlPrd, 4)
    wsCtlMon = Right(wsCtlPrd, 2)
    
    
    wsSQL = "SELECT EXCCURR, EXCDESC, EXCRATE FROM mstEXCHANGERATE "
    wsSQL = wsSQL & "WHERE EXCYR = '" & Set_Quote(wsCtlYr) & "' "
    wsSQL = wsSQL & "AND EXCMN = '" & To_Value(wsCtlMon) & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, Curr, Excr
         Do While Not rsRcd.EOF
             .AppendRows
             waResult(.UpperBound(1), Curr) = ReadRs(rsRcd, "EXCCURR")
             waResult(.UpperBound(1), CurrDesc) = ReadRs(rsRcd, "EXCDESC")
             waResult(.UpperBound(1), Excr) = Format(ReadRs(rsRcd, "EXCRATE"), gsExrFmt)
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
   
    LoadRecord = True
    
End Function

Private Function Chk_GrdRow(ByVal LastRow As Long) As Boolean

    On Error GoTo Chk_GrdRow_Err

    Chk_GrdRow = False
    
        If waResult.UpperBound(1) < 0 Then
            Chk_GrdRow = True
            Exit Function
        End If
    
        With tblDetail
            If To_Value(LastRow) > waResult.UpperBound(1) Then
                Chk_GrdRow = True
                Exit Function
            End If
            
            If chk_grdExchRate(waResult(LastRow, Excr)) = False Then
                .Col = Excr
                .Row = LastRow
                Exit Function
            End If
        End With
    
    Chk_GrdRow = True
    
    Exit Function

Chk_GrdRow_Err:
    MsgBox "Chk_GrdRow_Err!"
    
End Function

Private Function chk_grdExchRate(ByVal inExchRate As String) As Boolean

    chk_grdExchRate = False
    
        If To_Value(inExchRate) = 0 Then
            gsMsg = "Exchange Rate Can not equal to 0!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
    
        If To_Value(inExchRate) > wdExchangeRate Then
            gsMsg = "Exchange Rate too great!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
    
    chk_grdExchRate = True
    
End Function

Private Sub cboCusNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusNoFr
  
    wsSQL = "SELECT VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND VdrStatus ='1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VdrCode "
    Call Ini_Combo(2, wsSQL, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboCusNoFr_GotFocus()
    FocusMe cboCusNoFr
End Sub

Private Sub cboCusNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoFr, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboCusNoFr.Text) <> "" And _
            Trim(cboCusNoTo.Text) = "" Then
            cboCusNoTo.Text = cboCusNoFr.Text
        End If
        
        cboCusNoTo.SetFocus
    End If
End Sub

Private Sub cboCusNoTo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusNoTo
  
    wsSQL = "SELECT VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND VdrStatus ='1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VdrCode "
    Call Ini_Combo(2, wsSQL, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboCusNoTo_GotFocus()
    FocusMe cboCusNoTo
End Sub

Private Sub cboCusNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusNoTo = False Then
            cboCusNoTo.SetFocus
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub

Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboCusNoTo = True
End Function


Private Sub tblCommon_DblClick()
    
    wcCombo.Text = tblCommon.Columns(0).Text
    tblCommon.Visible = False
    wcCombo.SetFocus
    SendKeys "{Enter}"

End Sub

Private Sub tblCommon_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = vbKeyEscape Then
        KeyCode = vbDefault
        tblCommon.Visible = False
        wcCombo.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = vbDefault
        wcCombo.Text = tblCommon.Columns(0).Text
        tblCommon.Visible = False
        wcCombo.SetFocus
        SendKeys "{Enter}"
    End If

End Sub

Private Sub tblCommon_LostFocus()
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then wcCombo.SetFocus

End Sub

