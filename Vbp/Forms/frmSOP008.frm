VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSOP008 
   Caption         =   "SOP008"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   9000
      OleObjectBlob   =   "frmSOP008.frx":0000
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.TextBox txtCreditLimitTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   4
      Top             =   1560
      Width           =   1155
   End
   Begin VB.TextBox txtCreditLimitFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   3
      Top             =   1560
      Width           =   1155
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1812
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2790
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2040
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
            Picture         =   "frmSOP008.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP008.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   2790
      TabIndex        =   5
      Top             =   1920
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   13
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
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   870
      TabIndex        =   12
      Top             =   1965
      Width           =   1650
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "CUSNOTO"
      Height          =   225
      Left            =   4920
      TabIndex        =   11
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "CUSNOFR"
      Height          =   225
      Left            =   870
      TabIndex        =   10
      Top             =   1245
      Width           =   1890
   End
   Begin VB.Label lblCreditLimitTo 
      Caption         =   "CREDITLIMITTO"
      Height          =   225
      Left            =   4920
      TabIndex        =   9
      Top             =   1640
      Width           =   1095
   End
   Begin VB.Label lblCreditLimitFr 
      Caption         =   "CREDITLIMITFR"
      Height          =   225
      Left            =   870
      TabIndex        =   8
      Top             =   1640
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   7
      Top             =   760
      Width           =   1860
   End
End
Attribute VB_Name = "frmSOP008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Dim wcCombo As Control
Dim wgsTitle As String
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String

Private Sub cmdCancel()
    Ini_Scr
    cboCusNoFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsYYYY As String
    Dim wsMM As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    wsYYYY = Left(medPrdFr.Text, 4)
    wsMM = Right(medPrdFr.Text, 2)
    
    'Create Selection Criteria
    ReDim wsSelection(2)
    wsSelection(1) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo.Text)
    wsSelection(1) = lblPrdFr & " " & medPrdFr
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTSOP008 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSQL = wsSQL & IIf(Trim(txtCreditLimitFr) = "", 0, To_Value(txtCreditLimitFr.Text)) & ", "
    wsSQL = wsSQL & IIf(Trim(txtCreditLimitTo) = "", 0, To_Value(txtCreditLimitTo.Text)) & ", "
    wsSQL = wsSQL & "'" & Set_Quote(wsYYYY) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(wsMM) & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTSOP008"
    Else
    wsRptName = "RPTSOP008"
    End If
    
    NewfrmPrint.ReportID = "SOP008"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SOP008"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
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
    Call Ini_Scr

    MousePointer = vbDefault
End Sub

Private Sub Ini_Form()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "SOP008"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboCusNoFr.Text = ""
   cboCusNoTo.Text = ""
   txtCreditLimitFr = ""
   txtCreditLimitTo = ""
   
   Call SetPeriodMask(medPrdFr)
   
   medPrdFr.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))

   'txtCreditLimitFr.Text = Format("0", "##0.00")
   'txtCreditLimitTo.Text = Format(gsMaxVal, "##0.00")
   
   
   
   wgsTitle = "Customer Credit Check List"
    
End Sub

Private Function InputValidation() As Boolean
    InputValidation = False
    
    'If chk_cboMethodCodeTo = False Then
    '    cboMethodCodeTo.SetFocus
    '    Exit Function
    'End If
    If Not chk_txtCreditLimitTo Then Exit Function
    If Not chk_medPrdFr Then Exit Function
    
    InputValidation = True
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 3840
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set wcCombo = Nothing
   Set frmSOP008 = Nothing

End Sub

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
    
    
 On Error GoTo tblCommon_LostFocus_Err
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If
    
Exit Sub
tblCommon_LostFocus_Err:

Set wcCombo = Nothing

End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblCreditLimitFr.Caption = Get_Caption(waScrItm, "CREDITLIMITFR")
    lblCreditLimitTo.Caption = Get_Caption(waScrItm, "CREDITLIMITTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    
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

Private Sub txtCreditLimitFr_GotFocus()
    FocusMe txtCreditLimitFr
End Sub

Private Sub txtCreditLimitFr_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtCreditLimitFr, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_txtCreditLimitFr() Then
            txtCreditLimitTo.SetFocus
        End If
    End If
End Sub

Private Sub txtCreditLimitFr_LostFocus()
    FocusMe txtCreditLimitFr, True
End Sub

Private Sub txtCreditLimitTo_GotFocus()
    FocusMe txtCreditLimitTo
End Sub

Private Sub txtCreditLimitTo_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtCreditLimitTo, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_txtCreditLimitTo() = True Then
            medPrdFr.SetFocus
        End If
    End If
End Sub

Private Sub txtCreditLimitTo_LostFocus()
    FocusMe txtCreditLimitTo, True
End Sub

Private Sub txtTitle_GotFocus()
    FocusMe txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtTitle, 60, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboCusNoFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Sub cboCusNoFr_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusNoFr
    
    If gsLangID = "1" Then
        wsSQL = "SELECT CUSCODE, CUSNAME, CUSTEL, CUSFAX "
        wsSQL = wsSQL & " FROM MstCustomer "
        wsSQL = wsSQL & " WHERE CUSCODE LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        wsSQL = wsSQL & " AND CUSSTATUS  <> '2' "
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY CUSCODE "
    Else
        wsSQL = "SELECT CUSCODE, CUSNAME, CUSTEL, CUSFAX "
        wsSQL = wsSQL & " FROM MstCustomer "
        wsSQL = wsSQL & " WHERE CUSCODE LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        wsSQL = wsSQL & " AND CUSSTATUS  <> '2' "
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY CUSCODE "
    End If
    Call Ini_Combo(4, wsSQL, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboCusNoFr_GotFocus()
    FocusMe cboCusNoFr
    Set wcCombo = cboCusNoFr
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

Private Sub cboCusNoFr_LostFocus()
    FocusMe cboCusNoFr, True
End Sub

Private Sub cboCusNoTo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusNoTo
  
    If gsLangID = "1" Then
        wsSQL = "SELECT CUSCODE, CUSNAME, CUSTEL, CUSFAX "
        wsSQL = wsSQL & " FROM MstCustomer "
        wsSQL = wsSQL & " WHERE CUSCODE LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
        wsSQL = wsSQL & " AND CUSSTATUS  <> '2' "
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY CUSCODE "
    Else
        wsSQL = "SELECT CUSCODE, CUSNAME, CUSTEL, CUSFAX "
        wsSQL = wsSQL & " FROM MstCustomer "
        wsSQL = wsSQL & " WHERE CUSCODE LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
        wsSQL = wsSQL & " AND CUSSTATUS  <> '2' "
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY CUSCODE "
    End If
    Call Ini_Combo(4, wsSQL, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboCusNoTo_GotFocus()
    FocusMe cboCusNoTo
    Set wcCombo = cboCusNoTo
End Sub

Private Sub cboCusNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusNoTo = False Then
            cboCusNoTo.SetFocus
            Exit Sub
        End If
        
        txtCreditLimitFr.SetFocus
    End If
End Sub

Private Sub cboCusNoTo_LostFocus()
    FocusMe cboCusNoTo, True
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

Private Function chk_txtCreditLimitFr() As Boolean
    chk_txtCreditLimitFr = False
    
    If Trim(txtCreditLimitFr.Text) = "" Then
        chk_txtCreditLimitFr = True
        Exit Function
    End If
    
    If Len(txtCreditLimitFr.Text) > 8 Then
        wsMsg = "Credit Limit too large!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        txtCreditLimitFr.SetFocus
        Exit Function
    End If
    
    chk_txtCreditLimitFr = True
End Function

Private Function chk_txtCreditLimitTo() As Boolean
    chk_txtCreditLimitTo = False
    
    If Trim(txtCreditLimitTo.Text) = "" Then
      '  wsMsg = "Credit Limit must not be zero!"
      '  MsgBox wsMsg, vbOKOnly, gsTitle
      '  txtCreditLimitTo.SetFocus
        chk_txtCreditLimitTo = True
        Exit Function
    End If
    
    If Len(txtCreditLimitTo.Text) > 8 Then
        wsMsg = "Credit Limit too large!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        txtCreditLimitTo.SetFocus
        Exit Function
    End If
    
    If To_Value(txtCreditLimitFr.Text) > To_Value(txtCreditLimitTo.Text) Then
        wsMsg = "From > To!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        txtCreditLimitTo.SetFocus
        Exit Function
    End If
    
    chk_txtCreditLimitTo = True
End Function

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            medPrdFr.SetFocus
            Exit Sub
        End If

        cboCusNoFr.SetFocus
    End If
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/" Then
        gsMsg = "Must Input Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function

