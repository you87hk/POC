VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmARL002 
   Caption         =   "Material Master List"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   9195
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   3840
      OleObjectBlob   =   "frmARL002.frx":0000
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   600
      Width           =   4665
   End
   Begin VB.Frame Frame2 
      Height          =   492
      Left            =   2760
      TabIndex        =   19
      Top             =   3360
      Width           =   4455
      Begin VB.OptionButton optByDay 
         Caption         =   "BYDAY"
         Height          =   276
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   144
         Width           =   1335
      End
      Begin VB.OptionButton optByDay 
         Caption         =   "BYMONTH"
         Height          =   276
         Index           =   1
         Left            =   2520
         TabIndex        =   20
         Top             =   144
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Height          =   492
      Left            =   2760
      TabIndex        =   14
      Top             =   2760
      Width           =   4455
      Begin VB.OptionButton optByDate 
         Caption         =   "DOCDATE"
         Height          =   276
         Index           =   1
         Left            =   2520
         TabIndex        =   16
         Top             =   144
         Width           =   1530
      End
      Begin VB.OptionButton optByDate 
         Caption         =   "DUEDATE"
         Height          =   276
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   144
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkZero 
      Alignment       =   1  '靠右對齊
      Height          =   180
      Left            =   2565
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox cboCurr 
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Top             =   2115
      Width           =   1335
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   990
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   990
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2880
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
            Picture         =   "frmARL002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARL002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medAsAt 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   1725
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medIssue 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1365
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   18
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
   Begin VB.Label lblTitle 
      Caption         =   "SHIPPER"
      Height          =   240
      Left            =   840
      TabIndex        =   23
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label lblByDay 
      Caption         =   "BYDAY"
      Height          =   495
      Left            =   840
      TabIndex        =   22
      Top             =   3555
      Width           =   1560
   End
   Begin VB.Label lblByDate 
      Caption         =   "BYDATE"
      Height          =   495
      Left            =   840
      TabIndex        =   17
      Top             =   2955
      Width           =   1560
   End
   Begin VB.Label lblIssue 
      Caption         =   "ISSUE"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   1425
      Width           =   1560
   End
   Begin VB.Label lblZero 
      Caption         =   "ZERO"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2505
      Width           =   1680
   End
   Begin VB.Label LblCurr 
      Caption         =   "CURR"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   2175
      Width           =   1680
   End
   Begin VB.Label lblAsAt 
      Caption         =   "ASAT"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1785
      Width           =   1560
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "Customer"
      Height          =   225
      Left            =   840
      TabIndex        =   9
      Top             =   1005
      Width           =   1890
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   8
      Top             =   1005
      Width           =   375
   End
End
Attribute VB_Name = "frmARL002"
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
    cboDocNoFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(7)
    wsSelection(1) = lblDocNoFr.Caption & " " & Set_Quote(cboDocNoFr.Text) & " " & lblDocNoTo.Caption & " " & Set_Quote(cboDocNoTo.Text)
    wsSelection(2) = LblCurr.Caption & " " & Set_Quote(cboCurr)
    wsSelection(3) = lblIssue.Caption & " " & Set_Quote(medIssue)
    wsSelection(4) = lblAsAt.Caption & " " & Set_Quote(medAsAt)
    wsSelection(5) = lblZero.Caption & " " & IIf(chkZero = True, "Y", "N")
    wsSelection(6) = lblByDate.Caption & " " & IIf(optByDate(0).Value = True, "1", "2")
    wsSelection(7) = lblByDay.Caption & " " & IIf(optByDay(0).Value = True, "1", "2")
    
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTARL002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(10, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCurr) & "', "
    wsSQL = wsSQL & "'" & Left(medIssue, 4) & Right(medIssue, 2) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(medAsAt) & "', "
    wsSQL = wsSQL & "'" & IIf(chkZero.Value = 1, "Y", "N") & "', "
    wsSQL = wsSQL & "'" & IIf(optByDate(0).Value = True, "1", "2") & "', "
    wsSQL = wsSQL & "'" & IIf(optByDay(0).Value = True, "1", "2") & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTARL002"
    Else
    wsRptName = "RPTARL002"
    End If
    
    NewfrmPrint.ReportID = "ARL002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "ARL002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCurr_DropDown()
    
    Dim wsSQL As String
    Dim wsCtlDte As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCurr
    
    wsSQL = "SELECT DISTINCT EXCCURR FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboCurr.SelLength > 0, "", Set_Quote(cboCurr.Text)) & "%' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY EXCCURR "
    Call Ini_Combo(1, wsSQL, cboCurr.Left, cboCurr.Top + cboCurr.Height, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboCurr_GotFocus()
    FocusMe cboCurr
End Sub

Private Sub cboCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCurr, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCurr = False Then
            Exit Sub
        End If
        
        chkZero.SetFocus
        
    End If
End Sub

Private Function Chk_cboCurr() As Boolean
    
    Chk_cboCurr = False
     
    If Trim(cboCurr.Text) = "" Then
        Chk_cboCurr = True
        Exit Function
    End If
    
    If Chk_Curr(cboCurr, gsSystemDate) = False Then
        gsMsg = "沒有此貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCurr.SetFocus
       Exit Function
    End If
    
    Chk_cboCurr = True
    
End Function

Private Sub cboCurr_LostFocus()
    FocusMe cboCurr, True
End Sub

Private Sub cboDocNoTo_LostFocus()
    FocusMe cboDocNoTo, True
End Sub

Private Sub chkZero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call Opt_Setfocus(optByDate, 2, 0)
    End If
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
    
    wsFormID = "ARL002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   SetDateMask medAsAt
   SetPeriodMask medIssue
   
   medIssue.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))
   medAsAt.Text = Dsp_MedDate(gsSystemDate)
   
   optByDate(0).Value = True
   optByDay(0).Value = True

   
   wgsTitle = "State of Account"

End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    If Chk_medIssue = False Then Exit Function
        
    If Chk_medAsAt = False Then Exit Function
    
    If Chk_cboCurr = False Then Exit Function
    
    
    
    InputValidation = True
   
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 4500
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set wcCombo = Nothing
   Set frmARL002 = Nothing

End Sub

Private Sub medAsAt_GotFocus()
    FocusMe medAsAt
End Sub

Private Sub medAsAt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medAsAt Then
            cboCurr.SetFocus
        End If
    End If
End Sub

Private Sub medAsAt_LostFocus()
    FocusMe medAsAt, True
End Sub

Private Function Chk_medAsAt() As Boolean
    
    Chk_medAsAt = False
    
    If Trim(medAsAt.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medAsAt.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medAsAt) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medAsAt.SetFocus
        Exit Function
    End If
    
    Chk_medAsAt = True

End Function

Private Function Chk_medIssue() As Boolean
    
    Chk_medIssue = False
    
    If Trim(medIssue.Text) = "/" Then
        gsMsg = "月份錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medIssue.SetFocus
        Exit Function
    End If
    
    If Chk_Period(medIssue) = False Then
        gsMsg = "月份錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medIssue.SetFocus
        Exit Function
    End If
    
    Chk_medIssue = True

End Function

Private Sub medIssue_GotFocus()
    FocusMe medIssue
End Sub

Private Sub medIssue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medIssue Then
            medAsAt.SetFocus
        End If
    End If
End Sub

Private Sub medIssue_LostFocus()
    FocusMe medIssue, True
End Sub

Private Sub optByDate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call Opt_Setfocus(optByDay, 2, 0)
    End If
End Sub

Private Sub optByDay_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboDocNoFr.SetFocus
    End If
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then wcCombo.SetFocus

End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblDocNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    
    lblAsAt.Caption = Get_Caption(waScrItm, "ASAT")
    lblIssue.Caption = Get_Caption(waScrItm, "ISSUE")
    LblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblZero.Caption = Get_Caption(waScrItm, "ZERO")
    lblByDate.Caption = Get_Caption(waScrItm, "BYDATE")
    lblByDay.Caption = Get_Caption(waScrItm, "BYDAY")
    
    optByDate(0).Caption = Get_Caption(waScrItm, "BYDATE01")
    optByDate(1).Caption = Get_Caption(waScrItm, "BYDATE02")
    
    optByDay(0).Caption = Get_Caption(waScrItm, "BYDAY01")
    optByDay(1).Caption = Get_Caption(waScrItm, "BYDAY02")
    
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
End Sub

Private Function chk_cboDocNoTo() As Boolean
    chk_cboDocNoTo = False
    
    If UCase(cboDocNoFr.Text) > UCase(cboDocNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboDocNoTo = True
End Function

Private Sub cboDocNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSQL = "SELECT CusCode, CusName "
    wsSQL = wsSQL & " FROM MstCustomer "
    wsSQL = wsSQL & " WHERE CusCode LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND CusStatus ='1' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY CusCode "
    Call Ini_Combo(2, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboDocNoFr.Text) <> "" And _
            Trim(cboDocNoTo.Text) = "" Then
            cboDocNoTo.Text = cboDocNoFr.Text
        End If
        cboDocNoTo.SetFocus
    End If
End Sub

Private Sub cboDocNoFr_LostFocus()
    FocusMe cboDocNoFr, True
End Sub

Private Sub cboDocNoTo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoTo
  
    wsSQL = "SELECT CusCode, CusName "
    wsSQL = wsSQL & " FROM MstCustomer "
    wsSQL = wsSQL & " WHERE CusCode LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND CusStatus ='1' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY CusCode "
    Call Ini_Combo(2, wsSQL, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoTo_GotFocus()
    FocusMe cboDocNoTo
    Set wcCombo = cboDocNoTo
End Sub

Private Sub cboDocNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoTo = False Then
            Call cboDocNoTo_GotFocus
            Exit Sub
        End If
        
        medIssue.SetFocus
    End If
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

Private Sub txtTitle_GotFocus()
    FocusMe txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtTitle, 60, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboDocNoFr.SetFocus
        
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

