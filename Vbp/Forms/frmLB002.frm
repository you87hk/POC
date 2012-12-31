VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLB002 
   Caption         =   "Book Record Card"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmLB002.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3990
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   7680
      OleObjectBlob   =   "frmLB002.frx":030A
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboTypeNoTo 
      Height          =   300
      Left            =   5550
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboTypeNoFr 
      Height          =   300
      Left            =   2730
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   720
      TabIndex        =   21
      Top             =   480
      Width           =   6660
      Begin VB.OptionButton optBy 
         Caption         =   "CALLNO"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.OptionButton optBy 
         Caption         =   "BARCODE"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.ComboBox cboBatchNoFr 
      Height          =   300
      Left            =   2730
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1812
   End
   Begin VB.ComboBox cboBatchNoTo 
      Height          =   300
      Left            =   5550
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1230
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5550
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1950
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2730
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2160
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
            Picture         =   "frmLB002.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLB002.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   5550
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   2730
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Frame fraFormat 
      Caption         =   "Print Format"
      Height          =   1095
      Left            =   840
      TabIndex        =   20
      Top             =   2760
      Width           =   5775
      Begin VB.CheckBox ChkPrtAcc 
         Alignment       =   1  '靠右對齊
         Caption         =   "Print Acc. No. in Label:"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Width           =   4935
      End
      Begin VB.CheckBox chkPgeBrk 
         Alignment       =   1  '靠右對齊
         Caption         =   "New Page with Each Library Class  :"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   4935
      End
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   19
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
   Begin VB.Label lblTypeNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5190
      TabIndex        =   23
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label lblTypeNoFr 
      Caption         =   "Library From"
      Height          =   225
      Left            =   840
      TabIndex        =   22
      Top             =   1575
      Width           =   1890
   End
   Begin VB.Label lblBatchNoFr 
      Caption         =   "Document # From"
      Height          =   225
      Left            =   840
      TabIndex        =   18
      Top             =   1245
      Width           =   1890
   End
   Begin VB.Label lblBatchNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5190
      TabIndex        =   17
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5190
      TabIndex        =   15
      Top             =   2325
      Width           =   375
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5190
      TabIndex        =   14
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   840
      TabIndex        =   13
      Top             =   2325
      Width           =   1890
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "Customer Code From"
      Height          =   225
      Left            =   840
      TabIndex        =   12
      Top             =   1965
      Width           =   1890
   End
End
Attribute VB_Name = "frmLB002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Dim wcCombo As Control
Private wsFormCaption As String
Private waScrToolTip As New XArrayDB

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String


Private Sub cmdCancel()
    Ini_Scr
    Call Opt_Setfocus(optBy, 2)
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(3)
    wsSelection(1) = lblBatchNoFr.Caption & " " & Set_Quote(cboBatchNoFr.Text) & " " & lblBatchNoTo.Caption & " " & Set_Quote(cboBatchNoTo.Text)
    wsSelection(2) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo.Text)
    wsSelection(3) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    
     'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTLB002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboBatchNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboBatchNoTo.Text) = "", String(15, "z"), Set_Quote(cboBatchNoTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboTypeNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboTypeNoTo.Text) = "", String(10, "z"), Set_Quote(cboTypeNoTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "', "
    wsSql = wsSql & gsLangID
    
    If optBy(0) = True Then
    If chkPgeBrk.Value = 0 Then
        wsRptName = "RPTLB0021"
    Else
        wsRptName = "RPTLB0022"
    End If
    Else
    If chkPgeBrk.Value = 0 Then
        wsRptName = "RPTBAR021"
    Else
        wsRptName = "RPTBAR022"
    End If
    End If
    
    If ChkPrtAcc.Value = 0 Then
        wsRptName = wsRptName & "N"
    End If
    
    If gsLangID = "2" Then wsRptName = "C" + wsRptName
    
    NewfrmPrint.ReportID = "LB002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "LB002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND CusStatus <> '2' "
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr_GotFocus()
        FocusMe cboCusNoFr
    Set wcCombo = cboCusNoFr
End Sub

Private Sub cboCusNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoFr, 10, KeyAscii)
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
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
   wsSql = wsSql & " AND CusStatus <> '2' "
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoTo_GotFocus()
    FocusMe cboCusNoTo
    Set wcCombo = cboCusNoTo
End Sub

Private Sub cboCusNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusNoTo = False Then
            Call cboCusNoTo_GotFocus
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub



Private Sub cboCusNoTo_LostFocus()
FocusMe cboCusNoTo, True
End Sub

Private Sub cboBatchNoTo_LostFocus()
    FocusMe cboBatchNoTo, True
End Sub


Private Sub cboTypeNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboTypeNoFr
  
     wsSql = "SELECT CatCode, CatDesc "
    wsSql = wsSql & " FROM mstCategory "
    wsSql = wsSql & " WHERE CatCode LIKE '%" & IIf(cboTypeNoFr.SelLength > 0, "", Set_Quote(cboTypeNoFr.Text)) & "%' "
    wsSql = wsSql & " AND CATSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY CatCode "
    
    Call Ini_Combo(2, wsSql, cboTypeNoFr.Left, cboTypeNoFr.Top + cboTypeNoFr.Height, tblCommon, wsFormID, "TBLTypeNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboTypeNoFr_GotFocus()
    FocusMe cboTypeNoFr
    Set wcCombo = cboTypeNoFr
End Sub

Private Sub cboTypeNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboTypeNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboTypeNoFr.Text) <> "" And _
            Trim(cboTypeNoTo.Text) = "" Then
            cboTypeNoTo.Text = cboTypeNoFr.Text
        End If
        cboTypeNoTo.SetFocus
    End If
End Sub

Private Sub cboTypeNoFr_LostFocus()
    FocusMe cboTypeNoFr, True
End Sub

Private Sub cboTypeNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboTypeNoTo
 
    wsSql = "SELECT CatCode, CatDesc "
    wsSql = wsSql & " FROM mstCategory "
    wsSql = wsSql & " WHERE CatCode LIKE '%" & IIf(cboTypeNoTo.SelLength > 0, "", Set_Quote(cboTypeNoTo.Text)) & "%' "
    wsSql = wsSql & " AND CATSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY CatCode "
    
    Call Ini_Combo(2, wsSql, cboTypeNoTo.Left, cboTypeNoTo.Top + cboTypeNoTo.Height, tblCommon, wsFormID, "TBLTypeNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboTypeNoTo_GotFocus()
    FocusMe cboTypeNoTo
    Set wcCombo = cboTypeNoTo
End Sub

Private Sub cboTypeNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboTypeNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboTypeNoTo = False Then
            Exit Sub
        End If
        
        cboCusNoFr.SetFocus
    End If
End Sub


Private Sub cboTypeNoTo_LostFocus()
    FocusMe cboTypeNoTo, True
End Sub

Private Sub chkPgeBrk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
       
            
        ChkPrtAcc.SetFocus
        
    End If
End Sub



Private Sub ChkPrtAcc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
                  
       Call Opt_Setfocus(optBy, 2)
        
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
    
    wsFormID = "LB002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboBatchNoFr.Text = ""
   cboBatchNoTo.Text = ""
   cboTypeNoFr.Text = ""
   cboTypeNoTo.Text = ""
   cboCusNoFr.Text = ""
   cboCusNoTo.Text = ""
   Call SetPeriodMask(medPrdFr)
   Call SetPeriodMask(medPrdTo)

   chkPgeBrk.Value = 1
   ChkPrtAcc.Value = 1
   
   


End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboBatchNoTo = False Then
        cboBatchNoTo.SetFocus
        Exit Function
    End If
    
    If chk_cboCusNoTo = False Then
        cboCusNoTo.SetFocus
        Exit Function
    End If
    
    If chk_medPrdTo = False Then
        medPrdTo.SetFocus
        Exit Function
    End If
    
    
         
    InputValidation = True
   
End Function



Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 4395
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmLB002 = Nothing

End Sub




Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub




Private Sub optBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
       
            
        cboBatchNoFr.SetFocus
        
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
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblBatchNoFr.Caption = Get_Caption(waScrItm, "BatchNoFR")
    lblBatchNoTo.Caption = Get_Caption(waScrItm, "BatchNoTO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblTypeNoFr.Caption = Get_Caption(waScrItm, "TYPENOFR")
    lblTypeNoTo.Caption = Get_Caption(waScrItm, "TYPENOTO")
    
    chkPgeBrk.Caption = Get_Caption(waScrItm, "PGEBRK")
    ChkPrtAcc.Caption = Get_Caption(waScrItm, "PRTACC")
    fraFormat.Caption = Get_Caption(waScrItm, "FORMAT")
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    optBy(0).Caption = Get_Caption(waScrItm, "BY01")
    optBy(1).Caption = Get_Caption(waScrItm, "BY02")
    
    
End Sub



Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        wsMsg = "To Must > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
    
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function

Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboCusNoTo = True
End Function
Private Function chk_cboBatchNoTo() As Boolean
    chk_cboBatchNoTo = False
    
    If UCase(cboBatchNoFr.Text) > UCase(cboBatchNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        cboBatchNoFr.SetFocus
        Exit Function
    End If
    
    chk_cboBatchNoTo = True
End Function

Private Function chk_cboTypeNoTo() As Boolean
    chk_cboTypeNoTo = False
    
    If UCase(cboTypeNoFr.Text) > UCase(cboTypeNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        cboTypeNoFr.SetFocus
        Exit Function
    End If
    
    chk_cboTypeNoTo = True
End Function
Private Sub cboBatchNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboBatchNoFr
  
    wsSql = "SELECT STHBatchNo, STHLASTUPDDATE "
    wsSql = wsSql & " FROM ICSTKTRNHD "
    wsSql = wsSql & " WHERE STHBatchNo LIKE '%" & IIf(cboBatchNoFr.SelLength > 0, "", Set_Quote(cboBatchNoFr.Text)) & "%' "
    wsSql = wsSql & " AND STHSTATUS  <> '2' "
    wsSql = wsSql & " AND STHTRNCODE  = '2' "
    
    wsSql = wsSql & " ORDER BY STHBatchNo DESC "
    Call Ini_Combo(2, wsSql, cboBatchNoFr.Left, cboBatchNoFr.Top + cboBatchNoFr.Height, tblCommon, wsFormID, "TBLBatchNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboBatchNoFr_GotFocus()
    FocusMe cboBatchNoFr
    Set wcCombo = cboBatchNoFr
End Sub

Private Sub cboBatchNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboBatchNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboBatchNoFr.Text) <> "" And _
            Trim(cboBatchNoTo.Text) = "" Then
            cboBatchNoTo.Text = cboBatchNoFr.Text
        End If
        cboBatchNoTo.SetFocus
    End If
End Sub

Private Sub cboBatchNoFr_LostFocus()
    FocusMe cboBatchNoFr, True
End Sub

Private Sub cboBatchNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboBatchNoTo
  
    wsSql = "SELECT STHBatchNo, STHLASTUPDDATE "
    wsSql = wsSql & " FROM ICSTKTRNHD "
    wsSql = wsSql & " WHERE STHBatchNo LIKE '%" & IIf(cboBatchNoTo.SelLength > 0, "", Set_Quote(cboBatchNoTo.Text)) & "%' "
    wsSql = wsSql & " AND STHSTATUS  <> '2' "
    wsSql = wsSql & " AND STHTRNCODE  = '2' "
    Call Ini_Combo(2, wsSql, cboBatchNoTo.Left, cboBatchNoTo.Top + cboBatchNoTo.Height, tblCommon, wsFormID, "TBLBatchNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboBatchNoTo_GotFocus()
    FocusMe cboBatchNoTo
    Set wcCombo = cboBatchNoTo
End Sub

Private Sub cboBatchNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboBatchNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboBatchNoTo = False Then
            Exit Sub
        End If
        
        cboTypeNoFr.SetFocus
    End If
End Sub


Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub


Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        If Trim(medPrdFr) <> "/" And _
            Trim(medPrdTo) = "/" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub
Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdTo = False Then
            Exit Sub
        End If
        
       chkPgeBrk.SetFocus
    End If
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub
Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
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



