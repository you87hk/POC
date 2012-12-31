VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOS002 
   Caption         =   "O/S Purchase Order"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmOS002.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   7680
      OleObjectBlob   =   "frmOS002.frx":030A
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   744
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   744
      Width           =   1812
   End
   Begin VB.ComboBox cboVdrNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1812
   End
   Begin VB.ComboBox cboVdrNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1080
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
            Picture         =   "frmOS002.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOS002.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   15
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
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   5580
      TabIndex        =   5
      Top             =   1440
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medDueDate 
      Height          =   330
      Left            =   5580
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medWant 
      Height          =   330
      Left            =   2760
      TabIndex        =   6
      Top             =   1845
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Frame fraFormat 
      Caption         =   "Print Format"
      Height          =   1095
      Left            =   840
      TabIndex        =   18
      Top             =   2280
      Width           =   5775
      Begin VB.CheckBox chkGroupBy 
         Alignment       =   1  '靠右對齊
         Caption         =   "Group By Vendor"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   4935
      End
      Begin VB.CheckBox chkPgeBrk 
         Alignment       =   1  '靠右對齊
         Caption         =   "New Page with Each Vendor:"
         Height          =   180
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   4935
      End
   End
   Begin VB.Label lblDueDate 
      Caption         =   "DUEDATE"
      Height          =   225
      Left            =   3960
      TabIndex        =   17
      Top             =   1845
      Width           =   1635
   End
   Begin VB.Label lblWant 
      Caption         =   "WANTDATE"
      Height          =   225
      Left            =   840
      TabIndex        =   16
      Top             =   1845
      Width           =   1905
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "Document # From"
      Height          =   228
      Left            =   864
      TabIndex        =   14
      Top             =   768
      Width           =   1884
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   228
      Left            =   5220
      TabIndex        =   13
      Top             =   768
      Width           =   372
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   228
      Left            =   5220
      TabIndex        =   11
      Top             =   1488
      Width           =   372
   End
   Begin VB.Label lblVdrNoTo 
      Caption         =   "To"
      Height          =   228
      Left            =   5220
      TabIndex        =   10
      Top             =   1128
      Width           =   372
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   228
      Left            =   864
      TabIndex        =   9
      Top             =   1488
      Width           =   1884
   End
   Begin VB.Label lblVdrNoFr 
      Caption         =   "Vendor Code From"
      Height          =   228
      Left            =   864
      TabIndex        =   8
      Top             =   1128
      Width           =   1884
   End
End
Attribute VB_Name = "frmOS002"
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
Private waScrToolTip As New XArrayDB

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
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(5)
    wsSelection(1) = lblDocNoFr.Caption & " " & Set_Quote(cboDocNoFr.Text) & " " & lblDocNoTo.Caption & " " & Set_Quote(cboDocNoTo.Text)
    wsSelection(2) = lblVdrNoFr.Caption & " " & Set_Quote(cboVdrNoFr.Text) & " " & lblVdrNoTo.Caption & " " & Set_Quote(cboVdrNoTo.Text)
    wsSelection(3) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    wsSelection(4) = lblWant.Caption & " " & IIf(Trim(medWant.Text) = "/  /", "9999/99/99", medWant.Text)
    wsSelection(5) = lblDueDate.Caption & " " & IIf(Trim(medDueDate.Text) = "/  /", "9999/99/99", medDueDate.Text)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTOS002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboVdrNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboVdrNoTo.Text) = "", String(10, "z"), Set_Quote(cboVdrNoTo.Text)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medWant.Text) = "/  /", "9999/99/99", medWant.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medDueDate.Text) = "/  /", "9999/99/99", medDueDate.Text) & "', "
    wsSql = wsSql & gsLangID
    
    If chkGroupBy.Value = 0 Then
        wsRptName = "RPTOS002"
    Else
    If chkPgeBrk.Value = 0 Then
        wsRptName = "RPTOS0021"
    Else
        wsRptName = "RPTOS0022"
    End If
    End If
    
    If gsLangID = "2" Then wsRptName = "C" + wsRptName
    
    NewfrmPrint.ReportID = "OS002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "OS002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND VdrStatus <> '2' "
    wsSql = wsSql & " ORDER BY VdrCode "
    Call Ini_Combo(2, wsSql, cboVdrNoFr.Left, cboVdrNoFr.Top + cboVdrNoFr.Height, tblCommon, wsFormID, "TBLVdrNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrNoFr_GotFocus()
        FocusMe cboVdrNoFr
    Set wcCombo = cboVdrNoFr
End Sub

Private Sub cboVdrNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboVdrNoFr.Text) <> "" And _
            Trim(cboVdrNoTo.Text) = "" Then
            cboVdrNoTo.Text = cboVdrNoFr.Text
        End If
        cboVdrNoTo.SetFocus
    End If
End Sub


Private Sub cboVdrNoFr_LostFocus()
    FocusMe cboVdrNoFr, True
End Sub

Private Sub cboVdrNoTo_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND VdrStatus <> '2' "
    wsSql = wsSql & " ORDER BY VdrCode "
    Call Ini_Combo(2, wsSql, cboVdrNoTo.Left, cboVdrNoTo.Top + cboVdrNoTo.Height, tblCommon, wsFormID, "TBLVdrNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrNoTo_GotFocus()
    FocusMe cboVdrNoTo
    Set wcCombo = cboVdrNoTo
End Sub

Private Sub cboVdrNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboVdrNoTo = False Then
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub



Private Sub cboVdrNoTo_LostFocus()
FocusMe cboVdrNoTo, True
End Sub

Private Sub cboDocNoTo_LostFocus()
    FocusMe cboDocNoTo, True
End Sub



Private Sub chkGroupBy_Click()
        If chkGroupBy.Value = 0 Then
            chkPgeBrk.Value = 0
            chkPgeBrk.Enabled = False
        Else
            chkPgeBrk.Enabled = True
        End If
End Sub

Private Sub chkGroupBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chkGroupBy.Value = 0 Then
            chkPgeBrk.Value = 0
            chkPgeBrk.Enabled = False
        Else
            chkPgeBrk.Enabled = True
        End If
            
        If chkPgeBrk.Enabled = False Then
            cboDocNoFr.SetFocus
        Else
            chkPgeBrk.SetFocus
        End If
        
    End If
End Sub

Private Sub chkPgeBrk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
       
            
        cboDocNoFr.SetFocus
        
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
    
    wsFormID = "OS002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   cboVdrNoFr.Text = ""
   cboVdrNoTo.Text = ""
   Call SetPeriodMask(medPrdFr)
   Call SetPeriodMask(medPrdTo)
   Call SetDateMask(medDueDate)
   Call SetDateMask(medWant)
   chkGroupBy.Value = 0
   chkPgeBrk.Enabled = False
   chkPgeBrk.Value = 0
   
End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    If chk_cboVdrNoTo = False Then
        cboVdrNoTo.SetFocus
        Exit Function
    End If
    
    If chk_medPrdTo = False Then
        medPrdTo.SetFocus
        Exit Function
    End If
    
    
    If chk_medWant = False Then
        medWant.SetFocus
        Exit Function
    End If
    
    If Chk_medDueDate = False Then
        medDueDate.SetFocus
        Exit Function
    End If
        
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
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmOS002 = Nothing

End Sub



Private Sub medDueDate_GotFocus()
    FocusMe medDueDate
End Sub

Private Sub medDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_medDueDate = False Then
            Exit Sub
        End If
            
        chkGroupBy.SetFocus
    End If
End Sub

Private Sub medDueDate_LostFocus()
    FocusMe medDueDate, True
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub



Private Sub medWant_GotFocus()
    FocusMe medWant
End Sub

Private Sub medWant_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medWant = False Then
            Exit Sub
        End If
            
        medDueDate.SetFocus
    End If
End Sub

Private Sub medWant_LostFocus()
    FocusMe medWant, True
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
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblVdrNoFr.Caption = Get_Caption(waScrItm, "VdrNoFR")
    lblVdrNoTo.Caption = Get_Caption(waScrItm, "VdrNoTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblDueDate.Caption = Get_Caption(waScrItm, "DueDate")
    lblWant.Caption = Get_Caption(waScrItm, "WANTED")
    chkPgeBrk.Caption = Get_Caption(waScrItm, "PGEBRK")
    chkGroupBy.Caption = Get_Caption(waScrItm, "GROUPBY")
    fraFormat.Caption = Get_Caption(waScrItm, "FORMAT")
    
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
        
    wgsTitle = Get_Caption(waScrItm, "RPTTITLE")
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
Private Function Chk_medDueDate() As Boolean

    Chk_medDueDate = False
    
    If Trim(medDueDate.Text) = "/  /" Then
        Chk_medDueDate = True
        Exit Function
    End If
    
    
    If Chk_Date(medDueDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        medDueDate.SetFocus
        Exit Function
    End If
    
    Chk_medDueDate = True
    
End Function

Private Function chk_medWant() As Boolean

    chk_medWant = False
    
    
    If Trim(medWant.Text) = "/  /" Then
        chk_medWant = True
        Exit Function
    End If
    
    If Chk_Date(medWant) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        medWant.SetFocus
        Exit Function
    End If
    
    chk_medWant = True
    
End Function

Private Function chk_cboVdrNoTo() As Boolean
    chk_cboVdrNoTo = False
    
    If UCase(cboVdrNoFr.Text) > UCase(cboVdrNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        cboVdrNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboVdrNoTo = True
End Function
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
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSql = "SELECT POHDDOCNO, VdrCode, POHDDOCDATE "
    wsSql = wsSql & " FROM POPPOHD, mstVendor "
    wsSql = wsSql & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSql = wsSql & " AND POHDCUSID  = CUSID "
    wsSql = wsSql & " AND POHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY POHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoFr, 15, KeyAscii)
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
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoTo
  
    wsSql = "SELECT POHDDOCNO, VdrCode, POHDDOCDATE "
    wsSql = wsSql & " FROM POPPOHD, mstVendor "
    wsSql = wsSql & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSql = wsSql & " AND POHDCUSID  = CUSID "
    wsSql = wsSql & " AND POHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY POHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoTo_GotFocus()
    FocusMe cboDocNoTo
    Set wcCombo = cboDocNoTo
End Sub

Private Sub cboDocNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoTo = False Then
            Exit Sub
        End If
        
        cboVdrNoFr.SetFocus
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
        
        medWant.SetFocus
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


