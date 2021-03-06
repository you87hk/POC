VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGLP006 
   Caption         =   "Material Master List"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmGLP006.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   7800
      OleObjectBlob   =   "frmGLP006.frx":030A
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame Frame1 
      Height          =   492
      Left            =   2850
      TabIndex        =   22
      Top             =   2760
      Width           =   4455
      Begin VB.CheckBox chkPgeBrk 
         Alignment       =   1  '靠右對齊
         Caption         =   "New Page with Each Customer:"
         Height          =   180
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.ComboBox cboPfx 
      Height          =   300
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox cboUsrFr 
      Height          =   300
      Left            =   2880
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2085
      Width           =   1812
   End
   Begin VB.ComboBox cboUsrTo 
      Height          =   300
      Left            =   5700
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2085
      Width           =   1812
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2880
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   600
      Width           =   4665
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2880
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1350
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5700
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1350
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2640
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
            Picture         =   "frmGLP006.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP006.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   5700
      TabIndex        =   5
      Top             =   1725
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1725
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medUpdTo 
      Height          =   285
      Left            =   5700
      TabIndex        =   9
      Top             =   2445
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medUpdFr 
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Top             =   2445
      Width           =   1095
      _ExtentX        =   1931
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
      TabIndex        =   20
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
   Begin VB.Label lblPrtMrk 
      Caption         =   "Print Range"
      Height          =   255
      Left            =   960
      TabIndex        =   24
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Label lblPfx 
      Caption         =   "DOCNO"
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblUsrFr 
      Caption         =   "USRFR"
      Height          =   225
      Left            =   990
      TabIndex        =   19
      Top             =   2160
      Width           =   1890
   End
   Begin VB.Label lblUpdFr 
      Caption         =   "UPDFR"
      Height          =   225
      Left            =   990
      TabIndex        =   18
      Top             =   2535
      Width           =   1890
   End
   Begin VB.Label lblUsrTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   17
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblUpdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   16
      Top             =   2535
      Width           =   375
   End
   Begin VB.Label lblTitle 
      Caption         =   "SHIPPER"
      Height          =   240
      Left            =   990
      TabIndex        =   15
      Top             =   640
      Width           =   1860
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "Document # From"
      Height          =   225
      Left            =   990
      TabIndex        =   14
      Top             =   1410
      Width           =   1890
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   13
      Top             =   1410
      Width           =   375
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   11
      Top             =   1785
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   990
      TabIndex        =   10
      Top             =   1785
      Width           =   1890
   End
End
Attribute VB_Name = "frmGLP006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
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
    ReDim wsSelection(5)
    wsSelection(1) = lblDocNoFr.Caption & " " & Set_Quote(cboDocNoFr.Text) & " " & lblDocNoTo.Caption & " " & Set_Quote(cboDocNoTo.Text)
    wsSelection(2) = lblPfx.Caption & " " & Set_Quote(cboPfx.Text)
    wsSelection(3) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    wsSelection(4) = lblUsrFr.Caption & " " & Set_Quote(cboUsrFr.Text) & " " & lblUsrTo.Caption & " " & Set_Quote(cboUsrTo.Text)
    wsSelection(5) = lblUpdFr.Caption & " " & medUpdFr.Text & " " & lblUpdTo.Caption & " " & medUpdTo.Text
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTGLP006 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboPfx.Text) = "", "%", Set_Quote(cboPfx.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdFr.Text) = "/", "000000", Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdTo.Text) = "/", "999999", Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboUsrFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboUsrTo.Text) = "", String(10, "z"), Set_Quote(cboUsrTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medUpdFr.Text) = "/  /", "0000/00/00", medUpdFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medUpdTo.Text) = "/  /", "9999/99/99", medUpdTo.Text) & "', "
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTGLP006"
    Else
    wsRptName = "RPTGLP006"
    End If
    
        
    If chkPgeBrk.Value = 1 Then wsRptName = wsRptName + "P"
    
    NewfrmPrint.ReportID = "GLP006"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "GLP006"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub





Private Sub cboPfx_DropDown()
    
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboPfx
  
    
    wsSQL = "SELECT VOUPREFIX "
    wsSQL = wsSQL & " FROM SYSVOUNO "
    wsSQL = wsSQL & " WHERE VOUPREFIX LIKE '%" & IIf(cboPfx.SelLength > 0, "", Set_Quote(cboPfx.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY VOUPREFIX "
  
    
    Call Ini_Combo(1, wsSQL, cboPfx.Left, cboPfx.Top + cboPfx.Height, tblCommon, wsFormID, "TBLPFX", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub



Private Sub cboPfx_GotFocus()
FocusMe cboPfx
End Sub

Private Sub cboPfx_LostFocus()
FocusMe cboPfx, True
End Sub

Private Sub cboPfx_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboPfx, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboPfx() = False Then Exit Sub
        
        cboDocNoFr.SetFocus
        
    End If

End Sub

Private Function Chk_cboPfx() As Boolean

Dim rsRcd As New ADODB.Recordset
Dim wsSQL As String
Dim wsStatus As String
Dim wsUpdFlg As String
Dim wsTrnCode As String
Dim wsDocDate As String
Dim wsPgmNo As String
    
    Chk_cboPfx = False
    
    If Trim(cboPfx.Text) = "" Then
        Chk_cboPfx = True
        Exit Function
    End If
    
        
   If Chk_VouPrefix(cboPfx.Text) = False Then
            gsMsg = "This is not a valid prefix!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboPfx.SetFocus
            Exit Function
   End If
   Chk_cboPfx = True

End Function



Private Sub cboUsrFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboUsrFr
    
    Select Case gsLangID
        Case "1"
            wsSQL = "SELECT USRCODE, USRNAME FROM MstUser WHERE USRCODE LIKE '%" & IIf(cboUsrFr.SelLength > 0, "", Set_Quote(cboUsrFr.Text)) & "%' "
        Case "2"
            wsSQL = "SELECT USRCODE, USRNAME FROM MstUser WHERE USRCODE LIKE '%" & IIf(cboUsrFr.SelLength > 0, "", Set_Quote(cboUsrFr.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSQL = wsSQL & " ORDER BY USRCODE "
    Call Ini_Combo(2, wsSQL, cboUsrFr.Left, cboUsrFr.Top + cboUsrFr.Height, tblCommon, wsFormID, "TBLUSR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboUsrFr_GotFocus()
    FocusMe cboUsrFr
End Sub

Private Sub cboUsrFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboUsrFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboUsrFr.Text) <> "" And _
            Trim(cboUsrTo.Text) = "" Then
            cboUsrTo.Text = cboUsrFr.Text
        End If
        
        cboUsrTo.SetFocus
    End If
End Sub

Private Sub cboUsrFr_LostFocus()
    FocusMe cboUsrFr, True
End Sub

Private Sub cboUsrTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboUsrTo
    
    Select Case gsLangID
        Case "1"
            wsSQL = "SELECT USRCODE, USRNAME FROM MstUser WHERE USRCODE LIKE '%" & IIf(cboUsrTo.SelLength > 0, "", Set_Quote(cboUsrTo.Text)) & "%' "
        Case "2"
            wsSQL = "SELECT USRCODE, USRNAME FROM MstUser WHERE USRCODE LIKE '%" & IIf(cboUsrTo.SelLength > 0, "", Set_Quote(cboUsrTo.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSQL = wsSQL & " ORDER BY USRCODE "
    Call Ini_Combo(2, wsSQL, cboUsrTo.Left, cboUsrTo.Top + cboUsrTo.Height, tblCommon, wsFormID, "TBLUSR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboUsrTo_GotFocus()
    FocusMe cboUsrTo
End Sub

Private Sub cboUsrTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboUsrTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboUsrTo = False Then
            cboUsrTo.SetFocus
            Exit Sub
        End If
        
        medUpdFr.SetFocus
    End If
End Sub

Private Sub cboUsrTo_LostFocus()
    FocusMe cboUsrTo, True
End Sub



Private Sub chkPgeBrk_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        cboPfx.SetFocus
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
    
    wsFormID = "GLP006"
    
End Sub

Private Sub Ini_Scr()
    Dim wsCtlPrd As String
    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    cboDocNoFr.Text = ""
    cboDocNoTo.Text = ""
    cboPfx.Text = ""
    chkPgeBrk.Value = 1
    cboUsrFr.Text = ""
    cboUsrTo.Text = ""
    
    Call SetPeriodMask(medPrdFr)
    Call SetPeriodMask(medPrdTo)
    Call SetDateMask(medUpdFr)
    Call SetDateMask(medUpdTo)
    
    
   wsCtlPrd = getCtrlMth("GL")
'   medPrdFr.Text = Left(wsCtlPrd, 4) & "/" & Right(wsCtlPrd, 2)
'   medPrdTo.Text = Left(wsCtlPrd, 4) & "/" & Right(wsCtlPrd, 2)
   
    
'    medUpdFr = Left(gsSystemDate, 8) & "01"
'    medUpdTo = Format(DateAdd("D", -1, DateAdd("M", 1, CDate(medUpdFr.Text))), "YYYY/MM/DD")

End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    
    If Chk_cboPfx = False Then
        Exit Function
    End If
    
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
   
    
    If chk_medPrdFr = False Then
        medPrdFr.SetFocus
        Exit Function
    End If
    
    If chk_medPrdTo = False Then
        medPrdTo.SetFocus
        Exit Function
    End If
    
    
    If chk_medUpdFr = False Then
        medUpdFr.SetFocus
        Exit Function
    End If
    
    If chk_medUpdTo = False Then
        medUpdTo.SetFocus
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
   Set frmGLP006 = Nothing

End Sub


Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medUpdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medUpdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medUpdFr = False Then
            medUpdFr.SetFocus
            Exit Sub
        End If
        
        If Trim(medUpdFr) <> "/  /" And _
            Trim(medUpdTo) = "/  /" Then
            medUpdTo.Text = medUpdFr.Text
        End If
        
        medUpdTo.SetFocus
    End If
End Sub

Private Sub medUpdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medUpdTo_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medUpdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medUpdTo = False Then
            medUpdTo.SetFocus
            
            Exit Sub
        End If
        
        chkPgeBrk.SetFocus
    End If
End Sub

Private Sub medUpdTo_LostFocus()
    FocusMe medPrdFr, True
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
    lblPfx.Caption = Get_Caption(waScrItm, "PFXfr")
    chkPgeBrk.Caption = Get_Caption(waScrItm, "PGEBRK")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblUpdFr.Caption = Get_Caption(waScrItm, "UPDFR")
    lblUpdTo.Caption = Get_Caption(waScrItm, "UPDTO")
    lblUsrFr.Caption = Get_Caption(waScrItm, "USRFR")
    lblUsrTo.Caption = Get_Caption(waScrItm, "USRTO")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
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
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        wsMsg = "To Must > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
    
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function

Private Function chk_medUpdFr() As Boolean
    chk_medUpdFr = False
    
    If Trim(medUpdFr) = "/  /" Then
        chk_medUpdFr = True
        Exit Function
    End If
    
    If Chk_Date(medUpdFr) = False Then
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medUpdFr = True
End Function

Private Function chk_medUpdTo() As Boolean
    chk_medUpdTo = False
    
    If UCase(medUpdFr.Text) > UCase(medUpdTo.Text) Then
        wsMsg = "To Must > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If Trim(medUpdTo) = "/  /" Then
        chk_medUpdTo = True
        Exit Function
    End If

    If Chk_Date(medUpdTo) = False Then
    
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medUpdTo = True
End Function


Private Function chk_cboUsrTo() As Boolean
    chk_cboUsrTo = False
    
    If UCase(cboUsrFr.Text) > UCase(cboUsrTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboUsrTo = True
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
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSQL = "SELECT VOHDDOCNO, VOHDDOCDATE "
    wsSQL = wsSQL & "FROM GLVOHD "
    wsSQL = wsSQL & "WHERE VOHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & "AND VOHDPFX LIKE '%" & IIf(cboPfx.SelLength > 0, "", Set_Quote(cboPfx.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY VOHDDOCNO"
    Call Ini_Combo(2, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoTo
  
    wsSQL = "SELECT VOHDDOCNO, VOHDDOCDATE "
    wsSQL = wsSQL & "FROM GLVOHD "
    wsSQL = wsSQL & "WHERE VOHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & "AND VOHDPFX LIKE '%" & IIf(cboPfx.SelLength > 0, "", Set_Quote(cboPfx.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY VOHDDOCNO"
    Call Ini_Combo(2, wsSQL, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
            Call cboDocNoTo_GotFocus
            Exit Sub
        End If
        
       medPrdFr.SetFocus
    End If
End Sub

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Call medPrdFr_GotFocus
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
            medPrdTo.SetFocus
            
            Exit Sub
        End If
        
        cboUsrFr.SetFocus
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

Private Sub txtTitle_GotFocus()
    FocusMe txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtTitle, 60, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
       cboPfx.SetFocus
        
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub
