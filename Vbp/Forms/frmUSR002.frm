VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUSR002 
   Caption         =   "USR002"
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
      Left            =   1800
      OleObjectBlob   =   "frmUSR002.frx":0000
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboGrpCodeFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1812
   End
   Begin VB.ComboBox cboGrpCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1680
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
   Begin VB.ComboBox cboUsrCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1275
      Width           =   1812
   End
   Begin VB.ComboBox cboUsrCodeFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1275
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
            Picture         =   "frmUSR002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUSR002.frx":607B
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
   Begin VB.Label lblGrpCodeFr 
      Caption         =   "GRPCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   11
      Top             =   1710
      Width           =   1890
   End
   Begin VB.Label lblGrpCodeTo 
      Caption         =   "GRPCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   10
      Top             =   1710
      Width           =   375
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   9
      Top             =   760
      Width           =   1860
   End
   Begin VB.Label lblUsrCodeTo 
      Caption         =   "USRCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   6
      Top             =   1305
      Width           =   375
   End
   Begin VB.Label lblUsrCodeFr 
      Caption         =   "USRCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   5
      Top             =   1305
      Width           =   1890
   End
End
Attribute VB_Name = "frmUSR002"
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
    cboUsrCodeFr.SetFocus
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
    ReDim wsSelection(2)
    wsSelection(1) = lblUsrCodeFr.Caption & " " & Set_Quote(cboUsrCodeFr.Text) & " " & lblUsrCodeTo.Caption & " " & Set_Quote(cboUsrCodeTo.Text)
    wsSelection(2) = lblGrpCodeFr.Caption & " " & Set_Quote(cboGrpCodeFr.Text) & " " & lblGrpCodeTo.Caption & " " & Set_Quote(cboGrpCodeTo.Text)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTUSR002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboUsrCodeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboUsrCodeTo.Text) = "", String(10, "z"), Set_Quote(cboUsrCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboGrpCodeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboGrpCodeTo.Text) = "", String(10, "z"), Set_Quote(cboGrpCodeTo.Text)) & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTUSR002"
    Else
    wsRptName = "RPTUSR002"
    End If
    
    NewfrmPrint.ReportID = "USR002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "USR002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboUsrCodeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboUsrCodeFr
    
    wsSQL = "SELECT UsrCode, UsrName, UsrGrpCode FROM MstUser WHERE UsrCode LIKE '%" & IIf(cboUsrCodeFr.SelLength > 0, "", Set_Quote(cboUsrCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY UsrCode "
    Call Ini_Combo(3, wsSQL, cboUsrCodeFr.Left, cboUsrCodeFr.Top + cboUsrCodeFr.Height, tblCommon, wsFormID, "TBLUSRCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboUsrCodeFr_GotFocus()
    FocusMe cboUsrCodeFr
    Set wcCombo = cboUsrCodeFr
End Sub

Private Sub cboUsrCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboUsrCodeFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboUsrCodeFr.Text) <> "" And _
            Trim(cboUsrCodeTo.Text) = "" Then
            
            cboUsrCodeTo.Text = cboUsrCodeFr.Text
        End If
        cboUsrCodeTo.SetFocus
    End If
End Sub

Private Sub cboUsrCodeFr_LostFocus()
    FocusMe cboUsrCodeFr, True
End Sub

Private Sub cboUsrCodeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboUsrCodeTo
    
    Select Case gsLangID
        Case "1"
            wsSQL = "SELECT UsrCode, UsrName, UsrGrpCode FROM MstUser WHERE UsrCode LIKE '%" & IIf(cboUsrCodeTo.SelLength > 0, "", Set_Quote(cboUsrCodeTo.Text)) & "%' "
        Case "2"
            wsSQL = "SELECT UsrCode, UsrName, UsrGrpCode FROM MstUser WHERE UsrCode LIKE '%" & IIf(cboUsrCodeTo.SelLength > 0, "", Set_Quote(cboUsrCodeTo.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSQL = wsSQL & " ORDER BY UsrCode "
    Call Ini_Combo(3, wsSQL, cboUsrCodeTo.Left, cboUsrCodeTo.Top + cboUsrCodeTo.Height, tblCommon, wsFormID, "TBLUSRCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboUsrCodeTo_GotFocus()
    FocusMe cboUsrCodeTo
    Set wcCombo = cboUsrCodeTo
End Sub

Private Sub cboUsrCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboUsrCodeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboUsrCodeTo = False Then
            cboUsrCodeTo.SetFocus
            Exit Sub
        End If
        
        cboGrpCodeFr.SetFocus
    End If
End Sub

Private Sub cboUsrCodeTo_LostFocus()
    FocusMe cboUsrCodeTo, True
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
    
    wsFormID = "USR002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboUsrCodeFr.Text = ""
   cboUsrCodeTo.Text = ""
   cboGrpCodeFr.Text = ""
   cboGrpCodeTo.Text = ""
   
   wgsTitle = "User List"
    
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboUsrCodeTo = False Then
        cboUsrCodeTo.SetFocus
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
   Set wcCombo = Nothing
   Set frmUSR002 = Nothing

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
    lblUsrCodeFr.Caption = Get_Caption(waScrItm, "USRCODEFR")
    lblUsrCodeTo.Caption = Get_Caption(waScrItm, "USRCODETO")
    lblGrpCodeFr.Caption = Get_Caption(waScrItm, "GRPCODEFR")
    lblGrpCodeTo.Caption = Get_Caption(waScrItm, "GRPCODETO")
    
End Sub

Private Function chk_cboUsrCodeTo() As Boolean
    chk_cboUsrCodeTo = False
    
    If UCase(cboUsrCodeFr.Text) > UCase(cboUsrCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboUsrCodeTo = True
End Function

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
        
        cboUsrCodeFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Sub cboGrpCodeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboGrpCodeFr
    
    wsSQL = "SELECT DISTINCT UsrGrpCode FROM MstUser WHERE UsrGrpCode LIKE '%" & IIf(cboGrpCodeFr.SelLength > 0, "", Set_Quote(cboGrpCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY UsrGrpCode "
    
    Call Ini_Combo(1, wsSQL, cboGrpCodeFr.Left, cboGrpCodeFr.Top + cboGrpCodeFr.Height, tblCommon, wsFormID, "TBLGRPCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboGrpCodeFr_GotFocus()
    FocusMe cboGrpCodeFr
    Set wcCombo = cboGrpCodeFr
End Sub

Private Sub cboGrpCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboGrpCodeFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboGrpCodeFr.Text) <> "" And _
            Trim(cboGrpCodeTo.Text) = "" Then
            
            cboGrpCodeTo.Text = cboGrpCodeFr.Text
        End If
        
        cboGrpCodeTo.SetFocus
    End If
End Sub

Private Sub cboGrpCodeFr_LostFocus()
    FocusMe cboGrpCodeFr, True
End Sub

Private Sub cboGrpCodeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboGrpCodeTo
    
   wsSQL = "SELECT DISTINCT UsrGrpCode FROM MstUser WHERE UsrGrpCode LIKE '%" & IIf(cboGrpCodeTo.SelLength > 0, "", Set_Quote(cboGrpCodeTo.Text)) & "%' "
   wsSQL = wsSQL & " ORDER BY UsrGrpCode "
   
    Call Ini_Combo(1, wsSQL, cboGrpCodeTo.Left, cboGrpCodeTo.Top + cboGrpCodeTo.Height, tblCommon, wsFormID, "TBLGRPCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboGrpCodeTo_GotFocus()
    FocusMe cboGrpCodeTo
    Set wcCombo = cboGrpCodeTo
End Sub

Private Sub cboGrpCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboGrpCodeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboGrpCodeTo = False Then
            cboGrpCodeTo.SetFocus
            Exit Sub
        End If
        
        cboUsrCodeFr.SetFocus
    End If
End Sub

Private Sub cboGrpCodeTo_LostFocus()
    FocusMe cboGrpCodeTo, True
End Sub

Private Function chk_cboGrpCodeTo() As Boolean
    chk_cboGrpCodeTo = False
    
    If UCase(cboGrpCodeFr.Text) > UCase(cboGrpCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboGrpCodeTo = True
End Function

