VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSTF002 
   Caption         =   "STF002"
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
      OleObjectBlob   =   "frmSTF002.frx":0000
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame Frame1 
      Height          =   492
      Left            =   2760
      TabIndex        =   10
      Top             =   1680
      Width           =   4455
      Begin VB.OptionButton optStaffType 
         Caption         =   "SALESMAN"
         Height          =   276
         Index           =   2
         Left            =   1560
         TabIndex        =   4
         Top             =   170
         Width           =   1290
      End
      Begin VB.OptionButton optStaffType 
         Caption         =   "WORKER"
         Height          =   276
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   170
         Width           =   1335
      End
      Begin VB.OptionButton optStaffType 
         Caption         =   "ALL"
         Height          =   276
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   170
         Width           =   1050
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2790
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin VB.ComboBox cboStaffCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1275
      Width           =   1812
   End
   Begin VB.ComboBox cboStaffCodeFr 
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
            Picture         =   "frmSTF002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTF002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   12
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
   Begin VB.Label lblStaffType 
      Caption         =   "STAFFTYPE"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1850
      Width           =   1800
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   9
      Top             =   780
      Width           =   1860
   End
   Begin VB.Label lblStaffCodeTo 
      Caption         =   "STAFFCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   7
      Top             =   1330
      Width           =   375
   End
   Begin VB.Label lblStaffCodeFr 
      Caption         =   "STAFFCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   6
      Top             =   1330
      Width           =   1890
   End
End
Attribute VB_Name = "frmSTF002"
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
    cboStaffCodeFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    Dim wsStaffType As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(2)
    wsSelection(1) = lblStaffCodeFr.Caption & " " & Set_Quote(cboStaffCodeFr.Text) & " " & lblStaffCodeTo.Caption & " " & Set_Quote(cboStaffCodeTo.Text)
    
    wsStaffType = CStr(Opt_Getfocus(optStaffType, 3, 0))
    
    If wsStaffType = "0" Then
        wsStaffType = "0"
    ElseIf wsStaffType = "1" Then
        wsStaffType = "01"
    ElseIf wsStaffType = "2" Then
        wsStaffType = "02"
    End If
    
    wsSelection(2) = lblStaffType.Caption & " " & wsStaffType
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTSTF002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboStaffCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboStaffCodeTo.Text) = "", String(10, "z"), Set_Quote(cboStaffCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & wsStaffType & "', "
    wsSql = wsSql & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTSTF002"
    Else
    wsRptName = "RPTSTF002"
    End If
    
    NewfrmPrint.ReportID = "STF002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "STF002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboStaffCodeFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboStaffCodeFr
    
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT StaffCode, StaffName, StaffType FROM MstStaff WHERE StaffCode LIKE '%" & IIf(cboStaffCodeFr.SelLength > 0, "", Set_Quote(cboStaffCodeFr.Text)) & "%' AND StaffStatus <>'2' "
        Case "2"
            wsSql = "SELECT StaffCode, StaffName, StaffType FROM MstStaff WHERE StaffCode LIKE '%" & IIf(cboStaffCodeFr.SelLength > 0, "", Set_Quote(cboStaffCodeFr.Text)) & "%' AND StaffStatus <>'2' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY StaffCode "
    Call Ini_Combo(3, wsSql, cboStaffCodeFr.Left, cboStaffCodeFr.Top + cboStaffCodeFr.Height, tblCommon, wsFormID, "TBLSTAFFCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboStaffCodeFr_GotFocus()
    FocusMe cboStaffCodeFr
    Set wcCombo = cboStaffCodeFr
End Sub

Private Sub cboStaffCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboStaffCodeFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboStaffCodeFr.Text) <> "" And _
            Trim(cboStaffCodeTo.Text) = "" Then
            
            cboStaffCodeTo.Text = cboStaffCodeFr.Text
        End If
        cboStaffCodeTo.SetFocus
    End If
End Sub

Private Sub cboStaffCodeFr_LostFocus()
    FocusMe cboStaffCodeFr, True
End Sub

Private Sub cboStaffCodeTo_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboStaffCodeTo
    
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT StaffCode, StaffName, StaffType FROM MstStaff WHERE StaffCode LIKE '%" & IIf(cboStaffCodeTo.SelLength > 0, "", Set_Quote(cboStaffCodeTo.Text)) & "%' AND StaffStatus <>'2' "
        Case "2"
            wsSql = "SELECT StaffCode, StaffName, StaffType FROM MstStaff WHERE StaffCode LIKE '%" & IIf(cboStaffCodeTo.SelLength > 0, "", Set_Quote(cboStaffCodeTo.Text)) & "%' AND StaffStatus <>'2' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY StaffCode "
    Call Ini_Combo(3, wsSql, cboStaffCodeTo.Left, cboStaffCodeTo.Top + cboStaffCodeTo.Height, tblCommon, wsFormID, "TBLSTAFFCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboStaffCodeTo_GotFocus()
    FocusMe cboStaffCodeTo
    Set wcCombo = cboStaffCodeTo
End Sub

Private Sub cboStaffCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboStaffCodeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboStaffCodeTo = False Then
            cboStaffCodeTo.SetFocus
            Exit Sub
        End If
        
        Call Opt_Setfocus(optStaffType, 3, 0)
    End If
End Sub

Private Sub cboStaffCodeTo_LostFocus()
    FocusMe cboStaffCodeTo, True
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
    
    wsFormID = "STF002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboStaffCodeFr.Text = ""
   cboStaffCodeTo.Text = ""
   
   optStaffType(0).Value = True
   
   wgsTitle = "Staff List"
    
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboStaffCodeTo = False Then
        cboStaffCodeTo.SetFocus
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
   Set frmSTF002 = Nothing

End Sub

Private Sub optStaffType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cboStaffCodeFr.SetFocus
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
    lblStaffCodeFr.Caption = Get_Caption(waScrItm, "STAFFCODEFR")
    lblStaffCodeTo.Caption = Get_Caption(waScrItm, "STAFFCODETO")
    
    lblStaffType.Caption = Get_Caption(waScrItm, "STAFFTYPE")
    
    optStaffType(0).Caption = Get_Caption(waScrItm, "STAFFTYPE0")
    optStaffType(1).Caption = Get_Caption(waScrItm, "STAFFTYPE1")
    optStaffType(2).Caption = Get_Caption(waScrItm, "STAFFTYPE2")
End Sub

Private Function chk_cboStaffCodeTo() As Boolean
    chk_cboStaffCodeTo = False
    
    If UCase(cboStaffCodeFr.Text) > UCase(cboStaffCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboStaffCodeTo = True
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
        
        cboStaffCodeFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub
