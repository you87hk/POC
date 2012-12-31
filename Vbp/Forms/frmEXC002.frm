VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmEXC002 
   Caption         =   "EXC002"
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
      OleObjectBlob   =   "frmEXC002.frx":0000
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboExcCurr 
      Height          =   300
      Left            =   2760
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboExcYr 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1812
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
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
            Picture         =   "frmEXC002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEXC002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
   Begin VB.Label lblExcCurr 
      Caption         =   "EXCCURR"
      Height          =   225
      Left            =   840
      TabIndex        =   7
      Top             =   1590
      Width           =   1890
   End
   Begin VB.Label lblExcYr 
      Caption         =   "EXCYR"
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   1230
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   840
      TabIndex        =   4
      Top             =   760
      Width           =   1860
   End
End
Attribute VB_Name = "frmEXC002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Dim wcCombo As Control
Dim wgsTitle As String
Dim wsTrnCd As String
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String

Private Sub cmdCancel()
    Ini_Scr
    cboExcYr.SetFocus
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
    ReDim wsSelection(2)
    wsSelection(1) = lblExcYr.Caption & " " & Set_Quote(cboExcYr.Text)
    wsSelection(2) = lblExcCurr.Caption & " " & Set_Quote(cboExcCurr.Text)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTEXC002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboExcYr.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboExcCurr.Text) & "', "
    wsSql = wsSql & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTEXC002"
    Else
    wsRptName = "RPTEXC002"
    End If
    
    NewfrmPrint.ReportID = "EXC002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "EXC002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
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
    
    wsFormID = "EXC002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboExcYr.Text = ""
   cboExcCurr.Text = ""
   
   wgsTitle = "Exchange Rate List"
    
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If Chk_cboExcYr = False Then
        cboExcYr.SetFocus
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
   Set frmEXC002 = Nothing

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
    lblExcYr.Caption = Get_Caption(waScrItm, "EXCYR")
    lblExcCurr.Caption = Get_Caption(waScrItm, "EXCCURR")
    
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
        
        cboExcYr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Sub cboExcYr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboExcYr
    
    wsSql = "SELECT DISTINCT ExcYr FROM MstExchangeRate WHERE ExcStatus <> '2' "
    wsSql = wsSql & "ORDER BY ExcYr"
    Call Ini_Combo(1, wsSql, cboExcYr.Left, cboExcYr.Top + cboExcYr.Height, tblCommon, wsFormID, "TBLEXCYR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboExcYr_GotFocus()
    FocusMe cboExcYr
End Sub

Private Sub cboExcYr_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, cboExcYr, False, False)
    Call chk_InpLen(cboExcYr, 4, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboExcCurr.SetFocus
    End If
End Sub

Private Sub cboExcYr_LostFocus()
    FocusMe cboExcYr, True
End Sub

Private Function Chk_cboExcYr() As Boolean
    Chk_cboExcYr = False
    
    If Len(Trim(cboExcYr)) <> 4 Then
        gsMsg = "年份錯誤! 請輸入四位數字之年份!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboExcYr.SetFocus
        Exit Function
    End If
    
    Chk_cboExcYr = True
End Function

Private Sub cboExcCurr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboExcCurr
    
    wsSql = "SELECT DISTINCT ExcCurr FROM MstExchangeRate WHERE ExcStatus <> '2' "
    wsSql = wsSql & "ORDER BY ExcCurr"
    
    Call Ini_Combo(1, wsSql, cboExcCurr.Left, cboExcCurr.Top + cboExcCurr.Height, tblCommon, wsFormID, "TBLCURR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboExcCurr_GotFocus()
    FocusMe cboExcCurr
End Sub

Private Sub cboExcCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboExcCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboExcCurr = True Then
            cboExcYr.SetFocus
        End If
    End If
End Sub

Private Sub cboExcCurr_LostFocus()
    FocusMe cboExcCurr, True
End Sub

Private Function Chk_cboExcCurr() As Boolean
    Dim wsStatus As String
    
    Chk_cboExcCurr = False
    
    If Len(Trim(cboExcCurr)) = 0 And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入需要資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboExcCurr.SetFocus
        Exit Function
    End If
    
    'If Chk_ExcCurr(cboExcYr, cboExcCurr, wsStatus) = False Then
    '    gsMsg = "貨幣不存在!"
    '    MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    '    cboExcCurr.SetFocus
    '    Exit Function
    'Else
    '    If wsStatus = "2" Then
    '        gsMsg = "貨幣已存在但已無效!"
    '        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    '        cboExcCurr.SetFocus
    '        Exit Function
    '    Else
    '        gsMsg = "貨幣已存在!"
    '        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    '        cboExcCurr.SetFocus
    '        Exit Function
    '    End If
    '
    '    If wsStatus = "2" Then
    '        gsMsg = "貨幣已存在但已無效!"
    '        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    '        cboExcCurr.SetFocus
    '        Exit Function
    '    End If
    'End If
    
    Chk_cboExcCurr = True
End Function

Private Function Chk_ExcCurr(ByVal inCode As String, ByVal inCode1 As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_ExcCurr = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT ExcStatus "
    wsSql = wsSql & " FROM MstExchangeRate WHERE ExcYr = '" & Set_Quote(inCode) & "' AND ExcCurr = '" & Set_Quote(inCode1) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "ExcStatus")
    
    Chk_ExcCurr = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

