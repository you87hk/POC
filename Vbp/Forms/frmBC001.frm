VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmBC001 
   Caption         =   "Book List"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmBC001.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   7680
      OleObjectBlob   =   "frmBC001.frx":030A
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   2760
      TabIndex        =   0
      Text            =   "ABCDEFGHIJKLMNOPQRS-"
      Top             =   600
      Width           =   4875
   End
   Begin VB.ComboBox cboTypeNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboTypeNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboItemNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1590
      Width           =   1812
   End
   Begin VB.ComboBox cboItemNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1560
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
            Picture         =   "frmBC001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBC001.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   10
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
   Begin VB.Frame fraFormat 
      Caption         =   "Print Format"
      Height          =   1095
      Left            =   840
      TabIndex        =   12
      Top             =   2160
      Width           =   5775
      Begin VB.CheckBox chkPgeBrk 
         Alignment       =   1  '靠右對齊
         Caption         =   "New Page with Each Library Class"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   4935
      End
   End
   Begin VB.Label lblTitle 
      Caption         =   "Period From"
      Height          =   330
      Left            =   840
      TabIndex        =   11
      Top             =   600
      Width           =   1890
   End
   Begin VB.Label lblTypeNoFr 
      Caption         =   "Library From"
      Height          =   225
      Left            =   870
      TabIndex        =   9
      Top             =   1125
      Width           =   1890
   End
   Begin VB.Label lblTypeNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   8
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label lblItemNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   6
      Top             =   1605
      Width           =   375
   End
   Begin VB.Label lblItemNoFr 
      Caption         =   "ISBN From"
      Height          =   225
      Left            =   870
      TabIndex        =   5
      Top             =   1605
      Width           =   1890
   End
End
Attribute VB_Name = "frmBC001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wcCombo As Control
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String


Private Sub cmdCancel()
    Ini_Scr
    cboTypeNoFr.SetFocus
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
    wsSelection(1) = lblTypeNoFr.Caption & " " & Set_Quote(cboTypeNoFr.Text) & " " & lblTypeNoTo.Caption & " " & Set_Quote(cboTypeNoTo.Text)
    wsSelection(2) = lblItemNoFr.Caption & " " & Set_Quote(cboItemNoFr.Text) & " " & lblItemNoTo.Caption & " " & Set_Quote(cboItemNoTo.Text)
     'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTBC001 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboTypeNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboTypeNoTo.Text) = "", String(15, "z"), Set_Quote(cboTypeNoTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboItemNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboItemNoTo.Text) = "", String(10, "z"), Set_Quote(cboItemNoTo.Text)) & "', "
    wsSql = wsSql & gsLangID
    
    If chkPgeBrk.Value = 0 Then
    wsRptName = "RPTBC001"
    Else
    wsRptName = "RPTBC0011"
    End If
    
    If gsLangID = "2" Then wsRptName = "C" + wsRptName
    
    NewfrmPrint.ReportID = "BC001"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "BC001"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItemNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT ItmCode, ItmEngName FROM mstItem WHERE ItmCode LIKE '%" & IIf(cboItemNoFr.SelLength > 0, "", Set_Quote(cboItemNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT ItmCode, ItmChiName FROM mstItem WHERE ItmCode LIKE '%" & IIf(cboItemNoFr.SelLength > 0, "", Set_Quote(cboItemNoFr.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND ItmStatus <> '2' "
    wsSql = wsSql & " ORDER BY Itmcode "
    
    Call Ini_Combo(2, wsSql, cboItemNoFr.Left, cboItemNoFr.Top + cboItemNoFr.Height, tblCommon, wsFormID, "TBLItemNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItemNoFr_GotFocus()
        FocusMe cboItemNoFr
    Set wcCombo = cboItemNoFr
End Sub

Private Sub cboItemNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItemNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItemNoFr.Text) <> "" And _
            Trim(cboItemNoTo.Text) = "" Then
            cboItemNoTo.Text = cboItemNoFr.Text
        End If
        cboItemNoTo.SetFocus
    End If
End Sub


Private Sub cboItemNoFr_LostFocus()
    FocusMe cboItemNoFr, True
End Sub

Private Sub cboItemNoTo_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT ItmCode, ItmEngName FROM mstItem WHERE ItmCode LIKE '%" & IIf(cboItemNoTo.SelLength > 0, "", Set_Quote(cboItemNoTo.Text)) & "%' "
        Case "2"
            wsSql = "SELECT ItmCode, ItmChiName FROM mstItem WHERE ItmCode LIKE '%" & IIf(cboItemNoTo.SelLength > 0, "", Set_Quote(cboItemNoTo.Text)) & "%' "
        Case Else
        
    End Select
    wsSql = wsSql & " AND ItmStatus <> '2' "
    wsSql = wsSql & " ORDER BY Itmcode "
    
     Call Ini_Combo(2, wsSql, cboItemNoTo.Left, cboItemNoTo.Top + cboItemNoTo.Height, tblCommon, wsFormID, "TBLItemNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItemNoTo_GotFocus()
    FocusMe cboItemNoTo
    Set wcCombo = cboItemNoTo
End Sub

Private Sub cboItemNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItemNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItemNoTo = False Then
            Exit Sub
        End If
        
        chkPgeBrk.SetFocus
    End If
End Sub



Private Sub cboItemNoTo_LostFocus()
FocusMe cboItemNoTo, True
End Sub

Private Sub cboTypeNoTo_LostFocus()
    FocusMe cboTypeNoTo, True
End Sub







Private Sub chkPgeBrk_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
      
        cboTypeNoFr.SetFocus
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
    
    wsFormID = "BC001"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboTypeNoFr.Text = ""
   cboTypeNoTo.Text = ""
   cboItemNoFr.Text = ""
   cboItemNoTo.Text = ""
   chkPgeBrk.Value = 1


End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboTypeNoTo = False Then
        cboTypeNoTo.SetFocus
        Exit Function
    End If
    
    If chk_cboItemNoTo = False Then
        cboItemNoTo.SetFocus
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
   Set frmBC001 = Nothing

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
    lblTypeNoFr.Caption = Get_Caption(waScrItm, "TypeNoFR")
    lblTypeNoTo.Caption = Get_Caption(waScrItm, "TypeNoTO")
    lblItemNoFr.Caption = Get_Caption(waScrItm, "ItemNoFR")
    lblItemNoTo.Caption = Get_Caption(waScrItm, "ItemNoTO")
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
    chkPgeBrk.Caption = Get_Caption(waScrItm, "PGEBRK")
    txtTitle.Text = Get_Caption(waScrItm, "TITLECNT")
    
    fraFormat.Caption = Get_Caption(waScrItm, "PRINTFORMAT")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    
End Sub




Private Function chk_cboItemNoTo() As Boolean
    chk_cboItemNoTo = False
    
    If UCase(cboItemNoFr.Text) > UCase(cboItemNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        cboItemNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboItemNoTo = True
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
        
        cboItemNoFr.SetFocus
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

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtTitle, 50, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        cboTypeNoFr.SetFocus
    End If
End Sub

Private Sub txtTitle_GotFocus()
    FocusMe txtTitle
End Sub
Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

