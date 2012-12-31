VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmV002 
   Caption         =   "V002"
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
      OleObjectBlob   =   "V002.frx":0000
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2790
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin VB.CheckBox chkActive 
      Alignment       =   1  '靠右對齊
      Height          =   180
      Left            =   2640
      TabIndex        =   4
      Top             =   1965
      Width           =   375
   End
   Begin VB.ComboBox cboVdrNoFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboVdrNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboSaleCode 
      Height          =   300
      Left            =   2790
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1515
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
            Picture         =   "V002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   9
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
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   11
      Top             =   760
      Width           =   1860
   End
   Begin VB.Label lblActive 
      Caption         =   "ACTIVE"
      Height          =   225
      Left            =   870
      TabIndex        =   10
      Top             =   1965
      Width           =   1770
   End
   Begin VB.Label lblVdrNoFr 
      Caption         =   "VDRNOFR"
      Height          =   225
      Left            =   870
      TabIndex        =   8
      Top             =   1155
      Width           =   1890
   End
   Begin VB.Label lblVdrNoTo 
      Caption         =   "VDRNOTO"
      Height          =   225
      Left            =   5220
      TabIndex        =   7
      Top             =   1155
      Width           =   375
   End
   Begin VB.Label lblSaleCode 
      Caption         =   "SALECODE"
      Height          =   225
      Left            =   870
      TabIndex        =   5
      Top             =   1545
      Width           =   1890
   End
End
Attribute VB_Name = "frmV002"
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
    cboVdrNoFr.SetFocus
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
    ReDim wsSelection(3)
    wsSelection(1) = lblVdrNoFr.Caption & " " & Set_Quote(cboVdrNoFr.Text) & " " & lblVdrNoTo.Caption & " " & Set_Quote(cboVdrNoTo.Text)
    wsSelection(2) = lblSaleCode.Caption & " " & Set_Quote(cboSaleCode.Text)
    wsSelection(3) = lblActive.Caption & " " & IIf(chkActive.Value = 1, "Y", "N")
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTV002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Get_TableInfo("MstSalesman", "SaleCode= '" & Set_Quote(cboSaleCode) & "' AND SaleStatus ='1'", "SaleID") & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboVdrNoTo.Text) = "", String(15, "z"), Get_TableInfo("MstSalesman", "SaleCode= '" & Set_Quote(cboSaleCode) & "' AND SaleStatus ='1'", "SaleID")) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboSaleCode.Text) = "", "0", Get_TableInfo("MstSalesman", "SaleCode= '" & Set_Quote(cboSaleCode) & "' AND SaleStatus ='1'", "SaleID")) & "', "
    wsSQL = wsSQL & "'" & IIf(chkActive.Value = 1, "Y", "N") & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTV002"
    Else
    wsRptName = "RPTV002"
    End If
    
    NewfrmPrint.ReportID = "V002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "V002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboSaleCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboSaleCode
    
    Select Case gsLangID
        Case "1"
            wsSQL = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' AND SaleStatus <>'2' "
        Case "2"
            wsSQL = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' AND SaleStatus <>'2' "
        Case Else
        
    End Select
   
    wsSQL = wsSQL & " ORDER BY SaleCode "
    Call Ini_Combo(2, wsSQL, cboSaleCode.Left, cboSaleCode.Top + cboSaleCode.Height, tblCommon, wsFormID, "TBLSALECODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboSaleCode_GotFocus()
        FocusMe cboSaleCode
    Set wcCombo = cboSaleCode
End Sub

Private Sub cboSaleCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        chkActive.SetFocus
    End If
End Sub

Private Sub cboSaleCode_LostFocus()
    FocusMe cboSaleCode, True
End Sub

Private Sub cboVdrNoTo_LostFocus()
    FocusMe cboVdrNoTo, True
End Sub

Private Sub chkActive_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboVdrNoFr.SetFocus
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
    
    wsFormID = "V002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboVdrNoFr.Text = ""
   cboVdrNoTo.Text = ""
   cboSaleCode.Text = ""
   
   wgsTitle = "Vendor List"
   chkActive.Value = 0
    
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboVdrNoTo = False Then
        cboVdrNoTo.SetFocus
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
   Set frmV002 = Nothing

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
    lblVdrNoFr.Caption = Get_Caption(waScrItm, "VDRNOFR")
    lblVdrNoTo.Caption = Get_Caption(waScrItm, "VDRNOTO")
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblActive.Caption = Get_Caption(waScrItm, "ACTIVE")
    
End Sub

Private Function chk_cboVdrNoTo() As Boolean
    chk_cboVdrNoTo = False
    
    If UCase(cboVdrNoFr.Text) > UCase(cboVdrNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboVdrNoTo = True
End Function

Private Sub cboVdrNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboVdrNoFr
    
    If gsLangID = "1" Then
        wsSQL = "SELECT VDRCODE, VDRNAME, VDRTEL, VDRFAX "
        wsSQL = wsSQL & " FROM MstVendor "
        wsSQL = wsSQL & " WHERE VDRCODE LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        wsSQL = wsSQL & " AND VDRSTATUS  <> '2' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY VDRCODE "
    Else
        wsSQL = "SELECT VDRCODE, VDRNAME, VDRTEL, VDRFAX "
        wsSQL = wsSQL & " FROM MstVendor "
        wsSQL = wsSQL & " WHERE VDRCODE LIKE '%" & IIf(cboVdrNoFr.SelLength > 0, "", Set_Quote(cboVdrNoFr.Text)) & "%' "
        wsSQL = wsSQL & " AND VDRSTATUS  <> '2' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY VDRCODE "
    End If
    Call Ini_Combo(4, wsSQL, cboVdrNoFr.Left, cboVdrNoFr.Top + cboVdrNoFr.Height, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboVdrNoFr_GotFocus()
    FocusMe cboVdrNoFr
    Set wcCombo = cboVdrNoFr
End Sub

Private Sub cboVdrNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrNoFr, 15, KeyAscii)
    
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
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboVdrNoTo
  
    If gsLangID = "1" Then
        wsSQL = "SELECT VDRCODE, VDRNAME, VDRTEL, VDRFAX "
        wsSQL = wsSQL & " FROM MstVendor "
        wsSQL = wsSQL & " WHERE VDRCODE LIKE '%" & IIf(cboVdrNoTo.SelLength > 0, "", Set_Quote(cboVdrNoTo.Text)) & "%' "
        wsSQL = wsSQL & " AND VDRSTATUS  <> '2' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY VDRCODE "
    Else
        wsSQL = "SELECT VDRCODE, VDRNAME, VDRTEL, VDRFAX "
        wsSQL = wsSQL & " FROM MstVendor "
        wsSQL = wsSQL & " WHERE VDRCODE LIKE '%" & IIf(cboVdrNoTo.SelLength > 0, "", Set_Quote(cboVdrNoTo.Text)) & "%' "
        wsSQL = wsSQL & " AND VDRSTATUS  <> '2' "
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & " ORDER BY VDRCODE "
    End If
    Call Ini_Combo(4, wsSQL, cboVdrNoTo.Left, cboVdrNoTo.Top + cboVdrNoTo.Height, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboVdrNoTo_GotFocus()
    FocusMe cboVdrNoTo
    Set wcCombo = cboVdrNoTo
End Sub

Private Sub cboVdrNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrNoTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboVdrNoTo = False Then
            cboVdrNoTo.SetFocus
            Exit Sub
        End If
        
        cboSaleCode.SetFocus
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
        
        cboVdrNoFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub
