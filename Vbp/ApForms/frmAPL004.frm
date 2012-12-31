VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAPL004 
   Caption         =   "Material Master List"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   1800
      OleObjectBlob   =   "frmAPL004.frx":0000
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2280
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
            Picture         =   "frmAPL004.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL004.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   5550
      TabIndex        =   7
      Top             =   1500
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
      TabIndex        =   8
      Top             =   1500
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
      TabIndex        =   11
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
      Left            =   840
      TabIndex        =   10
      Top             =   1550
      Width           =   1890
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   9
      Top             =   1550
      Width           =   375
   End
   Begin VB.Label lblTitle 
      Caption         =   "SHIPPER"
      Height          =   240
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "Customer"
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   1150
      Width           =   1890
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   4
      Top             =   1150
      Width           =   375
   End
End
Attribute VB_Name = "frmAPL004"
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
    ReDim wsSelection(2)
    wsSelection(1) = lblDocNoFr.Caption & " " & Set_Quote(cboDocNoFr.Text) & " " & lblDocNoTo.Caption & " " & Set_Quote(cboDocNoTo.Text)
    wsSelection(2) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTAPL004 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(10, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdFr.Text) = "/  /", "0000/00/00", medPrdFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdTo.Text) = "/  /", "9999/99/99", medPrdTo.Text) & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTAPL004"
    Else
    wsRptName = "RPTAPL004"
    End If
    
    NewfrmPrint.ReportID = "APL004"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "APL004"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboDocNoTo_LostFocus()
    FocusMe cboDocNoTo, True
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
    
    wsFormID = "APL004"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   SetDateMask medPrdFr
   SetDateMask medPrdTo
   
   medPrdFr.Text = Left(gsSystemDate, 8) & "01"
   medPrdTo.Text = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(medPrdFr.Text))), "yyyy/mm/dd")
   
   
   wgsTitle = "SALES LEDGER"

End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
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
   Set frmAPL004 = Nothing

End Sub

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
        
        If Trim(medPrdFr) <> "/  /" And _
            Trim(medPrdTo) = "/  /" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub

Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdTo = False Then
            medPrdTo.SetFocus
            Exit Sub
        End If
        
        cboDocNoFr.SetFocus
        
    End If
End Sub

Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
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
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    
    'LblCurr.Caption = Get_Caption(waScrItm, "CURR")
    
    'lblPrtMrk.Caption = Get_Caption(waScrItm, "PRTMRK")
    
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
  
    wsSQL = "SELECT VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND VdrStatus ='1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VdrCode "
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
  
    wsSQL = "SELECT VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND VdrStatus ='1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VdrCode "
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
        
        medPrdFr.SetFocus
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

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/  /" Then
        wsMsg = "必需輸入日期!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        chk_medPrdFr = False
        Exit Function
    End If
    
    If Chk_Date(medPrdFr) = False Then
        wsMsg = "Wrong Date!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If Trim(medPrdTo) = "/  /" Then
        wsMsg = "必需輸入日期!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        chk_medPrdTo = False
        Exit Function
    End If
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        wsMsg = "To Must > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If Chk_Date(medPrdTo) = False Then
    
        wsMsg = "Wrong Date!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function

