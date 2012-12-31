VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAPL005 
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
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   6600
      OleObjectBlob   =   "frmAPL005.frx":0000
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.CheckBox chkPgeBrk 
      Alignment       =   1  '靠右對齊
      Height          =   180
      Left            =   840
      TabIndex        =   13
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox chkZero 
      Alignment       =   1  '靠右對齊
      Height          =   180
      Left            =   840
      TabIndex        =   5
      Top             =   2355
      Width           =   2175
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1140
      Width           =   1815
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1140
      Width           =   1815
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
            Picture         =   "frmAPL005.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPL005.frx":607B
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
   Begin MSMask.MaskEdBox medAsAt 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medStartMn 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   1940
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblStartMn 
      Caption         =   "STARTMN"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   2010
      Width           =   1680
   End
   Begin VB.Label lblAsAt 
      Caption         =   "ASAT"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1620
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Caption         =   "SHIPPER"
      Height          =   240
      Left            =   840
      TabIndex        =   9
      Top             =   760
      Width           =   1860
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "Customer"
      Height          =   225
      Left            =   840
      TabIndex        =   8
      Top             =   1180
      Width           =   1890
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   7
      Top             =   1180
      Width           =   375
   End
End
Attribute VB_Name = "frmAPL005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Dim wcCombo As Control
Dim wgsTitle As String
Dim wlMonthID As Integer
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String

Private Sub cmdCancel()
    Ini_Scr
    cboCusNoFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsBaseCurCd As String
    
    If InputValidation = False Then Exit Sub
    
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(5)
    wsSelection(1) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo.Text)
    wsSelection(2) = lblStartMn.Caption & " " & Set_Quote(medStartMn)
    wsSelection(3) = lblAsAt.Caption & " " & Set_Quote(medAsAt)
    wsSelection(4) = chkZero.Caption & " " & IIf(chkZero.Value = 1, "Y", "N")
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTAPL005 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSQL = wsSQL & "'', "
    wsSQL = wsSQL & "'" & Set_Quote(medStartMn) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(medAsAt) & "', "
    wsSQL = wsSQL & "'" & IIf(chkZero.Value = 1, "Y", "N") & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTAPL005"
    Else
    wsRptName = "RPTAPL005"
    End If
    
    If chkPgeBrk.Value = 1 Then
    wsRptName = wsRptName + "P"
    End If
    
    NewfrmPrint.ReportID = "APL005"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "APL005"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub



Private Sub cboCusNoTo_LostFocus()
    FocusMe cboCusNoTo, True
End Sub

Private Sub chkZero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        chkPgeBrk.SetFocus
    End If
End Sub

Private Sub chkPgeBrk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboCusNoFr.SetFocus
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
    
    wsFormID = "APL005"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboCusNoFr.Text = ""
   cboCusNoTo.Text = ""
   
   SetDateMask medAsAt
   SetDateMask medStartMn
   
   
   medAsAt.Text = gsSystemDate
   medStartMn.Text = Left(gsSystemDate, 8) & "01"
  ' chkZero.Value = 1
   
   wgsTitle = "CUSTOMER LEDGER"

End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboCusNoTo = False Then
        cboCusNoTo.SetFocus
        Exit Function
    End If
    
    If Chk_medAsAt = False Then
        Exit Function
    End If
    
    If Chk_medStartMn = False Then
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
   Set frmAPL005 = Nothing

End Sub

Private Sub medAsAt_GotFocus()
    FocusMe medAsAt
End Sub

Private Sub medAsAt_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medAsAt Then
            medStartMn.SetFocus
        End If
    End If
End Sub

Private Sub medAsAt_LostFocus()
    FocusMe medAsAt, True
End Sub

Private Function Chk_medAsAt() As Boolean
    Chk_medAsAt = False
    
    If Trim(medAsAt.Text) = "/  /" Then
        gsMsg = "必需輸入日期!"
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

Private Sub medStartMn_GotFocus()
    FocusMe medStartMn
End Sub

Private Sub medStartMn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medStartMn Then
           chkZero.SetFocus
        End If
    End If
End Sub

Private Sub medStartMn_LostFocus()
    FocusMe medStartMn, True
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
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    

    
    lblAsAt.Caption = Get_Caption(waScrItm, "ASAT")
    lblStartMn.Caption = Get_Caption(waScrItm, "STARTMN")
    chkZero.Caption = Get_Caption(waScrItm, "ZERO")
    chkPgeBrk.Caption = Get_Caption(waScrItm, "PGEBRK")
    
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")

End Sub

Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboCusNoTo = True
End Function

Private Sub cboCusNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusNoFr
  
    wsSQL = "SELECT VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND VdrStatus ='1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VdrCode "
    Call Ini_Combo(2, wsSQL, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
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
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusNoTo
  
    wsSQL = "SELECT VdrCode, VdrName "
    wsSQL = wsSQL & " FROM MstVendor "
    wsSQL = wsSQL & " WHERE VdrCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND VdrStatus ='1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY VdrCode "
    
    Call Ini_Combo(2, wsSQL, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
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
        
        medAsAt.SetFocus
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
        
        cboCusNoFr.SetFocus
        
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Function Chk_medStartMn() As Boolean
    Chk_medStartMn = False
    
    If Trim(medStartMn.Text) = "/  /" Then
        gsMsg = "必需輸入日期!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medStartMn.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medStartMn) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medStartMn.SetFocus
        Exit Function
    End If
    
    Chk_medStartMn = True
End Function

