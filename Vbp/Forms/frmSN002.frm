VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSN002 
   Caption         =   "Material Master List"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmSN002.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2910
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   600
      Width           =   4665
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   1800
      OleObjectBlob   =   "frmSN002.frx":030A
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame Frame1 
      Height          =   492
      Left            =   2880
      TabIndex        =   17
      Top             =   2040
      Width           =   4455
      Begin VB.OptionButton optPrtMrk 
         Caption         =   "All"
         Height          =   276
         Index           =   2
         Left            =   3000
         TabIndex        =   20
         Top             =   150
         Width           =   1050
      End
      Begin VB.OptionButton optPrtMrk 
         Caption         =   "UnPrint"
         Height          =   276
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   996
      End
      Begin VB.OptionButton optPrtMrk 
         Caption         =   "Printed"
         Height          =   276
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   150
         Width           =   1530
      End
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2910
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   990
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5700
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   990
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5700
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1350
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2910
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
            Picture         =   "frmSN002.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSN002.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '�������W��
      Height          =   360
      Left            =   0
      TabIndex        =   16
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
      Left            =   5700
      TabIndex        =   6
      Top             =   1680
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
      Left            =   2910
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTitle 
      Caption         =   "SHIPPER"
      Height          =   240
      Left            =   960
      TabIndex        =   19
      Top             =   600
      Width           =   1860
   End
   Begin VB.Label lblPrtMrk 
      Caption         =   "Print Range"
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "Document # From"
      Height          =   225
      Left            =   990
      TabIndex        =   15
      Top             =   1005
      Width           =   1890
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   14
      Top             =   1005
      Width           =   375
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   12
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   11
      Top             =   1365
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   990
      TabIndex        =   10
      Top             =   1725
      Width           =   1890
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "Customer Code From"
      Height          =   225
      Left            =   990
      TabIndex        =   9
      Top             =   1365
      Width           =   1890
   End
End
Attribute VB_Name = "frmSN002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim wsFormID As String
Dim wsTrnCd As String

Dim waScrItm As New XArrayDB
'Private waScrToolTip As New XArrayDB
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
    Dim wsTitle2 As String
    
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = lblDocNoFr.Caption & " " & Set_Quote(cboDocNoFr.Text) & " " & lblDocNoTo.Caption & " " & Set_Quote(cboDocNoTo.Text)
    wsSelection(2) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo.Text)
    wsSelection(3) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    wsSelection(4) = lblPrtMrk.Caption & " " & IIf(optPrtMrk(2) = True, "%", IIf(optPrtMrk(0) = True, "N", "Y"))
    

    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPT" & wsFormID & " '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & wsTrnCd & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdFr.Text) = "/", "000000", Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdTo.Text) = "/", "999999", Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "', "
    wsSQL = wsSQL & "'" & IIf(optPrtMrk(2) = True, "%", IIf(optPrtMrk(0) = True, "N", "Y")) & "', "
    wsSQL = wsSQL & "'N', "
    wsSQL = wsSQL & gsLangID
    
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPT" + wsFormID
    Else
    wsRptName = "RPT" + wsFormID
    End If
    
    If wsFormID = "SN002" Then
    wsRptName = wsRptName + "A"
    End If
    
    NewfrmPrint.ReportID = wsFormID
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = wsFormID
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    Select Case wsTrnCd
        Case "PO", "PV", "PR"
            wsSQL = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
             wsSQL = wsSQL & " ORDER BY Vdrcode "
        Case Else
            wsSQL = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
             wsSQL = wsSQL & " ORDER BY Cuscode "
        
    End Select
   
   
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
    Select Case wsTrnCd
        Case "PO", "PV", "PR"
            wsSQL = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
             wsSQL = wsSQL & " ORDER BY Vdrcode "
        Case Else
            wsSQL = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
             wsSQL = wsSQL & " ORDER BY Cuscode "
    End Select
   
   
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
        
        medPrdFr.SetFocus
    End If
End Sub



Private Sub cboCusNoTo_LostFocus()
FocusMe cboCusNoTo, True
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
    
 '   wsFormID = "SN002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   cboCusNoFr.Text = ""
   cboCusNoTo.Text = ""
   Call SetPeriodMask(medPrdFr)
   Call SetPeriodMask(medPrdTo)
   
   'medPrdFr = Left(gsSystemDate, 8) & "01"
   'medPrdTo = Format(DateAdd("D", -1, DateAdd("M", 1, CDate(medPrdFr.Text))), "YYYY/MM/DD")
   
   optPrtMrk(2) = True

End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    If chk_cboCusNoTo = False Then
        cboCusNoTo.SetFocus
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
 '  Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmSN002 = Nothing

End Sub





Private Sub optPrtMrk_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        cboDocNoFr.SetFocus
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
    'Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblPrtMrk.Caption = Get_Caption(waScrItm, "PRTMRK")
    optPrtMrk(0).Caption = Get_Caption(waScrItm, "UNPRINT")
    optPrtMrk(1).Caption = Get_Caption(waScrItm, "PRINTED")
    optPrtMrk(2).Caption = Get_Caption(waScrItm, "ALL")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrItm, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrItm, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrItm, tcExit) & "(F12)"
    
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

Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboCusNoTo = True
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
  
    Select Case wsTrnCd
    Case "SN"
    
    wsSQL = "SELECT SNHDDOCNO, CUSCODE, SNHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSNHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SNHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SNHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SNHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY SNHDDOCNO "
    
   Case "SO"
   
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
    
   Case "SP"
   
    wsSQL = "SELECT SPHDDOCNO, CUSCODE, SPHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSPHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SPHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SPHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SPHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY SPHDDOCNO "
    
   Case "SD"
   
    wsSQL = "SELECT SDHDDOCNO, CUSCODE, SDHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSDHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SDHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SDHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SDHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY SDHDDOCNO "
    
   Case "IV"
   
    wsSQL = "SELECT IVHDDOCNO, CUSCODE, IVHDDOCDATE "
    wsSQL = wsSQL & " FROM soaIVHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND IVHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO "
    
   Case "PO"
   
    wsSQL = "SELECT POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSQL = wsSQL & " FROM POPPOHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND POHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO "
    
    
   Case "PV"
   
    wsSQL = "SELECT PVHDDOCNO, VDRCODE, PVHDDOCDATE "
    wsSQL = wsSQL & " FROM POPPVHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE PVHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND PVHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND PVHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY PVHDDOCNO "
    
   Case "GR"
   
    wsSQL = "SELECT GRHDDOCNO, VDRCODE, GRHDDOCDATE "
    wsSQL = wsSQL & " FROM POPGRHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE GRHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND GRHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND GRHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY GRHDDOCNO "
    
   Case "PR"
   
    wsSQL = "SELECT PRHDDOCNO, VDRCODE, PRHDDOCDATE "
    wsSQL = wsSQL & " FROM POPPRHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE PRHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND PRHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND PRHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY PRHDDOCNO "
    
    End Select
    
    Call Ini_Combo(3, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
  
    Select Case wsTrnCd
    Case "SN"
    wsSQL = "SELECT SNHDDOCNO, CUSCODE, SNHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSNHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SNHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SNHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SNHDSTATUS NOT IN ( '2','3') "
    wsSQL = wsSQL & " ORDER BY SNHDDOCNO "
    
    Case "SO"
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS NOT IN ( '2','3') "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
     
    Case "SP"
    wsSQL = "SELECT SPHDDOCNO, CUSCODE, SPHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSPHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SPHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SPHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SPHDSTATUS NOT IN ( '2','3') "
    wsSQL = wsSQL & " ORDER BY SPHDDOCNO "
    
    Case "SD"
    wsSQL = "SELECT SDHDDOCNO, CUSCODE, SDHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSDHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SDHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SDHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SDHDSTATUS NOT IN ( '2','3') "
    wsSQL = wsSQL & " ORDER BY SDHDDOCNO "
    
    Case "IV"
    wsSQL = "SELECT IVHDDOCNO, CUSCODE, IVHDDOCDATE "
    wsSQL = wsSQL & " FROM soaIVHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND IVHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS NOT IN ( '2','3') "
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO "
    
    Case "PO"
   
    wsSQL = "SELECT POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSQL = wsSQL & " FROM POPPOHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND POHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO "
    
    
   Case "PV"
   
    wsSQL = "SELECT PVHDDOCNO, VDRCODE, PVHDDOCDATE "
    wsSQL = wsSQL & " FROM POPPVHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE PVHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND PVHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND PVHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY PVHDDOCNO "
    
   Case "GR"
   
    wsSQL = "SELECT GRHDDOCNO, VDRCODE, GRHDDOCDATE "
    wsSQL = wsSQL & " FROM POPGRHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE GRHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND GRHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND GRHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY GRHDDOCNO "
    
   Case "PR"
   
    wsSQL = "SELECT PRHDDOCNO, VDRCODE, PRHDDOCDATE "
    wsSQL = wsSQL & " FROM POPPRHD, MSTVENDOR "
    wsSQL = wsSQL & " WHERE PRHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND PRHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND PRHDSTATUS NOT IN ('2','3') "
    wsSQL = wsSQL & " ORDER BY PRHDDOCNO "
   
    End Select
    
    Call Ini_Combo(3, wsSQL, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
        
        cboCusNoFr.SetFocus
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

Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo = False Then
            Call medPrdTo_GotFocus
            Exit Sub
        End If
        Call Opt_Setfocus(optPrtMrk, 3, 0)
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
        
        
       cboDocNoFr.SetFocus
        
    End If
End Sub

Private Sub txtTitle_LostFocus()
FocusMe txtTitle, True
End Sub


Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property


Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property

