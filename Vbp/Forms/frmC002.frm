VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmC002 
   Caption         =   "C002"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   Begin VB.ComboBox cboCusCurrTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   31
      Top             =   2920
      Width           =   1812
   End
   Begin VB.ComboBox cboCusCurrFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   30
      Top             =   2920
      Width           =   1812
   End
   Begin VB.ComboBox cboTypCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   27
      Top             =   2560
      Width           =   1812
   End
   Begin VB.ComboBox cboTypCodeFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   26
      Top             =   2560
      Width           =   1812
   End
   Begin VB.ComboBox cboCusPayCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   23
      Top             =   1830
      Width           =   1812
   End
   Begin VB.ComboBox cboCusPayCodeFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   22
      Top             =   1830
      Width           =   1812
   End
   Begin VB.ComboBox cboSaleCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   19
      Top             =   1470
      Width           =   1812
   End
   Begin VB.ComboBox cboSaleCodeFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   18
      Top             =   1470
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNatureCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   15
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNatureCodeFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   14
      Top             =   1110
      Width           =   1812
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   9000
      OleObjectBlob   =   "frmC002.frx":0000
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame Frame1 
      Height          =   492
      Left            =   2760
      TabIndex        =   10
      Top             =   3480
      Width           =   4455
      Begin VB.OptionButton optPrtMrk 
         Caption         =   "UnPrint"
         Height          =   276
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   144
         Width           =   996
      End
      Begin VB.OptionButton optPrtMrk 
         Caption         =   "Printed"
         Height          =   276
         Index           =   1
         Left            =   2520
         TabIndex        =   11
         Top             =   144
         Width           =   1530
      End
   End
   Begin VB.ComboBox cboCusTerrFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   0
      Top             =   2190
      Width           =   1812
   End
   Begin VB.ComboBox cboCusTerrTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   1
      Top             =   2190
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   3
      Top             =   750
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   2
      Top             =   750
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
            Picture         =   "frmC002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmC002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   1058
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
   Begin VB.Label lblCusCurrTo 
      Caption         =   "CUSCURRTO"
      Height          =   225
      Left            =   5220
      TabIndex        =   33
      Top             =   2980
      Width           =   375
   End
   Begin VB.Label lblCusCurrFr 
      Caption         =   "CUSCURRFR"
      Height          =   225
      Left            =   870
      TabIndex        =   32
      Top             =   2980
      Width           =   1890
   End
   Begin VB.Label lblTypCodeTo 
      Caption         =   "TYPCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   29
      Top             =   2625
      Width           =   375
   End
   Begin VB.Label lblTypCodeFr 
      Caption         =   "TYPCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   28
      Top             =   2625
      Width           =   1890
   End
   Begin VB.Label lblCusPayCodeTo 
      Caption         =   "CUSPAYCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   25
      Top             =   1885
      Width           =   375
   End
   Begin VB.Label lblCusPayCodeFr 
      Caption         =   "CUSPAYCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   24
      Top             =   1885
      Width           =   1890
   End
   Begin VB.Label lblSaleCodeTo 
      Caption         =   "SALECODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   21
      Top             =   1525
      Width           =   375
   End
   Begin VB.Label lblSaleCodeFr 
      Caption         =   "SALECODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   20
      Top             =   1525
      Width           =   1890
   End
   Begin VB.Label lblCusNatureCodeTo 
      Caption         =   "CUSNATURECODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   17
      Top             =   1165
      Width           =   375
   End
   Begin VB.Label lblCusNatureCodeFr 
      Caption         =   "CUSNATURECODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   16
      Top             =   1165
      Width           =   1890
   End
   Begin VB.Label lblPrtMrk 
      Caption         =   "Print Range"
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Label lblCusTerrFr 
      Caption         =   "CUSTERRFR"
      Height          =   225
      Left            =   870
      TabIndex        =   8
      Top             =   2245
      Width           =   1890
   End
   Begin VB.Label lblCusTerrTo 
      Caption         =   "CUSTERRTO"
      Height          =   225
      Left            =   5220
      TabIndex        =   7
      Top             =   2245
      Width           =   375
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "CUSCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   5
      Top             =   805
      Width           =   375
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "CUSCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   4
      Top             =   805
      Width           =   1890
   End
End
Attribute VB_Name = "frmC002"
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
    cboCusNoFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(8)
    wsSelection(1) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo)
    wsSelection(2) = lblCusNatureCodeFr.Caption & " " & Set_Quote(cboCusNatureCodeFr.Text) & " " & lblCusNatureCodeTo.Caption & " " & Set_Quote(cboCusNatureCodeTo)
    wsSelection(3) = lblSaleCodeFr.Caption & " " & Set_Quote(cboSaleCodeFr) & " " & lblSaleCodeTo.Caption & " " & Set_Quote(cboSaleCodeTo)
    wsSelection(4) = lblCusPayCodeFr.Caption & " " & Set_Quote(cboCusPayCodeFr) & " " & lblCusPayCodeTo.Caption & " " & Set_Quote(cboCusPayCodeTo)
    wsSelection(5) = lblCusTerrFr.Caption & " " & Set_Quote(cboCusTerrFr) & " " & lblCusTerrTo.Caption & " " & Set_Quote(cboCusTerrTo)
    wsSelection(6) = lblTypCodeFr.Caption & " " & Set_Quote(cboTypCodeFr.Text) & " " & lblTypCodeTo.Caption & " " & Set_Quote(cboTypCodeTo)
    wsSelection(7) = lblCusCurrFr.Caption & " " & Set_Quote(cboCusCurrFr.Text) & " " & lblCusCurrTo.Caption & " " & Set_Quote(cboCusCurrTo)
    wsSelection(8) = lblPrtMrk.Caption & " " & IIf(optPrtMrk(0) = True, "N", "Y")
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTC002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(15, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusNatureCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusNatureCodeTo.Text) = "", String(10, "z"), Set_Quote(cboCusNatureCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboSaleCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboSaleCodeTo.Text) = "", String(10, "z"), Set_Quote(cboSaleCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusPayCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusPayCodeTo.Text) = "", String(10, "z"), Set_Quote(cboCusPayCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusTerrFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusTerrTo.Text) = "", String(10, "z"), Set_Quote(cboCusTerrTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboTypCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboTypCodeTo.Text) = "", String(10, "z"), Set_Quote(cboTypCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusCurrFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusCurrTo.Text) = "", String(10, "z"), Set_Quote(cboCusCurrTo.Text)) & "', "
    wsSql = wsSql & "'" & IIf(optPrtMrk(0) = True, "N", "Y") & "', "
    wsSql = wsSql & gsLangID
    
    NewfrmPrint.ReportID = "C002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "C002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = "RPTC002"
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, "SN002", "TBLCUSNO", Me.Width, Me.Height)
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
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, "SN002", "TBLCUSNO", Me.Width, Me.Height)
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
    
    wsFormID = "SN002"
    
End Sub

Private Sub Ini_Scr()
   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   optPrtMrk(0) = True
   wgsTitle = "Customer List"
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
   Set frmSN002 = Nothing

End Sub



Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then wcCombo.SetFocus

End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item("SN002", waScrItm)
   
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
    
    If UCase(medPrdTo.Text) > UCase(medPrdTo.Text) Then
        wsMsg = "To > From!"
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
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSql = "SELECT SNHDDOCNO, CUSCODE, SNHDDOCDATE "
    wsSql = wsSql & " FROM soaSNHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE SNHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSql = wsSql & " AND SNHDCUSID  = CUSID "
    wsSql = wsSql & " AND SNHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY SNHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, "SN002", "TBLDOCNO", Me.Width, Me.Height)
    
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
  
    wsSql = "SELECT SNHDDOCNO, CUSCODE, SNHDDOCDATE "
    wsSql = wsSql & " FROM soaSNHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE SNHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSql = wsSql & " AND SNHDCUSID  = CUSID "
    wsSql = wsSql & " AND SNHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY SNHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, "SN002", "TBLDOCNO", Me.Width, Me.Height)
    
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
            Call medPrdTo_GotFocus
            Exit Sub
        End If
        Call Opt_Setfocus(optPrtMrk, 2, 0)
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


