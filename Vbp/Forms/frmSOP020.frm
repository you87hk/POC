VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSOP020 
   Caption         =   "SOP020"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   9195
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   6960
      OleObjectBlob   =   "frmSOP020.frx":0000
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame fraSelect 
      Height          =   1935
      Left            =   7440
      TabIndex        =   38
      Top             =   1080
      Width           =   540
      Begin VB.OptionButton optSel 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton optSel 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton optSel 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton optSel 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton optSel 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2640
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5490
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2640
      Width           =   1812
   End
   Begin VB.ComboBox cboAccTypeCodeFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1812
   End
   Begin VB.ComboBox cboAccTypeCodeTo 
      Height          =   300
      Left            =   5490
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1812
   End
   Begin VB.ComboBox cboLevelCodeFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1812
   End
   Begin VB.ComboBox cboLevelCodeTo 
      Height          =   300
      Left            =   5490
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeCodeFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeCodeTo 
      Height          =   300
      Left            =   5490
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboItmCodeTo 
      Height          =   300
      Left            =   5490
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1812
   End
   Begin VB.ComboBox cboItmCodeFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1200
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
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2040
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
            Picture         =   "frmSOP020.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP020.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   24
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
   Begin VB.TextBox txtPayYear 
      Height          =   300
      Left            =   2760
      TabIndex        =   21
      Top             =   4680
      Width           =   885
   End
   Begin VB.TextBox txtPayQuarter 
      Height          =   300
      Left            =   2760
      TabIndex        =   19
      Top             =   4200
      Width           =   885
   End
   Begin VB.TextBox txtPayMonth 
      Height          =   300
      Left            =   2760
      TabIndex        =   17
      Top             =   3720
      Width           =   885
   End
   Begin VB.OptionButton optBy 
      Caption         =   "BYYEAR"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   20
      Top             =   4680
      Width           =   1335
   End
   Begin VB.OptionButton optBy 
      Caption         =   "BYMONTH"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.OptionButton optBy 
      Caption         =   "BYQUARTER"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   18
      Top             =   4230
      Width           =   1455
   End
   Begin VB.Frame fraRange 
      Caption         =   "RANGE"
      Height          =   615
      Left            =   840
      TabIndex        =   33
      Top             =   3480
      Width           =   3180
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   840
      TabIndex        =   34
      Top             =   3960
      Width           =   3180
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   840
      TabIndex        =   35
      Top             =   4440
      Width           =   3180
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "CUSNOFR"
      Height          =   225
      Left            =   840
      TabIndex        =   37
      Top             =   2690
      Width           =   1890
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "CUSNOTO"
      Height          =   225
      Left            =   4890
      TabIndex        =   36
      Top             =   2690
      Width           =   375
   End
   Begin VB.Label lblAccTypeCodeFr 
      Caption         =   "ACCTYPECODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   32
      Top             =   2325
      Width           =   1890
   End
   Begin VB.Label lblAccTypeCodeTo 
      Caption         =   "ACCTYPECODETO"
      Height          =   225
      Left            =   4890
      TabIndex        =   31
      Top             =   2325
      Width           =   375
   End
   Begin VB.Label lblLevelCodeFr 
      Caption         =   "LEVELCODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   30
      Top             =   1965
      Width           =   1890
   End
   Begin VB.Label lblLevelCodeTo 
      Caption         =   "LEVELCODETO"
      Height          =   225
      Left            =   4890
      TabIndex        =   29
      Top             =   1965
      Width           =   375
   End
   Begin VB.Label lblItmTypeCodeFr 
      Caption         =   "ITMTYPECODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   28
      Top             =   1605
      Width           =   1890
   End
   Begin VB.Label lblItmTypeCodeTo 
      Caption         =   "ITMTYPECODETO"
      Height          =   225
      Left            =   4890
      TabIndex        =   27
      Top             =   1605
      Width           =   375
   End
   Begin VB.Label lblItmCodeTo 
      Caption         =   "ITMCODETO"
      Height          =   225
      Left            =   4890
      TabIndex        =   26
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label lblItmCodeFr 
      Caption         =   "ITMCODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   25
      Top             =   1245
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   23
      Top             =   760
      Width           =   1860
   End
End
Attribute VB_Name = "frmSOP020"
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
    cboItmCodeFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wiSel As Integer
    Dim wiBy As Integer
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(7)
    wsSelection(1) = lblItmCodeFr.Caption & " " & Set_Quote(cboItmCodeFr.Text) & " " & lblItmCodeTo.Caption & " " & Set_Quote(cboItmCodeTo.Text)
    wsSelection(2) = lblItmTypeCodeFr.Caption & " " & Set_Quote(cboItmTypeCodeFr.Text) & " " & lblItmTypeCodeTo.Caption & " " & Set_Quote(cboItmTypeCodeTo.Text)
    wsSelection(3) = lblLevelCodeFr.Caption & " " & Set_Quote(cboLevelCodeFr.Text) & " " & lblLevelCodeTo.Caption & " " & Set_Quote(cboLevelCodeTo.Text)
    wsSelection(4) = lblAccTypeCodeFr.Caption & " " & Set_Quote(cboAccTypeCodeFr.Text) & " " & lblAccTypeCodeTo.Caption & " " & Set_Quote(cboAccTypeCodeTo.Text)
    wsSelection(5) = lblCusNoFr.Caption & " " & Set_Quote(cboCusNoFr.Text) & " " & lblCusNoTo.Caption & " " & Set_Quote(cboCusNoTo.Text)
    
    wiSel = Opt_Getfocus(optBy, 3, 0)
    
    Select Case wiBy
    Case 0
        wsSelection(6) = optBy(0).Caption & " " & Set_Quote(txtPayMonth.Text)
    Case 1
        wsSelection(6) = optBy(1).Caption & " " & Set_Quote(txtPayQuarter.Text)
    Case 2
        wsSelection(6) = optBy(2).Caption & " " & Set_Quote(txtPayYear.Text)
    End Select
    
    
    wiBy = Opt_Getfocus(optSel, 5, 0)
    Select Case wiBy
    Case 0
        wsSelection(7) = lblItmCodeFr.Caption
    Case 1
        wsSelection(7) = lblItmTypeCodeFr.Caption
    Case 2
        wsSelection(7) = lblLevelCodeFr.Caption
    Case 3
        wsSelection(7) = lblAccTypeCodeFr.Caption
    Case 4
        wsSelection(7) = lblCusNoFr.Caption
    End Select
        
        
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTSOP020 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboItmCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), Set_Quote(cboItmCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboItmTypeCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboItmTypeCodeTo.Text) = "", String(10, "z"), Set_Quote(cboItmTypeCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(Me.cboLevelCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboLevelCodeTo.Text) = "", String(10, "z"), Set_Quote(cboLevelCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboAccTypeCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboAccTypeCodeTo.Text) = "", String(10, "z"), Set_Quote(cboAccTypeCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSql = wsSql & wiSel & ", "
    
    
    If wiSel = 2 Then
        wsSql = wsSql & Set_Quote(txtPayYear.Text) & ", "
    ElseIf wiSel = 1 Then
        wsSql = wsSql & Set_Quote(txtPayQuarter.Text) & ", "
    ElseIf wiSel = 0 Then
        wsSql = wsSql & Set_Quote(txtPayMonth.Text) & ", "
    End If
    
    wsSql = wsSql & wiBy & ", "
    wsSql = wsSql & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTSOP020"
    Else
    wsRptName = "RPTSOP020"
    End If
    
    NewfrmPrint.ReportID = "SOP020"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SOP020"
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
    
    wsFormID = "SOP020"
    
End Sub

Private Sub Ini_Scr()
    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    cboItmCodeFr.Text = ""
    cboItmCodeTo.Text = ""
    cboItmTypeCodeFr.Text = ""
    cboItmTypeCodeTo.Text = ""
    cboLevelCodeFr.Text = ""
    cboLevelCodeTo.Text = ""
    cboAccTypeCodeFr.Text = ""
    cboAccTypeCodeTo.Text = ""
    cboCusNoFr.Text = ""
    cboCusNoTo.Text = ""
    'txtYear = Format(gsSystemDate, "YYYY")
    
    optBy(0).Value = True
    optSel(0).Value = True
    
    
    txtPayMonth.Text = To_Value(Mid(gsSystemDate, 6, 2))
    txtPayYear.Text = To_Value(Left(gsSystemDate, 4))
   
    If To_Value(txtPayMonth.Text) < 4 Then
        txtPayQuarter.Text = "1"
    ElseIf To_Value(txtPayMonth.Text) >= 4 And To_Value(txtPayMonth.Text) < 7 Then
        txtPayQuarter.Text = "2"
    ElseIf To_Value(txtPayMonth.Text) >= 7 And To_Value(txtPayMonth.Text) < 10 Then
        txtPayQuarter.Text = "3"
    Else
        txtPayQuarter.Text = "4"
    End If
   
    wgsTitle = "Sales Analysis Report"
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    'If chk_cboMethodCodeTo = False Then
    '    cboMethodCodeTo.SetFocus
    '    Exit Function
    'End If
    
    InputValidation = True
   
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 5760
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set waScrItm = Nothing
    Set wcCombo = Nothing
    Set frmSOP020 = Nothing
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
    
    lblItmCodeFr.Caption = Get_Caption(waScrItm, "ITMCODEFR")
    lblItmCodeTo.Caption = Get_Caption(waScrItm, "ITMCODETO")
    lblItmTypeCodeFr.Caption = Get_Caption(waScrItm, "ITMTYPECODEFR")
    lblItmTypeCodeTo.Caption = Get_Caption(waScrItm, "ITMTYPECODETO")
    lblAccTypeCodeFr.Caption = Get_Caption(waScrItm, "ACCTYPECODEFR")
    lblAccTypeCodeTo.Caption = Get_Caption(waScrItm, "ACCTYPECODETO")
    lblLevelCodeFr.Caption = Get_Caption(waScrItm, "LEVELCODEFR")
    lblLevelCodeTo.Caption = Get_Caption(waScrItm, "LEVELCODETO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    
    fraRange.Caption = Get_Caption(waScrItm, "RANGE")
    
    optBy(0).Caption = Get_Caption(waScrItm, "PAYMONTH")
    optBy(1).Caption = Get_Caption(waScrItm, "PAYQUARTER")
    optBy(2).Caption = Get_Caption(waScrItm, "PAYYEAR")
    
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
        
        cboItmCodeFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Sub cboItmCodeFr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeFr
    
    wsSql = "SELECT ITMCODE, ITMBARCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & ", ITMITMTYPECODE, ITMCLASS "
    wsSql = wsSql & " FROM MstItem "
    wsSql = wsSql & " WHERE ITMCODE LIKE '%" & IIf(cboItmCodeFr.SelLength > 0, "", Set_Quote(cboItmCodeFr.Text)) & "%' "
    wsSql = wsSql & " AND ITMSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY ITMCODE "
    Call Ini_Combo(5, wsSql, cboItmCodeFr.Left, cboItmCodeFr.Top + cboItmCodeFr.Height, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeFr_GotFocus()
    FocusMe cboItmCodeFr
    Set wcCombo = cboItmCodeFr
End Sub

Private Sub cboItmCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeFr, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboItmCodeFr.Text) <> "" And _
            Trim(cboItmCodeTo.Text) = "" Then
            cboItmCodeTo.Text = cboItmCodeFr.Text
        End If
        
        cboItmCodeTo.SetFocus
    End If
End Sub

Private Sub cboItmCodeFr_LostFocus()
    FocusMe cboItmCodeFr, True
End Sub

Private Sub cboItmCodeTo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeTo
  
    wsSql = "SELECT ITMCODE, ITMBARCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & ", ITMITMTYPECODE, ITMCLASS "
    wsSql = wsSql & " FROM MstItem "
    wsSql = wsSql & " WHERE ITMCODE LIKE '%" & IIf(cboItmCodeTo.SelLength > 0, "", Set_Quote(cboItmCodeTo.Text)) & "%' "
    wsSql = wsSql & " AND ITMSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY ITMCODE "
    
    Call Ini_Combo(5, wsSql, cboItmCodeTo.Left, cboItmCodeTo.Top + cboItmCodeTo.Height, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeTo_GotFocus()
    FocusMe cboItmCodeTo
    Set wcCombo = cboItmCodeTo
End Sub

Private Sub cboItmCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmCodeTo = False Then
            cboItmCodeTo.SetFocus
            Exit Sub
        End If
        
        cboItmTypeCodeFr.SetFocus
    End If
End Sub

Private Sub cboItmCodeTo_LostFocus()
    FocusMe cboItmCodeTo, True
End Sub

Private Function chk_cboItmCodeTo() As Boolean
    chk_cboItmCodeTo = False
    
    If UCase(cboItmCodeFr.Text) > UCase(cboItmCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboItmCodeTo = True
End Function

Private Sub optBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Index = 0 Then
            txtPayMonth.SetFocus
        ElseIf Index = 1 Then
            txtPayQuarter.SetFocus
        ElseIf Index = 2 Then
            txtPayYear.SetFocus
        End If
    End If
End Sub

Private Sub optSel_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        

        Call Opt_Setfocus(optBy, 3, 0)
        
    End If
End Sub

Private Sub txtPayMonth_GotFocus()
    FocusMe txtPayMonth
End Sub

Private Sub txtPayMonth_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayMonth, False, False)
    Call chk_InpLen(txtPayMonth, 2, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayMonth Then
            cboItmCodeFr.SetFocus
        End If
    End If
End Sub

Private Function Chk_txtPayMonth() As Boolean
    Chk_txtPayMonth = False

    If Trim(txtPayMonth) = "" Then
        Chk_txtPayMonth = True
        Exit Function
    End If

    If To_Value(txtPayMonth) < 1 Or To_Value(txtPayMonth) > 12 Then
        gsMsg = "月份錯誤, 請重新輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPayMonth.SetFocus
        Exit Function
    End If
    
    Chk_txtPayMonth = True
End Function

Private Sub txtPayMonth_LostFocus()
    FocusMe txtPayMonth, True
End Sub

Private Sub txtPayQuarter_GotFocus()
    FocusMe txtPayQuarter
End Sub

Private Sub txtPayQuarter_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayQuarter, False, False)
    Call chk_InpLen(txtPayQuarter, 1, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayQuarter Then
            cboItmCodeFr.SetFocus
        End If
    End If
End Sub

Private Function Chk_txtPayQuarter() As Boolean
    Chk_txtPayQuarter = False

    If Trim(txtPayQuarter) = "" Then
        Chk_txtPayQuarter = True
        Exit Function
    End If

    If To_Value(txtPayQuarter) < 1 Or To_Value(txtPayQuarter) > 4 Then
        gsMsg = "季節錯誤, 請重新輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPayQuarter.SetFocus
        Exit Function
    End If
    
    Chk_txtPayQuarter = True
End Function

Private Sub txtPayQuarter_LostFocus()
    FocusMe txtPayQuarter, True
End Sub

Private Sub txtPayYear_GotFocus()
    FocusMe txtPayYear
End Sub

Private Sub txtPayYear_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayYear, False, False)
    Call chk_InpLen(txtPayYear, 4, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayYear Then
            cboItmCodeFr.SetFocus
        End If
    End If
End Sub

Private Function Chk_txtPayYear() As Boolean
    Chk_txtPayYear = False

    If Trim(txtPayYear) = "" Then
        Chk_txtPayYear = True
        Exit Function
    End If

    If Len(txtPayYear) <> 4 Then
        gsMsg = "年份必須為四位數, 請重新輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPayYear.SetFocus
        Exit Function
    End If
    
    If txtPayYear < To_Value(1990) Or To_Value(txtPayYear) > Format(gsSystemDate, "YYYY") Then
        gsMsg = "年份錯誤, 請重新輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtPayYear.SetFocus
        Exit Function
    End If
    
    Chk_txtPayYear = True
End Function

Private Sub txtPayYear_LostFocus()
    FocusMe txtPayYear, True
End Sub

Private Sub cboItmTypeCodeFr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeFr
    
    wsSql = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC") & " FROM MstItemType WHERE ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeFr.SelLength > 0, "", Set_Quote(cboItmTypeCodeFr.Text)) & "%' AND ItmTypeStatus <>'2' "
    Call Ini_Combo(2, wsSql, cboItmTypeCodeFr.Left, cboItmTypeCodeFr.Top + cboItmTypeCodeFr.Height, tblCommon, wsFormID, "TBLITMTYPECODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmTypeCodeFr_GotFocus()
    FocusMe cboItmTypeCodeFr
    Set wcCombo = cboItmTypeCodeFr
End Sub

Private Sub cboItmTypeCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCodeFr, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboItmTypeCodeFr.Text) <> "" And _
            Trim(cboItmTypeCodeTo.Text) = "" Then
            cboItmTypeCodeTo.Text = cboItmTypeCodeFr.Text
        End If
        
        cboItmTypeCodeTo.SetFocus
    End If
End Sub

Private Sub cboItmTypeCodeFr_LostFocus()
    FocusMe cboItmTypeCodeFr, True
End Sub

Private Sub cboItmTypeCodeTo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeTo
  
    wsSql = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC") & " FROM MstItemType WHERE ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeTo.SelLength > 0, "", Set_Quote(cboItmTypeCodeTo.Text)) & "%' AND ItmTypeStatus <>'2' "
    Call Ini_Combo(2, wsSql, cboItmTypeCodeTo.Left, cboItmTypeCodeTo.Top + cboItmTypeCodeTo.Height, tblCommon, wsFormID, "TBLITMTYPECODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmTypeCodeTo_GotFocus()
    FocusMe cboItmTypeCodeTo
    Set wcCombo = cboItmTypeCodeTo
End Sub

Private Sub cboItmTypeCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCodeTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmTypeCodeTo = False Then
            cboItmTypeCodeTo.SetFocus
            Exit Sub
        End If
        
        cboLevelCodeFr.SetFocus
    End If
End Sub

Private Sub cboItmTypeCodeTo_LostFocus()
    FocusMe cboItmTypeCodeTo, True
End Sub

Private Sub cboLevelCodeFr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboLevelCodeFr
    
    wsSql = "SELECT ITEMCLASSCODE, " & IIf(gsLangID = "1", "ITEMCLASSEDESC", "ITEMCLASSCDESC") & " "
    wsSql = wsSql & " FROM MSTITEMCLASS "
    wsSql = wsSql & " WHERE ITEMCLASSCODE LIKE '%" & IIf(cboLevelCodeFr.SelLength > 0, "", Set_Quote(cboLevelCodeFr.Text)) & "%' "
    
    Call Ini_Combo(2, wsSql, cboLevelCodeFr.Left, cboLevelCodeFr.Top + cboLevelCodeFr.Height, tblCommon, wsFormID, "TBLLEVELCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboLevelCodeFr_GotFocus()
    FocusMe cboLevelCodeFr
    Set wcCombo = cboLevelCodeFr
End Sub

Private Sub cboLevelCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboLevelCodeFr, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboLevelCodeFr.Text) <> "" And _
            Trim(cboLevelCodeTo.Text) = "" Then
            cboLevelCodeTo.Text = cboLevelCodeFr.Text
        End If
        
        cboLevelCodeTo.SetFocus
    End If
End Sub

Private Sub cboLevelCodeFr_LostFocus()
    FocusMe cboLevelCodeFr, True
End Sub

Private Sub cboLevelCodeTo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboLevelCodeTo
  
    wsSql = "SELECT ITEMCLASSCODE, " & IIf(gsLangID = "1", "ITEMCLASSEDESC", "ITEMCLASSCDESC") & " "
    wsSql = wsSql & " FROM MSTITEMCLASS "
    wsSql = wsSql & " WHERE ITEMCLASSCODE LIKE '%" & IIf(cboLevelCodeTo.SelLength > 0, "", Set_Quote(cboLevelCodeTo.Text)) & "%' "
    
    Call Ini_Combo(2, wsSql, cboLevelCodeTo.Left, cboLevelCodeTo.Top + cboLevelCodeTo.Height, tblCommon, wsFormID, "TBLLEVELCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboLevelCodeTo_GotFocus()
    FocusMe cboLevelCodeTo
    Set wcCombo = cboLevelCodeTo
End Sub

Private Sub cboLevelCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboLevelCodeTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboLevelCodeTo = False Then
            cboLevelCodeTo.SetFocus
            Exit Sub
        End If
        
        cboAccTypeCodeFr.SetFocus
    End If
End Sub

Private Sub cboLevelCodeTo_LostFocus()
    FocusMe cboLevelCodeTo, True
End Sub

Private Sub cboAccTypeCodeFr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboAccTypeCodeFr
    
    wsSql = "SELECT AccTypeCode, AccTypeDesc FROM MstAccountType WHERE AccTypeCode LIKE '%" & IIf(cboAccTypeCodeFr.SelLength > 0, "", Set_Quote(cboAccTypeCodeFr.Text)) & "%' AND AccTypeStatus <>'2' "
    Call Ini_Combo(2, wsSql, cboAccTypeCodeFr.Left, cboAccTypeCodeFr.Top + cboAccTypeCodeFr.Height, tblCommon, wsFormID, "TBLACCTYPECODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboAccTypeCodeFr_GotFocus()
    FocusMe cboAccTypeCodeFr
    Set wcCombo = cboAccTypeCodeFr
End Sub

Private Sub cboAccTypeCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccTypeCodeFr, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboAccTypeCodeFr.Text) <> "" And _
            Trim(cboAccTypeCodeTo.Text) = "" Then
            cboAccTypeCodeTo.Text = cboAccTypeCodeFr.Text
        End If
        
        cboAccTypeCodeTo.SetFocus
    End If
End Sub

Private Sub cboAccTypeCodeFr_LostFocus()
    FocusMe cboAccTypeCodeFr, True
End Sub

Private Sub cboAccTypeCodeTo_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboAccTypeCodeTo
  
    wsSql = "SELECT AccTypeCode, AccTypeDesc FROM MstAccountType WHERE AccTypeCode LIKE '%" & IIf(cboAccTypeCodeTo.SelLength > 0, "", Set_Quote(cboAccTypeCodeTo.Text)) & "%' AND AccTypeStatus <>'2' "
    Call Ini_Combo(2, wsSql, cboAccTypeCodeTo.Left, cboAccTypeCodeTo.Top + cboAccTypeCodeTo.Height, tblCommon, wsFormID, "TBLACCTYPECODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboAccTypeCodeTo_GotFocus()
    FocusMe cboAccTypeCodeTo
    Set wcCombo = cboAccTypeCodeTo
End Sub

Private Sub cboAccTypeCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccTypeCodeTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboAccTypeCodeTo = False Then
            cboAccTypeCodeTo.SetFocus
            Exit Sub
        End If
        
        cboCusNoFr.SetFocus
        'Call Opt_Setfocus(optBy, 3, 0)
    End If
End Sub

Private Sub cboAccTypeCodeTo_LostFocus()
    FocusMe cboAccTypeCodeTo, True
End Sub

Private Function chk_cboItmTypeCodeTo() As Boolean
    chk_cboItmTypeCodeTo = False
    
    If UCase(cboItmTypeCodeFr.Text) > UCase(cboItmTypeCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboItmTypeCodeTo = True
End Function

Private Function chk_cboLevelCodeTo() As Boolean
    chk_cboLevelCodeTo = False
    
    If UCase(cboLevelCodeFr.Text) > UCase(cboLevelCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboLevelCodeTo = True
End Function

Private Function chk_cboAccTypeCodeTo() As Boolean
    chk_cboAccTypeCodeTo = False
    
    If UCase(cboAccTypeCodeFr.Text) > UCase(cboAccTypeCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboAccTypeCodeTo = True
End Function

Private Sub cboCusNoFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    
    wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
    wsSql = wsSql & " ORDER BY Cuscode "
    
    Call Ini_Combo(2, wsSql, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
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
    
    wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
    wsSql = wsSql & " ORDER BY Cuscode "
    
    Call Ini_Combo(2, wsSql, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
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
            cboCusNoTo.SetFocus
            Exit Sub
        End If
        
        Call Opt_Setfocus(optSel, 5, 0)
    End If
End Sub

Private Sub cboCusNoTo_LostFocus()
    FocusMe cboCusNoTo, True
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

Private Function GetOptBy() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 2
        If optBy(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetOptBy = "3"
            
        Case 1
            GetOptBy = "2"
        
        Case 2
            GetOptBy = "1"
    End Select
End Function

