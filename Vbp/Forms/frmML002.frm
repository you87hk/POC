VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmML002 
   Caption         =   "IP0021"
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
      Left            =   8520
      OleObjectBlob   =   "frmML002.frx":0000
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame FraMLType 
      Caption         =   "MLTYPE"
      Height          =   1335
      Left            =   600
      TabIndex        =   16
      Top             =   1920
      Width           =   8055
      Begin VB.OptionButton optMLType 
         Caption         =   "BANK"
         Height          =   255
         Index           =   5
         Left            =   5880
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optMLType 
         Caption         =   "CHEQUE"
         Height          =   255
         Index           =   4
         Left            =   2880
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optMLType 
         Caption         =   "A/P"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optMLType 
         Caption         =   "A/R"
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optMLType 
         Caption         =   "PURCHASE"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optMLType 
         Caption         =   "SALES"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
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
   Begin VB.ComboBox cboMLCodeFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboMLCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboCOAAccCode 
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
            Picture         =   "frmML002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmML002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   14
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
      TabIndex        =   15
      Top             =   760
      Width           =   1860
   End
   Begin VB.Label lblMLCodeFr 
      Caption         =   "MLCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   13
      Top             =   1155
      Width           =   1890
   End
   Begin VB.Label lblMLCodeTo 
      Caption         =   "MLCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   12
      Top             =   1155
      Width           =   375
   End
   Begin VB.Label lblCOAAccCode 
      Caption         =   "COAACCCODE"
      Height          =   225
      Left            =   870
      TabIndex        =   10
      Top             =   1545
      Width           =   1890
   End
End
Attribute VB_Name = "frmML002"
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
    cboMLCodeFr.SetFocus
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
    wsSelection(1) = lblMLCodeFr.Caption & " " & Set_Quote(cboMLCodeFr.Text) & " " & lblMLCodeTo.Caption & " " & Set_Quote(cboMLCodeTo.Text)
    wsSelection(2) = lblCOAAccCode.Caption & " " & Set_Quote(cboCOAAccCode.Text)
    wsSelection(3) = FraMLType.Caption & " " & Set_Quote(Opt_Getfocus(optMLType, 6, 0))
    
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTML002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboMLCodeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboMLCodeTo.Text) = "", String(15, "z"), Set_Quote(cboMLCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCOAAccCode.Text) = "", "0", Get_TableInfo("MstCOA", "COAAccCode= '" & Set_Quote(cboCOAAccCode) & "' AND COAStatus ='1'", "COAAccID")) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(GetMLType()) & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTML002"
    Else
    wsRptName = "RPTML002"
    End If
    
    NewfrmPrint.ReportID = "ML002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "ML002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCOAAccCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCOAAccCode
    
    Select Case gsLangID
        Case "1"
            wsSQL = "SELECT COAAccCode, COADesc FROM MstCOA WHERE COAAccCode LIKE '%" & IIf(cboCOAAccCode.SelLength > 0, "", Set_Quote(cboCOAAccCode.Text)) & "%' AND COAStatus <>'2' "
        Case Else
            wsSQL = "SELECT COAAccCode, COACDesc FROM MstCOA WHERE COAAccCode LIKE '%" & IIf(cboCOAAccCode.SelLength > 0, "", Set_Quote(cboCOAAccCode.Text)) & "%' AND COAStatus <>'2' "
    End Select
   
    wsSQL = wsSQL & " ORDER BY COAAccCode "
    Call Ini_Combo(2, wsSQL, cboCOAAccCode.Left, cboCOAAccCode.Top + cboCOAAccCode.Height, tblCommon, wsFormID, "TBLCOAACCCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCOAAccCode_GotFocus()
    FocusMe cboCOAAccCode
    Set wcCombo = cboCOAAccCode
End Sub

Private Sub cboCOAAccCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCOAAccCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call Opt_Setfocus(optMLType, 6, 0)
    End If
End Sub

Private Sub cboCOAAccCode_LostFocus()
    FocusMe cboCOAAccCode, True
End Sub

Private Sub cboMLCodeTo_LostFocus()
    FocusMe cboMLCodeTo, True
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
    
    wsFormID = "ML002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboMLCodeFr.Text = ""
   cboMLCodeTo.Text = ""
   cboCOAAccCode.Text = ""
   
   optMLType(0).Value = True
   
   wgsTitle = "Merchandise Class List"

End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboMLCodeTo = False Then
        cboMLCodeTo.SetFocus
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
   Set frmML002 = Nothing

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
    
    lblTitle.Caption = Get_Caption(waScrItm, "TITLE")
    txtTitle.Text = Get_Caption(waScrItm, "RPTTITLE")
    lblMLCodeFr.Caption = Get_Caption(waScrItm, "MLCODEFR")
    lblMLCodeTo.Caption = Get_Caption(waScrItm, "MLCODETO")
    lblCOAAccCode.Caption = Get_Caption(waScrItm, "COAACCCODE")
    
    optMLType(0).Caption = Get_Caption(waScrItm, "OPTMLTYPE0")
    optMLType(1).Caption = Get_Caption(waScrItm, "OPTMLTYPE1")
    optMLType(2).Caption = Get_Caption(waScrItm, "OPTMLTYPE2")
    optMLType(3).Caption = Get_Caption(waScrItm, "OPTMLTYPE3")
    optMLType(4).Caption = Get_Caption(waScrItm, "OPTMLTYPE4")
    optMLType(5).Caption = Get_Caption(waScrItm, "OPTMLTYPE5")
    
    FraMLType.Caption = Get_Caption(waScrItm, "FRAMLTYPE")
    
End Sub

Private Function chk_cboMLCodeTo() As Boolean
    chk_cboMLCodeTo = False
    
    If UCase(cboMLCodeFr.Text) > UCase(cboMLCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboMLCodeTo = True
End Function

Private Sub cboMLCodeFr_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboMLCodeFr
    
    If gsLangID = "1" Then
        wsSQL = "SELECT MLCODE, MLDESC "
        wsSQL = wsSQL & " FROM MstMerchClass "
        wsSQL = wsSQL & " WHERE MLCODE LIKE '%" & IIf(cboMLCodeFr.SelLength > 0, "", Set_Quote(cboMLCodeFr.Text)) & "%' "
        wsSQL = wsSQL & " AND MLSTATUS  <> '2' "
        wsSQL = wsSQL & " ORDER BY MLCODE "
    Else
        wsSQL = "SELECT MLCODE, MLDESC "
        wsSQL = wsSQL & " FROM MstMerchClass "
        wsSQL = wsSQL & " WHERE MLCODE LIKE '%" & IIf(cboMLCodeTo.SelLength > 0, "", Set_Quote(cboMLCodeTo.Text)) & "%' "
        wsSQL = wsSQL & " AND MLSTATUS  <> '2' "
        wsSQL = wsSQL & " ORDER BY MLCODE "
    End If
    Call Ini_Combo(2, wsSQL, cboMLCodeFr.Left, cboMLCodeFr.Top + cboMLCodeFr.Height, tblCommon, wsFormID, "TBLMLCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboMLCodeFr_GotFocus()
    FocusMe cboMLCodeFr
    Set wcCombo = cboMLCodeFr
End Sub

Private Sub cboMLCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboMLCodeFr, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboMLCodeFr.Text) <> "" And _
            Trim(cboMLCodeTo.Text) = "" Then
            cboMLCodeTo.Text = cboMLCodeFr.Text
        End If
        
        cboMLCodeTo.SetFocus
    End If
End Sub

Private Sub cboMLCodeFr_LostFocus()
    FocusMe cboMLCodeFr, True
End Sub

Private Sub cboMLCodeTo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboMLCodeTo
  
    If gsLangID = "1" Then
        wsSQL = "SELECT MLCODE, MLDESC "
        wsSQL = wsSQL & " FROM MstMerchClass "
        wsSQL = wsSQL & " WHERE MLCODE LIKE '%" & IIf(cboMLCodeTo.SelLength > 0, "", Set_Quote(cboMLCodeTo.Text)) & "%' "
        wsSQL = wsSQL & " AND MLSTATUS  <> '2' "
        wsSQL = wsSQL & " ORDER BY MLCODE "
    Else
        wsSQL = "SELECT MLCODE, MLDESC "
        wsSQL = wsSQL & " FROM MstMerchClass "
        wsSQL = wsSQL & " WHERE MLCODE LIKE '%" & IIf(cboMLCodeTo.SelLength > 0, "", Set_Quote(cboMLCodeTo.Text)) & "%' "
        wsSQL = wsSQL & " AND MLSTATUS  <> '2' "
        wsSQL = wsSQL & " ORDER BY MLCODE "
    End If
    Call Ini_Combo(2, wsSQL, cboMLCodeTo.Left, cboMLCodeTo.Top + cboMLCodeTo.Height, tblCommon, wsFormID, "TBLMLCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboMLCodeTo_GotFocus()
    FocusMe cboMLCodeTo
    Set wcCombo = cboMLCodeTo
End Sub

Private Sub cboMLCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboMLCodeTo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboMLCodeTo = False Then
            cboMLCodeTo.SetFocus
            Exit Sub
        End If
        
        cboCOAAccCode.SetFocus
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
        
        cboMLCodeFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub

Private Sub optMLType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        cboMLCodeFr.SetFocus
    End If
End Sub

Private Function GetMLType() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 5
        If optMLType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetMLType = "S"
            
        Case 1
            GetMLType = "P"
        
        Case 2
            GetMLType = "A"
        
        Case 3
            GetMLType = "R"
            
        Case 4
            GetMLType = "G"
        
        Case 5
            GetMLType = "B"
    End Select
End Function

