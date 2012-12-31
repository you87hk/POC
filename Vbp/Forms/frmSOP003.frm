VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSOP003 
   Caption         =   "SOP003"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   9195
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   7920
      OleObjectBlob   =   "frmSOP003.frx":0000
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboItmTypeCode 
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeCodeTo 
      Height          =   300
      Left            =   5610
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboCusCodeTo 
      Height          =   300
      Left            =   5640
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1155
      Width           =   1812
   End
   Begin VB.TextBox txtPayYear 
      Height          =   300
      Left            =   2760
      TabIndex        =   10
      Top             =   3360
      Width           =   885
   End
   Begin VB.TextBox txtPayQuarter 
      Height          =   300
      Left            =   2760
      TabIndex        =   9
      Top             =   2880
      Width           =   885
   End
   Begin VB.TextBox txtPayMonth 
      Height          =   300
      Left            =   2760
      TabIndex        =   8
      Top             =   2400
      Width           =   885
   End
   Begin VB.OptionButton optBy 
      Caption         =   "BYYEAR"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.OptionButton optBy 
      Caption         =   "BYMONTH"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optBy 
      Caption         =   "BYQUARTER"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   6
      Top             =   2910
      Width           =   1455
   End
   Begin VB.Frame fraRange 
      Caption         =   "RANGE"
      Height          =   615
      Left            =   840
      TabIndex        =   15
      Top             =   2160
      Width           =   3180
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2760
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin VB.ComboBox cboCusCode 
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1155
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   2520
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
            Picture         =   "frmSOP003.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP003.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   13
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
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   840
      TabIndex        =   16
      Top             =   2640
      Width           =   3180
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   840
      TabIndex        =   17
      Top             =   3120
      Width           =   3180
   End
   Begin VB.Label lblItmTypeCode 
      Caption         =   "METHODCODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   20
      Top             =   1635
      Width           =   1890
   End
   Begin VB.Label lblItmTypeCodeTo 
      Caption         =   "METHODCODETO"
      Height          =   225
      Left            =   5250
      TabIndex        =   19
      Top             =   1635
      Width           =   375
   End
   Begin VB.Label lblCusCodeTo 
      Caption         =   "METHODCODETO"
      Height          =   225
      Left            =   5280
      TabIndex        =   18
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   14
      Top             =   760
      Width           =   1860
   End
   Begin VB.Label lblCusCode 
      Caption         =   "METHODCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   11
      Top             =   1230
      Width           =   1890
   End
End
Attribute VB_Name = "frmSOP003"
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
    cboCusCode.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wiSel As Integer
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(3)
    wsSelection(1) = lblCusCode.Caption & " " & Set_Quote(cboCusCode.Text) & " " & lblCusCodeTo.Caption & " " & Set_Quote(cboCusCodeTo.Text)
    wsSelection(2) = lblItmTypeCode.Caption & " " & Set_Quote(cboItmTypeCode.Text) & " " & lblCusCodeTo.Caption & " " & Set_Quote(cboItmTypeCodeTo.Text)
    
    wiSel = Opt_Getfocus(optBy, 3, 0)
    
    If wiSel = 2 Then
        wsSelection(3) = optBy(2).Caption & " " & Set_Quote(txtPayYear.Text)
    ElseIf wiSel = 1 Then
        wsSelection(3) = optBy(1).Caption & " " & Set_Quote(txtPayQuarter.Text)
    ElseIf wiSel = 0 Then
        wsSelection(3) = optBy(0).Caption & " " & Set_Quote(txtPayMonth.Text)
    End If
    
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTSOP003 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCusCode.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCusCodeTo.Text) = "", String(10, "z"), Set_Quote(cboCusCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmTypeCode.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmTypeCodeTo.Text) = "", String(10, "z"), Set_Quote(cboItmTypeCodeTo.Text)) & "', "
    wsSQL = wsSQL & wiSel & ", "
    If wiSel = 2 Then
        wsSQL = wsSQL & Set_Quote(txtPayYear.Text) & ", "
    ElseIf wiSel = 1 Then
        wsSQL = wsSQL & Set_Quote(txtPayQuarter.Text) & ", "
    ElseIf wiSel = 0 Then
        wsSQL = wsSQL & Set_Quote(txtPayMonth.Text) & ", "
    End If
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTSOP003"
    Else
    wsRptName = "RPTSOP003"
    End If
    
    NewfrmPrint.ReportID = "SOP003"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SOP003"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusCode
    
    wsSQL = "SELECT CusCode, CusName FROM MstCustomer WHERE CusCode LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' AND CusStatus <>'2' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY CusCode "
    Call Ini_Combo(2, wsSQL, cboCusCode.Left, cboCusCode.Top + cboCusCode.Height, tblCommon, wsFormID, "TBLCusCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusCode_GotFocus()
    FocusMe cboCusCode
    Set wcCombo = cboCusCode
End Sub

Private Sub cboCusCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboCusCode.Text) <> "" And _
            Trim(cboCusCodeTo.Text) = "" Then
            
            cboCusCodeTo.Text = cboCusCode.Text
        End If
        
        cboCusCodeTo.SetFocus
    End If
End Sub

Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
End Sub

Private Sub cboCusCodeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusCodeTo
    
    wsSQL = "SELECT CusCode, CusName FROM MstCustomer WHERE CusCode LIKE '%" & IIf(cboCusCodeTo.SelLength > 0, "", Set_Quote(cboCusCodeTo.Text)) & "%' AND CusStatus <>'2' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY CusCode "
    Call Ini_Combo(2, wsSQL, cboCusCodeTo.Left, cboCusCodeTo.Top + cboCusCodeTo.Height, tblCommon, wsFormID, "TBLCusCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusCodeTo_GotFocus()
    FocusMe cboCusCodeTo
    Set wcCombo = cboCusCodeTo
End Sub

Private Sub cboCusCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusCodeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusCodeTo = False Then
            cboCusCodeTo.SetFocus
            Exit Sub
        End If
        
        cboItmTypeCode.SetFocus
       ' Call Opt_Setfocus(optBy, 3, 0)
    End If
End Sub

Private Sub cboCusCodeTo_LostFocus()
    FocusMe cboCusCodeTo, True
End Sub

Private Function chk_cboCusCodeTo() As Boolean
    chk_cboCusCodeTo = False
    
    If UCase(cboCusCode.Text) > UCase(cboCusCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboCusCodeTo = True
End Function

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
    
    wsFormID = "SOP003"
    
End Sub

Private Sub Ini_Scr()
    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    cboCusCode.Text = ""
    cboCusCodeTo.Text = ""
    cboItmTypeCode.Text = ""
    cboItmTypeCodeTo.Text = ""
    
    
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
    
    optBy(0).Value = True
    
    wgsTitle = "Sales Report"
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    'If chk_cboCusCodeTo = False Then
    '    cboCusCodeTo.SetFocus
    '    Exit Function
    'End If
    
    InputValidation = True
   
End Function

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 4275
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set wcCombo = Nothing
   Set frmSOP003 = Nothing

End Sub

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
    lblCusCode.Caption = Get_Caption(waScrItm, "CusCodefr")
    lblCusCodeTo.Caption = Get_Caption(waScrItm, "CusCodeTO")
    lblItmTypeCode.Caption = Get_Caption(waScrItm, "ItmTypeCodefr")
    lblItmTypeCodeTo.Caption = Get_Caption(waScrItm, "ItmTypeCodeTO")
    
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

Private Sub txtPayMonth_GotFocus()
    FocusMe txtPayMonth
End Sub

Private Sub txtPayMonth_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayMonth, False, False)
    Call chk_InpLen(txtPayMonth, 2, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayMonth Then
            cboCusCode.SetFocus
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

Private Sub txtTitle_GotFocus()
    FocusMe txtTitle
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtTitle, 60, KeyAscii)
 
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboCusCode.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
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
            cboCusCode.SetFocus
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
            cboCusCode.SetFocus
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

Private Function GetPrd() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 2
        If Me.optBy(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetPrd = "3"
            
        Case 1
            GetPrd = "2"
        
        Case 2
            GetPrd = "1"
    End Select
    
End Function

Private Sub cboItmTypeCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCode
    
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC")
    wsSQL = wsSQL & " FROM MstItemType WHERE ItmTypeCode LIKE '%" & IIf(cboItmTypeCode.SelLength > 0, "", Set_Quote(cboItmTypeCode.Text)) & "%' AND ItmTypeStatus <>'2' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeCode.Left, cboItmTypeCode.Top + cboItmTypeCode.Height, tblCommon, wsFormID, "TBLItmTypeCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
    
End Sub

Private Sub cboItmTypeCode_GotFocus()
    FocusMe cboItmTypeCode
    Set wcCombo = cboItmTypeCode
End Sub

Private Sub cboItmTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItmTypeCode.Text) <> "" And _
            Trim(cboItmTypeCodeTo.Text) = "" Then
            
            cboItmTypeCodeTo.Text = cboItmTypeCode.Text
        End If
        
        cboItmTypeCodeTo.SetFocus
    End If
End Sub

Private Sub cboItmTypeCode_LostFocus()
    FocusMe cboItmTypeCode, True
End Sub

Private Sub cboItmTypeCodeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeTo
    
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC")
    wsSQL = wsSQL & " FROM MstItemType WHERE ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeTo.SelLength > 0, "", Set_Quote(cboItmTypeCodeTo.Text)) & "%' AND ItmTypeStatus <>'2' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeCodeTo.Left, cboItmTypeCodeTo.Top + cboItmTypeCodeTo.Height, tblCommon, wsFormID, "TBLItmTypeCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmTypeCodeTo_GotFocus()
    FocusMe cboItmTypeCodeTo
    Set wcCombo = cboItmTypeCodeTo
End Sub

Private Sub cboItmTypeCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCodeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmTypeCodeTo = False Then
            cboItmTypeCodeTo.SetFocus
            Exit Sub
        End If
        
        Call Opt_Setfocus(optBy, 3, 0)
    End If
End Sub

Private Sub cboItmTypeCodeTo_LostFocus()
    FocusMe cboItmTypeCodeTo, True
End Sub

Private Function chk_cboItmTypeCodeTo() As Boolean
    chk_cboItmTypeCodeTo = False
    
    If UCase(cboItmTypeCode.Text) > UCase(cboItmTypeCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboItmTypeCodeTo = True
End Function
