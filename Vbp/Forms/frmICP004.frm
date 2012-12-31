VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmICP004 
   Caption         =   "ICP003"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   Begin VB.CheckBox chkShowBin 
      Alignment       =   1  '靠右對齊
      Caption         =   "暫停發貨 :"
      Height          =   180
      Left            =   3480
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   7920
      OleObjectBlob   =   "frmICP004.frx":0000
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.CheckBox chkNoZero 
      Alignment       =   1  '靠右對齊
      Caption         =   "暫停發貨 :"
      Height          =   180
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.ComboBox cboItmTypeCodeTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeCodeFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1812
   End
   Begin VB.ComboBox cboWhsCodeFr 
      Height          =   300
      Left            =   2760
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1812
   End
   Begin VB.ComboBox cboWhsCodeTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1812
   End
   Begin VB.ComboBox cboItmCodeTo 
      Height          =   300
      Left            =   5520
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
            Picture         =   "frmICP004.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP004.frx":607B
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
   Begin VB.Label lblItmTypeCodeFr 
      Caption         =   "ITMTYPECODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   16
      Top             =   1680
      Width           =   1890
   End
   Begin VB.Label lblItmTypeCodeTo 
      Caption         =   "ITMTYPECODETO"
      Height          =   225
      Left            =   4920
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblWhsCodeFr 
      Caption         =   "WHSCODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   14
      Top             =   2205
      Width           =   1890
   End
   Begin VB.Label lblWhsCodeTo 
      Caption         =   "WHSCODETO"
      Height          =   225
      Left            =   4920
      TabIndex        =   13
      Top             =   2205
      Width           =   615
   End
   Begin VB.Label lblItmCodeTo 
      Caption         =   "ITMCODETO"
      Height          =   225
      Left            =   4920
      TabIndex        =   12
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label lblItmCodeFr 
      Caption         =   "ITMCODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   11
      Top             =   1245
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   9
      Top             =   760
      Width           =   1860
   End
End
Attribute VB_Name = "frmICP004"
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
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(3)
    wsSelection(1) = lblItmCodeFr.Caption & " " & Set_Quote(cboItmCodeFr.Text) & " " & lblItmCodeTo.Caption & " " & Set_Quote(cboItmCodeTo.Text)
    wsSelection(2) = lblItmTypeCodeFr.Caption & " " & Set_Quote(Me.cboItmTypeCodeFr.Text) & " " & lblItmTypeCodeTo.Caption & " " & Set_Quote(cboItmTypeCodeTo.Text)
    wsSelection(3) = lblWhsCodeFr.Caption & " " & Set_Quote(cboWhsCodeFr.Text) & " " & lblWhsCodeTo.Caption & " " & Set_Quote(cboWhsCodeTo.Text)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTICP004 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmCodeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), Set_Quote(cboItmCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmTypeCodeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmTypeCodeTo.Text) = "", String(10, "z"), Set_Quote(cboItmTypeCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboWhsCodeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboWhsCodeTo.Text) = "", String(10, "z"), Set_Quote(cboWhsCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(chkNoZero.Value = 1, "Y", "N") & "', "
    wsSQL = wsSQL & "'" & IIf(chkShowBin.Value = 1, "Y", "N") & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTICP004"
    Else
    wsRptName = "RPTICP004"
    End If
    
    If chkShowBin.Value = 1 Then
    wsRptName = wsRptName & "B"
    End If
    
    NewfrmPrint.ReportID = "ICP004"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "ICP004"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub


Private Sub chkNoZero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboItmCodeFr.SetFocus
    End If
End Sub

Private Sub chkShowBin_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboItmCodeFr.SetFocus
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
    
    wsFormID = "ICP004"
    
End Sub

Private Sub Ini_Scr()

    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    cboItmCodeFr.Text = ""
    cboItmCodeTo.Text = ""

    cboItmTypeCodeFr.Text = ""
    cboItmTypeCodeTo.Text = ""
    cboWhsCodeFr.Text = Get_WorkStation_Info("WSWHSCODE")
    cboWhsCodeTo.Text = cboWhsCodeFr.Text
    chkNoZero.Value = 1
    chkShowBin.Value = 0
   
    wgsTitle = "Stock In/Out Ledger"
    
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
        Me.Height = 3840
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set waScrItm = Nothing
    Set wcCombo = Nothing
    Set frmICP004 = Nothing
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
    lblWhsCodeFr.Caption = Get_Caption(waScrItm, "WHSCODEFR")
    lblWhsCodeTo.Caption = Get_Caption(waScrItm, "WHSCODETO")
    chkNoZero.Caption = Get_Caption(waScrItm, "NOZERO")
    chkShowBin.Caption = Get_Caption(waScrItm, "SHOWBIN")
    
    
    
    
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
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeFr
  
    wsSQL = "SELECT ItmCode, ItmItmTypeCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & " "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmCode LIKE '%" & IIf(cboItmCodeFr.SelLength > 0, "", Set_Quote(cboItmCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    
    Call Ini_Combo(3, wsSQL, cboItmCodeFr.Left, cboItmCodeFr.Top + cboItmCodeFr.Height, tblCommon, wsFormID, "TBLItmCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeFr_GotFocus()
    FocusMe cboItmCodeFr
    Set wcCombo = cboItmCodeFr
End Sub

Private Sub cboItmCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeFr, 30, KeyAscii)
    
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
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeTo
  
    wsSQL = "SELECT ItmCode, ItmItmTypeCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & " "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmCode LIKE '%" & IIf(cboItmCodeTo.SelLength > 0, "", Set_Quote(cboItmCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    
    Call Ini_Combo(3, wsSQL, cboItmCodeTo.Left, cboItmCodeTo.Top + cboItmCodeTo.Height, tblCommon, wsFormID, "TBLItmCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeTo_GotFocus()
    FocusMe cboItmCodeTo
    Set wcCombo = cboItmCodeTo
End Sub

Private Sub cboItmCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeTo, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmCodeTo = False Then
            Call cboItmCodeTo_GotFocus
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
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboItmCodeTo = True
End Function


Private Sub cboItmTypeCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeFr
  
    If gsLangID = "1" Then
    wsSQL = "SELECT ItmTypeCode, ItmTypeEngDesc FROM MstItemType "
    Else
    wsSQL = "SELECT ItmTypeCode, ItmTypeChiDesc FROM MstItemType "
    End If
    wsSQL = wsSQL & " WHERE ItmTypeStatus = '1'"
    wsSQL = wsSQL & " AND ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeFr.SelLength > 0, "", Set_Quote(cboItmTypeCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeCodeFr.Left, cboItmTypeCodeFr.Top + cboItmTypeCodeFr.Height, tblCommon, wsFormID, "TBLItmTypeCode", Me.Width, Me.Height)
    
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
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeTo
  
    If gsLangID = "1" Then
    wsSQL = "SELECT ItmTypeCode, ItmTypeEngDesc FROM MstItemType "
    Else
    wsSQL = "SELECT ItmTypeCode, ItmTypeChiDesc FROM MstItemType "
    End If
    wsSQL = wsSQL & " WHERE ItmTypeStatus = '1'"
    wsSQL = wsSQL & " AND ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeTo.SelLength > 0, "", Set_Quote(cboItmTypeCodeTo.Text)) & "%' "
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
    Call chk_InpLen(cboItmTypeCodeTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmTypeCodeTo = False Then
            Call cboItmTypeCodeTo_GotFocus
            Exit Sub
        End If
        
        cboWhsCodeFr.SetFocus
        
        
    End If
End Sub

Private Sub cboItmTypeCodeTo_LostFocus()
    FocusMe cboItmTypeCodeTo, True
End Sub

Private Function chk_cboItmTypeCodeTo() As Boolean
    chk_cboItmTypeCodeTo = False
    
    If UCase(cboItmTypeCodeFr.Text) > UCase(cboItmTypeCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboItmTypeCodeTo = True
End Function

Private Sub cboWhsCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWhsCodeFr
  
    wsSQL = "SELECT WhsCode, WhsDesc FROM MstWarehouse "
    wsSQL = wsSQL & " WHERE WhsStatus = '1'"
    wsSQL = wsSQL & " AND WhsCode LIKE '%" & IIf(cboWhsCodeFr.SelLength > 0, "", Set_Quote(cboWhsCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY WhsCode "
    
    Call Ini_Combo(2, wsSQL, cboWhsCodeFr.Left, cboWhsCodeFr.Top + cboWhsCodeFr.Height, tblCommon, wsFormID, "TBLWhsCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboWhsCodeFr_GotFocus()
    FocusMe cboWhsCodeFr
    Set wcCombo = cboWhsCodeFr
End Sub

Private Sub cboWhsCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWhsCodeFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboWhsCodeFr.Text) <> "" And _
            Trim(cboWhsCodeTo.Text) = "" Then
            cboWhsCodeTo.Text = cboWhsCodeFr.Text
        End If
        cboWhsCodeTo.SetFocus
    End If
End Sub

Private Sub cboWhsCodeFr_LostFocus()
    FocusMe cboWhsCodeFr, True
End Sub

Private Sub cboWhsCodeTo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWhsCodeTo
  
    wsSQL = "SELECT WhsCode, WhsDesc FROM MstWarehouse "
    wsSQL = wsSQL & " WHERE WhsStatus = '1'"
    wsSQL = wsSQL & " AND WhsCode LIKE '%" & IIf(cboWhsCodeTo.SelLength > 0, "", Set_Quote(cboWhsCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY WhsCode "
    Call Ini_Combo(2, wsSQL, cboWhsCodeTo.Left, cboWhsCodeTo.Top + cboWhsCodeTo.Height, tblCommon, wsFormID, "TBLWhsCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboWhsCodeTo_GotFocus()
    FocusMe cboWhsCodeTo
    Set wcCombo = cboWhsCodeTo
End Sub

Private Sub cboWhsCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWhsCodeTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboWhsCodeTo = False Then
            cboWhsCodeTo.SetFocus
            Exit Sub
        End If
        
        chkNoZero.SetFocus
        
    End If
End Sub

Private Sub cboWhsCodeTo_LostFocus()
    FocusMe cboWhsCodeTo, True
End Sub

Private Function chk_cboWhsCodeTo() As Boolean
    chk_cboWhsCodeTo = False
    
    If UCase(cboWhsCodeFr.Text) > UCase(cboWhsCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboWhsCodeTo = True
End Function


