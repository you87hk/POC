VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmICP001 
   Caption         =   "ICP001"
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
      Left            =   9000
      OleObjectBlob   =   "frmICP001.frx":0000
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboICTrnType 
      Height          =   300
      Left            =   2760
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1812
   End
   Begin VB.ComboBox cboICTrnCode 
      Height          =   300
      Left            =   2760
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1812
   End
   Begin VB.ComboBox cboICSrcCode 
      Height          =   300
      Left            =   2760
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1560
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
            Picture         =   "frmICP001.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmICP001.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   5
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
   Begin VB.Label lblICTrnType 
      Caption         =   "ICTRNTYPE"
      Height          =   225
      Left            =   840
      TabIndex        =   13
      Top             =   2310
      Width           =   1890
   End
   Begin VB.Label lblICTrnCode 
      Caption         =   "ICTRNCODE"
      Height          =   225
      Left            =   840
      TabIndex        =   11
      Top             =   1950
      Width           =   1890
   End
   Begin VB.Label lblICSrcCode 
      Caption         =   "ICSRCCODE"
      Height          =   225
      Left            =   840
      TabIndex        =   9
      Top             =   1590
      Width           =   1890
   End
   Begin VB.Label lblItmCodeTo 
      Caption         =   "ITMCODETO"
      Height          =   225
      Left            =   4920
      TabIndex        =   7
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label lblItmCodeFr 
      Caption         =   "ITMCODEFR"
      Height          =   225
      Left            =   840
      TabIndex        =   6
      Top             =   1245
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Caption         =   "TITLE"
      Height          =   240
      Left            =   870
      TabIndex        =   4
      Top             =   760
      Width           =   1860
   End
End
Attribute VB_Name = "frmICP001"
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
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = lblItmCodeFr.Caption & " " & Set_Quote(cboItmCodeFr.Text) & " " & lblItmCodeTo.Caption & " " & Set_Quote(cboItmCodeTo.Text)
    wsSelection(2) = lblICSrcCode.Caption & " " & Set_Quote(cboICSrcCode.Text)
    wsSelection(3) = lblICTrnCode.Caption & " " & Set_Quote(cboICTrnCode.Text)
    wsSelection(4) = lblICTrnType.Caption & " " & Set_Quote(cboICTrnType.Text)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTICP001 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboItmCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), Set_Quote(cboItmCodeTo.Text)) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboICSrcCode.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboICTrnCode.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboICTrnType.Text) & "', "
    wsSql = wsSql & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTICP001"
    Else
    wsRptName = "RPTICP001"
    End If
    
    NewfrmPrint.ReportID = "ICP001"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "ICP001"
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
    
    wsFormID = "ICP001"
    
End Sub

Private Sub Ini_Scr()

    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    cboItmCodeFr.Text = ""
    cboItmCodeTo.Text = ""
    cboICSrcCode.Text = ""
    cboICTrnCode.Text = ""
    cboICTrnType.Text = ""
   
    wgsTitle = "Stock Transaction Register"
    
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboICSRCCode = False Then
        cboICSrcCode.SetFocus
        Exit Function
    End If
    
    If chk_cboICTrnCode = False Then
        cboICTrnCode.SetFocus
        Exit Function
    End If
    
    If Chk_cboICTrnType = False Then
        cboICTrnType.SetFocus
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
    Set frmICP001 = Nothing
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
    
    lblICSrcCode.Caption = Get_Caption(waScrItm, "ICSRCCODE")
    lblICTrnCode.Caption = Get_Caption(waScrItm, "ICTRNCODE")
    lblICTrnType.Caption = Get_Caption(waScrItm, "ICTRNTYPE")
    
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
  
    wsSql = "SELECT ItmCode, ItmBarCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & " "
    wsSql = wsSql & " FROM mstItem "
    wsSql = wsSql & " WHERE ItmCode LIKE '%" & IIf(cboItmCodeFr.SelLength > 0, "", Set_Quote(cboItmCodeFr.Text)) & "%' "
    wsSql = wsSql & " AND ItmSTATUS = '1' "
    wsSql = wsSql & " ORDER BY ItmCode "
    
    Call Ini_Combo(3, wsSql, cboItmCodeFr.Left, cboItmCodeFr.Top + cboItmCodeFr.Height, tblCommon, wsFormID, "TBLItmCode", Me.Width, Me.Height)
    
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
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeTo
  
    wsSql = "SELECT ItmCode, ItmBarCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & " "
    wsSql = wsSql & " FROM mstItem "
    wsSql = wsSql & " WHERE ItmCode LIKE '%" & IIf(cboItmCodeTo.SelLength > 0, "", Set_Quote(cboItmCodeTo.Text)) & "%' "
    wsSql = wsSql & " AND ItmSTATUS = '1' "
    wsSql = wsSql & " ORDER BY ItmCode "
    
    Call Ini_Combo(3, wsSql, cboItmCodeTo.Left, cboItmCodeTo.Top + cboItmCodeTo.Height, tblCommon, wsFormID, "TBLItmCode", Me.Width, Me.Height)
    
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
        
       cboICSrcCode.SetFocus
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

Private Sub cboICSRCCode_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    wsSql = "SELECT ICSRCCODE, SCDDESC "
    wsSql = wsSql & " FROM ICINVENTORY, SYSCODEDESC "
    wsSql = wsSql & " WHERE SCDLANGID = " & gsLangID
    wsSql = wsSql & " AND SCDCODE = ICSRCCODE "
    wsSql = wsSql & " GROUP BY ICSRCCODE, SCDDESC "
    
    Call Ini_Combo(2, wsSql, cboICSrcCode.Left, cboICSrcCode.Top + cboICSrcCode.Height, tblCommon, wsFormID, "TBLICSRCCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboICSRCCode_GotFocus()
        FocusMe cboICSrcCode
    Set wcCombo = cboICSrcCode
End Sub

Private Sub cboICSRCCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboICSrcCode, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboICSRCCode = False Then Exit Sub
        
        cboICTrnCode.Text = ""
        cboICTrnCode.SetFocus
        
    End If
End Sub


Private Sub cboICSRCCode_LostFocus()
    FocusMe cboICSrcCode, True
End Sub

Private Sub cboICTrnCode_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboICTrnCode
  
    wsSql = "SELECT ICTRNCODE, SCDDESC "
    wsSql = wsSql & " FROM ICINVENTORY, SYSCODEDESC "
    wsSql = wsSql & " WHERE SCDLANGID = " & gsLangID
    wsSql = wsSql & " AND ICSRCCODE = '" & Set_Quote(cboICSrcCode.Text) & "'"
    wsSql = wsSql & " AND SCDCODE = ICTRNCODE "
    wsSql = wsSql & " GROUP BY ICTRNCODE, SCDDESC "
    
    
    Call Ini_Combo(2, wsSql, cboICTrnCode.Left, cboICTrnCode.Top + cboICTrnCode.Height, tblCommon, wsFormID, "TBLICTRNCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboICTrnCode_GotFocus()
    FocusMe cboICTrnCode
    Set wcCombo = cboICTrnCode
End Sub

Private Sub cboICTrnCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboICTrnCode, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboICTrnCode = False Then Exit Sub
        
        cboICTrnType = ""
        cboICTrnType.SetFocus
        
    End If
End Sub

Private Sub cboICTrnCode_LostFocus()
    FocusMe cboICTrnCode, True
End Sub

Private Function chk_cboICTrnCode() As Boolean
    chk_cboICTrnCode = False
    
    If Trim(cboICTrnCode.Text) = "" Then
        gsMsg = "Must Input Transaction Code!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboICTrnCode.SetFocus
        Exit Function
    End If
    
    chk_cboICTrnCode = True
End Function

Private Sub cboICTrnType_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboICTrnType
  
    wsSql = "SELECT ICTRNTYPE, SCDDESC "
    wsSql = wsSql & " FROM ICINVENTORY, SYSCODEDESC "
    wsSql = wsSql & " WHERE SCDLANGID = " & gsLangID
    wsSql = wsSql & " AND ICSRCCODE = '" & Set_Quote(cboICSrcCode.Text) & "'"
    wsSql = wsSql & " AND ICTRNCODE = '" & Set_Quote(cboICTrnCode.Text) & "'"
    wsSql = wsSql & " AND SCDCODE = ICTRNTYPE "
    wsSql = wsSql & " GROUP BY ICTRNTYPE, SCDDESC "
    
    Call Ini_Combo(2, wsSql, cboICTrnType.Left, cboICTrnType.Top + cboICTrnType.Height, tblCommon, wsFormID, "TBLICTRNTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboICTrnType_GotFocus()
    FocusMe cboICTrnType
    Set wcCombo = cboICTrnType
End Sub

Private Sub cboICTrnType_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboICTrnType, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboICTrnType = False Then Exit Sub
        
        cboItmCodeFr.SetFocus
    End If
End Sub

Private Sub cboICTrnType_LostFocus()
    FocusMe cboICTrnType, True
End Sub
        
Private Function Chk_cboICTrnType() As Boolean
    Chk_cboICTrnType = False
    
    If Trim(cboICTrnType.Text) = "" Then
        gsMsg = "Must Input Transaction Type!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboICTrnType.SetFocus
        Exit Function
    End If
    
    Chk_cboICTrnType = True
End Function

Private Function chk_cboICSRCCode() As Boolean
    chk_cboICSRCCode = False
    
    If Trim(cboICSrcCode.Text) = "" Then
        gsMsg = "Must IC Source Code!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboICSrcCode.SetFocus
        Exit Function
    End If
    
    chk_cboICSRCCode = True
End Function

