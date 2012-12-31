VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPR002 
   Caption         =   "PR002"
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
      OleObjectBlob   =   "frmPR002.frx":0000
      TabIndex        =   5
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
   Begin VB.ComboBox cboPrcCodeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1275
      Width           =   1812
   End
   Begin VB.ComboBox cboPrcCodeFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1275
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
            Picture         =   "frmPR002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPR002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   760
      Width           =   1860
   End
   Begin VB.Label lblPrcCodeTo 
      Caption         =   "PRCCODETO"
      Height          =   225
      Left            =   5220
      TabIndex        =   4
      Top             =   1305
      Width           =   375
   End
   Begin VB.Label lblPrcCodeFr 
      Caption         =   "PRCCODEFR"
      Height          =   225
      Left            =   870
      TabIndex        =   3
      Top             =   1305
      Width           =   1890
   End
End
Attribute VB_Name = "frmPR002"
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
    cboPrcCodeFr.SetFocus
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
    ReDim wsSelection(1)
    wsSelection(1) = lblPrcCodeFr.Caption & " " & Set_Quote(cboPrcCodeFr.Text) & " " & lblPrcCodeTo.Caption & " " & Set_Quote(cboPrcCodeTo.Text)
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTPR002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & Set_Quote(txtTitle.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboPrcCodeFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboPrcCodeTo.Text) = "", String(10, "z"), Set_Quote(cboPrcCodeTo.Text)) & "', "
    wsSql = wsSql & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTPR002"
    Else
    wsRptName = "RPTPR002"
    End If
    
    NewfrmPrint.ReportID = "PR002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "PR002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboPrcCodeFr_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboPrcCodeFr
    
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT PrcCode, PrcDesc FROM MstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCodeFr.SelLength > 0, "", Set_Quote(cboPrcCodeFr.Text)) & "%' AND PrcStatus <>'2' "
        Case "2"
            wsSql = "SELECT PrcCode, PrcDesc FROM MstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCodeFr.SelLength > 0, "", Set_Quote(cboPrcCodeFr.Text)) & "%' AND PrcStatus <>'2' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY PrcCode "
    Call Ini_Combo(2, wsSql, cboPrcCodeFr.Left, cboPrcCodeFr.Top + cboPrcCodeFr.Height, tblCommon, wsFormID, "TBLPrcCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboPrcCodeFr_GotFocus()
    FocusMe cboPrcCodeFr
    Set wcCombo = cboPrcCodeFr
End Sub

Private Sub cboPrcCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboPrcCodeFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboPrcCodeFr.Text) <> "" And _
            Trim(cboPrcCodeTo.Text) = "" Then
            
            cboPrcCodeTo.Text = cboPrcCodeFr.Text
        End If
        cboPrcCodeTo.SetFocus
    End If
End Sub

Private Sub cboPrcCodeFr_LostFocus()
    FocusMe cboPrcCodeFr, True
End Sub

Private Sub cboPrcCodeTo_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboPrcCodeTo
    
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT PrcCode, PrcDesc FROM MstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCodeTo.SelLength > 0, "", Set_Quote(cboPrcCodeTo.Text)) & "%' AND PrcStatus <>'2' "
        Case "2"
            wsSql = "SELECT PrcCode, PrcDesc FROM MstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCodeTo.SelLength > 0, "", Set_Quote(cboPrcCodeTo.Text)) & "%' AND PrcStatus <>'2' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY PrcCode "
    Call Ini_Combo(2, wsSql, cboPrcCodeTo.Left, cboPrcCodeTo.Top + cboPrcCodeTo.Height, tblCommon, wsFormID, "TBLPrcCODE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboPrcCodeTo_GotFocus()
    FocusMe cboPrcCodeTo
    Set wcCombo = cboPrcCodeTo
End Sub

Private Sub cboPrcCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboPrcCodeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboPrcCodeTo = False Then
            cboPrcCodeTo.SetFocus
            Exit Sub
        End If
        
        cboPrcCodeFr.SetFocus
    End If
End Sub

Private Sub cboPrcCodeTo_LostFocus()
    FocusMe cboPrcCodeTo, True
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
    
    wsFormID = "PR002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboPrcCodeFr.Text = ""
   cboPrcCodeTo.Text = ""
   
   wgsTitle = "Prc Term List"
    
End Sub

Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboPrcCodeTo = False Then
        cboPrcCodeTo.SetFocus
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
   Set frmPR002 = Nothing

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
    lblPrcCodeFr.Caption = Get_Caption(waScrItm, "PRCCODEFR")
    lblPrcCodeTo.Caption = Get_Caption(waScrItm, "PRCCODETO")
    
End Sub

Private Function chk_cboPrcCodeTo() As Boolean
    chk_cboPrcCodeTo = False
    
    If UCase(cboPrcCodeFr.Text) > UCase(cboPrcCodeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboPrcCodeTo = True
End Function

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
        
        cboPrcCodeFr.SetFocus
    End If
End Sub

Private Sub txtTitle_LostFocus()
    FocusMe txtTitle, True
End Sub
