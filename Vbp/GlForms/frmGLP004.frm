VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGLP004 
   Caption         =   "Material Master List"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmGLP004.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   2880
      TabIndex        =   0
      Text            =   "01234567890123457890"
      Top             =   720
      Width           =   4665
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   1920
      OleObjectBlob   =   "frmGLP004.frx":030A
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Frame Frame1 
      Height          =   492
      Left            =   2880
      TabIndex        =   13
      Top             =   1920
      Width           =   4455
      Begin VB.CheckBox chkPgeBrk 
         Alignment       =   1  '靠右對齊
         Caption         =   "New Page with Each Customer:"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.ComboBox cboAccNoFr 
      Height          =   300
      Left            =   2880
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin VB.ComboBox cboAccNoTo 
      Height          =   300
      Left            =   5700
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1110
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   240
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
            Picture         =   "frmGLP004.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGLP004.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
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
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   5700
      TabIndex        =   4
      Top             =   1560
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
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
      TabIndex        =   14
      Top             =   780
      Width           =   1860
   End
   Begin VB.Label lblPrtMrk 
      Caption         =   "Print Range"
      Height          =   255
      Left            =   990
      TabIndex        =   12
      Top             =   2040
      Width           =   1800
   End
   Begin VB.Label lblAccNoFr 
      Caption         =   "Document # From"
      Height          =   225
      Left            =   990
      TabIndex        =   10
      Top             =   1170
      Width           =   1890
   End
   Begin VB.Label lblAccNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   9
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5340
      TabIndex        =   7
      Top             =   1605
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   990
      TabIndex        =   6
      Top             =   1605
      Width           =   1890
   End
End
Attribute VB_Name = "frmGLP004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wcCombo As Control
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String


Private Sub cmdCancel()
    Ini_Scr
    cboAccNoFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsCtlFr As String
    Dim wsCtlTo As String
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    wsCtlFr = Get_FiscalPeriod(medPrdFr & "/01")
    wsCtlTo = Get_FiscalPeriod(medPrdTo & "/01")
    
    'Create Selection Criteria
    ReDim wsSelection(3)
    wsSelection(1) = lblAccNoFr.Caption & " " & Set_Quote(cboAccNoFr.Text) & " " & lblAccNoTo.Caption & " " & Set_Quote(cboAccNoTo.Text)
    wsSelection(2) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    wsSelection(3) = lblPrtMrk.Caption & " " & IIf(chkPgeBrk.Value = 1, "N", "Y")
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSql = "EXEC usp_RPTGLP004 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & txtTitle.Text & "', "
    wsSql = wsSql & "'" & Set_Quote(cboAccNoFr.Text) & "', "
    wsSql = wsSql & "'" & IIf(Trim(cboAccNoTo.Text) = "", String(15, "z"), Set_Quote(cboAccNoTo.Text)) & "', "
    wsSql = wsSql & "'" & wsCtlFr & "', "
    wsSql = wsSql & "'" & wsCtlTo & "', "
    wsSql = wsSql & "'" & IIf(chkPgeBrk.Value = 1, "N", "Y") & "', "
    wsSql = wsSql & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTGLP004"
    Else
    wsRptName = "RPTGLP004"
    End If
    
    If chkPgeBrk.Value = 1 Then wsRptName = wsRptName + "P"
    
    NewfrmPrint.ReportID = "GLP004"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "GLP004"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboAccNoTo_LostFocus()
    FocusMe cboAccNoTo, True
End Sub



Private Sub chkPgeBrk_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        cboAccNoFr.SetFocus
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
    
    wsFormID = "GLP004"
    
End Sub

Private Sub Ini_Scr()
Dim wsCtlPrd As String
   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboAccNoFr.Text = ""
   cboAccNoTo.Text = ""
   Call SetPeriodMask(medPrdFr)
   Call SetPeriodMask(medPrdTo)
   wsCtlPrd = getCtrlMth("GL")
   medPrdFr.Text = Left(wsCtlPrd, 4) & "/" & Right(wsCtlPrd, 2)
   medPrdTo.Text = medPrdFr.Text
   chkPgeBrk.Value = 1

End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboAccNoTo = False Then
        cboAccNoTo.SetFocus
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
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmGLP004 = Nothing

End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
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
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblAccNoFr.Caption = Get_Caption(waScrItm, "ACCNoFR")
    lblAccNoTo.Caption = Get_Caption(waScrItm, "ACCNoTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblPrtMrk.Caption = Get_Caption(waScrItm, "PRTMRK")
    chkPgeBrk.Caption = Get_Caption(waScrItm, "PGEBRK")
    
    lblTitle = Get_Caption(waScrItm, "LBLTITLE")
    
    txtTitle.Text = Get_Caption(waScrItm, "TITLE")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
End Sub



Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    
    If Chk_Period(medPrdFr) = False Then
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    
    End If
    
    
    If chk_RetPrd(medPrdFr) = False Then
        wsMsg = "Period Must > Minimin Date!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_RetPrd(inMedDte As Date) As Boolean
    Dim wsStrDte As String
    Dim wsEndDte As String
    Dim wsCtlDte As String
    Dim wiRetValue As Integer
    Dim wdMinDte As Date
    chk_RetPrd = False

    wiRetValue = Get_TableInfo("sysMonCtl", "MCMODNO = 'GL'", "MCKeepMn")
    
    wsCtlDte = getCtrlMth("GL", wsStrDte, wsEndDte)
    wdMinDte = DateAdd("M", To_Value(wiRetValue) * -1, (CDate(wsStrDte)))
    If CDate(inMedDte) < wdMinDte Then Exit Function
    
    chk_RetPrd = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
     If Chk_Period(medPrdTo) = False Then
        wsMsg = "Wrong Period!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        wsMsg = "To must > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If CDate(medPrdFr.Text) > gsDateTo Then
        wsMsg = "Period Must < Max Date!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
   
    
    chk_medPrdTo = True
End Function


Private Function chk_cboAccNoTo() As Boolean
    chk_cboAccNoTo = False
    
    If UCase(cboAccNoFr.Text) > UCase(cboAccNoTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        cboAccNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboAccNoTo = True
End Function
Private Sub cboAccNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboAccNoFr
  
    wsSql = "SELECT COAACCCODE, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " "
    wsSql = wsSql & " FROM mstCOA "
    wsSql = wsSql & " WHERE COAAccCode LIKE '%" & IIf(cboAccNoFr.SelLength > 0, "", Set_Quote(cboAccNoFr.Text)) & "%' "
    wsSql = wsSql & " AND COASTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY COAAccCode "
    Call Ini_Combo(2, wsSql, cboAccNoFr.Left, cboAccNoFr.Top + cboAccNoFr.Height, tblCommon, wsFormID, "TBLACCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboAccNoFr_GotFocus()
    FocusMe cboAccNoFr
    Set wcCombo = cboAccNoFr
End Sub

Private Sub cboAccNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboAccNoFr.Text) <> "" And _
            Trim(cboAccNoTo.Text) = "" Then
            cboAccNoTo.Text = cboAccNoFr.Text
        End If
        cboAccNoTo.SetFocus
    End If
End Sub

Private Sub cboAccNoFr_LostFocus()
    FocusMe cboAccNoFr, True
End Sub

Private Sub cboAccNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboAccNoTo
  
    wsSql = "SELECT COAACCCODE, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " "
    wsSql = wsSql & " FROM mstCOA "
    wsSql = wsSql & " WHERE COAAccCode LIKE '%" & IIf(cboAccNoTo.SelLength > 0, "", Set_Quote(cboAccNoTo.Text)) & "%' "
    wsSql = wsSql & " AND COASTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY COAAccCode "
   Call Ini_Combo(2, wsSql, cboAccNoTo.Left, cboAccNoTo.Top + cboAccNoTo.Height, tblCommon, wsFormID, "TBLAccNo", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboAccNoTo_GotFocus()
    FocusMe cboAccNoTo
    Set wcCombo = cboAccNoTo
End Sub

Private Sub cboAccNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboAccNoTo = False Then
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub

Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
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
            Exit Sub
        End If
        chkPgeBrk.SetFocus
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


Private Sub txtTitle_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccNoFr, 50, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        cboAccNoFr.SetFocus
    End If
End Sub
