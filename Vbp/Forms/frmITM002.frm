VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmITM002 
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
      Left            =   1800
      OleObjectBlob   =   "frmITM002.frx":0000
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.ComboBox cboItmClassFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboItmClassTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   744
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   744
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1150
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1150
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
            Picture         =   "frmITM002.frx":2703
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":2FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":38B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":3D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":415B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":4475
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":48C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":4D19
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":5033
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":534D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":579F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM002.frx":607B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   15
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
      Left            =   5580
      TabIndex        =   7
      Top             =   1920
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
      Left            =   2790
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblItmClassFr 
      Caption         =   "Item Category Code From"
      Height          =   225
      Left            =   870
      TabIndex        =   17
      Top             =   1620
      Width           =   1890
   End
   Begin VB.Label lblItmClassTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   16
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "ISBN # From"
      Height          =   228
      Left            =   870
      TabIndex        =   14
      Top             =   800
      Width           =   1884
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   228
      Left            =   5220
      TabIndex        =   13
      Top             =   800
      Width           =   372
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   11
      Top             =   1980
      Width           =   375
   End
   Begin VB.Label lblItmTypeTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   10
      Top             =   1180
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   870
      TabIndex        =   9
      Top             =   1980
      Width           =   1890
   End
   Begin VB.Label lblItmTypeFr 
      Caption         =   "Item Type Code From"
      Height          =   225
      Left            =   870
      TabIndex        =   8
      Top             =   1180
      Width           =   1890
   End
End
Attribute VB_Name = "frmITM002"
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
    cboDocNoFr.SetFocus
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
    ReDim wsSelection(4)
    wsSelection(1) = lblDocNoFr.Caption & " " & Set_Quote(cboDocNoFr.Text) & " " & lblDocNoTo.Caption & " " & Set_Quote(cboDocNoTo.Text)
    wsSelection(2) = lblItmTypeFr.Caption & " " & Set_Quote(cboItmTypeFr.Text) & " " & lblItmTypeTo.Caption & " " & Set_Quote(cboItmTypeTo.Text)
    wsSelection(3) = lblItmClassFr.Caption & " " & Set_Quote(cboItmClassFr.Text) & " " & lblItmClassTo.Caption & " " & Set_Quote(cboItmClassTo.Text)
    wsSelection(4) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTITM002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(13, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmTypeFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmTypeTo.Text) = "", String(10, "z"), Set_Quote(cboItmTypeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmClassFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmClassTo.Text) = "", String(10, "z"), Set_Quote(cboItmClassTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdFr.Text) = "/  /", "", medPrdFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdTo.Text) = "/  /", "99999999", medPrdTo.Text) & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
    wsRptName = "C" & "RPTITM002"
    Else
    wsRptName = "RPTITM002"
    End If
    
    NewfrmPrint.ReportID = "ITM002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "ITM002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmClassFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmClassFr
    
    wsSQL = "SELECT ItemClassCode, " & IIf(gsLangID = "1", "ItemClassEDesc", "ItemClassCDesc") & " FROM MstItemClass WHERE ItemClassCode LIKE '%" & IIf(cboItmClassFr.SelLength > 0, "", Set_Quote(cboItmClassFr.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY ItemClassCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeFr.Left, cboItmTypeFr.Top + cboItmTypeFr.Height, tblCommon, wsFormID, "TBLItmClass", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmClassFr_GotFocus()
    FocusMe cboItmClassFr
End Sub

Private Sub cboItmClassFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmClassFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItmClassFr.Text) <> "" And _
            Trim(cboItmClassTo.Text) = "" Then
            cboItmClassTo.Text = cboItmClassFr.Text
        End If
        
        cboItmClassTo.SetFocus
    End If

End Sub

Private Sub cboItmClassFr_LostFocus()
    FocusMe cboItmClassFr, True
End Sub

Private Sub cboItmClassTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmClassTo
    
    wsSQL = "SELECT ItemClassCode, " & IIf(gsLangID = "1", "ItemClassEDesc", "ItemClassCDesc") & " FROM MstItemClass WHERE ItemClassCode LIKE '%" & IIf(cboItmClassTo.SelLength > 0, "", Set_Quote(cboItmClassTo.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY ItemClassCode "
    
     Call Ini_Combo(2, wsSQL, cboItmClassTo.Left, cboItmClassTo.Top + cboItmClassTo.Height, tblCommon, wsFormID, "TBLItmClass", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmClassTo_GotFocus()
    FocusMe cboItmClassTo
End Sub

Private Sub cboItmClassTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmClassTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmClassTo = False Then
            cboItmClassTo.SetFocus
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub

Private Sub cboItmClassTo_LostFocus()
    FocusMe cboItmClassTo, True
End Sub

Private Sub cboItmTypeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeFr
    
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ItmTypeEngDesc", "ItmTypeChiDesc") & " FROM MstItemType WHERE ItmTypeCode LIKE '%" & IIf(cboItmTypeFr.SelLength > 0, "", Set_Quote(cboItmTypeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeFr.Left, cboItmTypeFr.Top + cboItmTypeFr.Height, tblCommon, wsFormID, "TBLITMTYPE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmTypeFr_GotFocus()
        FocusMe cboItmTypeFr
    Set wcCombo = cboItmTypeFr
End Sub

Private Sub cboItmTypeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeFr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItmTypeFr.Text) <> "" And _
            Trim(cboItmTypeTo.Text) = "" Then
            cboItmTypeTo.Text = cboItmTypeFr.Text
        End If
        cboItmTypeTo.SetFocus
    End If
End Sub


Private Sub cboItmTypeFr_LostFocus()
    FocusMe cboItmTypeFr, True
End Sub

Private Sub cboItmTypeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeTo
    
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ItmTypeEngDesc", "ItmTypeChiDesc") & " FROM MstItemType WHERE ItmTypeCode LIKE '%" & IIf(cboItmTypeTo.SelLength > 0, "", Set_Quote(cboItmTypeTo.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeTo.Left, cboItmTypeTo.Top + cboItmTypeTo.Height, tblCommon, wsFormID, "TBLITMTYPE", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmTypeTo_GotFocus()
    FocusMe cboItmTypeTo
    Set wcCombo = cboItmTypeTo
End Sub

Private Sub cboItmTypeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeTo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmTypeTo = False Then
            cboItmTypeTo.SetFocus
            Exit Sub
        End If
        
        cboItmClassFr.SetFocus
    End If
End Sub



Private Sub cboItmTypeTo_LostFocus()
FocusMe cboItmTypeTo, True
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
    
    wsFormID = "ITM002"
    
End Sub

Private Sub Ini_Scr()

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   cboItmTypeFr.Text = ""
   cboItmTypeTo.Text = ""
   cboItmClassFr.Text = ""
   cboItmClassTo.Text = ""
   Call SetDateMask(medPrdFr)
   Call SetDateMask(medPrdTo)
   

    
End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
    If chk_cboDocNoTo = False Then
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    If chk_cboItmTypeTo = False Then
        cboItmTypeTo.SetFocus
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
   Set frmITM002 = Nothing

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
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblItmTypeFr.Caption = Get_Caption(waScrItm, "TYPEFR")
    lblItmTypeTo.Caption = Get_Caption(waScrItm, "TYPETO")
    lblItmClassFr.Caption = Get_Caption(waScrItm, "CLASSFR")
    lblItmClassTo.Caption = Get_Caption(waScrItm, "CLASSTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    wgsTitle = Get_Caption(waScrItm, "RPTTITLE")
    
End Sub

Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/  /" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Date(medPrdFr) = False Then
        wsMsg = "Wrong Date!"
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
    
    If Trim(medPrdTo) = "/  /" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Date(medPrdTo) = False Then
    
        wsMsg = "Wrong Date!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function

Private Function chk_cboItmTypeTo() As Boolean
    chk_cboItmTypeTo = False
    
    If UCase(cboItmTypeFr.Text) > UCase(cboItmTypeTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboItmTypeTo = True
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

Private Function chk_cboItmClassTo() As Boolean
    chk_cboItmClassTo = False
    
    If UCase(cboItmClassFr.Text) > UCase(cboItmClassTo.Text) Then
        wsMsg = "To > From!"
        MsgBox wsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_cboItmClassTo = True
End Function

Private Sub cboDocNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
    
     wsSQL = "SELECT ITMCODE, ITMBARCODE,  " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & ", " & IIf(gsLangID = "1", "ItmTypeEngDesc", "ItmTypeChiDesc") & ", " & IIf(gsLangID = "1", "ItemClassEDesc", "ItemClassCDesc") & " "
     wsSQL = wsSQL & " FROM MstItem, MstItemType, MstItemClass "
     wsSQL = wsSQL & " WHERE ITMCODE LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
     wsSQL = wsSQL & " AND ItemClassCODE = ITMCLASS"
     wsSQL = wsSQL & " AND ITMITMTYPECODE = ITMTYPECODE "
     wsSQL = wsSQL & " AND ITMSTATUS  <> '2' "
     wsSQL = wsSQL & " ORDER BY ITMCODE "
    Call Ini_Combo(5, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoFr, 13, KeyAscii)
    
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
  
    wsSQL = "SELECT ITMCODE, ITMBARCODE,  " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & ", " & IIf(gsLangID = "1", "ItmTypeEngDesc", "ItmTypeChiDesc") & ", " & IIf(gsLangID = "1", "ItemClassEDesc", "ItemClassCDesc") & " "
     wsSQL = wsSQL & " FROM MstItem, MstItemType, MstItemClass "
     wsSQL = wsSQL & " WHERE ITMCODE LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
     wsSQL = wsSQL & " AND ItemClassCODE = ITMCLASS"
     wsSQL = wsSQL & " AND ITMITMTYPECODE = ITMTYPECODE "
     wsSQL = wsSQL & " AND ITMSTATUS  <> '2' "
     wsSQL = wsSQL & " ORDER BY ITMCODE "
    Call Ini_Combo(5, wsSQL, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoTo_GotFocus()
    FocusMe cboDocNoTo
    Set wcCombo = cboDocNoTo
End Sub

Private Sub cboDocNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoTo, 13, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoTo = False Then
            Call cboDocNoTo_GotFocus
            Exit Sub
        End If
        
        cboItmTypeFr.SetFocus
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
            medPrdTo.SetFocus
            Exit Sub
        End If
        
         cboDocNoFr.SetFocus
         
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


