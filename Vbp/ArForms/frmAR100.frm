VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAR100 
   Caption         =   "AR Update"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmAR100.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   1470
      Left            =   4560
      OleObjectBlob   =   "frmAR100.frx":030A
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.CheckBox chkSettle 
      Alignment       =   1  '靠右對齊
      Caption         =   "Settlement"
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cboCusNoFr2 
      Height          =   300
      Left            =   2790
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3030
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo2 
      Height          =   300
      Left            =   5580
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   3030
      Width           =   1812
   End
   Begin VB.ComboBox cboChqNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2670
      Width           =   1812
   End
   Begin VB.ComboBox cboChqNoFr 
      Height          =   300
      Left            =   2790
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2670
      Width           =   1812
   End
   Begin VB.CheckBox chkAR 
      Alignment       =   1  '靠右對齊
      Caption         =   "AR Transaction"
      Height          =   180
      Left            =   840
      TabIndex        =   21
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   990
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   990
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5580
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1350
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2784
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1350
      Width           =   1812
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   120
      Top             =   4080
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
            Picture         =   "frmAR100.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAR100.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   20
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
   Begin MSMask.MaskEdBox medPrdFr 
      Height          =   285
      Left            =   2790
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrdTo2 
      Height          =   285
      Left            =   5580
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrdFr2 
      Height          =   285
      Left            =   2790
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCusNoFr2 
      Caption         =   "Customer Code From"
      Height          =   225
      Left            =   870
      TabIndex        =   27
      Top             =   3045
      Width           =   1890
   End
   Begin VB.Label lblPrdFr2 
      Caption         =   "Period From"
      Height          =   225
      Left            =   870
      TabIndex        =   26
      Top             =   3405
      Width           =   1890
   End
   Begin VB.Label lblCusNoTo2 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   25
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label lblPrdTo2 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   24
      Top             =   3405
      Width           =   375
   End
   Begin VB.Label lblChqNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   23
      Top             =   2685
      Width           =   375
   End
   Begin VB.Label lblChqNoFr 
      Caption         =   "Document # From"
      Height          =   225
      Left            =   870
      TabIndex        =   22
      Top             =   2685
      Width           =   1890
   End
   Begin VB.Label lblDocNoFr 
      Caption         =   "Document # From"
      Height          =   225
      Left            =   870
      TabIndex        =   19
      Top             =   1005
      Width           =   1890
   End
   Begin VB.Label lblDocNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   18
      Top             =   1005
      Width           =   375
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   16
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label lblCusNoTo 
      Caption         =   "To"
      Height          =   225
      Left            =   5220
      TabIndex        =   15
      Top             =   1365
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   870
      TabIndex        =   14
      Top             =   1725
      Width           =   1890
   End
   Begin VB.Label lblCusNoFr 
      Caption         =   "Customer Code From"
      Height          =   225
      Left            =   870
      TabIndex        =   13
      Top             =   1365
      Width           =   1890
   End
End
Attribute VB_Name = "frmAR100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
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
    Dim wsSql As String
    Dim adcmdSave As New ADODB.Command

On Error GoTo cmdSave_Err

    wsDteTim = gsSystemDate
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
    
    If chkAR.Value = 1 Then
    
        
    adcmdSave.CommandText = "USP_AR100A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, Change_SQLDate(wsDteTim))
    Call SetSPPara(adcmdSave, 3, wsDteTim)
    Call SetSPPara(adcmdSave, 4, cboDocNoFr.Text)
    Call SetSPPara(adcmdSave, 5, IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), cboDocNoTo.Text))
    Call SetSPPara(adcmdSave, 6, cboCusNoFr.Text)
    Call SetSPPara(adcmdSave, 7, IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), cboCusNoTo.Text))
    Call SetSPPara(adcmdSave, 8, medPrdFr.Text)
    Call SetSPPara(adcmdSave, 9, medPrdTo.Text)
    Call SetSPPara(adcmdSave, 10, "")
    
    
    adcmdSave.Execute
    
    End If
    
    If chkSettle.Value = 1 Then
    
        
    adcmdSave.CommandText = "USP_AR100B"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, Change_SQLDate(wsDteTim))
    Call SetSPPara(adcmdSave, 3, wsDteTim)
    Call SetSPPara(adcmdSave, 4, cboChqNoFr.Text)
    Call SetSPPara(adcmdSave, 5, IIf(Trim(cboChqNoTo.Text) = "", String(15, "z"), cboChqNoTo.Text))
    Call SetSPPara(adcmdSave, 6, cboCusNoFr2.Text)
    Call SetSPPara(adcmdSave, 7, IIf(Trim(cboCusNoTo2.Text) = "", String(10, "z"), cboCusNoTo2.Text))
    Call SetSPPara(adcmdSave, 8, medPrdFr2.Text)
    Call SetSPPara(adcmdSave, 9, medPrdTo2.Text)
    
    
    adcmdSave.Execute
    
    End If
    
    cnCon.CommitTrans 'Create Stored Procedure String
    Set adcmdSave = Nothing
    Me.MousePointer = vbDefault
    
    gsMsg = "Update Process is completed!"
    MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        
    Call cmdCancel
    
    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
End Sub

Private Sub cboChqNoTo_LostFocus()
    FocusMe cboChqNoTo, True
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
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
        Case Else
        
    End Select
   
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



Private Sub chkAR_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = vbDefault
    cboDocNoFr.SetFocus
End If

End Sub

Private Sub chkSettle_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = vbDefault
    cboChqNoFr.SetFocus
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
    
    wsFormID = "AR100"
    
End Sub

Private Sub Ini_Scr()
Dim wsFromDate As String
Dim wsToDate As String

   Me.Caption = wsFormCaption

   tblCommon.Visible = False
   
    wsFromDate = getCtrlMth("AR")
    wsFromDate = Left(wsFromDate, 4) & "/" & Mid(wsFromDate, 5, 2) & "/" & "01"
    wsToDate = Format(DateAdd("D", -1, CDate(DateAdd("M", 1, CDate(wsFromDate)))), "yyyy/mm/dd")
   
   
   chkAR.Value = 1
   
   cboDocNoFr.Text = ""
   cboDocNoTo.Text = ""
   cboCusNoFr.Text = ""
   cboCusNoTo.Text = ""
   Call SetDateMask(medPrdFr)
   Call SetDateMask(medPrdTo)
   
   cboChqNoFr.Text = ""
   cboChqNoTo.Text = ""
   cboCusNoFr2.Text = ""
   cboCusNoTo2.Text = ""
   Call SetDateMask(medPrdFr2)
   Call SetDateMask(medPrdTo2)
   
   chkSettle.Value = 1
   
   medPrdFr.Text = wsFromDate
   medPrdFr2.Text = wsFromDate
   medPrdTo.Text = wsToDate
   medPrdTo2.Text = wsToDate
   
   
End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
If chkAR.Value = 1 Then
    
    If chk_cboDocNoTo = False Then
        Exit Function
    End If
    
    If chk_cboCusNoTo = False Then
        Exit Function
    End If
    
    
    If chk_medPrdFr = False Then
        Exit Function
    End If
    
    If chk_medPrdTo = False Then
        Exit Function
    End If
    
End If
If chkSettle.Value = 1 Then
    
    If chk_cboChqNoTo = False Then
        Exit Function
    End If
    
    If chk_cboCusNoTo2 = False Then
        Exit Function
    End If
    
    
    If chk_medPrdFr2 = False Then
        Exit Function
    End If
    
    If chk_medPrdTo2 = False Then
        Exit Function
    End If
    
End If

If chkAR.Value = 0 And chkSettle.Value = 0 Then
       gsMsg = "Please Select Type Of Update!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       chkAR.SetFocus
       Exit Function
End If

    InputValidation = True
   
End Function



Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 5190
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmAR100 = Nothing

End Sub



Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub



Private Sub medPrdFr2_LostFocus()
    FocusMe medPrdFr2, True
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
    Call Get_Scr_Item("AR100", waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    lblChqNoFr.Caption = Get_Caption(waScrItm, "CHQNOFR")
    lblChqNoTo.Caption = Get_Caption(waScrItm, "CHQNOTO")
    lblCusNoFr2.Caption = Get_Caption(waScrItm, "CUSNOFR2")
    lblCusNoTo2.Caption = Get_Caption(waScrItm, "CUSNOTO2")
    lblPrdFr2.Caption = Get_Caption(waScrItm, "PRDFR2")
    lblPrdTo2.Caption = Get_Caption(waScrItm, "PRDTO2")
    chkAR.Caption = Get_Caption(waScrItm, "CHKAR")
    chkSettle.Caption = Get_Caption(waScrItm, "CHKSETTLE")
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    

    
End Sub



Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    
    
    If Chk_Date(medPrdFr) = False Then
       gsMsg = "Invalid Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr.SetFocus
       Exit Function
    End If
                
    If medPrdFr.Text < gsDateFrom Or medPrdTo.Text > gsDateTo Then
       gsMsg = "Out Of date range!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr.SetFocus
       Exit Function
    End If
        
    If medPrdFr.Text > medPrdTo.Text Then
       gsMsg = "To Date must greater From Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo.SetFocus
        Exit Function
    End If
    
    
    chk_medPrdFr = True
    
End Function

Private Function chk_medPrdFr2() As Boolean
    chk_medPrdFr2 = False
    
    
    
    If Chk_Date(medPrdFr2) = False Then
       gsMsg = "Invalid Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr2.SetFocus
       Exit Function
    End If
                
    If medPrdFr2.Text < gsDateFrom Or medPrdTo2.Text > gsDateTo Then
       gsMsg = "Out Of date range!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdFr2.SetFocus
       Exit Function
    End If
        
    If medPrdFr2.Text > medPrdTo2.Text Then
       gsMsg = "To Date must greater From Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo2.SetFocus
        Exit Function
    End If
    
    
    chk_medPrdFr2 = True
    
End Function



Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If Chk_Date(medPrdTo) = False Then
       gsMsg = "Invalid Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo.SetFocus
       Exit Function
    End If
                
    If medPrdTo.Text < gsDateFrom Or medPrdTo.Text > gsDateTo Then
       gsMsg = "Out Of date range!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo.SetFocus
       Exit Function
    End If
        
    If medPrdFr.Text > medPrdTo.Text Then
       gsMsg = "To Date must greater From Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo.SetFocus
        Exit Function
    End If
    
    chk_medPrdTo = True
End Function

Private Function chk_medPrdTo2() As Boolean
    chk_medPrdTo2 = False
    
    If Chk_Date(medPrdTo2) = False Then
       gsMsg = "Invalid Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo2.SetFocus
       Exit Function
    End If
                
    If medPrdTo2.Text < gsDateFrom Or medPrdTo2.Text > gsDateTo Then
       gsMsg = "Out Of date range!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo2.SetFocus
       Exit Function
    End If
        
    If medPrdFr2.Text > medPrdTo2.Text Then
       gsMsg = "To Date must greater From Date!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       medPrdTo2.SetFocus
        Exit Function
    End If
    
    chk_medPrdTo2 = True
End Function

Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusNoFr.SetFocus
        Exit Function
    End If
    
    chk_cboCusNoTo = True
End Function

Private Function chk_cboCusNoTo2() As Boolean
    chk_cboCusNoTo2 = False
    
    If UCase(cboCusNoFr2.Text) > UCase(cboCusNoTo2.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusNoFr2.SetFocus
        Exit Function
    End If
    
    chk_cboCusNoTo2 = True
End Function
Private Function chk_cboDocNoTo() As Boolean
    chk_cboDocNoTo = False
    
    If UCase(cboDocNoFr.Text) > UCase(cboDocNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboDocNoTo = True
End Function

Private Function chk_cboChqNoTo() As Boolean
    chk_cboChqNoTo = False
    
    If UCase(cboChqNoFr.Text) > UCase(cboChqNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboChqNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboChqNoTo = True
End Function

Private Sub cboDocNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    wsSql = "SELECT INHDDOCNO, CUSCODE, INHDDOCDATE "
    wsSql = wsSql & " FROM ARINHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE INHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSql = wsSql & " AND INHDCUSID  = CUSID "
    wsSql = wsSql & " AND INHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY INHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
  
    wsSql = "SELECT INHDDOCNO, CUSCODE, INHDDOCDATE "
    wsSql = wsSql & " FROM ARINHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE INHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSql = wsSql & " AND INHDCUSID  = CUSID "
    wsSql = wsSql & " AND INHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY INHDDOCNO "
    
    Call Ini_Combo(3, wsSql, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
        chkSettle.SetFocus
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


Private Sub cboChqNoFr_DropDown()
   Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboChqNoFr
  
    wsSql = "SELECT ARCQCHQNO, CUSCODE, ARCQCHQDATE "
    wsSql = wsSql & " FROM ARCHEQUE, mstCUSTOMER "
    wsSql = wsSql & " WHERE ARCQCHQNO LIKE '%" & IIf(cboChqNoFr.SelLength > 0, "", Set_Quote(cboChqNoFr.Text)) & "%' "
    wsSql = wsSql & " AND ARCQCUSID  = CUSID "
    wsSql = wsSql & " AND ARCQSTATUS  <> '2' "
   ' wsSql = wsSql & " ORDER BY ARCQCHQNO "
    wsSql = wsSql & " UNION "
    wsSql = wsSql & " SELECT ARSHDOCNO, CUSCODE, ARSHDOCDATE "
    wsSql = wsSql & " FROM ARSTHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE ARSHDOCNO LIKE '%" & IIf(cboChqNoFr.SelLength > 0, "", Set_Quote(cboChqNoFr.Text)) & "%' "
    wsSql = wsSql & " AND ARSHCUSID  = CUSID "
    wsSql = wsSql & " AND ARSHSTATUS  <> '2' "
   ' wsSql = wsSql & " ORDER BY ARSHDOCNO "
    
    Call Ini_Combo(3, wsSql, cboChqNoFr.Left, cboChqNoFr.Top + cboChqNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboChqNoFr_GotFocus()
    FocusMe cboChqNoFr
    Set wcCombo = cboChqNoFr
End Sub

Private Sub cboChqNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboChqNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboChqNoFr.Text) <> "" And _
            Trim(cboChqNoTo.Text) = "" Then
            cboChqNoTo.Text = cboChqNoFr.Text
        End If
        cboChqNoTo.SetFocus
    End If
End Sub

Private Sub cboChqNoFr_LostFocus()
    FocusMe cboChqNoFr, True
End Sub

Private Sub cboChqNoTo_DropDown()
Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboChqNoTo
  
    wsSql = "SELECT ARCQCHQNO, CUSCODE, ARCQCHQDATE "
    wsSql = wsSql & " FROM ARCHEQUE, mstCUSTOMER "
    wsSql = wsSql & " WHERE ARCQCHQNO LIKE '%" & IIf(cboChqNoTo.SelLength > 0, "", Set_Quote(cboChqNoTo.Text)) & "%' "
    wsSql = wsSql & " AND ARCQCUSID  = CUSID "
    wsSql = wsSql & " AND ARCQSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY ARCQCHQNO "
    
    Call Ini_Combo(3, wsSql, cboChqNoTo.Left, cboChqNoTo.Top + cboChqNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboChqNoTo_GotFocus()
    FocusMe cboChqNoTo
    Set wcCombo = cboChqNoTo
End Sub

Private Sub cboChqNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboChqNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboChqNoTo = False Then
            Exit Sub
        End If
        
        cboCusNoFr2.SetFocus
    End If
End Sub


Private Sub medPrdFr2_GotFocus()
    FocusMe medPrdFr2
End Sub


Private Sub medPrdFr2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr2 = False Then
            Exit Sub
        End If
        
        If Trim(medPrdFr2) <> "/" And _
            Trim(medPrdTo2) = "/" Then
            medPrdTo2.Text = medPrdFr2.Text
        End If
        medPrdTo2.SetFocus
    End If
End Sub
Private Sub medPrdTo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo2 = False Then
            Exit Sub
        End If
        chkAR.SetFocus
    End If
End Sub

Private Sub medPrdTo2_GotFocus()
    FocusMe medPrdTo2
End Sub
Private Sub medPrdTo2_LostFocus()
    FocusMe medPrdTo2, True
End Sub


Private Sub cboCusNoFr2_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr2.SelLength > 0, "", Set_Quote(cboCusNoFr2.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr2.SelLength > 0, "", Set_Quote(cboCusNoFr2.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoFr2.Left, cboCusNoFr2.Top + cboCusNoFr2.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr2_GotFocus()
        FocusMe cboCusNoFr2
    Set wcCombo = cboCusNoFr2
End Sub

Private Sub cboCusNoFr2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoFr2, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboCusNoFr2.Text) <> "" And _
            Trim(cboCusNoTo2.Text) = "" Then
            cboCusNoTo2.Text = cboCusNoFr2.Text
        End If
        cboCusNoTo2.SetFocus
    End If
End Sub


Private Sub cboCusNoFr2_LostFocus()
    FocusMe cboCusNoFr2, True
End Sub



Private Sub cboCusNoTo2_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass
    Select Case gsLangID
        Case "1"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr2.SelLength > 0, "", Set_Quote(cboCusNoFr2.Text)) & "%' "
        Case "2"
            wsSql = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr2.SelLength > 0, "", Set_Quote(cboCusNoFr2.Text)) & "%' "
        Case Else
        
    End Select
   
    wsSql = wsSql & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSql, cboCusNoTo2.Left, cboCusNoTo2.Top + cboCusNoTo2.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoTo2_GotFocus()
    FocusMe cboCusNoTo2
    Set wcCombo = cboCusNoTo2
End Sub

Private Sub cboCusNoTo2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoTo2, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusNoTo2 = False Then
            Exit Sub
        End If
        
        medPrdFr2.SetFocus
    End If
End Sub



Private Sub cboCusNoTo2_LostFocus()
FocusMe cboCusNoTo2, True
End Sub


