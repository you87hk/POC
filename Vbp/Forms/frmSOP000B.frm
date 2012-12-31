VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSOP000B 
   Caption         =   "AR Update"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmSOP000B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox ChkPrtAcc 
      Alignment       =   1  '靠右對齊
      Caption         =   "Print Acc. No. in Label:"
      Height          =   180
      Left            =   720
      TabIndex        =   9
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   6375
      Begin VB.Label lblWarning 
         Alignment       =   2  '置中對齊
         Caption         =   "Period From"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   5730
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  '置中對齊
         Caption         =   "Period From"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   5730
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  '置中對齊
         Caption         =   "Period From"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   5730
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   7560
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":1910
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":1D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":207C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":24CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":2F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":33A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":3C82
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSOP000B.frx":3FAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   2
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
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Go"
            Object.ToolTipText     =   "Go (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSMask.MaskEdBox medPrdTo 
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   2475
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
      Left            =   2520
      TabIndex        =   0
      Top             =   2475
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "####/##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblPrdTo 
      Caption         =   "To"
      Height          =   330
      Left            =   3840
      TabIndex        =   4
      Top             =   2475
      Width           =   375
   End
   Begin VB.Label lblPrdFr 
      Caption         =   "Period From"
      Height          =   225
      Left            =   720
      TabIndex        =   3
      Top             =   2475
      Width           =   1890
   End
End
Attribute VB_Name = "frmSOP000B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim wsFormID As String
Dim wsTrnCd As String

Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wcCombo As Control
Dim wgsTitle As String
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcPrint = "Print"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String


Private Sub cmdCancel()
    Ini_Scr
    medPrdFr.SetFocus
End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim adcmdSave As New ADODB.Command
    Dim wsActFlg As String
    
On Error GoTo cmdOK_Err

    wsDteTim = gsSystemDate
    
    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
    
    wsActFlg = IIf(ChkPrtAcc.Value = 0, "0", "1")
    wsDteTim = Change_SQLDate(Now)
    
        
    adcmdSave.CommandText = "USP_SOP000A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wsActFlg)
    Call SetSPPara(adcmdSave, 2, IIf(Trim(medPrdFr.Text) = "/", String(6, "0"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)))
    Call SetSPPara(adcmdSave, 3, IIf(Trim(medPrdTo.Text) = "/", String(6, "9"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)))
    Call SetSPPara(adcmdSave, 4, gsUserID)
    Call SetSPPara(adcmdSave, 5, gsSystemDate)
    Call SetSPPara(adcmdSave, 6, wsDteTim)
    Call SetSPPara(adcmdSave, 7, gsLangID)
    
    
    adcmdSave.Execute
    
    
    cnCon.CommitTrans 'Create Stored Procedure String
    Set adcmdSave = Nothing
    Me.MousePointer = vbDefault
    
    gsMsg = "Update Process is completed!"
    MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        
    Call cmdCancel
    
    Exit Sub
    
cmdOK_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
End Sub











Private Sub ChkPrtAcc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
                  
      medPrdFr.SetFocus
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
   Select Case KeyCode
        
         Case vbKeyF10
        
           Call cmdPrint
        
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
    
    wsFormID = "SOP000B"
    
End Sub

Private Sub Ini_Scr()
Dim wsFromDate As String
Dim wsToDate As String

   Me.Caption = wsFormCaption

  


   Call SetPeriodMask(medPrdFr)
   Call SetPeriodMask(medPrdTo)

    
   medPrdFr.Text = Left(gsSystemDate, 4) & "/" & Mid(gsSystemDate, 6, 2)
   medPrdTo.Text = medPrdFr.Text
   
   
End Sub
Private Function InputValidation() As Boolean

    InputValidation = False

       
    If chk_medPrdFr = False Then
        Exit Function
    End If
    
    If chk_medPrdTo = False Then
        Exit Function
    End If
    

    

    InputValidation = True
   
End Function



Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 4995
        Me.Width = 9315
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Set waScrItm = Nothing
   Set waScrToolTip = Nothing
   Set wcCombo = Nothing
   Set frmSOP000B = Nothing

End Sub




Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
   wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
   lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
   lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
   ChkPrtAcc.Caption = Get_Caption(waScrItm, "REFRESH")
   
   lblWarning(0).Caption = Get_Caption(waScrItm, "WARN0")
   lblWarning(1).Caption = Get_Caption(waScrItm, "WARN1")
   lblWarning(2).Caption = Get_Caption(waScrItm, "WARN2")
   
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F10)"
    
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    

    
End Sub



Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr.Text) = "/" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
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





Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If Trim(medPrdTo.Text) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If
    
    If Chk_Period(medPrdTo) = False Then
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

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub
Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo = False Then
            Exit Sub
        End If
       medPrdFr.SetFocus
    End If
End Sub

Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
        
        Case tcGo
            Call cmdOK
        Case tcPrint
            Call cmdPrint
        Case tcCancel
                Call cmdCancel
        Case tcExit
            Unload Me
    End Select
    
End Sub

Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSql As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsTitle As String
    Dim wsActFlg As String

    If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(1)
    wsSelection(1) = lblPrdFr.Caption & " " & medPrdFr.Text & " " & lblPrdTo.Caption & " " & medPrdTo.Text
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsTitle = "會計入數 "
    
    wsActFlg = IIf(ChkPrtAcc.Value = 0, "0", "1")
   
    
    wsSql = "EXEC usp_SOP000B '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wsTitle & "', "
    wsSql = wsSql & "" & wsActFlg & ", "
    wsSql = wsSql & "'" & IIf(Trim(medPrdFr.Text) = "/", "000000", Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "', "
    wsSql = wsSql & "'" & IIf(Trim(medPrdTo.Text) = "/", "999999", Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "', "
    wsSql = wsSql & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTSOP000B"
    Else
    wsRptName = "RPTSOP000B"
    End If
    
    
    NewfrmPrint.ReportID = "SOP000B"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SOP000B"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub


