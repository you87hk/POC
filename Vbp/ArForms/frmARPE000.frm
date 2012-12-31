VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmARPE000 
   Caption         =   "Period End Process"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9195
   Icon            =   "frmARPE000.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9195
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox chkGL 
      Alignment       =   1  '靠右對齊
      Caption         =   "General Ledger Transaction"
      Height          =   420
      Left            =   6960
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox chkAP 
      Alignment       =   1  '靠右對齊
      Caption         =   "Account Payable Transaction"
      Height          =   420
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox chkAR 
      Alignment       =   1  '靠右對齊
      Caption         =   "Account Receiable Transaction"
      Height          =   420
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
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
            Picture         =   "frmARPE000.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":1910
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":1D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":207C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":24CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":2C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":2F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":33A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARPE000.frx":3C82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   3
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
   Begin VB.Label lblDspAPCtlPrd 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4080
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblDspGLCtlPrd 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   7080
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblDspARCtlPrd 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblCtlPrd 
      Caption         =   "CUSFAX"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblWarning 
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
      Height          =   1305
      Left            =   2640
      TabIndex        =   4
      Top             =   1200
      Width           =   4650
   End
End
Attribute VB_Name = "frmARPE000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim wsFormID As String
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Dim wgsTitle As String
Private wsFormCaption As String

Private Const tcGo = "Go"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private wsMsg As String
Private wsARToDate As String
Private wsAPToDate As String
Private wsGLToDate As String


Private Sub cmdCancel()
    Ini_Scr

End Sub

Private Sub cmdOK()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim adcmdSave As New ADODB.Command

On Error GoTo cmdSave_Err

    wsDteTim = gsSystemDate
    
   ' If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
    
    If chkAR.Value = 1 Then
    If Chk_ARUpdFlg = True Then
    
        
    adcmdSave.CommandText = "USP_ARPE000"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wsARToDate)
    Call SetSPPara(adcmdSave, 2, gsUserID)
    Call SetSPPara(adcmdSave, 3, wsDteTim)
    
    
    adcmdSave.Execute
    
    End If
    End If
    
    If chkAP.Value = 1 Then
    If Chk_APUpdFlg = True Then
        
    adcmdSave.CommandText = "USP_APPE000"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wsAPToDate)
    Call SetSPPara(adcmdSave, 2, gsUserID)
    Call SetSPPara(adcmdSave, 3, wsDteTim)
    
    
    adcmdSave.Execute
    
    End If
    End If
    
    If chkGL.Value = 1 Then
    
        
    adcmdSave.CommandText = "USP_GLPE000"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wsGLToDate)
    Call SetSPPara(adcmdSave, 2, gsUserID)
    Call SetSPPara(adcmdSave, 3, wsDteTim)
    
    
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



Private Sub chkAR_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = vbDefault
    chkAP.SetFocus
End If

End Sub
Private Sub chkAP_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = vbDefault
    chkGL.SetFocus
End If

End Sub

Private Sub chkGL_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = vbDefault
    chkAR.SetFocus
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
    
    wsFormID = "ARPE000"
    
End Sub

Private Sub Ini_Scr()
Dim wsFromDate As String
Dim wsARCtlPrd As String
Dim wsAPCtlPrd As String
Dim wsGLCtlPrd As String

   Me.Caption = wsFormCaption

   
    wsARCtlPrd = getCtrlMth("AR")
    wsFromDate = Left(wsARCtlPrd, 4) & "/" & Mid(wsARCtlPrd, 5, 2) & "/" & "01"
    wsARToDate = DateAdd("D", -1, CDate(DateAdd("M", 1, CDate(wsFromDate))))
   
   
    wsAPCtlPrd = getCtrlMth("AP")
    wsFromDate = Left(wsAPCtlPrd, 4) & "/" & Mid(wsAPCtlPrd, 5, 2) & "/" & "01"
    wsAPToDate = DateAdd("D", -1, CDate(DateAdd("M", 1, CDate(wsFromDate))))
   
   
   wsGLCtlPrd = getCtrlMth("GL")
   wsFromDate = Left(wsGLCtlPrd, 4) & "/" & Mid(wsGLCtlPrd, 5, 2) & "/" & "01"
   wsGLToDate = DateAdd("D", -1, CDate(DateAdd("M", 1, CDate(wsFromDate))))
   
   chkAR.Value = 1
   chkAP.Value = 1
   chkGL.Value = 1
   
   lblDspARCtlPrd.Caption = Left(wsARCtlPrd, 4) & "/" & Right(wsARCtlPrd, 2)
   lblDspAPCtlPrd.Caption = Left(wsAPCtlPrd, 4) & "/" & Right(wsAPCtlPrd, 2)
   lblDspGLCtlPrd.Caption = Left(wsGLCtlPrd, 4) & "/" & Right(wsGLCtlPrd, 2)
    
   
End Sub
Private Function InputValidation() As Boolean

    InputValidation = False
    
If chkAR.Value = 1 Then
    If Chk_ARUpdFlg = False Then
        Exit Function
    End If
End If
If chkAP.Value = 1 Then
    If Chk_APUpdFlg = False Then
        Exit Function
    End If
End If
If chkGL.Value = 1 Then
'    If chk_GLUpdFlg = False Then
'        Exit Function
'    End If
End If
    
 
If chkAR.Value = 0 And chkAP.Value = 0 And chkGL.Value = 0 Then
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
   Set frmARPE000 = Nothing

End Sub



Private Sub Ini_Caption()
    Call Get_Scr_Item("ARPE000", waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
   
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    chkAR.Caption = Get_Caption(waScrItm, "CHKAR")
    chkAP.Caption = Get_Caption(waScrItm, "CHKAP")
    chkGL.Caption = Get_Caption(waScrItm, "CHKGL")
    lblWarning.Caption = Get_Caption(waScrItm, "WARN1") & Chr(13) & Chr(10) & _
                         Get_Caption(waScrItm, "WARN2")
                         
    lblCtlPrd.Caption = Get_Caption(waScrItm, "CTLPRD")
    
    
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcGo).ToolTipText = Get_Caption(waScrToolTip, tcGo) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    

    
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

Private Function Chk_ARUpdFlg() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    Chk_ARUpdFlg = False
    
    wsSQL = "SELECT Count(*) RecCnt FROM ARINHD "
    wsSQL = wsSQL & "WHERE INHDUPDFLG = 'N' "
    wsSQL = wsSQL & "And INHDDOCDATE <= '" & wsARToDate & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "RecCnt")) > 0 Then
           gsMsg = "All Invoice/Credit/Debit data must posted!"
           MsgBox gsMsg, vbOKOnly, gsTitle
           chkAR.SetFocus
           rsRcd.Close
           Set rsRcd = Nothing
           Exit Function
        End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    wsSQL = "SELECT Count(*) RecCnt FROM ARCHEQUE "
    wsSQL = wsSQL & "WHERE ARCQUPDFLG = 'N' "
    wsSQL = wsSQL & "And ARCQCHQDATE <= '" & wsARToDate & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "RecCnt")) > 0 Then
           gsMsg = "All Cheque data must posted!"
           MsgBox gsMsg, vbOKOnly, gsTitle
           chkAR.SetFocus
           rsRcd.Close
           Set rsRcd = Nothing
           Exit Function
        End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    wsSQL = "SELECT Count(*) RecCnt FROM ARSTHD "
    wsSQL = wsSQL & "WHERE ARSHUPDFLG = 'N' "
    wsSQL = wsSQL & "And ARSHDOCDATE <= '" & wsARToDate & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "RecCnt")) > 0 Then
           gsMsg = "All Settlement data must posted!"
           MsgBox gsMsg, vbOKOnly, gsTitle
           chkAR.SetFocus
           rsRcd.Close
           Set rsRcd = Nothing
           Exit Function
        End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_ARUpdFlg = True

End Function
Private Function Chk_APUpdFlg() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    Chk_APUpdFlg = False
    
    wsSQL = "SELECT Count(*) RecCnt FROM APIPHD "
    wsSQL = wsSQL & "WHERE IPHDUPDFLG = 'N' "
    wsSQL = wsSQL & "And IPHDDOCDATE <= '" & wsAPToDate & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "RecCnt")) > 0 Then
           gsMsg = "All Invoice/Credit/Debit data must posted!"
           MsgBox gsMsg, vbOKOnly, gsTitle
           chkAP.SetFocus
           rsRcd.Close
           Set rsRcd = Nothing
           Exit Function
        End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    wsSQL = "SELECT Count(*) RecCnt FROM APCHEQUE "
    wsSQL = wsSQL & "WHERE APCQUPDFLG = 'N' "
    wsSQL = wsSQL & "And APCQCHQDATE <= '" & wsAPToDate & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "RecCnt")) > 0 Then
           gsMsg = "All Cheque data must posted!"
           MsgBox gsMsg, vbOKOnly, gsTitle
           chkAP.SetFocus
           rsRcd.Close
           Set rsRcd = Nothing
           Exit Function
        End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    wsSQL = "SELECT Count(*) RecCnt FROM APSTHD "
    wsSQL = wsSQL & "WHERE APSHUPDFLG = 'N' "
    wsSQL = wsSQL & "And APSHDOCDATE <= '" & wsAPToDate & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        If To_Value(ReadRs(rsRcd, "RecCnt")) > 0 Then
           gsMsg = "All Settlement data must posted!"
           MsgBox gsMsg, vbOKOnly, gsTitle
           chkAP.SetFocus
           rsRcd.Close
           Set rsRcd = Nothing
           Exit Function
        End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_APUpdFlg = True

End Function

