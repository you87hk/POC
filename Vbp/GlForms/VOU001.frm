VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmVOU001 
   BackColor       =   &H8000000A&
   Caption         =   "憑單"
   ClientHeight    =   4440
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "VOU001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "VOU001.frx":08CA
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboVouType 
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   3540
      Width           =   2730
   End
   Begin VB.ComboBox cboVouPrefix 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "憑單"
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtVouSpa 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Top             =   2760
         Width           =   570
      End
      Begin VB.CheckBox chkVouYrFlg 
         Alignment       =   1  '靠右對齊
         Caption         =   "VOUYRFLG"
         Height          =   180
         Left            =   340
         TabIndex        =   4
         Top             =   2160
         Width           =   1520
      End
      Begin VB.CheckBox chkVouMnFlg 
         Alignment       =   1  '靠右對齊
         Caption         =   "VOUMNFLG"
         Height          =   180
         Left            =   340
         TabIndex        =   5
         Top             =   2520
         Width           =   1520
      End
      Begin VB.TextBox txtVouDesc 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   6450
      End
      Begin VB.TextBox txtVouLastKey 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   1680
         Width           =   810
      End
      Begin VB.TextBox txtVouLen 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   1320
         Width           =   810
      End
      Begin VB.TextBox txtVouPrefix 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   8
         Top             =   600
         Width           =   2730
      End
      Begin VB.Label lblVouSpa 
         Caption         =   "VOUSPA"
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   2840
         Width           =   1380
      End
      Begin VB.Label lblVouType 
         Caption         =   "VOUTYPE"
         Height          =   240
         Left            =   380
         TabIndex        =   16
         Top             =   3240
         Width           =   1380
      End
      Begin VB.Label lblVouDesc 
         Caption         =   "DOCTYPEDESC"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label lblVouLastKey 
         Caption         =   "VOULASTKEY"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label lblVouLen 
         Caption         =   "VOULEN"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label lblVouPrefix 
         Caption         =   "VOUPREFIX"
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   660
         Width           =   1380
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   720
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
            Picture         =   "VOU001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "VOU001.frx":6945
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
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "開新視窗 (F6)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "新增 (F2)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Object.ToolTipText     =   "修改 (F5)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "刪除 (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "儲存 (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "尋找 (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmVOU001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wsFormCaption As String
Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
 
Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"

Private wiAction As Integer

Private wcCombo As Control

Private wsActNam(4) As String
'Row Lock Variable

Private Const wsKeyType = "sysVouNo"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboVouPrefix_LostFocus()
    FocusMe cboVouPrefix, True
End Sub

Private Sub cboVouType_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVouType
    
    wsSQL = "SELECT MCModNo FROM sysMonCtl "
    wsSQL = wsSQL & "ORDER BY MCModNo"
    Call Ini_Combo(1, wsSQL, cboVouType.Left, cboVouType.Top + cboVouType.Height, tblCommon, wsFormID, "TBLVOUTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVouType_GotFocus()
    FocusMe cboVouType
End Sub

Private Sub cboVouType_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVouType, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboVouType() = False Then
            Exit Sub
        End If
           
        txtVouDesc.SetFocus
    End If
End Sub

Private Sub cboVouType_LostFocus()
    FocusMe cboVouType, True
End Sub

Private Sub chkVouMnFlg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtVouSpa.SetFocus
    End If
End Sub

Private Sub chkVouSpa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cboVouType.SetFocus
    End If
End Sub

Private Sub chkVouYrFlg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        chkVouMnFlg.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        Case vbKeyF6
            Call cmdOpen
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
        
        Case vbKeyF5
            If wiAction = DefaultPage Then Call cmdEdit
        
        Case vbKeyF3
            If wiAction = DefaultPage Then Call cmdDel
        
         Case vbKeyF9
            If tbrProcess.Buttons(tcFind).Enabled = True Then
                Call cmdFind
            End If
            
        Case vbKeyF10
            If tbrProcess.Buttons(tcSave).Enabled = True Then
                Call cmdSave
            End If
            
        Case vbKeyF11
            If wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec Then Call cmdCancel
        
        Case vbKeyF12
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    
    MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Scr
    
    MousePointer = vbDefault
  
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 4845
        Me.Width = 8700
    End If
End Sub

'-- Set toolbar buttons status in different mode, Default, AddEdit, None.
Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = True
                .Buttons(tcEdit).Enabled = True
                .Buttons(tcDelete).Enabled = True
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = False
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
            End With
            
        Case "AfrActAdd"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "AfrActEdit"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = True
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "AfrKey"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "ReadOnly"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = False
                .Buttons(tcFind).Enabled = True
                .Buttons(tcExit).Enabled = True
            
            End With
    End Select
End Sub

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            Me.txtVouPrefix.Enabled = False
            Me.txtVouDesc.Enabled = False
            Me.txtVouLen.Enabled = False
            Me.txtVouLastKey.Enabled = False
            
            Me.cboVouPrefix.Enabled = False
            Me.cboVouPrefix.Visible = False
            Me.txtVouPrefix.Visible = True
            Me.txtVouPrefix.Enabled = False
            
            chkVouYrFlg.Enabled = False
            chkVouMnFlg.Enabled = False
            txtVouSpa.Enabled = False
            cboVouType.Enabled = False
            
            
        Case "AfrActAdd"
            Me.cboVouPrefix.Enabled = False
            Me.cboVouPrefix.Visible = False
            
            Me.txtVouPrefix.Enabled = True
            Me.txtVouPrefix.Visible = True
            
        Case "AfrActEdit"
            Me.cboVouPrefix.Enabled = True
            Me.cboVouPrefix.Visible = True
            
            Me.txtVouPrefix.Enabled = False
            Me.txtVouPrefix.Visible = False
            
        Case "AfrKey"
            Me.cboVouPrefix.Enabled = False
            Me.txtVouPrefix.Enabled = False
            
            
            Me.txtVouPrefix.Enabled = True
            Me.txtVouDesc.Enabled = True
            
            Me.txtVouLen.Enabled = True
            Me.txtVouLastKey.Enabled = True
            
            chkVouYrFlg.Enabled = True
            chkVouMnFlg.Enabled = True
            txtVouSpa.Enabled = True
            cboVouType.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    'If Chk_txtVouPrefix() = False Then
    '    Exit Function
    'End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From sysVouNo "
    wsSQL = wsSQL + "WHERE (((sysVouNo.VouPrefix)='" + Set_Quote(cboVouPrefix.Text) + "'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False

    Else
        Me.cboVouPrefix = ReadRs(rsRcd, "VouPrefix")
        Me.txtVouLen = ReadRs(rsRcd, "VouLen")
        Me.txtVouLastKey = ReadRs(rsRcd, "VouLastKey")
        Me.txtVouDesc = ReadRs(rsRcd, "VouDesc")
        Me.cboVouType = ReadRs(rsRcd, "VouType")
        Me.txtVouSpa = ReadRs(rsRcd, "VouSpa")
        
        If ReadRs(rsRcd, "VouYrFlg") = "Y" Then
            chkVouYrFlg.Value = 1
        End If
        
        If ReadRs(rsRcd, "VouMnFlg") = "Y" Then
            chkVouMnFlg.Value = 1
        End If
        
        LoadRecord = True
    End If
    
    rsRcd.Close
    
    Set rsRcd = Nothing
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    Call UnLockAll(wsConnTime, wsFormID)
    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set frmVOU001 = Nothing
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case tcOpen
            Call cmdOpen
        
        Case tcAdd
            Call cmdNew
            
        Case tcEdit
            Call cmdEdit
        
        Case tcDelete
            
            Call cmdDel
            
        Case tcSave
            
            Call cmdSave
            
        Case tcCancel
            
             If tbrProcess.Buttons(tcSave).Enabled = True Then
                gsMsg = "你是否確定儲存現時之變更而離開?"
                If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
                    Call cmdCancel
                End If
            Else
                Call cmdCancel
            End If
        
        Case tcFind
            
            Call cmdFind
            
        Case tcExit
        
            Unload Me
            
    End Select
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
  '  Me.Left = 0
  '  Me.Top = 0
  '  Me.Width = Screen.Width
  '  Me.Height = Screen.Height
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "VOU001"
    wsTrnCd = ""
    
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblVouPrefix.Caption = Get_Caption(waScrItm, "VOUPREFIX")
    lblVouLen.Caption = Get_Caption(waScrItm, "VOULEN")
    lblVouLastKey.Caption = Get_Caption(waScrItm, "VOULASTKEY")
    lblVouDesc.Caption = Get_Caption(waScrItm, "VOUDESC")
    lblVouType.Caption = Get_Caption(waScrItm, "VOUTYPE")
    
    chkVouYrFlg.Caption = Get_Caption(waScrItm, "YRFLG")
    chkVouMnFlg.Caption = Get_Caption(waScrItm, "MNFLG")
    lblVouSpa.Caption = Get_Caption(waScrItm, "SPA")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")

    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "VOUADD")
    wsActNam(2) = Get_Caption(waScrItm, "VOUEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "VOUDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub Ini_Scr()
    Dim MyControl As Control
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
            Case "ComboBox"
                MyControl.Clear
            Case "TextBox"
                MyControl.Text = ""
            Case "TDBGrid"
                MyControl.ClearFields
            Case "Label"
                If UCase(MyControl.Name) Like "LBLDSP*" Then
                    MyControl.Caption = ""
                End If
            Case "RichTextBox"
                MyControl.Text = ""
            Case "CheckBox"
                MyControl.Value = 0
        End Select
    Next
    
    wiAction = DefaultPage

    chkVouYrFlg.Value = 0
    chkVouMnFlg.Value = 0

    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    tblCommon.Visible = False
    Me.Caption = wsFormCaption
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
              
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
        txtVouPrefix.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboVouPrefix.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboVouPrefix.SetFocus
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Select Case wiAction
    
        Case CorRec, DelRec

            If LoadRecord() = False Then
                gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
                MsgBox gsMsg, vbOKOnly, gsTitle
                Exit Sub
            Else
                If RowLock(wsConnTime, wsKeyType, cboVouPrefix, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtVouDesc.SetFocus
End Sub

Private Function Chk_VouPrefix(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_VouPrefix = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT VouPrefix "
    wsSQL = wsSQL & " FROM sysVouNo WHERE VouPrefix = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "VouPrefix")
    
    Chk_VouPrefix = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtVouPrefix() As Boolean
    Dim wsStatus As String

    Chk_txtVouPrefix = False
    
    If Trim(txtVouPrefix.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVouPrefix.SetFocus
        Exit Function
    End If

    If Chk_VouPrefix(txtVouPrefix.Text, wsStatus) = True Then
        gsMsg = "文件號已存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVouPrefix.SetFocus
        Exit Function
    End If
    
    Chk_txtVouPrefix = True
End Function

Private Function Chk_cboVouPrefix() As Boolean
    Dim wsStatus As String
 
    Chk_cboVouPrefix = False
    
    If Trim(cboVouPrefix.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVouPrefix.SetFocus
        Exit Function
    End If

    If Chk_VouPrefix(cboVouPrefix.Text, wsStatus) = False Then
        gsMsg = "文件號不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVouPrefix.SetFocus
        Exit Function
    End If
    
    Chk_cboVouPrefix = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmVOU001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show
End Sub

Private Sub cmdNew()
    wiAction = AddRec
    Ini_Scr_AfrAct
End Sub

Private Sub cmdEdit()
    wiAction = CorRec
    Ini_Scr_AfrAct
End Sub

Private Sub cmdDel()
    wiAction = DelRec
    Ini_Scr_AfrAct
End Sub

Private Sub cmdCancel()
    If tbrProcess.Buttons(tcSave).Enabled = True Then
        Select Case wiAction
            Case AddRec
                Call Ini_Scr
                Call cmdNew
                
            Case CorRec
                Call UnLockAll(wsConnTime, wsFormID)
                Call Ini_Scr
                Call cmdEdit
                
            Case DelRec
                Call UnLockAll(wsConnTime, wsFormID)
                Call Ini_Scr
                Call cmdDel
        End Select
    Else
        Call Ini_Scr
    End If
End Sub

Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim wsNo As String
    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboVouPrefix, wsFormID) Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
   
    If wiAction = DelRec Then
        If MsgBox("你是否確定要刪除此記錄?", vbYesNo, gsTitle) = vbNo Then
            cmdCancel
            MousePointer = vbDefault
            Exit Function
        End If
    Else
        If InputValidation() = False Then
            MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_VOU001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtVouPrefix.Text, cboVouPrefix.Text))
    Call SetSPPara(adcmdSave, 3, txtVouDesc)
    Call SetSPPara(adcmdSave, 4, txtVouLen)
    Call SetSPPara(adcmdSave, 5, txtVouLastKey)
    Call SetSPPara(adcmdSave, 6, IIf(chkVouYrFlg.Value = 0, "N", "Y"))
    Call SetSPPara(adcmdSave, 7, IIf(chkVouMnFlg.Value = 0, "N", "Y"))
    Call SetSPPara(adcmdSave, 8, txtVouSpa)
    Call SetSPPara(adcmdSave, 9, cboVouType)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 10)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - VOU001!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
        If wiAction = DelRec Then
            gsMsg = "已成功刪除!"
        Else
            gsMsg = "已成功儲存!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    End If
    
    Call cmdCancel
    
    Set adcmdSave = Nothing
    cmdSave = True
    
    MousePointer = vbDefault
    
    Exit Function
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Function

Private Sub cmdFind()
     Call OpenPromptForm
End Sub

Private Function SaveData() As Boolean
    SaveData = False
    
    If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) And tbrProcess.Buttons(tcSave).Enabled = True Then
       gsMsg = "你是否確定要儲存現時之作業?"
       If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
            Exit Function
        Else
                If cmdSave = True Then
                    Exit Function
                End If
        End If
        SaveData = True
    Else
        SaveData = False
    End If
End Function

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim wsSQL As String
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "文件號編碼"
    vFilterAry(1, 2) = "VouPrefix"
    
    vFilterAry(2, 1) = "長度"
    vFilterAry(2, 2) = "VouLen"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "文件號編碼"
    vAry(1, 2) = "VouPrefix"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "長度"
    vAry(2, 2) = "VouLen"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT sysVouNo.VouPrefix, sysVouNo.VouLen "
        wsSQL = wsSQL + "FROM sysVouNo "
        .sBindSQL = wsSQL
        .sBindWhereSQL = ""
        .sBindOrderSQL = "ORDER BY sysVouNo.VouPrefix"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboVouPrefix Then
        cboVouPrefix = Trim(frmShareSearch.Tag)
        cboVouPrefix.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
End Sub

Private Sub txtVouDesc_GotFocus()
    FocusMe txtVouDesc
End Sub

Private Sub txtVouDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVouDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
            txtVouLen.SetFocus
    End If
End Sub

Private Sub txtVouDesc_LostFocus()
    FocusMe txtVouDesc, True
End Sub

Private Sub txtVouLastKey_GotFocus()
    FocusMe txtVouLastKey
End Sub

Private Sub txtVouLastKey_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtVouLastKey, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        chkVouYrFlg.SetFocus
    End If
End Sub

Private Sub txtVouLastKey_LostFocus()
    FocusMe txtVouLastKey, True
End Sub

Private Sub txtVouLen_GotFocus()
    FocusMe txtVouLen
End Sub

Private Sub txtVouLen_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtVouLen, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtVouLen() = True Then
            txtVouLastKey.SetFocus
        End If
    End If
End Sub

Private Sub txtVouLen_LostFocus()
    FocusMe txtVouLen, True
End Sub

Private Sub txtVouPrefix_GotFocus()
    FocusMe txtVouPrefix
End Sub

Private Sub txtVouPrefix_LostFocus()
    FocusMe txtVouPrefix, True
End Sub

Private Sub txtVouPrefix_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVouPrefix, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtVouPrefix() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Function Chk_txtVouLen() As Boolean
    Chk_txtVouLen = False
    
    If Trim(txtVouLen.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVouLen.SetFocus
        Exit Function
    End If
    
    If Not (txtVouLen >= 0 And txtVouLen <= 12) Then
        gsMsg = "長度必須少於十二!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVouLen.SetFocus
        Exit Function
    End If
    
    Chk_txtVouLen = True
End Function

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
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If

End Sub

Private Sub cboVouPrefix_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(cboVouPrefix, 3, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboVouPrefix() = True Then
            Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub cboVouPrefix_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboVouPrefix
    
    wsSQL = "SELECT VouPrefix, VouDesc FROM sysVouNo WHERE "
    wsSQL = wsSQL & " VouPrefix LIKE '%" & IIf(cboVouPrefix.SelLength > 0, "", Set_Quote(cboVouPrefix.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY VouPrefix "
    Call Ini_Combo(2, wsSQL, cboVouPrefix.Left, cboVouPrefix.Top + cboVouPrefix.Height, tblCommon, wsFormID, "TBLVOUPREFIX", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboVouPrefix_GotFocus()
    FocusMe cboVouPrefix
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT VouPrefix FROM SysVouNo WHERE VouPrefix = '" & Set_Quote(txtVouPrefix) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_KeyExist = True
    Else
        Chk_KeyExist = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "VouPrefix"
        .KeyLen = 10
        Set .ctlKey = txtVouPrefix
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function Chk_cboVouType() As Boolean
    Chk_cboVouType = False
    
    If UCase(cboVouType) = "" Then
        gsMsg = "必需輸入文件類別!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVouType.SetFocus
        Exit Function
    End If
    
    If Chk_VouType(cboVouType.Text) = False Then
        gsMsg = "文件類別不正確!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVouType.SetFocus
        Exit Function
    End If
    
    Chk_cboVouType = True
End Function

Private Function Chk_VouType(ByVal inCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_VouType = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT MCModNo "
    wsSQL = wsSQL & " FROM sysMonCtl WHERE sysMonCtl.MCModNo = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        Chk_VouType = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub txtVouSpa_GotFocus()
    FocusMe txtVouSpa
End Sub

Private Sub txtVouSpa_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVouSpa, 1, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboVouType.SetFocus
    End If
End Sub

Private Sub txtVouSpa_LostFocus()
    FocusMe txtVouSpa, True
End Sub

