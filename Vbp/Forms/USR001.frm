VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUSR001 
   BackColor       =   &H8000000A&
   Caption         =   "USR001"
   ClientHeight    =   3450
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "USR001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "USR001.frx":0E42
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboUsrGrpCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   2730
   End
   Begin VB.ComboBox cboUsrCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "USRINFO"
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtUsrPwd1 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  '暫止
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtUsrPwd 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  '暫止
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtUsrCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtUsrName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lblUsrPwd1 
         Caption         =   "USRPWD1"
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   2100
         Width           =   1380
      End
      Begin VB.Label lblUsrPwd 
         Caption         =   "USRPWD"
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label lblUsrLastUpd 
         Caption         =   "USRLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   2565
         Width           =   1500
      End
      Begin VB.Label lblUsrLastUpdDate 
         Caption         =   "USRLASTUPDDATE"
         Height          =   240
         Left            =   4320
         TabIndex        =   13
         Top             =   2565
         Width           =   1500
      End
      Begin VB.Label lblDspUsrLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   2520
         Width           =   2265
      End
      Begin VB.Label lblDspUsrLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5880
         TabIndex        =   11
         Top             =   2520
         Width           =   2265
      End
      Begin VB.Label lblUsrCode 
         Caption         =   "USRCODE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblUsrGrpCode 
         Caption         =   "USRGRPCODE"
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label lblUsrName 
         Caption         =   "USRNAME"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1020
         Width           =   1215
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
            Picture         =   "USR001.frx":3545
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":3E1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":46F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":4B4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":4F9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":52B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":5709
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":5B5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":5E75
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":618F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":65E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "USR001.frx":6EBD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   7
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
Attribute VB_Name = "frmUSR001"
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

Private Const wsKeyType = "MstUser"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboUsrCode_LostFocus()
    FocusMe cboUsrCode, True
End Sub



Private Sub cboUsrGrpCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboUsrGrpCode
    
    wsSQL = "SELECT DISTINCT UsrGrpCode FROM MstUser "
    wsSQL = wsSQL & "ORDER BY UsrGrpCode "
    Call Ini_Combo(1, wsSQL, cboUsrGrpCode.Left, cboUsrGrpCode.Top + cboUsrGrpCode.Height, tblCommon, wsFormID, "TBLGRP", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
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
        Me.Height = 3855
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
            Me.txtUsrName.Enabled = False
            Me.cboUsrGrpCode.Enabled = False
            Me.txtUsrPwd.Enabled = False
            Me.txtUsrPwd1.Enabled = False
            
            Me.cboUsrCode.Enabled = False
            Me.cboUsrCode.Visible = False
            Me.txtUsrCode.Visible = True
            Me.txtUsrCode.Enabled = False
            
        Case "AfrActAdd"
            Me.cboUsrCode.Enabled = False
            Me.cboUsrCode.Visible = False
            
            Me.txtUsrCode.Enabled = True
            Me.txtUsrCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboUsrCode.Enabled = True
            Me.cboUsrCode.Visible = True
            
            Me.txtUsrCode.Enabled = False
            Me.txtUsrCode.Visible = False
            
        Case "AfrKey"
            Me.cboUsrCode.Enabled = False
            Me.txtUsrCode.Enabled = False
            
            Me.txtUsrName.Enabled = True
            Me.cboUsrGrpCode.Enabled = True
            Me.txtUsrPwd.Enabled = True
            Me.txtUsrPwd1.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtUsrName = False Then
        Exit Function
    End If
    
    If Chk_cboUsrGrpCode = False Then
        Exit Function
    End If
    
    If Chk_txtUsrPwd = False Then
        Exit Function
    End If
    
    If Chk_txtUsrPwd1 = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From MstUser "
    wsSQL = wsSQL + "WHERE (MstUser.UsrCode)='" + Set_Quote(cboUsrCode.Text) + "'"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False

    Else
        Me.cboUsrCode = ReadRs(rsRcd, "UsrCode")
        Me.txtUsrName = ReadRs(rsRcd, "UsrName")
        Me.cboUsrGrpCode = ReadRs(rsRcd, "UsrGrpCode")
        Me.lblDspUsrLastUpd = ReadRs(rsRcd, "UsrLastUpd")
        Me.lblDspUsrLastUpdDate = ReadRs(rsRcd, "UsrLastUpdDate")
        
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
    Set frmUSR001 = Nothing
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

'    Me.Left = 0
'    Me.Top = 0
'    Me.Width = Screen.Width
'    Me.Height = Screen.Height
    
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "USR001"
    wsTrnCd = ""
    
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblUsrCode.Caption = Get_Caption(waScrItm, "USRCODE")
    lblUsrName.Caption = Get_Caption(waScrItm, "USRNAME")
    lblUsrGrpCode.Caption = Get_Caption(waScrItm, "USRGRPCODE")
    lblUsrPwd.Caption = Get_Caption(waScrItm, "USRPWD")
    lblUsrPwd1.Caption = Get_Caption(waScrItm, "USRPWD1")
    lblUsrLastUpd.Caption = Get_Caption(waScrItm, "USRLASTUPD")
    lblUsrLastUpdDate.Caption = Get_Caption(waScrItm, "USRLASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "USRADD")
    wsActNam(2) = Get_Caption(waScrItm, "USREDIT")
    wsActNam(3) = Get_Caption(waScrItm, "USRDELETE")
    
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
        txtUsrCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboUsrCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboUsrCode.SetFocus
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
                If RowLock(wsConnTime, wsKeyType, cboUsrCode, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtUsrName.SetFocus
End Sub

Private Function Chk_UsrCode(ByVal inCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_UsrCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT UsrCode "
    wsSQL = wsSQL & " FROM MstUser WHERE UsrCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    
    Chk_UsrCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtUsrCode() As Boolean
    Chk_txtUsrCode = False
    
        If Trim(txtUsrCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtUsrCode.SetFocus
            Exit Function
        End If
    
        If Chk_UsrCode(txtUsrCode.Text) = True Then
            gsMsg = "用戶已存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtUsrCode.SetFocus
            Exit Function
        End If
    
    Chk_txtUsrCode = True
End Function

Private Function Chk_cboUsrCode() As Boolean
    Chk_cboUsrCode = False
    
    If Trim(cboUsrCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboUsrCode.SetFocus
        Exit Function
    End If

    If Chk_UsrCode(cboUsrCode.Text) = False Then
        gsMsg = "用戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboUsrCode.SetFocus
        Exit Function
    End If

    Chk_cboUsrCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmUSR001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboUsrCode, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_USR001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, UCase(IIf(wiAction = AddRec, txtUsrCode, cboUsrCode)))
    Call SetSPPara(adcmdSave, 3, UCase(txtUsrName))
    Call SetSPPara(adcmdSave, 4, UCase(cboUsrGrpCode))
    Call SetSPPara(adcmdSave, 5, Encrypt(UCase(Set_Quote(txtUsrPwd))))
    Call SetSPPara(adcmdSave, 6, gsUserID)
    Call SetSPPara(adcmdSave, 7, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 8)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - USR001!"
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
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "用戶編碼"
    vFilterAry(1, 2) = "UsrCode"
    
    vFilterAry(2, 1) = "名稱"
    vFilterAry(2, 2) = "UsrName"
    
    vFilterAry(3, 1) = "群組"
    vFilterAry(3, 2) = "UsrGrpCode"
    
    ReDim vAry(3, 3)
    vAry(1, 1) = "用戶編碼"
    vAry(1, 2) = "UsrCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "名稱"
    vAry(2, 2) = "UsrName"
    vAry(2, 3) = "5000"
    
    vAry(3, 1) = "群組"
    vAry(3, 2) = "UsrGrpCode"
    vAry(3, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT MstUser.UsrCode, MstUser.UsrName, MstUser.UsrGrpCode "
        wsSQL = wsSQL + "FROM MstUser "
        .sBindSQL = wsSQL
        .sBindWhereSQL = ""
        .sBindOrderSQL = "ORDER BY MstUser.UsrCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboUsrCode Then
        cboUsrCode = Trim(frmShareSearch.Tag)
        cboUsrCode.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
    
End Sub

Private Sub txtUsrCode_GotFocus()
    FocusMe txtUsrCode
End Sub

Private Sub txtUsrCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtUsrCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtUsrCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtUsrCode_LostFocus()
    FocusMe txtUsrCode, True
End Sub

Private Sub txtUsrName_LostFocus()
    FocusMe txtUsrName, True
End Sub
Private Sub cboUsrGrpCode_GotFocus()
    FocusMe cboUsrGrpCode
End Sub

Private Sub cboUsrGrpCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboUsrGrpCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboUsrGrpCode() = True Then
            txtUsrPwd.SetFocus
        End If
    End If
End Sub

Private Sub cboUsrGrpCode_LostFocus()
    FocusMe cboUsrGrpCode, True
End Sub

Private Sub txtUsrName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtUsrName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtUsrName() = True Then
            cboUsrGrpCode.SetFocus
        End If
    End If
End Sub


Private Sub txtUsrName_GotFocus()
    FocusMe txtUsrName
End Sub

Private Function Chk_txtUsrName() As Boolean
    
    Chk_txtUsrName = False
    
    If Trim(txtUsrName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtUsrName.SetFocus
        Exit Function
    End If
    
    Chk_txtUsrName = True
End Function

Private Function Chk_cboUsrGrpCode() As Boolean
    Chk_cboUsrGrpCode = False
    
    If Trim(cboUsrGrpCode.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboUsrGrpCode.SetFocus
        Exit Function
    End If
    
    Chk_cboUsrGrpCode = True
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

Private Sub cboUsrCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboUsrCode, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboUsrCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboUsrCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboUsrCode
    
    wsSQL = "SELECT UsrCode, UsrName FROM MstUser "
    wsSQL = wsSQL & "ORDER BY UsrCode "
    Call Ini_Combo(2, wsSQL, cboUsrCode.Left, cboUsrCode.Top + cboUsrCode.Height, tblCommon, wsFormID, "TBLUSR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboUsrCode_GotFocus()
    FocusMe cboUsrCode
End Sub



Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT UsrCode FROM MstUser WHERE UsrCode = '" & Set_Quote(txtUsrCode) & "'"
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
        .TableKey = "UsrCode"
        .KeyLen = 10
        Set .ctlKey = txtUsrCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub txtUsrPwd_GotFocus()
    FocusMe txtUsrPwd
End Sub

Private Sub txtUsrPwd_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtUsrPwd, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtUsrPwd() = True Then
            txtUsrPwd1.SetFocus
        End If
    End If
End Sub

Private Sub txtUsrPwd_LostFocus()
    FocusMe txtUsrPwd, True
End Sub

Private Function Chk_txtUsrPwd() As Boolean
    Chk_txtUsrPwd = False
    
    If Trim(txtUsrPwd.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtUsrPwd.SetFocus
        Exit Function
    End If
    
    Chk_txtUsrPwd = True
End Function

Private Function Chk_txtUsrPwd1() As Boolean
    Chk_txtUsrPwd1 = False
    
    If Trim(txtUsrPwd1.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtUsrPwd1.SetFocus
        Exit Function
    End If
    
    If Trim(txtUsrPwd1.Text) <> Trim(txtUsrPwd.Text) Then
        gsMsg = "密碼確認失敗!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtUsrPwd1.SetFocus
        Exit Function
    End If
    
    Chk_txtUsrPwd1 = True
End Function

Private Sub txtUsrPwd1_LostFocus()
    FocusMe txtUsrPwd1, True
End Sub

Private Sub txtUsrPwd1_GotFocus()
    FocusMe txtUsrPwd1
End Sub

Private Sub txtUsrPwd1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtUsrPwd1, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtUsrPwd1() = True Then
            txtUsrName.SetFocus
        End If
    End If
End Sub

