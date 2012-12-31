VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSYS002 
   BackColor       =   &H8000000A&
   Caption         =   "文件號"
   ClientHeight    =   3450
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "SYS002.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "SYS002.frx":08CA
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboDocType 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "文件號"
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtDocLastYr 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   3720
         TabIndex        =   7
         Top             =   2400
         Width           =   690
      End
      Begin VB.CheckBox chkDocYear 
         Alignment       =   1  '靠右對齊
         Caption         =   "DOCYEAR"
         Height          =   180
         Left            =   360
         TabIndex        =   6
         Top             =   2480
         Width           =   2415
      End
      Begin VB.TextBox txtDocLastKey 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   2040
         Width           =   2730
      End
      Begin VB.TextBox txtDocLen 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   1680
         Width           =   2730
      End
      Begin VB.TextBox txtDocType 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtDocPrefix 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   2730
      End
      Begin VB.TextBox txtDocTypeDesc 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   6450
      End
      Begin VB.Label lblDocLastYr 
         Caption         =   "DOCLASTYR"
         Height          =   180
         Left            =   2880
         TabIndex        =   16
         Top             =   2475
         Width           =   780
      End
      Begin VB.Label lblDocLastKey 
         Caption         =   "DOCLASTKEY"
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   2100
         Width           =   1380
      End
      Begin VB.Label lblDocLen 
         Caption         =   "DOCLEN"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label lblDocType 
         Caption         =   "DOCTYPE"
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
         TabIndex        =   12
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblDocPrefix 
         Caption         =   "DOCPREFIX"
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label lblDocTypeDesc 
         Caption         =   "DOCTYPEDESC"
         Height          =   255
         Left            =   360
         TabIndex        =   10
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
            Picture         =   "SYS002.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS002.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   9
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
Attribute VB_Name = "frmSYS002"
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

Private Const wsKeyType = "SysDocNo"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboDocType_LostFocus()
    FocusMe cboDocType, True
End Sub

Private Sub chkDocYear_Click()
    If chkDocYear.Value = 1 Then
        txtDocLastYr.Enabled = True
    Else
        txtDocLastYr.Enabled = False
    End If
End Sub

Private Sub chkDocYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtDocLastYr.Enabled = True Then
            txtDocLastYr.SetFocus
        Else
            txtDocTypeDesc.SetFocus
        End If
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
            Me.txtDocTypeDesc.Enabled = False
            Me.txtDocPrefix.Enabled = False
            Me.txtDocLen.Enabled = False
            Me.txtDocLastKey.Enabled = False
            
            Me.cboDocType.Enabled = False
            Me.cboDocType.Visible = False
            Me.txtDocType.Visible = True
            Me.txtDocType.Enabled = False
            
        Case "AfrActAdd"
            Me.cboDocType.Enabled = False
            Me.cboDocType.Visible = False
            
            Me.txtDocType.Enabled = True
            Me.txtDocType.Visible = True
            
        Case "AfrActEdit"
            Me.cboDocType.Enabled = True
            Me.cboDocType.Visible = True
            
            Me.txtDocType.Enabled = False
            Me.txtDocType.Visible = False
            
        Case "AfrKey"
            Me.cboDocType.Enabled = False
            Me.txtDocType.Enabled = False
            
            Me.txtDocTypeDesc.Enabled = True
            Me.txtDocPrefix.Enabled = True
            Me.txtDocLen.Enabled = True
            Me.txtDocLastKey.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtDocTypeDesc() = False Then
        Exit Function
    End If
    
    If chkDocYear.Value = 1 Then
        If Chk_txtDocLastYr() = False Then
            Exit Function
        End If
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From sysDocNo "
    wsSQL = wsSQL + "WHERE (((sysDocNo.DocType)='" + Set_Quote(cboDocType.Text) + "'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False

    Else
        Me.cboDocType = ReadRs(rsRcd, "DocType")
        Me.txtDocTypeDesc = ReadRs(rsRcd, "DocTypeDesc")
        Me.txtDocPrefix = ReadRs(rsRcd, "DocPrefix")
        Me.txtDocLen = ReadRs(rsRcd, "DocLen")
        Me.txtDocLastKey = ReadRs(rsRcd, "DocLastKey")
        
        If ReadRs(rsRcd, "DocYear") = "Y" Then
            chkDocYear.Value = 1
            txtDocLastYr.Enabled = True
            txtDocLastYr = ReadRs(rsRcd, "DocLastYr")
        Else
            chkDocYear.Value = 0
            txtDocLastYr.Enabled = False
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
    Set frmSYS002 = Nothing
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
    wsFormID = "SYS002"
    wsTrnCd = ""
    
End Sub


Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocType.Caption = Get_Caption(waScrItm, "DOCTYPE")
    lblDocTypeDesc.Caption = Get_Caption(waScrItm, "DOCTYPEDESC")
    lblDocPrefix.Caption = Get_Caption(waScrItm, "DOCPREFIX")
    lblDocLen.Caption = Get_Caption(waScrItm, "DOCLEN")
    lblDocLastKey.Caption = Get_Caption(waScrItm, "DOCLASTKEY")
    chkDocYear.Caption = Get_Caption(waScrItm, "DOCYEAR")
    lblDocLastYr.Caption = Get_Caption(waScrItm, "DOCLASTYR")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")

    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"

    wsActNam(1) = Get_Caption(waScrItm, "SYSADD")
    wsActNam(2) = Get_Caption(waScrItm, "SYSEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SYSDELETE")
    
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
    chkDocYear.Value = 0
    
    
    tblCommon.Visible = False
    Me.Caption = wsFormCaption
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
              
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
        txtDocType.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboDocType.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboDocType.SetFocus
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
                If RowLock(wsConnTime, wsKeyType, cboDocType, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtDocTypeDesc.SetFocus
End Sub

Private Function Chk_DocType(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_DocType = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT DocType "
    wsSQL = wsSQL & " FROM SysDocNo WHERE DocType = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "DocType")
    
    Chk_DocType = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtDocType() As Boolean
    Dim wsStatus As String

    Chk_txtDocType = False
    
        If Trim(txtDocType.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtDocType.SetFocus
            Exit Function
        End If
    
        If Chk_DocType(txtDocType.Text, wsStatus) = True Then
            gsMsg = "文件號已存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtDocType.SetFocus
            Exit Function
        End If
    
    Chk_txtDocType = True
End Function

Private Function Chk_cboDocType() As Boolean
    Dim wsStatus As String
 
    Chk_cboDocType = False
    
    If Trim(cboDocType.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboDocType.SetFocus
        Exit Function
    End If

    If Chk_DocType(cboDocType.Text, wsStatus) = False Then
        gsMsg = "文件號不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboDocType.SetFocus
        Exit Function
    End If
    
    Chk_cboDocType = True
    
End Function

Private Sub cmdOpen()
    Dim newForm As New frmSYS002
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboDocType, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_SYS002"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtDocType.Text, cboDocType.Text))
    Call SetSPPara(adcmdSave, 3, txtDocTypeDesc)
    Call SetSPPara(adcmdSave, 4, txtDocPrefix)
    Call SetSPPara(adcmdSave, 5, txtDocLen)
    Call SetSPPara(adcmdSave, 6, txtDocLastKey)
    Call SetSPPara(adcmdSave, 7, IIf(chkDocYear.Value = 1, "Y", "N"))
    Call SetSPPara(adcmdSave, 8, IIf(chkDocYear.Value = 1, txtDocLastYr, "50"))
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 9)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - SYS002!"
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
    vFilterAry(1, 2) = "DocType"
    
    vFilterAry(2, 1) = "註解"
    vFilterAry(2, 2) = "DocTypeDesc"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "文件號編碼"
    vAry(1, 2) = "DocTypeCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "註解"
    vAry(2, 2) = "DocTypeDesc"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT sysDocNo.DocType, sysDocNo.DocTypeDesc "
        wsSQL = wsSQL + "FROM sysDocNo "
        .sBindSQL = wsSQL
        .sBindWhereSQL = ""
        .sBindOrderSQL = "ORDER BY sysDocNo.DocType"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboDocType Then
        cboDocType = Trim(frmShareSearch.Tag)
        cboDocType.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
    
End Sub

Private Sub txtDocLastKey_GotFocus()
    FocusMe txtDocLastKey
End Sub

Private Sub txtDocLastKey_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtDocLastKey, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        chkDocYear.SetFocus
    End If
End Sub

Private Sub txtDocLastKey_LostFocus()
    FocusMe txtDocLastKey, True
End Sub

Private Sub txtDocLastYr_GotFocus()
    FocusMe txtDocLastYr
End Sub

Private Sub txtDocLastYr_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtDocLastYr, False, False)
    Call chk_InpLen(txtDocLastYr, 2, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtDocLastYr() = True Then
            txtDocTypeDesc.SetFocus
        End If
    End If
End Sub

Private Sub txtDocLastYr_LostFocus()
    FocusMe txtDocLastYr, True
End Sub

Private Sub txtDocLen_GotFocus()
    FocusMe txtDocLen
End Sub

Private Sub txtDocLen_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtDocType, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtDocLen() = True Then
            txtDocLastKey.SetFocus
        End If
    End If
End Sub

Private Sub txtDocLen_LostFocus()
    FocusMe txtDocLen, True
End Sub

Private Sub txtDocPrefix_GotFocus()
    FocusMe txtDocPrefix
End Sub

Private Sub txtDocPrefix_LostFocus()
    FocusMe txtDocPrefix, True
End Sub

Private Sub txtDocType_GotFocus()
    FocusMe txtDocType
End Sub

Private Sub txtDocType_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtDocType, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtDocType() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtDocType_LostFocus()
    FocusMe txtDocType, True
End Sub

Private Sub txtDocTypeDesc_GotFocus()
    FocusMe txtDocTypeDesc
End Sub

Private Sub txtDocTypeDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtDocTypeDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtDocTypeDesc() = True Then
            txtDocPrefix.SetFocus
        End If
    End If
End Sub

Private Sub txtDocTypeDesc_LostFocus()
    FocusMe txtDocTypeDesc, True
End Sub

Private Sub txtDocPrefix_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtDocPrefix, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtDocLen.SetFocus
    End If
End Sub

Private Function Chk_txtDocTypeDesc() As Boolean
    Chk_txtDocTypeDesc = False
    
    If Trim(txtDocTypeDesc.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtDocTypeDesc.SetFocus
        Exit Function
    End If
    
    Chk_txtDocTypeDesc = True
End Function

Private Function Chk_txtDocLen() As Boolean
    Chk_txtDocLen = False
    
    If Trim(txtDocLen.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtDocLen.SetFocus
        Exit Function
    End If
    
    If Not (txtDocLen >= 0 And txtDocLen <= 12) Then
        gsMsg = "長度必須少於十二!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtDocLen.SetFocus
        Exit Function
    End If
    
    Chk_txtDocLen = True
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

Private Sub cboDocType_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocType, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboDocType() = True Then
            Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub cboDocType_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboDocType
    
    wsSQL = "SELECT DocType, DocTypeDesc FROM sysDocNo WHERE "
    wsSQL = wsSQL & " DocType LIKE '%" & IIf(cboDocType.SelLength > 0, "", Set_Quote(cboDocType.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY DocType "
    Call Ini_Combo(2, wsSQL, cboDocType.Left, cboDocType.Top + cboDocType.Height, tblCommon, "SYS002", "TBLDOCTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocType_GotFocus()
    FocusMe cboDocType
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT DocType FROM SysDocNo WHERE DocType = '" & Set_Quote(txtDocType) & "'"
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
        .TableKey = "DocType"
        .KeyLen = 10
        Set .ctlKey = txtDocType
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function Chk_txtDocLastYr() As Boolean
    Chk_txtDocLastYr = False
    
    If Trim(txtDocLastYr.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtDocLastYr.SetFocus
        Exit Function
    End If
    
    If Len(txtDocLastYr) <> 2 Then
        gsMsg = "長度等於二!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtDocLastYr.SetFocus
        Exit Function
    End If
    
    Chk_txtDocLastYr = True
End Function

