VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmCR001 
   BackColor       =   &H8000000A&
   Caption         =   "CR001"
   ClientHeight    =   3450
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "CR001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "CR001.frx":08CA
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCRegCusCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtCRegPrefix 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   1680
         Width           =   1290
      End
      Begin VB.TextBox txtCRegRegNo 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   1320
         Width           =   1290
      End
      Begin VB.TextBox txtCRegLen 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label lblCRegPrefix 
         Caption         =   "CREGPREFIX"
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label lblCRegLastUpd 
         Caption         =   "CREGLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   2445
         Width           =   1500
      End
      Begin VB.Label lblCRegLastUpdDate 
         Caption         =   "CREGLASTUPDDATE"
         Height          =   240
         Left            =   4320
         TabIndex        =   11
         Top             =   2445
         Width           =   1500
      End
      Begin VB.Label lblDspCRegLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   10
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label lblDspCRegLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5880
         TabIndex        =   9
         Top             =   2400
         Width           =   2265
      End
      Begin VB.Label lblCusCode 
         Caption         =   "CUSCODE"
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
         TabIndex        =   8
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblCRegRegNo 
         Caption         =   "CREGREGNO"
         Height          =   240
         Left            =   360
         TabIndex        =   7
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label lblCRegLen 
         Caption         =   "CREGLEN"
         Height          =   255
         Left            =   360
         TabIndex        =   6
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
            Picture         =   "CR001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CR001.frx":6945
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
Attribute VB_Name = "frmCR001"
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

Private Const wsKeyType = "MstCusReg"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboCRegCusCode_LostFocus()
    FocusMe cboCRegCusCode, True
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
            Me.txtCRegRegNo.Enabled = False
            Me.txtCRegLen.Enabled = False
            Me.txtCRegPrefix.Enabled = False
            
            Me.cboCRegCusCode.Enabled = False
            
            txtCRegRegNo.Text = "0"
            txtCRegLen.Text = "0"
            
        Case "AfrActAdd"
            Me.cboCRegCusCode.Enabled = True
            
        Case "AfrActEdit"
            Me.cboCRegCusCode.Enabled = True
            
        Case "AfrKey"
            Me.cboCRegCusCode.Enabled = False
            
            Me.txtCRegRegNo.Enabled = True
            Me.txtCRegLen.Enabled = True
            Me.txtCRegPrefix.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    InputValidation = False
    
    If Chk_txtCRegLen = False Then
        Exit Function
    End If
    
    If Chk_txtCRegRegNo = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    Dim wsCRegCusID As String
    
    LoadCusIDByCode wsCRegCusID, cboCRegCusCode

    wsSql = "SELECT MstCustomer.CusCode, MstCusReg.* "
    wsSql = wsSql + "From MstCustomer, MstCusReg "
    wsSql = wsSql + "WHERE (((MstCusReg.CRegCusID)='" + Set_Quote(wsCRegCusID) + "' AND CRegStatus = '1'))"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        'Me.cboCRegCusCode = ReadRs(rsRcd, "CusCode")
        Me.txtCRegLen = To_Value(ReadRs(rsRcd, "CRegLen"))
        Me.txtCRegRegNo = To_Value(ReadRs(rsRcd, "CRegRegNo"))
        Me.txtCRegPrefix = ReadRs(rsRcd, "CRegPrefix")
        
        Me.lblDspCRegLastUpd = ReadRs(rsRcd, "CRegLastUpd")
        Me.lblDspCRegLastUpdDate = ReadRs(rsRcd, "CRegLastUpdDate")
        
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
    
    Set frmCR001 = Nothing
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
                gsMsg = "你是否確定要放棄現時之作業?"
                If MsgBox(gsMsg, vbYesNo, gsTitle) = vbYes Then
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
    wsFormID = "CR001"
    wsTrnCd = ""
    
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCRegLen.Caption = Get_Caption(waScrItm, "CREGLEN")
    lblCRegRegNo.Caption = Get_Caption(waScrItm, "CREGREGNO")
    lblCRegPrefix.Caption = Get_Caption(waScrItm, "CREGPREFIX")
    lblCRegLastUpd.Caption = Get_Caption(waScrItm, "CREGLASTUPD")
    lblCRegLastUpdDate.Caption = Get_Caption(waScrItm, "CREGLASTUPDDATE")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
   
    wsActNam(1) = Get_Caption(waScrItm, "CRADD")
    wsActNam(2) = Get_Caption(waScrItm, "CREDIT")
    wsActNam(3) = Get_Caption(waScrItm, "CRDELETE")
    
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
        cboCRegCusCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCRegCusCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCRegCusCode.SetFocus
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
                If RowLock(wsConnTime, wsKeyType, cboCRegCusCode, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtCRegLen.SetFocus
End Sub

Private Function Chk_CRegCusCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_CRegCusCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT CusStatus "
    wsSql = wsSql & " FROM MstCustomer WHERE CusCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "CusStatus")
    
    Chk_CRegCusCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCRegCusCode() As Boolean
    Dim wsStatus As String
 
    Chk_cboCRegCusCode = False
    
    If Trim(cboCRegCusCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCRegCusCode.SetFocus
        Exit Function
    End If

    If Chk_CRegCusCode(cboCRegCusCode.Text, wsStatus) = False Then
        gsMsg = "客戶不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCRegCusCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "客戶已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCRegCusCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboCRegCusCode = True
End Function


Private Sub cmdOpen()
    Dim newForm As New frmCR001
    
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
    Dim wsCRegCusID As String
    
    LoadCusIDByCode wsCRegCusID, cboCRegCusCode
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboCRegCusCode, wsFormID) Then
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
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_CR001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsCRegCusID)
    Call SetSPPara(adcmdSave, 3, txtCRegLen)
    Call SetSPPara(adcmdSave, 4, txtCRegRegNo)
    Call SetSPPara(adcmdSave, 5, txtCRegPrefix)
    Call SetSPPara(adcmdSave, 6, gsUserID)
    Call SetSPPara(adcmdSave, 7, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 8)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - CR001!"
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
       gsMsg = "你是否確定不儲存現時之變更而離開?"
       If MsgBox(gsMsg, vbYesNo, gsTitle) = vbYes Then
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
    Dim wsSql As String
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "客戶編碼"
    vFilterAry(1, 2) = "CusCode"
    
    vFilterAry(2, 1) = "名稱"
    vFilterAry(2, 2) = "CusName"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "客戶編碼"
    vAry(1, 2) = "CusCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "名稱"
    vAry(2, 2) = "CusName"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSql = "SELECT MstCustomer.CusCode, MstCustomer.CusName "
        wsSql = wsSql + "FROM MstCustomer, MstCusReg "
        .sBindSQL = wsSql
        .sBindWhereSQL = "WHERE MstCusReg.CRegStatus = '1' AND MstCustomer.CusID = MstCusReg.CRegCusID "
        .sBindOrderSQL = "ORDER BY MstCustomer.CusCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboCRegCusCode Then
        cboCRegCusCode = Trim(frmShareSearch.Tag)
        cboCRegCusCode.SetFocus
        SendKeys "{Enter}"
    End If
End Sub

Private Sub txtCRegLen_LostFocus()
    FocusMe txtCRegLen, True
End Sub



Private Sub txtCRegLen_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtCRegLen, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtCRegLen() = True Then
            txtCRegRegNo.SetFocus
        End If
    End If
End Sub

Private Sub txtCRegRegNo_GotFocus()
    FocusMe txtCRegRegNo
End Sub
Private Sub txtCRegRegNo_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtCRegRegNo, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtCRegRegNo() = True Then
            txtCRegPrefix.SetFocus
        End If
    End If
End Sub
Private Sub txtCRegRegNo_LostFocus()
    FocusMe txtCRegRegNo, True
End Sub

Private Sub txtCRegPrefix_GotFocus()
    FocusMe txtCRegPrefix
End Sub
Private Sub txtCRegPrefix_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCRegPrefix, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
            txtCRegLen.SetFocus
        
    End If
End Sub
Private Sub txtCRegPrefix_LostFocus()
    FocusMe txtCRegPrefix, True
End Sub

Private Sub txtCRegLen_GotFocus()
    FocusMe txtCRegLen
End Sub

Private Function Chk_txtCRegLen() As Boolean
    Chk_txtCRegLen = False
    
    If Trim(txtCRegLen.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCRegLen.SetFocus
        Exit Function
    End If
    
    If txtCRegLen.Text > 15 Or txtCRegLen.Text < 0 Then
        gsMsg = "輸入之資料不得小於零或大於十五!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCRegLen.SetFocus
        Exit Function
    End If
    
    Chk_txtCRegLen = True
End Function

Private Function Chk_txtCRegRegNo() As Boolean
    Chk_txtCRegRegNo = False
    
    If Trim(txtCRegRegNo.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCRegRegNo.SetFocus
        Exit Function
    End If
    
    Chk_txtCRegRegNo = True
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

Private Sub cboCRegCusCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCRegCusCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCRegCusCode() = True Then
            Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub cboCRegCusCode_DropDown()
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCRegCusCode
    
    If wiAction = AddRec Then
        wsSql = "SELECT MstCustomer.CusCode, MstCustomer.CusName FROM MstCustomer WHERE NOT EXISTS (SELECT CRegCusID FROM MstCusReg WHERE CRegCusID = MstCustomer.CusID)"
        wsSql = wsSql & "ORDER BY MstCustomer.CusCode "
    Else
        wsSql = "SELECT MstCustomer.CusCode, MstCustomer.CusName FROM MstCustomer, MstCusReg WHERE CRegStatus = '1' "
        wsSql = wsSql & "AND MstCustomer.CusID = MstCusReg.CRegCusID "
        wsSql = wsSql & "ORDER BY MstCustomer.CusCode "
    End If
    
    Call Ini_Combo(2, wsSql, cboCRegCusCode.Left, cboCRegCusCode.Top + cboCRegCusCode.Height, tblCommon, "CR001", "TBLCUS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCRegCusCode_GotFocus()
    FocusMe cboCRegCusCode
End Sub



Public Sub LoadCusIDByCode(outText As String, ByVal inCusCode)
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT CusID FROM MstCustomer WHERE CusCode ='" + Set_Quote(inCusCode) + "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
          outText = ReadRs(rsRcd, "CusID")
    Else
          outText = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
End Sub


