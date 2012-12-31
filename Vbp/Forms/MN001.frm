VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMN001 
   BackColor       =   &H8000000A&
   Caption         =   "Master Nature Master Maintenance"
   ClientHeight    =   3450
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "MN001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   Begin VB.Frame fraDetailInfo 
      Caption         =   "Method Nature"
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtMethodNatureCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Tag             =   "K"
         Top             =   600
         Width           =   2730
      End
      Begin VB.CommandButton btnMethodNatureCode 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Tag             =   "K"
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtMethodNatureDesc 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label lblMethodNatureLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   360
         TabIndex        =   9
         Top             =   2445
         Width           =   1140
      End
      Begin VB.Label lblMethodNatureLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4320
         TabIndex        =   8
         Top             =   2445
         Width           =   1260
      End
      Begin VB.Label lblDspMethodNatureLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   2400
         Width           =   2505
      End
      Begin VB.Label lblDspMethodNatureLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5640
         TabIndex        =   6
         Top             =   2400
         Width           =   2505
      End
      Begin VB.Label lblMethodNatureCode 
         Caption         =   "編碼 :"
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblMethodNatureDesc 
         Caption         =   "註解 :"
         Height          =   240
         Left            =   360
         TabIndex        =   4
         Top             =   1020
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
            Picture         =   "MN001.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":1ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":2322
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":263C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":2A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":2EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":31FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":3514
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":3966
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MN001.frx":4242
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
            Object.Visible         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "尋找 (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
Attribute VB_Name = "frmMN001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public oCalledForm As Form

Private isPrompted As Boolean
Private wiSalesmanID As Integer
Private wsFormCaption As String
 
Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"

Dim wiAction As Integer
Private wsKey As String
Private wsDocTyp As String
Private wsFormID As String
Private wsConnTime As String

Private Const wsKeyType = "soaM"
Dim wsUsrId As String


Private Sub btnMethodNatureCode_Click()
    ReDim vAry(3, 3)
    vAry(1, 1) = ""
    vAry(1, 2) = "MethodNatureCode"
    vAry(1, 3) = "30"
    
    vAry(2, 1) = "Method Nature "
    vAry(2, 2) = "MethodNatureCode"
    vAry(2, 3) = "100"
    
    vAry(3, 1) = "註解"
    vAry(3, 2) = "MethodNatureDesc"
    vAry(3, 3) = "600"
    
    Me.MousePointer = vbHourglass
    With frmPSearch
        Set .oTextBox = txtMethodNatureCode
        .sBindSQL = "SELECT MethodNatureCode, MethodNatureDesc FROM MstMethodNature WHERE MethodNatureStatus = '1' Order By MethodNatureCode"
        .vHeadDataAry = vAry
        .lKeyPos = 1
        .lSearchPos = 3
        .sInputType = "N"
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    txtMethodNatureCode.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        'Case vbKeyPageDown
        '    KeyCode = 0
        '    If tabDetailInfo.Tab < tabDetailInfo.Tabs - 1 Then
        '        tabDetailInfo.Tab = tabDetailInfo.Tab + 1
        '        Exit Sub
        '    End If
        'Case vbKeyPageUp
        '    KeyCode = 0
        '    If tabDetailInfo.Tab > 0 Then
        '        tabDetailInfo.Tab = tabDetailInfo.Tab - 1
        '        Exit Sub
        '    End If
        
        Case vbKeyF6
            Call cmdOpen
        
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
            
        
        Case vbKeyF5
            If wiAction = DefaultPage Then Call cmdEdit
       
        
        Case vbKeyF3
            If wiAction = DefaultPage Then Call cmdDel
        
         Case vbKeyF9
        
            'If wiAction = DefaultPage Then Call cmdFind
            
        Case vbKeyF10
        
            If wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec Then Call cmdSave
            
        Case vbKeyF11
        
            If wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec Then Call cmdCancel
        
        Case vbKeyF12
        
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    Dim iCounter As Integer
    Dim iTabs As Integer
    Dim vToolTip As Variant
    
    MousePointer = vbHourglass
  
    wsFormCaption = Me.Caption
    
    IniForm
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
            Me.txtMethodNatureCode.Enabled = False
            Me.txtMethodNatureDesc.Enabled = False
            
            Me.btnMethodNatureCode.Enabled = False
            
        Case "AfrActAdd"
            Me.txtMethodNatureCode.Enabled = True
            Me.btnMethodNatureCode.Enabled = False
            
        Case "AfrActEdit"
            Me.txtMethodNatureCode.Enabled = True
            Me.btnMethodNatureCode.Enabled = True
            
        Case "AfrKey"
            Me.txtMethodNatureCode.Enabled = False
            Me.btnMethodNatureCode.Enabled = False
            
            Me.txtMethodNatureDesc.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    Dim sMsg As String
        
    InputValidation = False
    
    If Chk_txtMethodNatureCode = False Then
        Exit Function
    End If
    
    If Chk_txtMethodNatureDesc = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim Criteria As String
    Dim rsRcd As New ADODB.Recordset
    
    Criteria = "SELECT MstMethodNature.MethodNatureCode, MstMethodNature.* "
    Criteria = Criteria + "From MstMethodNature "
    Criteria = Criteria + "WHERE (((MstMethodNature.MethodNatureCode)='" + txtMethodNatureCode + "') "
    Criteria = Criteria + "AND ((MethodNatureStatus )= '1'));"

    rsRcd.Open Criteria, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        'wsKey = ""
    Else
        'wsKey = ReadRs(rsRcd, "AccTypeCode")
        
        Me.txtMethodNatureCode = ReadRs(rsRcd, "MethodNatureCode")
        Me.txtMethodNatureDesc = ReadRs(rsRcd, "MethodNatureDesc")
        Me.lblDspMethodNatureLastUpd = ReadRs(rsRcd, "MethodNatureLastUpd")
        Me.lblDspMethodNatureLastUpdDate = ReadRs(rsRcd, "MethodNatureLastUpdDate")
        
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
   ' Call UnLockAll(wsConnTime, wsFormID)
   ' Set waResult = Nothing
   ' Set waScrItm = Nothing
   ' Set waPgmItm = Nothing
    Set frmMN001 = Nothing
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
            
            If MsgBox("你是否確定要放棄現時之作業?", vbYesNo, sTitle) = vbYes Then
                Call cmdCancel
            End If
        
        Case tcFind
            
            Call OpenPromptForm
            
        Case tcExit
        
            Unload Me
            
    End Select
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
    Me.Left = 0
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "MN001"
    wsDocTyp = "MstMethodNature"
End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
            Case "ComboBox"
                MyControl.Clear
            Case "TextBox"
                MyControl.Text = ""
        '    Case "TDBGrid"
        '        MyControl.ClearFields
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
    wsKey = ""
    wiSalesmanID = 0
    
    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    
    Me.Caption = wsFormCaption
   
  '  wbFinChk = False
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
              
       Me.Caption = wsFormCaption + " - ADD"
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
       
    Case CorRec
           
        Me.Caption = wsFormCaption + " - EDIT"
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
       
    
    Case DelRec
    
        Me.Caption = wsFormCaption + " - DELETE"
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
    End Select
    
    txtMethodNatureCode.SetFocus
End Sub

Private Sub Ini_Scr_AfrKey()
    Dim Ctrl As Control
    Dim sMsg As String
    
    Select Case wiAction
    
    Case CorRec, DelRec

        If LoadRecord() = False Then
            sMsg = "存取檔案失敗! 請聯絡系統管理員或無限系統顧問!"
            MsgBox sMsg, vbOKOnly, sTitle
            Exit Sub
        Else
            If RowLock(wsConnTime, wsKeyType, cbmethodnatureCode, wsFormID, wsUsrId) = False Then
                wsMsg = "Record Lock By " & wsUsrId
                MsgBox wsMsg, vbOKOnly, sTitle
            Else
                    Call SetFieldStatus("AfrKey")
                    Call SetButtonStatus("AfrKey")
                    txtMethodNatureDesc.SetFocus
            End If
        End If
    End Select
End Sub

Private Function Chk_MethodNatureCode(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim sMsg As String
    Dim Criteria As String
    
    Chk_MethodNatureCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    Criteria = "SELECT MethodNatureCode "
    Criteria = Criteria & " FROM MstMethodNature WHERE MethodNatureCode = '" & Set_Quote(inCode) & "' AND MethodNatureStatus = '1'"
    
    rsRcd.Open Criteria, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Exit Function
    End If
    
    Chk_MethodNatureCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtMethodNatureCode() As Boolean
    Dim sMsg As String
    
    Chk_txtMethodNatureCode = False
    
    If wiAction = AddRec Then
        If Trim(txtMethodNatureCode.Text) = "" And Chk_AutoGen("") = "N" Then
            sMsg = "沒有輸入須要之資料!"
            MsgBox sMsg, vbInformation + vbOKOnly, sTitle
            txtMethodNatureCode.SetFocus
            Exit Function
        End If
    
        If Chk_MethodNatureCode(txtMethodNatureCode.Text) = True Then
            sMsg = "MethodNature 編碼已存在!"
            MsgBox sMsg, vbInformation + vbOKOnly, sTitle
            txtMethodNatureCode.SetFocus
            Exit Function
        End If
    Else
        If Chk_MethodNatureCode(txtMethodNatureCode.Text) = False Then
            sMsg = "MethodNature 編碼不存在!"
            MsgBox sMsg, vbInformation + vbOKOnly, sTitle
            txtMethodNatureCode.SetFocus
            Exit Function
        End If
    End If
    Chk_txtMethodNatureCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmMN001
    
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
                Call Ini_Scr
                Call cmdEdit
                
            Case DelRec
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
    Dim wsMsg As String
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    
  '  If wiAction <> AddRec Then
  '      If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Then
  '          Call Dsp_Err("R0003", "", "E", Me.Caption)
  '          MousePointer = vbDefault
  '          Exit Function
  '      End If
  '  End If
   
    
    If wiAction = DelRec Then
        If MsgBox("你是否確定要刪除此檔案?", vbYesNo, sTitle) = vbNo Then
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
        
    adcmdSave.CommandText = "USP_MN001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    'Call SetSPPara(adcmdSave, 2, wsKey)
    Call SetSPPara(adcmdSave, 2, txtMethodNatureCode)
    Call SetSPPara(adcmdSave, 3, txtMethodNatureDesc)
    Call SetSPPara(adcmdSave, 4, sUserName)
    Call SetSPPara(adcmdSave, 5, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 6)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        wsMsg = "儲存失敗, 請檢查 Store Procedure - MN001!"
        MsgBox wsMsg, vbInformation + vbOKOnly, sTitle
    Else
        wsMsg = "已成功儲存!"
        MsgBox wsMsg, vbInformation + vbOKOnly, sTitle
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

Private Function SaveData() As Boolean

    Dim wiRet As Long
    
    SaveData = False
    
    If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) Then
       If MsgBox("你是否確定不儲存現時之變更而離開?", vbYesNo, sTitle) = vbYes Then
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
End Sub

Private Sub txtMethodNatureCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtMethodNatureCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtMethodNatureCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtMethodNatureCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            KeyCode = 0
            Call btnMethodNatureCode_Click
    End Select
End Sub

Private Sub txtMethodNatureDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtMethodNatureDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Chk_txtMethodNatureDesc
    End If
End Sub

Private Sub txtMethodNatureCode_GotFocus()
    Call SelObj(txtMethodNatureCode)
End Sub

Private Sub txtMethodNatureDesc_GotFocus()
    Call SelObj(txtMethodNatureDesc)
End Sub

Private Function Chk_txtMethodNatureDesc() As Boolean
    Dim sMsg As String
    
    Chk_txtMethodNatureDesc = False
    
    If Trim(txtMethodNatureDesc.Text) = "" Then
        sMsg = "沒有輸入須要之資料!"
        MsgBox sMsg, vbInformation + vbOKOnly, sTitle
        txtMethodNatureDesc.SetFocus
        Exit Function
    End If
    
    Chk_txtMethodNatureDesc = True
End Function

