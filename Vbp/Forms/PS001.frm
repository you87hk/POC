VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmPS001 
   BackColor       =   &H8000000A&
   Caption         =   "列印大小"
   ClientHeight    =   3450
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "PS001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "PS001.frx":08CA
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboPrintSizeCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "列印大小"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtPrintSizeCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Tag             =   "K"
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtPrintSizeDesc 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label lblPrintSizeLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   360
         TabIndex        =   8
         Top             =   2445
         Width           =   1140
      End
      Begin VB.Label lblPrintSizeLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4320
         TabIndex        =   7
         Top             =   2445
         Width           =   1260
      End
      Begin VB.Label lblDspPrintSizeLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Top             =   2400
         Width           =   2505
      End
      Begin VB.Label lblDspPrintSizeLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5640
         TabIndex        =   5
         Top             =   2400
         Width           =   2505
      End
      Begin VB.Label lblPrintSizeCode 
         Caption         =   "大小編碼 :"
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
         TabIndex        =   4
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblPrintSizeDesc 
         Caption         =   "大小註解 :"
         Height          =   240
         Left            =   360
         TabIndex        =   3
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
            Picture         =   "PS001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PS001.frx":6945
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
Attribute VB_Name = "frmPS001"
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

Private wsActNam(4) As String

Private wiAction As Integer
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control

Private Const wsKeyType = "MstPrintSize"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboPrintSizeCode_LostFocus()
    FocusMe cboPrintSizeCode, True
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
            Me.txtPrintSizeDesc.Enabled = False
            
            Me.cboPrintSizeCode.Enabled = False
            Me.cboPrintSizeCode.Visible = False
            Me.txtPrintSizeCode.Visible = True
            Me.txtPrintSizeCode.Enabled = False
            
        Case "AfrActAdd"
            Me.cboPrintSizeCode.Enabled = False
            Me.cboPrintSizeCode.Visible = False
            
            Me.txtPrintSizeCode.Enabled = True
            Me.txtPrintSizeCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboPrintSizeCode.Enabled = True
            Me.cboPrintSizeCode.Visible = True
            
            Me.txtPrintSizeCode.Enabled = False
            Me.txtPrintSizeCode.Visible = False
            
        Case "AfrKey"
            Me.txtPrintSizeDesc.Enabled = True
            
            Me.cboPrintSizeCode.Enabled = False
            Me.txtPrintSizeCode.Enabled = False
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtPrintSizeDesc = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstPrintSize.* "
    wsSql = wsSql + "From MstPrintSize "
    wsSql = wsSql + "WHERE (((MstPrintSize.PrintSizeCode)='" + Set_Quote(cboPrintSizeCode) + "') "
    wsSql = wsSql + "AND ((MstPrintSize.PrintSizeStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        
        Me.cboPrintSizeCode = ReadRs(rsRcd, "PrintSizeCode")
        Me.txtPrintSizeDesc = ReadRs(rsRcd, "PrintSizeDesc")
        Me.lblDspPrintSizeLastUpd = ReadRs(rsRcd, "PrintSizeLastUpd")
        Me.lblDspPrintSizeLastUpdDate = ReadRs(rsRcd, "PrintSizeLastUpdDate")
        
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
    Set frmPS001 = Nothing
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
    wsFormID = "PS001"
    wsTrnCd = ""
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
        txtPrintSizeCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboPrintSizeCode.SetFocus
       
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboPrintSizeCode.SetFocus
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Dim Ctrl As Control
    
    Select Case wiAction
    
    Case CorRec, DelRec

        If LoadRecord() = False Then
            gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Sub
        Else
            If RowLock(wsConnTime, wsKeyType, cboPrintSizeCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtPrintSizeDesc.SetFocus
End Sub

Private Function Chk_PrintSizeCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_PrintSizeCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT PrintSizeStatus "
    wsSql = wsSql & " FROM MstPrintSize WHERE PrintSizeCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "PrintSizeStatus")
    
    Chk_PrintSizeCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboPrintSizeCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboPrintSizeCode = False
    
    If Trim(cboPrintSizeCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboPrintSizeCode.SetFocus
        Exit Function
    End If
    
    If Chk_PrintSizeCode(cboPrintSizeCode.Text, wsStatus) = False Then
        gsMsg = "列印大小編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboPrintSizeCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "列印大小分類已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboPrintSizeCode.SetFocus
            Exit Function
        End If
    End If
    Chk_cboPrintSizeCode = True
End Function

Private Function Chk_txtPrintSizeCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtPrintSizeCode = False
    
    If Trim(txtPrintSizeCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPrintSizeCode.SetFocus
        Exit Function
    End If
    
    If Chk_PrintSizeCode(txtPrintSizeCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "列印大小編碼已存在但已無效!"
        Else
            gsMsg = "列印大小編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPrintSizeCode.SetFocus
        Exit Function
    End If
    
    Chk_txtPrintSizeCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmPS001
    
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

Private Sub cmdFind()
     Call OpenPromptForm
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboPrintSizeCode, wsFormID) Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If wiAction = DelRec Then
        gsMsg = "你是否確定要刪除此記錄?"
        If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
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
        
    adcmdSave.CommandText = "USP_PS001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtPrintSizeCode, cboPrintSizeCode))
    Call SetSPPara(adcmdSave, 3, txtPrintSizeDesc)
    Call SetSPPara(adcmdSave, 4, gsUserID)
    Call SetSPPara(adcmdSave, 5, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 6)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - PS001!"
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

Private Function SaveData() As Boolean

    Dim wiRet As Long
    
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
    Dim sSQL As String
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "開度編碼"
    vFilterAry(1, 2) = "PrintSizeCode"
    
    vFilterAry(2, 1) = "註解"
    vFilterAry(2, 2) = "PrintSizeDesc"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "開度編碼"
    vAry(1, 2) = "PrintSizeCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "註解"
    vAry(2, 2) = "PrintSizeDesc"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstPrintSize.PrintSizeCode, MstPrintSize.PrintSizeDesc "
        sSQL = sSQL + "FROM MstPrintSize "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE MstPrintSize.PrintSizeStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstPrintSize.PrintSizeCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboPrintSizeCode Then
        cboPrintSizeCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtPrintSizeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtPrintSizeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPrintSizeCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtPrintSizeCode_LostFocus()
    FocusMe txtPrintSizeCode, True
End Sub

Private Sub txtPrintSizeDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtPrintSizeDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Chk_txtPrintSizeDesc
    End If
End Sub

Private Sub txtPrintSizeCode_GotFocus()
    FocusMe txtPrintSizeCode
End Sub

Private Sub txtPrintSizeDesc_GotFocus()
    FocusMe txtPrintSizeDesc
End Sub

Private Function Chk_txtPrintSizeDesc() As Boolean
    
    Chk_txtPrintSizeDesc = False
    
    If Trim(txtPrintSizeDesc.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPrintSizeDesc.SetFocus
        Exit Function
    End If
    
    Chk_txtPrintSizeDesc = True
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

Private Sub cboPrintSizeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboPrintSizeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPrintSizeCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboPrintSizeCode_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboPrintSizeCode
    
    wsSql = "SELECT PrintSizeCode, PrintSizeDesc FROM MstPrintSize WHERE PrintSizeStatus = '1'"
    wsSql = wsSql & " AND PrintSizeCode LIKE '%" & IIf(cboPrintSizeCode.SelLength > 0, "", Set_Quote(cboPrintSizeCode.Text)) & "%' "
   
    wsSql = wsSql & "ORDER BY PrintSizeCode "
    Call Ini_Combo(2, wsSql, cboPrintSizeCode.Left, cboPrintSizeCode.Top + cboPrintSizeCode.Height, tblCommon, "PS001", "TBLPS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboPrintSizeCode_GotFocus()
    FocusMe cboPrintSizeCode
End Sub


Private Sub txtPrintSizeDesc_LostFocus()
    FocusMe txtPrintSizeDesc, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT PrintSizeStatus FROM MstPrintSize WHERE PrintSizeCode = '" & Set_Quote(txtPrintSizeCode) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
        .TableKey = "PrintSizeCode"
        .KeyLen = 10
        Set .ctlKey = txtPrintSizeCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblPrintSizeCode.Caption = Get_Caption(waScrItm, "PRINTSIZECODE")
    lblPrintSizeDesc.Caption = Get_Caption(waScrItm, "PRINTSIZEDESC")
    lblPrintSizeLastUpd.Caption = Get_Caption(waScrItm, "PRINTSIZELASTUPD")
    lblPrintSizeLastUpdDate.Caption = Get_Caption(waScrItm, "PRINTSIZELASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "PSADD")
    wsActNam(2) = Get_Caption(waScrItm, "PSEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PSDELETE")

    
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

