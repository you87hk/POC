VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSYS001 
   BackColor       =   &H8000000A&
   Caption         =   "系統設定"
   ClientHeight    =   4170
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "SYS001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "SYS001.frx":08CA
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "系統設定"
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   8355
      Begin VB.Frame FraSyPStkVal 
         Caption         =   "SYPSTKVAL"
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   7455
         Begin VB.OptionButton optSyPStkVal 
            Caption         =   "AVERAGECOST"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optSyPStkVal 
            Caption         =   "LASTPOCOST"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optSyPStkVal 
            Caption         =   "BYLOT"
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.TextBox txtSyPRecID 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtSyPHisLog 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   2730
      End
      Begin MSMask.MaskEdBox medSyPMinDte 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSyPMaxDte 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblDspSyPLUpDte 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5640
         TabIndex        =   18
         Top             =   3240
         Width           =   2265
      End
      Begin VB.Label lblDspSyPLUpUsr 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Top             =   3240
         Width           =   2265
      End
      Begin VB.Label lblSyPLUpDte 
         Caption         =   "SYPLUPDTE"
         Height          =   240
         Left            =   4200
         TabIndex        =   16
         Top             =   3285
         Width           =   1380
      End
      Begin VB.Label lblSyPLUpUsr 
         Caption         =   "SYPLUPUSR"
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   3285
         Width           =   1380
      End
      Begin VB.Label lblSyPMaxDte 
         Caption         =   "SYPMAXDTE"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   2820
         Width           =   1380
      End
      Begin VB.Label lblSyPMinDte 
         Caption         =   "SYPMINDTE"
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   2460
         Width           =   1380
      End
      Begin VB.Label lblSyPRecID 
         Caption         =   "SYPRECID"
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
      Begin VB.Label lblSyPHisLog 
         Caption         =   "SYPHISLOG"
         Height          =   240
         Left            =   360
         TabIndex        =   9
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
            Picture         =   "SYS001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   8
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
            Object.Visible         =   0   'False
            Key             =   "Open"
            Object.ToolTipText     =   "開新視窗 (F6)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Add"
            Object.ToolTipText     =   "新增 (F2)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Edit"
            Object.ToolTipText     =   "修改 (F5)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "刪除 (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
Attribute VB_Name = "frmSYS001"
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
        Me.Height = 4575
        Me.Width = 8700
    End If
End Sub

'-- Set toolbar buttons status in different mode, Default, AddEdit, None.
Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            With tbrProcess
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
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
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcExit).Enabled = True
            End With
        
        Case "AfrKey"
            With tbrProcess
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
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
            Me.txtSyPHisLog.Enabled = True
            Me.medSyPMinDte.Enabled = True
            Me.medSyPMaxDte.Enabled = True
            Me.optSyPStkVal(0).Enabled = True
            Me.optSyPStkVal(1).Enabled = True
            Me.optSyPStkVal(2).Enabled = True
            
            Me.txtSyPRecID.Enabled = False
            
        Case "AfrActAdd"
            
        Case "AfrActEdit"
            Me.txtSyPRecID.Enabled = False
            
            Me.txtSyPHisLog.Enabled = True
            Me.medSyPMinDte.Enabled = True
            Me.medSyPMaxDte.Enabled = True
            Me.optSyPStkVal(0).Enabled = True
            Me.optSyPStkVal(1).Enabled = True
            Me.optSyPStkVal(2).Enabled = True
            
        Case "AfrKey"
            Me.txtSyPRecID.Enabled = False
            
            Me.txtSyPHisLog.Enabled = True
            Me.medSyPMinDte.Enabled = True
            Me.medSyPMaxDte.Enabled = True
            Me.optSyPStkVal(0).Enabled = True
            Me.optSyPStkVal(1).Enabled = True
            Me.optSyPStkVal(2).Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_DateRange() = False Then
        Exit Function
    End If
    
    If Chk_txtSyPHisLog() = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From SysConst "
    wsSQL = wsSQL + "WHERE (((SysConst.SyPRecID)='01'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        Me.txtSyPRecID = ReadRs(rsRcd, "SyPRecID")
        Me.txtSyPHisLog = ReadRs(rsRcd, "SyPHisLog")
        Me.medSyPMinDte = Format(ReadRs(rsRcd, "SyPMinDte"), "YYYY/MM/DD")
        Me.medSyPMaxDte = Format(ReadRs(rsRcd, "SyPMaxDte"), "YYYY/MM/DD")
        Me.lblDspSyPLUpUsr = ReadRs(rsRcd, "SyPLUpUsr")
        Me.lblDspSyPLUpDte = ReadRs(rsRcd, "SyPLUpDte")
        
        SetStkVal ReadRs(rsRcd, "SyPStkVal")
        
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
    Set frmSYS001 = Nothing
End Sub

Private Sub optSyPStkVal_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        medSyPMinDte.SetFocus
    End If
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
    wsFormID = "SYS001"
    wsTrnCd = ""
    
End Sub


Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblSyPRecID.Caption = Get_Caption(waScrItm, "SYPRECID")
    FraSyPStkVal.Caption = Get_Caption(waScrItm, "SYPSTKVAL")
    lblSyPHisLog.Caption = Get_Caption(waScrItm, "SYPHISLOG")
    lblSyPMinDte.Caption = Get_Caption(waScrItm, "SYPMINDTE")
    lblSyPMaxDte.Caption = Get_Caption(waScrItm, "SYPMAXDTE")
    
    lblSyPLUpUsr.Caption = Get_Caption(waScrItm, "SYPLUPUSR")
    lblSyPLUpDte.Caption = Get_Caption(waScrItm, "SYPLUPDTE")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
    
    optSyPStkVal(0).Caption = Get_Caption(waScrItm, "SYPSTKVAL01")
    optSyPStkVal(1).Caption = Get_Caption(waScrItm, "SYPSTKVAL02")
    optSyPStkVal(2).Caption = Get_Caption(waScrItm, "SYPSTKVAL03")

    'tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    'tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    'tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    'tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    'tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
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
    
    Call SetDateMask(medSyPMinDte)
    Call SetDateMask(medSyPMaxDte)

    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    tblCommon.Visible = False
    Me.Caption = wsFormCaption
    
    If LoadRecord() = False Then
        gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Sub
    Else
        If RowLock(wsConnTime, wsKeyType, txtSyPRecID, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
        End If
    End If
    wiAction = CorRec
    Ini_Scr_AfrAct
    'Call SetFieldStatus("AfrKey")
    'Call SetButtonStatus("AfrKey")
    'txtSyPHisLog.SetFocus
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
        Case AddRec
                  
        Case CorRec
               
            Call SetFieldStatus("AfrActEdit")
            Call SetButtonStatus("AfrActEdit")
            'txtSyPHisLog.SetFocus
        
        Case DelRec
    
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub



Private Sub cmdOpen()
    Dim newForm As New frmSYS001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, txtSyPRecID, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_SYS001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, txtSyPRecID)
    Call SetSPPara(adcmdSave, 3, GetStkVal())
    Call SetSPPara(adcmdSave, 4, txtSyPHisLog)
    Call SetSPPara(adcmdSave, 5, medSyPMinDte)
    Call SetSPPara(adcmdSave, 6, medSyPMaxDte)
    Call SetSPPara(adcmdSave, 7, gsUserID)
    Call SetSPPara(adcmdSave, 8, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 9)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - SYS001!"
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



Private Function Chk_txtSyPHisLog() As Boolean
    Chk_txtSyPHisLog = False
    
    If Trim(txtSyPHisLog.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtSyPHisLog.SetFocus
        Exit Function
    End If
    
    If Not (txtSyPHisLog >= 0 And txtSyPHisLog <= 12) Then
        gsMsg = "長度必須少於十二!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtSyPHisLog.SetFocus
        Exit Function
    End If
    
    Chk_txtSyPHisLog = True
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





Private Sub SetStkVal(ByVal inCode As String)
    Select Case inCode
        Case "1"
            optSyPStkVal(0).Value = True
            
        Case "3"
            optSyPStkVal(1).Value = True
            
        Case "4"
            optSyPStkVal(2).Value = True
    End Select
End Sub

Private Function GetStkVal() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 2
        If optSyPStkVal(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetStkVal = "1"
            
        Case 1
            GetStkVal = "3"
        
        Case 2
            GetStkVal = "4"
        
    End Select
End Function

Private Sub txtSyPHisLog_GotFocus()
    FocusMe txtSyPHisLog
End Sub

Private Sub txtSyPHisLog_KeyPress(KeyAscii As Integer)
    Dim iCounter As Integer

    Call Chk_InpNum(KeyAscii, txtSyPHisLog, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtSyPHisLog() = True Then
            For iCounter = 0 To 5
                If optSyPStkVal(iCounter).Value = True Then
                    Exit For
                End If
            Next
    
            optSyPStkVal(iCounter).Value = True
            optSyPStkVal(iCounter).SetFocus
        End If
    End If
End Sub

Private Sub txtSyPHisLog_LostFocus()
    FocusMe txtSyPHisLog, True
End Sub

Private Sub medSyPMinDte_GotFocus()
    FocusMe medSyPMinDte
End Sub

Private Sub medSyPMinDte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medSyPMinDte Then
            medSyPMaxDte.SetFocus
        End If
    End If
End Sub

Private Sub medSyPMinDte_LostFocus()
    FocusMe medSyPMinDte, True
End Sub

Private Function Chk_medSyPMinDte() As Boolean
    Chk_medSyPMinDte = False
    
    If Trim(medSyPMinDte.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medSyPMinDte.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medSyPMinDte) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medSyPMinDte.SetFocus
        Exit Function
    End If
    
    Chk_medSyPMinDte = True
End Function

Private Sub medSyPMaxDte_GotFocus()
    FocusMe medSyPMaxDte
End Sub

Private Sub medSyPMaxDte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medSyPMaxDte Then
            txtSyPHisLog.SetFocus
        End If
    End If
End Sub

Private Sub medSyPMaxDte_LostFocus()
    FocusMe medSyPMaxDte, True
End Sub

Private Function Chk_medSyPMaxDte() As Boolean
    Chk_medSyPMaxDte = False
    
    If Trim(medSyPMaxDte.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medSyPMaxDte.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medSyPMaxDte) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medSyPMaxDte.SetFocus
        Exit Function
    End If
    
    Chk_medSyPMaxDte = True
End Function

Private Function Chk_DateRange() As Boolean
    Chk_DateRange = False

    If medSyPMinDte > medSyPMaxDte Then
        gsMsg = "最小日期大於最大日期!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medSyPMinDte.SetFocus
        Exit Function
    End If
    
    Chk_DateRange = True
End Function
