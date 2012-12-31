VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmLOC001 
   BackColor       =   &H8000000A&
   Caption         =   "LOCATION"
   ClientHeight    =   6600
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   9945
   Icon            =   "LOC001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9945
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "LOC001.frx":08CA
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboLocTerrCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   10
      Top             =   3960
      Width           =   2370
   End
   Begin VB.ComboBox cboLocCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "LOCATION"
      Height          =   6135
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   9675
      Begin VB.TextBox txtLocRemark 
         Enabled         =   0   'False
         Height          =   1020
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   4320
         Width           =   7305
      End
      Begin VB.TextBox txtLocEmail 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   11
         Top             =   3960
         Width           =   7280
      End
      Begin VB.TextBox txtLocFax 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   9
         Top             =   3240
         Width           =   2835
      End
      Begin VB.TextBox txtLocTel 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   8
         Top             =   3240
         Width           =   2835
      End
      Begin VB.PictureBox picLocAdr 
         BackColor       =   &H80000009&
         Height          =   1455
         Left            =   1680
         ScaleHeight     =   1395
         ScaleWidth      =   7215
         TabIndex        =   23
         Top             =   1680
         Width           =   7275
         Begin VB.TextBox txtLocAddress 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   4
            Left            =   0
            TabIndex        =   7
            Top             =   1035
            Width           =   7100
         End
         Begin VB.TextBox txtLocAddress 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   3
            Left            =   0
            TabIndex        =   6
            Top             =   690
            Width           =   7100
         End
         Begin VB.TextBox txtLocAddress 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   5
            Top             =   345
            Width           =   7100
         End
         Begin VB.TextBox txtLocAddress 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   7100
         End
      End
      Begin VB.TextBox txtLocContactPerson 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   7280
      End
      Begin VB.TextBox txtLocCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Tag             =   "K"
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtLocName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   7280
      End
      Begin VB.Label lblLocRemark 
         Caption         =   "LOCREMARK"
         Height          =   240
         Left            =   360
         TabIndex        =   30
         Top             =   4350
         Width           =   900
      End
      Begin VB.Label lblDspLocTerritory 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   4080
         TabIndex        =   29
         Top             =   3600
         Width           =   4890
      End
      Begin VB.Label lblLocTerrCode 
         Caption         =   "LOCTERRCODE"
         Height          =   240
         Left            =   360
         TabIndex        =   28
         Top             =   3660
         Width           =   1260
      End
      Begin VB.Label lblLocEmail 
         Caption         =   "LOCEMAIL"
         Height          =   240
         Left            =   360
         TabIndex        =   27
         Top             =   4020
         Width           =   1140
      End
      Begin VB.Label lblLocFax 
         Caption         =   "LOCFAX"
         Height          =   240
         Left            =   4800
         TabIndex        =   26
         Top             =   3300
         Width           =   1140
      End
      Begin VB.Label lblLocTel 
         Caption         =   "LOCTEL"
         Height          =   240
         Left            =   360
         TabIndex        =   25
         Top             =   3300
         Width           =   1260
      End
      Begin VB.Label lblLocAddress 
         Caption         =   "LOCADDRESS"
         Height          =   240
         Left            =   360
         TabIndex        =   24
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label lblLocContactPerson 
         Caption         =   "LOCCONTACTPERSON"
         Height          =   240
         Left            =   360
         TabIndex        =   22
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label lblLocLastUpd 
         Caption         =   "LOCLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   19
         Top             =   5685
         Width           =   1140
      End
      Begin VB.Label lblLocLastUpdDate 
         Caption         =   "LOCLASTUPDDATE"
         Height          =   240
         Left            =   4920
         TabIndex        =   18
         Top             =   5685
         Width           =   1260
      End
      Begin VB.Label lblDspLocLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1680
         TabIndex        =   17
         Top             =   5640
         Width           =   2505
      End
      Begin VB.Label lblDspLocLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6480
         TabIndex        =   16
         Top             =   5640
         Width           =   2505
      End
      Begin VB.Label lblLocCode 
         Caption         =   "LOCCODE"
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
         TabIndex        =   15
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblLocName 
         Caption         =   "LOCNAME"
         Height          =   240
         Left            =   360
         TabIndex        =   14
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
            Picture         =   "LOC001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LOC001.frx":6945
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
      Width           =   9945
      _ExtentX        =   17542
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
Attribute VB_Name = "frmLOC001"
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
Private wlKey As Long

Private Const wsKeyType = "MstLocation"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboLocCode_LostFocus()
    FocusMe cboLocCode, True
End Sub

Private Sub cboLocTerrCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboLocTerrCode
    
    wsSql = "SELECT TerrCode, TerrDesc FROM MstTerritory WHERE TerrStatus = '1'"
    wsSql = wsSql & "ORDER BY TerrCode "
    Call Ini_Combo(2, wsSql, cboLocTerrCode.Left, cboLocTerrCode.Top + cboLocTerrCode.Height, tblCommon, "LOC001", "TBLLOCTERR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboLocTerrCode_GotFocus()
    FocusMe cboLocTerrCode
End Sub

Private Sub cboLocTerrCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboLocTerrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Me.lblDspLocTerritory = LoadDescByCode("MstTerritory", "TerrCode", "TerrDesc", cboLocTerrCode, True)
        
        If Chk_cboLocTerrCode() = True Then
            txtLocEmail.SetFocus
        End If
    End If
End Sub

Private Sub cboLocTerrCode_LostFocus()
    FocusMe cboLocTerrCode, True
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
        Me.Height = 7005
        Me.Width = 10065
    End If
End Sub

'-- Set toolbar buttons status in different mode, Default, AddEdit, None.
Public Sub SetButtonStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
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
Public Sub SetFieldStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
        Case "Default"
            Me.txtLocName.Enabled = False
            Me.txtLocContactPerson.Enabled = False
            Me.picLocAdr.Enabled = False
            Me.txtLocTel.Enabled = False
            Me.txtLocFax.Enabled = False
            Me.txtLocEmail.Enabled = False
            Me.cboLocTerrCode.Enabled = False
            Me.txtLocRemark.Enabled = False
            
            Me.cboLocCode.Enabled = False
            Me.cboLocCode.Visible = False
            Me.txtLocCode.Visible = True
            Me.txtLocCode.Enabled = False
            
        Case "AfrActAdd"
            Me.cboLocCode.Enabled = False
            Me.cboLocCode.Visible = False
            
            Me.txtLocCode.Enabled = True
            Me.txtLocCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboLocCode.Enabled = True
            Me.cboLocCode.Visible = True
            
            Me.txtLocCode.Enabled = False
            Me.txtLocCode.Visible = False
            
        Case "AfrKey"
            Me.txtLocName.Enabled = True
            Me.txtLocContactPerson.Enabled = True
            Me.picLocAdr.Enabled = True
            Me.txtLocTel.Enabled = True
            Me.txtLocFax.Enabled = True
            Me.txtLocEmail.Enabled = True
            Me.cboLocTerrCode.Enabled = True
            Me.txtLocRemark.Enabled = True
            
            Me.cboLocCode.Enabled = False
            Me.txtLocCode.Enabled = False
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtLocName = False Then
        Exit Function
    End If
    
    If Chk_cboLocTerrCode = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstLocation.* "
    wsSql = wsSql + "From MstLocation "
    wsSql = wsSql + "WHERE (((MstLocation.LocCode)='" + Set_Quote(cboLocCode) + "') "
    wsSql = wsSql + "AND ((LocStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        wlKey = ReadRs(rsRcd, "LocID")
        Me.cboLocCode = ReadRs(rsRcd, "LocCode")
        Me.txtLocName = ReadRs(rsRcd, "LocName")
        Me.txtLocContactPerson = ReadRs(rsRcd, "LocContactPerson")
        Me.txtLocAddress(1) = ReadRs(rsRcd, "LocAddress1")
        Me.txtLocAddress(2) = ReadRs(rsRcd, "LocAddress2")
        Me.txtLocAddress(3) = ReadRs(rsRcd, "LocAddress3")
        Me.txtLocAddress(4) = ReadRs(rsRcd, "LocAddress4")
        Me.txtLocTel = ReadRs(rsRcd, "LocTel")
        Me.txtLocFax = ReadRs(rsRcd, "LocFax")
        Me.txtLocEmail = ReadRs(rsRcd, "LocEmail")
        Me.cboLocTerrCode = ReadRs(rsRcd, "LocTerrCode")
        Me.txtLocRemark = ReadRs(rsRcd, "LocRemark")
        
        Me.lblDspLocTerritory = LoadDescByCode("MstTerritory", "TerrCode", "TerrDesc", cboLocTerrCode, True)
        
        Me.lblDspLocLastUpd = ReadRs(rsRcd, "LocLastUpd")
        Me.lblDspLocLastUpdDate = ReadRs(rsRcd, "LocLastUpdDate")
        
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
    Set frmLOC001 = Nothing
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
    wsFormID = "LOC001"
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
        txtLocCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboLocCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboLocCode.SetFocus
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
            If RowLock(wsConnTime, wsKeyType, cboLocCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtLocName.SetFocus
End Sub

Private Function Chk_LocCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_LocCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT LocStatus "
    wsSql = wsSql & " FROM MstLocation WHERE LocCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "LocStatus")
    
    Chk_LocCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboLocCode() As Boolean
    Dim wsStatus As String

    Chk_cboLocCode = False
    
    If Trim(cboLocCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboLocCode.SetFocus
        Exit Function
    End If

    If Chk_LocCode(cboLocCode.Text, wsStatus) = False Then
        gsMsg = "書展/寄售編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboLocCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "書展/寄售已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboLocCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboLocCode = True
End Function

Private Function Chk_txtLocCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtLocCode = False
    
    If Trim(txtLocCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtLocCode.SetFocus
        Exit Function
    End If
    
    If Chk_LocCode(txtLocCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "書展/寄售編碼已存在但已無效!"
        Else
            gsMsg = "書展/寄售已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtLocCode.SetFocus
        Exit Function
    End If
    
    Chk_txtLocCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmLOC001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboLocCode, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_LOC001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, txtLocCode, cboLocCode))
    Call SetSPPara(adcmdSave, 4, txtLocName)
    Call SetSPPara(adcmdSave, 5, txtLocContactPerson)
    Call SetSPPara(adcmdSave, 6, txtLocAddress(1))
    Call SetSPPara(adcmdSave, 7, txtLocAddress(2))
    Call SetSPPara(adcmdSave, 8, txtLocAddress(3))
    Call SetSPPara(adcmdSave, 9, txtLocAddress(4))
    Call SetSPPara(adcmdSave, 10, txtLocTel)
    Call SetSPPara(adcmdSave, 11, txtLocFax)
    Call SetSPPara(adcmdSave, 12, txtLocEmail)
    Call SetSPPara(adcmdSave, 13, cboLocTerrCode)
    Call SetSPPara(adcmdSave, 14, txtLocRemark)
    Call SetSPPara(adcmdSave, 15, gsUserID)
    Call SetSPPara(adcmdSave, 16, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 17)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - LOC001!"
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
    vFilterAry(1, 1) = "書展/寄售編碼"
    vFilterAry(1, 2) = "LocCode"
    
    vFilterAry(2, 1) = "名稱"
    vFilterAry(2, 2) = "LocName"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "書展/寄售編碼"
    vAry(1, 2) = "LocCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "名稱"
    vAry(2, 2) = "LocName"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstLocation.LocCode, MstLocation.LocName "
        sSQL = sSQL + "FROM MstLocation "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE MstLocation.LocStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstLocation.LocCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboLocCode Then
        cboLocCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
End Sub

Private Sub txtLocAddress_GotFocus(Index As Integer)
    FocusMe txtLocAddress(Index)
End Sub

Private Sub txtLocAddress_KeyPress(Index As Integer, KeyAscii As Integer)
    Call chk_InpLen(txtLocAddress(Index), 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Index < 4 Then
            txtLocAddress(Index + 1).SetFocus
        Else
            txtLocTel.SetFocus
        End If
    End If
End Sub

Private Sub txtLocAddress_LostFocus(Index As Integer)
    FocusMe txtLocAddress(Index), True
End Sub

Private Sub txtLocCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtLocCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtLocCode_LostFocus()
    FocusMe txtLocCode, True
End Sub

Private Sub txtLocContactPerson_GotFocus()
    FocusMe txtLocContactPerson
End Sub

Private Sub txtLocContactPerson_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocContactPerson, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtLocAddress(1).SetFocus
    End If
End Sub

Private Sub txtLocContactPerson_LostFocus()
    FocusMe txtLocContactPerson, True
End Sub

Private Sub txtLocEmail_GotFocus()
    FocusMe txtLocEmail
End Sub

Private Sub txtLocEmail_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocEmail, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtLocRemark.SetFocus
    End If
End Sub

Private Sub txtLocEmail_LostFocus()
    FocusMe txtLocEmail, True
End Sub

Private Sub txtLocName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtLocName Then
            txtLocContactPerson.SetFocus
        End If
    End If
End Sub

Private Sub txtLocCode_GotFocus()
    FocusMe txtLocCode
End Sub

Private Sub txtLocName_GotFocus()
    FocusMe txtLocName
End Sub

Private Function Chk_txtLocName() As Boolean
    Chk_txtLocName = False
    
    If Trim(txtLocName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtLocName.SetFocus
        Exit Function
    End If
    
    Chk_txtLocName = True
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

Private Sub cboLocCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboLocCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboLocCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboLocCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboLocCode
    
    wsSql = "SELECT LocCode, LocName FROM MstLocation WHERE LocStatus = '1'"
    wsSql = wsSql & " AND LocCode LIKE '%" & IIf(cboLocCode.SelLength > 0, "", Set_Quote(cboLocCode.Text)) & "%' "
    wsSql = wsSql & "ORDER BY LocCode "
    Call Ini_Combo(2, wsSql, cboLocCode.Left, cboLocCode.Top + cboLocCode.Height, tblCommon, "LOC001", "TBLLOC", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboLocCode_GotFocus()
    FocusMe cboLocCode
End Sub

Private Sub txtLocName_LostFocus()
    FocusMe txtLocName, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT LocStatus FROM MstLocation WHERE LocCode = '" & Set_Quote(txtLocCode) & "'"
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
        .TableKey = "LocCode"
        .KeyLen = 10
        Set .ctlKey = txtLocCode
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
    
    lblLocCode.Caption = Get_Caption(waScrItm, "LOCCODE")
    lblLocName.Caption = Get_Caption(waScrItm, "LOCNAME")
    lblLocContactPerson.Caption = Get_Caption(waScrItm, "LOCCONTACTPERSON")
    lblLocAddress.Caption = Get_Caption(waScrItm, "LOCADDRESS")
    lblLocTel.Caption = Get_Caption(waScrItm, "LOCTEL")
    lblLocFax.Caption = Get_Caption(waScrItm, "LOCFAX")
    lblLocTerrCode.Caption = Get_Caption(waScrItm, "LOCTERRCODE")
    lblLocEmail.Caption = Get_Caption(waScrItm, "LOCEMAIL")
    lblLocRemark.Caption = Get_Caption(waScrItm, "LOCREMARK")
    lblLocLastUpd.Caption = Get_Caption(waScrItm, "LOCLASTUPD")
    lblLocLastUpdDate.Caption = Get_Caption(waScrItm, "LOCLASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "LOCADD")
    wsActNam(2) = Get_Caption(waScrItm, "LOCEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "LOCDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub txtLocRemark_GotFocus()
    FocusMe txtLocRemark
End Sub

Private Sub txtLocRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocRemark, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtLocName.SetFocus
    End If
End Sub

Private Sub txtLocRemark_LostFocus()
    FocusMe txtLocRemark, True
End Sub

Private Sub txtLocTel_GotFocus()
    FocusMe txtLocTel
End Sub

Private Sub txtLocTel_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocTel, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtLocFax.SetFocus
    End If
End Sub

Private Sub txtLocTel_LostFocus()
    FocusMe txtLocTel, True
End Sub

Private Sub txtLocFax_GotFocus()
    FocusMe txtLocFax
End Sub

Private Sub txtLocFax_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtLocFax, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboLocTerrCode.SetFocus
    End If
End Sub

Private Sub txtLocFax_LostFocus()
    FocusMe txtLocFax, True
End Sub

Private Function Chk_cboLocTerrCode() As Boolean
    Dim wsStatus As String

    Chk_cboLocTerrCode = False
    
    If Trim(cboLocTerrCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboLocTerrCode.SetFocus
        Exit Function
    End If

    If Chk_LocTerrCode(cboLocTerrCode.Text, wsStatus) = False Then
        gsMsg = "地區編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboLocTerrCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "地區已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboLocTerrCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboLocTerrCode = True
End Function

Private Function Chk_LocTerrCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_LocTerrCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT TerrStatus "
    wsSql = wsSql & " FROM MstTerritory WHERE TerrCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "TerrStatus")
    
    Chk_LocTerrCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

