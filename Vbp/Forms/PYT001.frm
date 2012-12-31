VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmPYT001 
   BackColor       =   &H8000000A&
   Caption         =   "付款條款"
   ClientHeight    =   3930
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "PYT001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "PYT001.frx":08CA
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboPayCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtPayDay 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3240
         TabIndex        =   4
         Top             =   1660
         Width           =   885
      End
      Begin VB.OptionButton optBy 
         Caption         =   "BYMONTH"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   2200
         Width           =   1335
      End
      Begin VB.OptionButton optBy 
         Caption         =   "BYDAY"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtPayClsDay 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7200
         TabIndex        =   6
         Top             =   2190
         Width           =   885
      End
      Begin VB.TextBox txtPayInvDay 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7200
         TabIndex        =   7
         Top             =   2540
         Width           =   885
      End
      Begin VB.TextBox txtPayMonth 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3240
         TabIndex        =   5
         Top             =   2190
         Width           =   885
      End
      Begin VB.TextBox txtPayCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   8
         Tag             =   "K"
         Top             =   360
         Width           =   2730
      End
      Begin VB.TextBox txtPayDesc 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   6495
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   8100
         Begin VB.Label lblPayDay 
            Caption         =   "PAYDAY"
            Height          =   255
            Left            =   1560
            TabIndex        =   24
            Top             =   280
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   8100
         Begin VB.Label lblPayMonth 
            Caption         =   "PAYMONTH"
            Height          =   255
            Left            =   1560
            TabIndex        =   23
            Top             =   310
            Width           =   1215
         End
         Begin VB.Label lblPayInvDay 
            Caption         =   "PAYINVDAY"
            Height          =   255
            Left            =   5400
            TabIndex        =   22
            Top             =   660
            Width           =   1575
         End
         Begin VB.Label lblPayClsDay 
            Caption         =   "PAYCLSDAY"
            Height          =   255
            Left            =   5400
            TabIndex        =   21
            Top             =   315
            Width           =   1575
         End
      End
      Begin VB.Label lblPayMethod 
         Caption         =   "PAYMETHOD"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblPayLastUpd 
         Caption         =   "PAYLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   3045
         Width           =   1740
      End
      Begin VB.Label lblPayLastUpdDate 
         Caption         =   "PAYLASTUPDDATE"
         Height          =   240
         Left            =   4320
         TabIndex        =   14
         Top             =   3045
         Width           =   1620
      End
      Begin VB.Label lblDspPayLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2160
         TabIndex        =   13
         Top             =   3000
         Width           =   2025
      End
      Begin VB.Label lblDspPayLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6000
         TabIndex        =   12
         Top             =   3000
         Width           =   2145
      End
      Begin VB.Label lblPayCode 
         Caption         =   "PAYCODE"
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
         TabIndex        =   11
         Top             =   420
         Width           =   1260
      End
      Begin VB.Label lblPayDesc 
         Caption         =   "PAYDESC"
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   780
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
            Picture         =   "PYT001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PYT001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   16
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
Attribute VB_Name = "frmPYT001"
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

Private Const wsKeyType = "MstPayTerm"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboPayCode_LostFocus()
    FocusMe cboPayCode, True
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
        Me.Height = 4335
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
            Me.txtPayDesc.Enabled = False
            
            Me.cboPayCode.Enabled = False
            Me.cboPayCode.Visible = False
            Me.txtPayCode.Visible = True
            Me.txtPayCode.Enabled = False
            
            Me.txtPayDay.Enabled = False
            
            Me.optBy(0).Enabled = False
            Me.optBy(1).Enabled = False
            
            Me.txtPayDay.Enabled = False
            Me.txtPayMonth.Enabled = False
            Me.txtPayInvDay.Enabled = False
            Me.txtPayClsDay.Enabled = False
            
            txtPayMonth.Text = "0"
            txtPayInvDay.Text = "0"
            txtPayClsDay.Text = "0"
            txtPayDay.Text = "0"
            
            
        Case "AfrActAdd"
            Me.cboPayCode.Enabled = False
            Me.cboPayCode.Visible = False
            
            Me.txtPayCode.Enabled = True
            Me.txtPayCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboPayCode.Enabled = True
            Me.cboPayCode.Visible = True
            
            Me.txtPayCode.Enabled = False
            Me.txtPayCode.Visible = False
            
        Case "AfrKey"
            Me.txtPayDesc.Enabled = True
            
            Me.txtPayDay.Enabled = True
            
            Me.optBy(0).Enabled = True
            Me.optBy(1).Enabled = True
            
            Me.txtPayMonth.Enabled = False
            Me.txtPayInvDay.Enabled = False
            Me.txtPayClsDay.Enabled = False
            
            Me.cboPayCode.Enabled = False
            Me.txtPayCode.Enabled = False
            
            If optBy(0).Value = True Then
                optBy_Click 0
            ElseIf optBy(1).Value = True Then
                optBy_Click 1
            End If
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtPayDesc = False Then
        Exit Function
    End If
    
    If optBy(0) = True Then
        If Chk_txtPayDay = False Then
            Exit Function
        End If
    ElseIf optBy(1) = True Then
        If Chk_txtPayMonth = False Then
            Exit Function
        End If
        
        If Chk_txtPayInvDay = False Then
            Exit Function
        End If
        
        If Chk_txtPayClsDay = False Then
            Exit Function
        End If
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim sMethod As String
    
    wsSQL = "SELECT MstPayTerm.* "
    wsSQL = wsSQL + "From MstPayTerm "
    wsSQL = wsSQL + "WHERE (((MstPayTerm.PayCode)='" + Set_Quote(cboPayCode) + "') "
    wsSQL = wsSQL + "AND ((MstPayTerm.PayStatus)='1'));"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        
        Me.cboPayCode = ReadRs(rsRcd, "PayCode")
        Me.txtPayDesc = ReadRs(rsRcd, "PayDesc")
        
        sMethod = ReadRs(rsRcd, "PayMethod")
        
        '1 : Month, 2 : Day
        If sMethod = 1 Then
            optBy(1).Value = True
            setOptButton optBy(1)
        ElseIf sMethod = 2 Then
            optBy(0).Value = True
            setOptButton optBy(0)
        End If
        
        Me.txtPayDay = Format(ReadRs(rsRcd, "PayDay"), gsQtyFmt)
        
        Me.txtPayMonth = Format(ReadRs(rsRcd, "PayMonth"), gsQtyFmt)
        Me.txtPayClsDay = Format(ReadRs(rsRcd, "PayClsDay"), gsQtyFmt)
        Me.txtPayInvDay = Format(ReadRs(rsRcd, "PayInvDay"), gsQtyFmt)
        
        Me.lblDspPayLastUpd = ReadRs(rsRcd, "PayLastUpd")
        Me.lblDspPayLastUpdDate = ReadRs(rsRcd, "PayLastUpdDate")
        
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
    Set frmPYT001 = Nothing
End Sub

Private Sub optBy_Click(Index As Integer)
    If Index = 0 Then
        lblPayDay.Enabled = True
        Me.txtPayDay.Enabled = True
        
        lblPayMonth.Enabled = False
        lblPayInvDay.Enabled = False
        lblPayClsDay.Enabled = False
        Me.txtPayMonth.Enabled = False
        Me.txtPayInvDay.Enabled = False
        Me.txtPayClsDay.Enabled = False
    ElseIf Index = 1 Then
        lblPayDay.Enabled = False
        Me.txtPayDay.Enabled = False
        
        lblPayMonth.Enabled = True
        lblPayInvDay.Enabled = True
        lblPayClsDay.Enabled = True
        Me.txtPayMonth.Enabled = True
        Me.txtPayInvDay.Enabled = True
        Me.txtPayClsDay.Enabled = True
    End If
End Sub

Private Sub optBy_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        optBy_Click Index
        
        If Index = 0 Then
        
            If optBy(0).Value = True Then
                txtPayDay.SetFocus
            ElseIf optBy(1).Value = True Then
                optBy(1).SetFocus
            End If
            
        ElseIf Index = 1 Then
            If optBy(0).Value = True Then
                optBy(1).SetFocus
            ElseIf optBy(1).Value = True Then
                txtPayMonth.SetFocus
            End If
        End If
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
    wsFormID = "PYT001"
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
        txtPayCode.SetFocus
        
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboPayCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboPayCode.SetFocus
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
            If RowLock(wsConnTime, wsKeyType, cboPayCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtPayDesc.SetFocus
End Sub

Private Function Chk_PayCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_PayCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT PayStatus "
    wsSQL = wsSQL & " FROM MstPayTerm WHERE PayCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "PayStatus")
    
    Chk_PayCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboPayCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboPayCode = False
    
    If Trim(cboPayCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboPayCode.SetFocus
        Exit Function
    End If
    
    If Chk_PayCode(cboPayCode.Text, wsStatus) = False Then
        gsMsg = "付款條款編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboPayCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "付款條款編碼已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboPayCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboPayCode = True
End Function

Private Function Chk_txtPayCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtPayCode = False
    
    If Trim(txtPayCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayCode.SetFocus
        Exit Function
    End If
    
    If Chk_PayCode(txtPayCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "付款條款編碼已存在但已無效!"
        Else
            gsMsg = "付款條款編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayCode.SetFocus
        Exit Function
    End If
    
    Chk_txtPayCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmPYT001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboPayCode, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_PYT001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtPayCode, cboPayCode))
    Call SetSPPara(adcmdSave, 3, txtPayDesc)
    
    If optBy(0).Value = True Then
        Call SetSPPara(adcmdSave, 4, 2)
        Call SetSPPara(adcmdSave, 5, txtPayDay)
        Call SetSPPara(adcmdSave, 6, 0)
        Call SetSPPara(adcmdSave, 7, 0)
        Call SetSPPara(adcmdSave, 8, 0)
    ElseIf optBy(1).Value = True Then
        Call SetSPPara(adcmdSave, 4, 1)
        Call SetSPPara(adcmdSave, 5, 0)
        Call SetSPPara(adcmdSave, 6, txtPayMonth)
        Call SetSPPara(adcmdSave, 7, txtPayClsDay)
        Call SetSPPara(adcmdSave, 8, txtPayInvDay)
    End If
    
    Call SetSPPara(adcmdSave, 9, gsUserID)
    Call SetSPPara(adcmdSave, 10, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 11)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - PYT001!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
        gsMsg = "已成功儲存!"
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
    Dim sSQL As String
    
    ReDim vFilterAry(2, 2)
    vFilterAry(1, 1) = "付款條款編碼"
    vFilterAry(1, 2) = "PayCode"
    
    vFilterAry(2, 1) = "註解"
    vFilterAry(2, 2) = "PayDesc"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "付款條款編碼"
    vAry(1, 2) = "PayCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "註解"
    vAry(2, 2) = "PayDesc"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstPayTerm.PayCode, MstPayTerm.PayDesc "
        sSQL = sSQL + "FROM MstPayTerm "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE MstPayTerm.PayStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstPayTerm.PayCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboPayCode Then
        cboPayCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
    Unload frmShareSearch
End Sub

Private Sub txtPayCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtPayCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtPayCode_LostFocus()
    FocusMe txtPayCode, True
End Sub

Private Sub txtPayDay_LostFocus()
    FocusMe txtPayDay, True
End Sub

Private Sub txtPayDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtPayDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayDesc = True Then
            If optBy(1).Value = True Then
                optBy(1).SetFocus
            Else
                optBy(0).SetFocus
            End If
            
        End If
    End If
End Sub

Private Sub txtPayCode_GotFocus()
    FocusMe txtPayCode
End Sub

Private Sub txtPayDesc_GotFocus()
    FocusMe txtPayDesc
End Sub

Private Function Chk_txtPayDesc() As Boolean
    
    Chk_txtPayDesc = False
    
    If Trim(txtPayDesc.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayDesc.SetFocus
        Exit Function
    End If
    
    Chk_txtPayDesc = True
End Function

Private Sub txtPayDay_GotFocus()
    FocusMe txtPayDay
End Sub

Private Sub txtPayDay_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayDay, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayDay() = True Then
            txtPayDesc.SetFocus
        End If
    End If
End Sub

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

Private Sub cboPayCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboPayCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPayCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboPayCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboPayCode
    
    wsSQL = "SELECT PayCode, PayDesc, PayMethod FROM MstPayTerm WHERE PayStatus = '1'"
    wsSQL = wsSQL & " AND PayCode LIKE '%" & IIf(cboPayCode.SelLength > 0, "", Set_Quote(cboPayCode.Text)) & "%' "
  
    wsSQL = wsSQL & "ORDER BY PayCode "
    Call Ini_Combo(3, wsSQL, cboPayCode.Left, cboPayCode.Top + cboPayCode.Height, tblCommon, "PYT001", "TBLPYT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboPayCode_GotFocus()
    FocusMe cboPayCode
End Sub


Private Sub txtPayDesc_LostFocus()
    FocusMe txtPayDesc, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT PayStatus FROM MstPayTerm WHERE PayCode = '" & Set_Quote(txtPayCode) & "'"
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
        .TableKey = "PayCode"
        .KeyLen = 10
        Set .ctlKey = txtPayCode
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
    
    lblPayCode.Caption = Get_Caption(waScrItm, "PAYCODE")
    lblPayDesc.Caption = Get_Caption(waScrItm, "PAYDESC")
    lblPayMethod.Caption = Get_Caption(waScrItm, "PAYMETHOD")
    lblPayDay.Caption = Get_Caption(waScrItm, "PAYDAY")
    lblPayMonth.Caption = Get_Caption(waScrItm, "PAYMONTH")
    lblPayClsDay.Caption = Get_Caption(waScrItm, "PAYCLSDAY")
    lblPayInvDay.Caption = Get_Caption(waScrItm, "PAYINVDAY")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    optBy(0).Caption = Get_Caption(waScrItm, "BYDAY")
    optBy(1).Caption = Get_Caption(waScrItm, "BYMONTH")
    
    lblPayLastUpd.Caption = Get_Caption(waScrItm, "PAYLASTUPD")
    lblPayLastUpdDate.Caption = Get_Caption(waScrItm, "PAYLASTUPDDATE")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "PYTADD")
    wsActNam(2) = Get_Caption(waScrItm, "PYTEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PYTDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub txtPayMonth_GotFocus()
    FocusMe txtPayMonth
End Sub

Private Sub txtPayMonth_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayMonth, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayMonth() = True Then
            txtPayClsDay.SetFocus
        End If
    End If
End Sub

Private Sub txtPayMonth_LostFocus()
    FocusMe txtPayMonth, True
End Sub

Private Sub txtPayClsDay_GotFocus()
    FocusMe txtPayClsDay
End Sub

Private Sub txtPayClsDay_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayClsDay, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayClsDay() = True Then
            txtPayInvDay.SetFocus
        End If
    End If
End Sub

Private Sub txtPayClsDay_LostFocus()
    FocusMe txtPayClsDay, True
End Sub

Private Sub txtPayInvDay_GotFocus()
    FocusMe txtPayInvDay
End Sub

Private Sub txtPayInvDay_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtPayInvDay, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtPayInvDay() = True Then
            txtPayDesc.SetFocus
        End If
    End If
End Sub

Private Sub txtPayInvDay_LostFocus()
    FocusMe txtPayInvDay, True
End Sub

Private Sub setOptButton(ctl As Control)
    ctl.Value = True
End Sub

Private Function Chk_txtPayDay() As Boolean
    Chk_txtPayDay = False
    
    If Trim(txtPayDay.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayDay.SetFocus
        Exit Function
    End If
    
    Chk_txtPayDay = True
End Function

Private Function Chk_txtPayMonth() As Boolean
    Chk_txtPayMonth = False
    
    If Trim(txtPayMonth.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayMonth.SetFocus
        Exit Function
    End If
    
    Chk_txtPayMonth = True
End Function

Private Function Chk_txtPayClsDay() As Boolean
    Chk_txtPayClsDay = False
    
    If Trim(txtPayClsDay.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayClsDay.SetFocus
        Exit Function
    End If
    
    Chk_txtPayClsDay = True
End Function

Private Function Chk_txtPayInvDay() As Boolean
    Chk_txtPayInvDay = False
    
    If Trim(txtPayInvDay.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtPayInvDay.SetFocus
        Exit Function
    End If
    
    Chk_txtPayInvDay = True
End Function

