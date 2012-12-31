VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmML001 
   BackColor       =   &H8000000A&
   Caption         =   "frmML001"
   ClientHeight    =   4575
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "ML001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "ML001.frx":08CA
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCOAAccCode 
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   1530
   End
   Begin VB.ComboBox cboMLCode 
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1530
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   4095
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   8355
      Begin VB.Frame FraMLType 
         Caption         =   "MLTYPE"
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   8055
         Begin VB.OptionButton optMLType 
            Caption         =   "SALES"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optMLType 
            Caption         =   "PURCHASE"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optMLType 
            Caption         =   "A/R"
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optMLType 
            Caption         =   "A/P"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optMLType 
            Caption         =   "CHEQUE"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   7
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton optMLType 
            Caption         =   "BANK"
            Height          =   255
            Index           =   5
            Left            =   5880
            TabIndex        =   8
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.TextBox txtMLCode 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Tag             =   "K"
         Top             =   600
         Width           =   1530
      End
      Begin VB.TextBox txtMLDesc 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1920
         TabIndex        =   9
         Top             =   2880
         Width           =   6255
      End
      Begin VB.Label lblDspCOADesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3480
         TabIndex        =   21
         Top             =   960
         Width           =   4665
      End
      Begin VB.Label lblCOAAccCode 
         Caption         =   "COAACCCODE"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   1035
         Width           =   1740
      End
      Begin VB.Label lblMLLastUpd 
         Caption         =   "MLLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   3645
         Width           =   1980
      End
      Begin VB.Label lblMLLastUpdDate 
         Caption         =   "MLLASTUPDDATE"
         Height          =   240
         Left            =   4320
         TabIndex        =   15
         Top             =   3645
         Width           =   2220
      End
      Begin VB.Label lblDspMLLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2520
         TabIndex        =   14
         Top             =   3600
         Width           =   1665
      End
      Begin VB.Label lblDspMLLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6600
         TabIndex        =   13
         Top             =   3600
         Width           =   1545
      End
      Begin VB.Label lblMLCode 
         Caption         =   "MLCODE"
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
         Left            =   120
         TabIndex        =   12
         Top             =   660
         Width           =   1740
      End
      Begin VB.Label lblMLDesc 
         Caption         =   "MLDESC"
         Height          =   240
         Left            =   360
         TabIndex        =   11
         Top             =   2955
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
            Picture         =   "ML001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ML001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   17
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
Attribute VB_Name = "frmML001"
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
Private wlMLAccID As Long
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control

Private Const wsKeyType = "MstMerchClass"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboMLCode_LostFocus()
    FocusMe cboMLCode, True
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
        Me.Height = 4980
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
            Me.txtMLDesc.Enabled = False
            
            Me.cboMLCode.Enabled = False
            Me.cboMLCode.Visible = False
            Me.txtMLCode.Visible = True
            Me.txtMLCode.Enabled = False
            cboCOAAccCode.Enabled = False
            
            optMLType(0).Enabled = False
            optMLType(1).Enabled = False
            optMLType(2).Enabled = False
            optMLType(3).Enabled = False
            optMLType(4).Enabled = False
            optMLType(5).Enabled = False
            
        Case "AfrActAdd"
            Me.cboMLCode.Enabled = False
            Me.cboMLCode.Visible = False
            
            Me.txtMLCode.Enabled = True
            Me.txtMLCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboMLCode.Enabled = True
            Me.cboMLCode.Visible = True
            
            Me.txtMLCode.Enabled = False
            Me.txtMLCode.Visible = False
            
        Case "AfrKey"
            Me.txtMLDesc.Enabled = True
            cboCOAAccCode.Enabled = True
            
            optMLType(0).Enabled = True
            optMLType(1).Enabled = True
            optMLType(2).Enabled = True
            optMLType(3).Enabled = True
            optMLType(4).Enabled = True
            optMLType(5).Enabled = True
            
            Me.cboMLCode.Enabled = False
            Me.txtMLCode.Enabled = False
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtMLDesc = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT MstMerchClass.* "
    wsSQL = wsSQL + "From MstMerchClass "
    wsSQL = wsSQL + "WHERE (((MstMerchClass.MLCode)='" + Set_Quote(cboMLCode) + "') "
    wsSQL = wsSQL + "AND ((MLStatus)='1'));"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        
        Me.cboMLCode = ReadRs(rsRcd, "MLCode")
        Me.txtMLDesc = ReadRs(rsRcd, "MLDesc")
        Me.lblDspMLLastUpd = ReadRs(rsRcd, "MLLastUpd")
        Me.lblDspMLLastUpdDate = ReadRs(rsRcd, "MLLastUpdDate")
        wlMLAccID = To_Value(ReadRs(rsRcd, "MLAccID"))
        cboCOAAccCode = Get_TableInfo("MstCOA", "COAAccID =" & wlMLAccID, "COAAccCode")
        SetMLType ReadRs(rsRcd, "MLType")
        lblDspCOADesc.Caption = Get_TableInfo("MstCOA", "COAAccCode = '" & Set_Quote(cboCOAAccCode.Text) & "'", "COADesc")
        
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
    Set frmML001 = Nothing
End Sub

Private Sub optMLType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        txtMLDesc.SetFocus
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
'    Me.Left = 0
'    Me.Top = 0
'    Me.Width = Screen.Width
'    Me.Height = Screen.Height
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "ML001"
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
        txtMLCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboMLCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboMLCode.SetFocus
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
            If RowLock(wsConnTime, wsKeyType, cboMLCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    cboCOAAccCode.SetFocus
End Sub

Private Function Chk_MLCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_MLCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT MLStatus "
    wsSQL = wsSQL & " FROM MstMerchClass WHERE MLCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "MLStatus")
    
    Chk_MLCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_COAAccCode(ByVal inCode As String, ByRef OutID As Long, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_COAAccCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT COAAccID, COAStatus "
    wsSQL = wsSQL & " FROM MstCOA WHERE COAAccCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        OutID = 0
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "COAStatus")
    OutID = ReadRs(rsRcd, "COAAccID")
    
    Chk_COAAccCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboMLCode() As Boolean
    Dim wsStatus As String

    Chk_cboMLCode = False
    
    If Trim(cboMLCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboMLCode.SetFocus
        Exit Function
    End If

    If Chk_MLCode(cboMLCode.Text, wsStatus) = False Then
        gsMsg = "買手編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboMLCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "買手已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboMLCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboMLCode = True
End Function

Private Function Chk_cboCOAAccCode() As Boolean
    Dim wsStatus As String

    Chk_cboCOAAccCode = False
    
    If Trim(cboCOAAccCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCOAAccCode.SetFocus
        Exit Function
    End If

    If Chk_COAAccCode(cboCOAAccCode.Text, wlMLAccID, wsStatus) = False Then
        gsMsg = "會計科目編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCOAAccCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "會計科目已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCOAAccCode.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboCOAAccCode = True
End Function

Private Function Chk_txtMLCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtMLCode = False
    
    If Trim(txtMLCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_MLCode(txtMLCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "買手編碼已存在但已無效!"
        Else
            gsMsg = "買手編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtMLCode.SetFocus
        Exit Function
    End If
    
    Chk_txtMLCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmML001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboMLCode, wsFormID) Then
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
        
        If Chk_MLExist(cboMLCode.Text) = True Then
            gsMsg = "記錄已用, 不能刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
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
        
    adcmdSave.CommandText = "USP_ML001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtMLCode, cboMLCode))
    Call SetSPPara(adcmdSave, 3, txtMLDesc)
    Call SetSPPara(adcmdSave, 4, wlMLAccID)
    Call SetSPPara(adcmdSave, 5, GetMLType())
    Call SetSPPara(adcmdSave, 6, gsUserID)
    Call SetSPPara(adcmdSave, 7, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 8)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - ML001!"
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
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "買手編碼"
    vFilterAry(1, 2) = "MLCode"
    
    vFilterAry(2, 1) = "會計項"
    vFilterAry(2, 2) = "MLType"
    
    vFilterAry(3, 1) = "註解"
    vFilterAry(3, 2) = "MLDesc"
    
    ReDim vAry(3, 3)
    vAry(1, 1) = "買手編碼"
    vAry(1, 2) = "MLCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "會計項"
    vAry(2, 2) = "MLType"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "註解"
    vAry(3, 2) = "MLDesc"
    vAry(3, 3) = "4000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstMerchClass.MLCode, MstMerchClass.MLType, MstMerchClass.MLDesc "
        sSQL = sSQL + "FROM MstMerchClass "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE MstMerchClass.MLStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstMerchClass.MLCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboMLCode Then
        cboMLCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
    Unload frmShareSearch
End Sub

Private Sub txtMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtMLCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtMLCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtMLCode_LostFocus()
    FocusMe txtMLCode, True
End Sub

Private Sub txtMLDesc_KeyPress(KeyAscii As Integer)
    Dim iCounter As Integer
    
    Call chk_InpLen(txtMLDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtMLDesc = True Then
              cboCOAAccCode.SetFocus
        End If
    End If
End Sub

Private Sub txtMLCode_GotFocus()
    FocusMe txtMLCode
End Sub

Private Sub txtMLDesc_GotFocus()
    FocusMe txtMLDesc
End Sub

Private Function Chk_txtMLDesc() As Boolean
    
    Chk_txtMLDesc = False
    
    If Trim(txtMLDesc.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtMLDesc.SetFocus
        Exit Function
    End If
    
    Chk_txtMLDesc = True
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

Private Sub cboMLCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboMLCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboMLCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "

    Call Ini_Combo(2, wsSQL, cboMLCode.Left, cboMLCode.Top + cboMLCode.Height, tblCommon, "ML001", "TBLML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboMLCode_GotFocus()
    FocusMe cboMLCode
End Sub

Private Sub txtMLDesc_LostFocus()
    FocusMe txtMLDesc, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT MLStatus FROM MstMerchClass WHERE MLCode = '" & Set_Quote(txtMLCode) & "'"
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
        .TableKey = "MLCode"
        .KeyLen = 10
        Set .ctlKey = txtMLCode
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
    
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblMLDesc.Caption = Get_Caption(waScrItm, "MLDESC")
    lblCOAAccCode.Caption = Get_Caption(waScrItm, "COAACCCODE")
    lblMLLastUpd.Caption = Get_Caption(waScrItm, "MLLASTUPD")
    lblMLLastUpdDate.Caption = Get_Caption(waScrItm, "MLLASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    optMLType(0).Caption = Get_Caption(waScrItm, "OPTMLTYPE0")
    optMLType(1).Caption = Get_Caption(waScrItm, "OPTMLTYPE1")
    optMLType(2).Caption = Get_Caption(waScrItm, "OPTMLTYPE2")
    optMLType(3).Caption = Get_Caption(waScrItm, "OPTMLTYPE3")
    optMLType(4).Caption = Get_Caption(waScrItm, "OPTMLTYPE4")
    optMLType(5).Caption = Get_Caption(waScrItm, "OPTMLTYPE5")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
    FraMLType.Caption = Get_Caption(waScrItm, "FRAMLTYPE")
   
    wsActNam(1) = Get_Caption(waScrItm, "MLADD")
    wsActNam(2) = Get_Caption(waScrItm, "MLEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "MLDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub cboCOAAccCode_KeyPress(KeyAscii As Integer)
    Dim iCounter As Integer
    
    Call chk_InpLen(cboCOAAccCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCOAAccCode() = True Then
        

            Call Opt_Setfocus(optMLType, 6, 0)
            lblDspCOADesc.Caption = Get_TableInfo("MstCOA", "COAAccCode = '" & Set_Quote(cboCOAAccCode.Text) & "'", "COADesc")
            
        End If
    End If
End Sub

Private Sub cboCOAAccCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCOAAccCode
    
    wsSQL = "SELECT COAAccCode, COADesc FROM MstCOA WHERE COAStatus = '1'"
    wsSQL = wsSQL & " AND COAAccCode LIKE '%" & IIf(cboCOAAccCode.SelLength > 0, "", Set_Quote(cboCOAAccCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY COAAccCode "
    Call Ini_Combo(2, wsSQL, cboCOAAccCode.Left, cboCOAAccCode.Top + cboCOAAccCode.Height, tblCommon, "ML001", "TBLCOA", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCOAAccCode_GotFocus()
    FocusMe cboCOAAccCode
End Sub

Private Sub cboCOAAccCode_LostFocus()
    FocusMe cboCOAAccCode, True
End Sub

Private Sub SetMLType(ByVal inCode As String)
    Select Case inCode
        Case "S"
            optMLType(0).Value = True
            
        Case "P"
            optMLType(1).Value = True
            
        Case "A"
            optMLType(2).Value = True
            
        Case "R"
            optMLType(3).Value = True
            
        Case "G"
            optMLType(4).Value = True
            
        Case "B"
            optMLType(5).Value = True
    End Select
End Sub

Private Function GetMLType() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 5
        If optMLType(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetMLType = "S"
            
        Case 1
            GetMLType = "P"
        
        Case 2
            GetMLType = "A"
        
        Case 3
            GetMLType = "R"
            
        Case 4
            GetMLType = "G"
        
        Case 5
            GetMLType = "B"
    End Select
End Function


Private Function Chk_MLExist(ByVal inCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
wsSQL = "SELECT MLCODE FROM"
wsSQL = wsSQL & " ( "
wsSQL = wsSQL & " SELECT IPHDMLCODE MLCODE FROM APIPHD WHERE IPHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT IPDTMLCODE MLCODE FROM APIPHD, APIPDT"
wsSQL = wsSQL & " WHERE IPHDSTATUS <> '2'"
wsSQL = wsSQL & " AND IPHDDOCID = IPDTDOCID"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT APCQBANKML MLCODE FROM APCHEQUE"
wsSQL = wsSQL & " WHERE APCQSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT APCQTMPML MLCODE FROM APCHEQUE"
wsSQL = wsSQL & " WHERE APCQSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT ApShMLCODE MLCODE FROM APSTHD"
wsSQL = wsSQL & " WHERE APSHSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT ApSDMLCODE MLCODE FROM APSTHD, APSTDT"
wsSQL = wsSQL & " WHERE APSHSTATUS <> '2'"
wsSQL = wsSQL & " AND APSHDOCID = APSDDOCID"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT INHDMLCODE MLCODE FROM ARINHD WHERE INHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT INDTMLCODE MLCODE FROM ARINHD, ARINDT"
wsSQL = wsSQL & " WHERE INHDSTATUS <> '2'"
wsSQL = wsSQL & " AND INHDDOCID = INDTDOCID"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT ARCQBANKML MLCODE FROM ARCHEQUE"
wsSQL = wsSQL & " WHERE ARCQSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT ARCQTMPML MLCODE FROM ARCHEQUE"
wsSQL = wsSQL & " WHERE ARCQSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT ARShMLCODE MLCODE FROM ARSTHD"
wsSQL = wsSQL & " WHERE ARSHSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT ARSDMLCODE MLCODE FROM ARSTHD, ARSTDT"
wsSQL = wsSQL & " WHERE ARSHSTATUS <> '2'"
wsSQL = wsSQL & " AND ARSHDOCID = ARSDDOCID"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT IVHDMLCODE MLCODE FROM SOAIVHD"
wsSQL = wsSQL & " WHERE IVHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT IVHDCRML MLCODE FROM SOAIVHD"
wsSQL = wsSQL & " WHERE IVHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT GRHDMLCODE MLCODE FROM POPGRHD"
wsSQL = wsSQL & " WHERE GRHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT GRHDCRML MLCODE FROM POPGRHD"
wsSQL = wsSQL & " WHERE GRHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT PRHDMLCODE MLCODE FROM POPPRHD"
wsSQL = wsSQL & " WHERE PRHDSTATUS <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT AccTypeSALML MLCODE"
wsSQL = wsSQL & " From MSTACCOUNTTYPE"
wsSQL = wsSQL & " WHERE AccTypeStatus <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT AccTypeCOSML MLCODE"
wsSQL = wsSQL & " From MSTACCOUNTTYPE"
wsSQL = wsSQL & " WHERE AccTypeStatus <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT AccTypeINVML MLCODE"
wsSQL = wsSQL & " From MSTACCOUNTTYPE"
wsSQL = wsSQL & " WHERE AccTypeStatus <> '2'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpSupMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpExgMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpExlMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpTIMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpTEMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpSamMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpDamMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpAdjMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " "
wsSQL = wsSQL & " Union"
wsSQL = wsSQL & " SELECT CmpTRMLCode MLCODE FROM MSTCOMPANY WHERE CMPID = '01'"
wsSQL = wsSQL & " ) A"
wsSQL = wsSQL & " WHERE ISNULL(MLCODE,'') <> ''"
wsSQL = wsSQL & " AND MLCode = '" & Set_Quote(inCode) & "' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        Chk_MLExist = True
    Else
        Chk_MLExist = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function

