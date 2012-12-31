VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmCOA001 
   BackColor       =   &H8000000A&
   Caption         =   "COA001"
   ClientHeight    =   5505
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8895
   Icon            =   "COA001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   8895
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "COA001.frx":08CA
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCOAAccCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   8715
      Begin VB.TextBox txtCOACDesc 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2280
         TabIndex        =   9
         Top             =   3000
         Width           =   6210
      End
      Begin VB.TextBox txtCOAConsCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2280
         TabIndex        =   10
         Top             =   3480
         Width           =   2730
      End
      Begin VB.Frame FRAACCTYPE 
         Caption         =   "ACCTYPE"
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   8415
         Begin VB.OptionButton optCOAClass 
            Caption         =   "SUSPEND"
            Height          =   255
            Index           =   5
            Left            =   5880
            TabIndex        =   7
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton optCOAClass 
            Caption         =   "EXPENSE"
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   6
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton optCOAClass 
            Caption         =   "REVENUE"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   5
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optCOAClass 
            Caption         =   "CAPITAL"
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optCOAClass 
            Caption         =   "LIABILITY"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   3
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optCOAClass 
            Caption         =   "ASSET"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   2
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCOADesc 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Top             =   2640
         Width           =   6210
      End
      Begin VB.TextBox txtCOAAccCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   600
         Width           =   2730
      End
      Begin VB.Label lblCOAcDesc 
         Caption         =   "COADESC"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3060
         Width           =   2055
      End
      Begin VB.Label lblCOAConsCode 
         Caption         =   "COACONSCODE"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3540
         Width           =   2175
      End
      Begin VB.Label lblCOAAccCode 
         Caption         =   "COAACCCODE"
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
         TabIndex        =   19
         Top             =   675
         Width           =   2100
      End
      Begin VB.Label lblCOALastUpd 
         Caption         =   "COALASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   4605
         Width           =   1860
      End
      Begin VB.Label lblCOaLastUpdDate 
         Caption         =   "COALASTUPDDATE"
         Height          =   240
         Left            =   4320
         TabIndex        =   16
         Top             =   4605
         Width           =   1980
      End
      Begin VB.Label lblDspCOALastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2280
         TabIndex        =   15
         Top             =   4560
         Width           =   1905
      End
      Begin VB.Label lblDspCOALastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6480
         TabIndex        =   14
         Top             =   4560
         Width           =   1665
      End
      Begin VB.Label lblCOADesc 
         Caption         =   "COADESC"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2700
         Width           =   2175
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
            Picture         =   "COA001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "COA001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
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
Attribute VB_Name = "frmCOA001"
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
Private wlCOAAccID As Long

Private wcCombo As Control

Private wsActNam(4) As String
'Row Lock Variable

Private Const wsKeyType = "MstCOA"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboCOAAccCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCOAAccCode
    

    wsSQL = "SELECT COAAccCode, " & IIf(gsLangID = "2", "COACDESC", "COADESC") & " FROM MstCOA WHERE COAStatus = '1'"
    wsSQL = wsSQL & " AND COAAccCode LIKE '%" & IIf(cboCOAAccCode.SelLength > 0, "", Set_Quote(cboCOAAccCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY COAAccCode "
    
    
    Call Ini_Combo(2, wsSQL, cboCOAAccCode.Left, cboCOAAccCode.Top + cboCOAAccCode.Height, tblCommon, "COA001", "TBLCOA", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCOAAccCode_GotFocus()
    FocusMe cboCOAAccCode
End Sub

Private Sub cboCOAAccCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCOAAccCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCOAAccCode() = True Then
            Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboCOAAccCode_LostFocus()
    FocusMe cboCOAAccCode, True
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
        Me.Height = 5910
        Me.Width = 9015
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
            txtCOAAccCode.Enabled = False
            cboCOAAccCode.Enabled = False
        
            optCOAClass(0).Enabled = False
            optCOAClass(1).Enabled = False
            optCOAClass(2).Enabled = False
            optCOAClass(3).Enabled = False
            optCOAClass(4).Enabled = False
            optCOAClass(5).Enabled = False
            
            txtCOADesc.Enabled = False
            txtCOACDesc.Enabled = False
            
            txtCOAConsCode.Enabled = False
            
            Me.cboCOAAccCode.Visible = False
            Me.txtCOAAccCode.Visible = True
            
            optCOAClass(0).Value = True
            
        Case "AfrActAdd"
            cboCOAAccCode.Enabled = False
            
            txtCOAAccCode.Enabled = True
            txtCOAAccCode.Visible = True
            
        Case "AfrActEdit"
            cboCOAAccCode.Enabled = True
            cboCOAAccCode.Visible = True
            
            txtCOAAccCode.Enabled = False
            txtCOAAccCode.Visible = False
            
        Case "AfrKey"
            txtCOAAccCode.Enabled = False
            cboCOAAccCode.Enabled = False
            
            optCOAClass(0).Enabled = True
            optCOAClass(1).Enabled = True
            optCOAClass(2).Enabled = True
            optCOAClass(3).Enabled = True
            optCOAClass(4).Enabled = True
            optCOAClass(5).Enabled = True
            
            txtCOADesc.Enabled = True
            txtCOACDesc.Enabled = True
            
            txtCOAConsCode.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    InputValidation = False
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim iCounter As Integer
        
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From MstCOA "
    wsSQL = wsSQL + "WHERE (((MstCOA.COAAccCode)='" + Set_Quote(cboCOAAccCode.Text) + "' AND COAStatus = '1'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        cboCOAAccCode = ReadRs(rsRcd, "COAAccCode")
        txtCOADesc = ReadRs(rsRcd, "COADesc")
        txtCOACDesc = ReadRs(rsRcd, "COACDesc")
        
        txtCOAConsCode = ReadRs(rsRcd, "COAConsCode")
        
        lblDspCOALastUpd = ReadRs(rsRcd, "COALastUpd")
        lblDspCOALastUpdDate = ReadRs(rsRcd, "COALastUpdDate")
        
        wlCOAAccID = ReadRs(rsRcd, "COAAccID")
        
        SetCOAClass ReadRs(rsRcd, "COAClass")
        
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
    Set frmCOA001 = Nothing
End Sub

Private Sub optCOAClass_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        txtCOADesc.SetFocus
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
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "COA001"
    wsTrnCd = ""
End Sub

Private Sub Ini_Caption()
On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblCOAAccCode.Caption = Get_Caption(waScrItm, "COAACCCODE")
    lblCOADesc.Caption = Get_Caption(waScrItm, "COADESC")
    lblCOAcDesc.Caption = Get_Caption(waScrItm, "COACDESC")
    lblCOAConsCode.Caption = Get_Caption(waScrItm, "COACONSCODE")
    lblCOALastUpd.Caption = Get_Caption(waScrItm, "COALASTUPD")
    lblCOaLastUpdDate.Caption = Get_Caption(waScrItm, "COALASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    optCOAClass(0).Caption = Get_Caption(waScrItm, "OPTCOACLASS0")
    optCOAClass(1).Caption = Get_Caption(waScrItm, "OPTCOACLASS1")
    optCOAClass(2).Caption = Get_Caption(waScrItm, "OPTCOACLASS2")
    optCOAClass(3).Caption = Get_Caption(waScrItm, "OPTCOACLASS3")
    optCOAClass(4).Caption = Get_Caption(waScrItm, "OPTCOACLASS4")
    optCOAClass(5).Caption = Get_Caption(waScrItm, "OPTCOACLASS5")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
    FRAACCTYPE.Caption = Get_Caption(waScrItm, "ACCTYPE")
    
    wsActNam(1) = Get_Caption(waScrItm, "COAADD")
    wsActNam(2) = Get_Caption(waScrItm, "COAEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "COADELETE")
    
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
    wlCOAAccID = 0
    
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
        txtCOAAccCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCOAAccCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCOAAccCode.SetFocus
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Dim iCounter As Integer
    
    Select Case wiAction
        Case CorRec, DelRec

            If LoadRecord() = False Then
                gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
                MsgBox gsMsg, vbOKOnly, gsTitle
                Exit Sub
            Else
                If RowLock(wsConnTime, wsKeyType, cboCOAAccCode, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    For iCounter = 0 To 5
        If optCOAClass(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    
    optCOAClass(iCounter).Value = True
    optCOAClass(iCounter).SetFocus
End Sub

Private Function Chk_txtCOAAccCode() As Boolean
    Dim wsStatus As String

    Chk_txtCOAAccCode = False
    
    If Trim(txtCOAAccCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCOAAccCode.SetFocus
        Exit Function
    End If

    If Chk_COAAccCode(txtCOAAccCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "COA 編碼已存在但已無效!"
        Else
            gsMsg = "COA 編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCOAAccCode.SetFocus
        Exit Function
    End If
    
    Chk_txtCOAAccCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmCOA001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboCOAAccCode, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_COA001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlCOAAccID)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, txtCOAAccCode.Text, cboCOAAccCode.Text))
    Call SetSPPara(adcmdSave, 4, GetCOAClass())
    Call SetSPPara(adcmdSave, 5, txtCOADesc)
    Call SetSPPara(adcmdSave, 6, txtCOACDesc)
    Call SetSPPara(adcmdSave, 7, txtCOAConsCode)
    Call SetSPPara(adcmdSave, 8, gsUserID)
    Call SetSPPara(adcmdSave, 9, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 10)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - COA001!"
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
    vFilterAry(1, 1) = "會計科目編碼"
    vFilterAry(1, 2) = "COAAccCode"
    
    vFilterAry(2, 1) = "類別"
    vFilterAry(2, 2) = "COAClass"
    
    vFilterAry(3, 1) = "註解"
    vFilterAry(3, 2) = "COADesc"
    
    ReDim vAry(3, 3)
    vAry(1, 1) = "會計科目編碼"
    vAry(1, 2) = "COAAccCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "類別"
    vAry(2, 2) = "COAClass"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "註解"
    vAry(3, 2) = "COADesc"
    vAry(3, 3) = "4000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT MstCOA.COAAccCode, MstCOA.COAClass, COADesc "
        wsSQL = wsSQL + "FROM MstCOA "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE MstCOA.COAStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstCOA.COAAccCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboCOAAccCode Then
        cboCOAAccCode = Trim(frmShareSearch.Tag)
        cboCOAAccCode.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
End Sub



Private Sub txtCOAAccCode_GotFocus()
    FocusMe txtCOAAccCode
End Sub

Private Sub txtCOAAccCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtCOAAccCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtCOAAccCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtCOAAccCode_LostFocus()
    FocusMe txtCOAAccCode, True
End Sub

Private Sub txtCOAConsCode_GotFocus()
    FocusMe txtCOAConsCode
End Sub

Private Sub txtCOAConsCode_KeyPress(KeyAscii As Integer)
    Dim iCounter As Integer
    
    Call chk_InpLen(txtCOAConsCode, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        For iCounter = 0 To 5
            If optCOAClass(iCounter).Value = True Then
                optCOAClass(iCounter).SetFocus
                Exit For
            End If
        Next
    End If
End Sub

Private Sub txtCOAConsCode_LostFocus()
    FocusMe txtCOAConsCode, True
End Sub

Private Sub txtCOADesc_GotFocus()
    FocusMe txtCOADesc
End Sub

Private Sub txtCOADesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCOADesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCOACDesc.SetFocus
    End If
End Sub

Private Sub txtCOADesc_LostFocus()
    FocusMe txtCOADesc, True
End Sub

Private Sub txtCOACDesc_GotFocus()
    FocusMe txtCOACDesc
End Sub

Private Sub txtCOACDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCOACDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCOAConsCode.SetFocus
    End If
End Sub

Private Sub txtCOACDesc_LostFocus()
    FocusMe txtCOACDesc, True
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

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT COAStatus FROM MstCOA WHERE COAAccCode = '" & Set_Quote(txtCOAAccCode) & "'"
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
        .TableKey = "COAAccCode"
        .KeyLen = 10
        Set .ctlKey = txtCOAAccCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function Chk_cboCOAAccCode() As Boolean
    Dim wsStatus As String

    Chk_cboCOAAccCode = False
    
    If Trim(cboCOAAccCode.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCOAAccCode.SetFocus
        Exit Function
    End If

    If Chk_COAAccCode(cboCOAAccCode.Text, wsStatus) = False Then
        gsMsg = "COA 編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCOAAccCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "COA 編碼已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCOAAccCode.SetFocus
            Exit Function
        End If
    End If

    Chk_cboCOAAccCode = True
End Function

Private Function Chk_COAAccCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_COAAccCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT COAStatus "
    wsSQL = wsSQL & " FROM MstCOA WHERE MstCOA.COAAccCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "COAStatus")
    
    Chk_COAAccCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_CusCode(ByVal inCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_CusCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT CusStatus "
    wsSQL = wsSQL & " FROM MstCustomer WHERE CusCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_CusCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub SetCOAClass(ByVal inCode As String)
    Select Case inCode
        Case "A"
            optCOAClass(0).Value = True
            
        Case "L"
            optCOAClass(1).Value = True
            
        Case "C"
            optCOAClass(2).Value = True
            
        Case "R"
            optCOAClass(3).Value = True
            
        Case "E"
            optCOAClass(4).Value = True
            
        Case "S"
            optCOAClass(5).Value = True
    End Select
End Sub

Private Function GetCOAClass() As String
    Dim iCounter As Integer
    
    For iCounter = 0 To 5
        If optCOAClass(iCounter).Value = True Then
            Exit For
        End If
    Next
    
    Select Case iCounter
        Case 0
            GetCOAClass = "A"
            
        Case 1
            GetCOAClass = "L"
        
        Case 2
            GetCOAClass = "C"
        
        Case 3
            GetCOAClass = "R"
            
        Case 4
            GetCOAClass = "E"
        
        Case 5
            GetCOAClass = "S"
    End Select
End Function
