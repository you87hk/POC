VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmAT001 
   BackColor       =   &H8000000A&
   Caption         =   "會計碼"
   ClientHeight    =   3915
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "AT001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3915
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "AT001.frx":08CA
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboAccTypeSALML 
      Height          =   300
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   2730
   End
   Begin VB.ComboBox cboAccTypeINVML 
      Height          =   300
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   2730
   End
   Begin VB.ComboBox cboAccTypeCOSML 
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   2730
   End
   Begin VB.ComboBox cboAccTypeCode 
      Height          =   300
      Left            =   2880
      TabIndex        =   0
      Top             =   960
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "會計版別"
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtAccTypeCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   2730
      End
      Begin VB.TextBox txtAccTypeDesc 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label lblAccTypeSALML 
         Caption         =   "ACCTYPEINVML"
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   2120
         Width           =   2220
      End
      Begin VB.Label lblAccTypeINVML 
         Caption         =   "ACCTYPEINVML"
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   1755
         Width           =   2220
      End
      Begin VB.Label lblAccTypeCOSML 
         Caption         =   "ACCTYPECOSML"
         Height          =   240
         Left            =   360
         TabIndex        =   15
         Top             =   1395
         Width           =   2220
      End
      Begin VB.Label lblAccTypeLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   2925
         Width           =   2220
      End
      Begin VB.Label lblAccTypeLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4320
         TabIndex        =   12
         Top             =   2925
         Width           =   2100
      End
      Begin VB.Label lblDspAccTypeLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2760
         TabIndex        =   11
         Top             =   2880
         Width           =   1425
      End
      Begin VB.Label lblDspAccTypeLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6600
         TabIndex        =   10
         Top             =   2880
         Width           =   1545
      End
      Begin VB.Label lblAccTypeCode 
         Caption         =   "會計編碼 :"
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
         TabIndex        =   9
         Top             =   660
         Width           =   2220
      End
      Begin VB.Label lblAccTypeDesc 
         Caption         =   "會計註解 :"
         Height          =   240
         Left            =   360
         TabIndex        =   8
         Top             =   1020
         Width           =   2220
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
            Picture         =   "AT001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AT001.frx":6945
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
Attribute VB_Name = "frmAT001"
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

Private Const wsKeyType = "MstAccountType"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboAccTypeCode_LostFocus()
    FocusMe cboAccTypeCode, True
End Sub

Private Sub cboAccTypeCOSML_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboAccTypeCOSML
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboAccTypeCOSML.SelLength > 0, "", Set_Quote(cboAccTypeCOSML.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboAccTypeCOSML.Left, cboAccTypeCOSML.Top + cboAccTypeCOSML.Height, tblCommon, "AT001", "TBLACCTYPECOSML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboAccTypeCOSML_GotFocus()
    FocusMe cboAccTypeCOSML
End Sub

Private Sub cboAccTypeCOSML_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccTypeCOSML, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboAccTypeCOSML() = True Then
            cboAccTypeINVML.SetFocus
        End If
        
    End If
End Sub

Private Sub cboAccTypeCOSML_LostFocus()
    FocusMe cboAccTypeCOSML, True
End Sub

Private Sub cboAccTypeINVML_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboAccTypeINVML
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboAccTypeINVML.SelLength > 0, "", Set_Quote(cboAccTypeINVML.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboAccTypeINVML.Left, cboAccTypeINVML.Top + cboAccTypeINVML.Height, tblCommon, "AT001", "TBLACCTYPEINVML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboAccTypeINVML_GotFocus()
    FocusMe cboAccTypeINVML
End Sub

Private Sub cboAccTypeINVML_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccTypeINVML, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboAccTypeINVML() = True Then
            cboAccTypeSALML.SetFocus
        End If
        
    End If
End Sub

Private Sub cboAccTypeINVML_LostFocus()
    FocusMe cboAccTypeINVML, True
End Sub

Private Sub cboAccTypeSALML_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboAccTypeSALML
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & " AND MLCode LIKE '%" & IIf(cboAccTypeSALML.SelLength > 0, "", Set_Quote(cboAccTypeSALML.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboAccTypeSALML.Left, cboAccTypeSALML.Top + cboAccTypeSALML.Height, tblCommon, "AT001", "TBLACCTYPESALML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboAccTypeSALML_GotFocus()
    FocusMe cboAccTypeSALML
End Sub

Private Sub cboAccTypeSALML_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccTypeSALML, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboAccTypeSALML() = True Then
            txtAccTypeDesc.SetFocus
        End If
    End If
End Sub

Private Sub cboAccTypeSALML_LostFocus()
    FocusMe cboAccTypeSALML, True
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
        Me.Height = 4320
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
            Me.txtAccTypeDesc.Enabled = False
            Me.cboAccTypeCOSML.Enabled = False
            Me.cboAccTypeINVML.Enabled = False
            Me.cboAccTypeSALML.Enabled = False
            
            Me.cboAccTypeCode.Enabled = False
            Me.cboAccTypeCode.Visible = False
            Me.txtAccTypeCode.Visible = True
            Me.txtAccTypeCode.Enabled = False
            
        Case "AfrActAdd"
            Me.cboAccTypeCode.Enabled = False
            Me.cboAccTypeCode.Visible = False
            
            Me.txtAccTypeCode.Enabled = True
            Me.txtAccTypeCode.Visible = True
            
        Case "AfrActEdit"
            Me.cboAccTypeCode.Enabled = True
            Me.cboAccTypeCode.Visible = True
            
            Me.txtAccTypeCode.Enabled = False
            Me.txtAccTypeCode.Visible = False
            
        Case "AfrKey"
            Me.cboAccTypeCode.Enabled = False
            Me.txtAccTypeCode.Enabled = False
            
            Me.cboAccTypeCOSML.Enabled = True
            Me.cboAccTypeINVML.Enabled = True
            Me.cboAccTypeSALML.Enabled = True
            Me.txtAccTypeDesc.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtAccTypeDesc = False Then
        Exit Function
    End If
    
    If Chk_cboAccTypeCOSML = False Then
        Exit Function
    End If
    
    If Chk_cboAccTypeINVML = False Then
        Exit Function
    End If
    
    If Chk_cboAccTypeSALML = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From MstAccountType "
    wsSQL = wsSQL + "WHERE (((MstAccountType.AccTypeCode)='" + Set_Quote(cboAccTypeCode.Text) + "' AND AccTypeStatus = '1'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False

    Else
        Me.cboAccTypeCode = ReadRs(rsRcd, "AccTypeCode")
        Me.txtAccTypeDesc = ReadRs(rsRcd, "AccTypeDesc")
        Me.cboAccTypeCOSML = ReadRs(rsRcd, "AccTypeCOSML")
        Me.cboAccTypeINVML = ReadRs(rsRcd, "AccTypeINVML")
        Me.cboAccTypeSALML = ReadRs(rsRcd, "AccTypeSALML")
        Me.lblDspAccTypeLastUpd = ReadRs(rsRcd, "AccTypeLastUpd")
        Me.lblDspAccTypeLastUpdDate = ReadRs(rsRcd, "AccTypeLastUpdDate")
        
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
    Set frmAT001 = Nothing
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
    wsFormID = "AT001"
    wsTrnCd = ""
    
End Sub


Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblAccTypeCode.Caption = Get_Caption(waScrItm, "ACCTYPECODE")
    lblAccTypeDesc.Caption = Get_Caption(waScrItm, "ACCTYPEDESC")
    lblAccTypeCOSML.Caption = Get_Caption(waScrItm, "ACCTYPECOSML")
    lblAccTypeINVML.Caption = Get_Caption(waScrItm, "ACCTYPEINVML")
    lblAccTypeSALML.Caption = Get_Caption(waScrItm, "ACCTYPESALML")
    lblAccTypeLastUpd.Caption = Get_Caption(waScrItm, "ACCTYPELASTUPD")
    lblAccTypeLastUpdDate.Caption = Get_Caption(waScrItm, "ACCTYPELASTUPDDATE")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")

    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"

    wsActNam(1) = Get_Caption(waScrItm, "ATADD")
    wsActNam(2) = Get_Caption(waScrItm, "ATEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "ATDELETE")
    
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
        txtAccTypeCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboAccTypeCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboAccTypeCode.SetFocus
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
                If RowLock(wsConnTime, wsKeyType, cboAccTypeCode, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtAccTypeDesc.SetFocus
End Sub

Private Function Chk_AccTypeCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_AccTypeCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT AccTypeStatus "
    wsSQL = wsSQL & " FROM MstAccountType WHERE AccTypeCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "AccTypeStatus")
    
    Chk_AccTypeCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtAccTypeCode() As Boolean
    Dim wsStatus As String

    Chk_txtAccTypeCode = False
    
        If Trim(txtAccTypeCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtAccTypeCode.SetFocus
            Exit Function
        End If
    
        If Chk_AccTypeCode(txtAccTypeCode.Text, wsStatus) = True Then
            
            If wsStatus = "2" Then
            gsMsg = "會計版別已存在但已無效!"
            Else
            gsMsg = "會計版別已存在!"
            End If
            
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtAccTypeCode.SetFocus
            Exit Function
            
        End If
    
    Chk_txtAccTypeCode = True
End Function

Private Function Chk_cboAccTypeCode() As Boolean
    Dim wsStatus As String
 
    Chk_cboAccTypeCode = False
    
        If Trim(cboAccTypeCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboAccTypeCode.SetFocus
            Exit Function
        End If
    
        If Chk_AccTypeCode(cboAccTypeCode.Text, wsStatus) = False Then
                      
            gsMsg = "會計版別不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboAccTypeCode.SetFocus
            Exit Function
            
        Else
            
            If wsStatus = "2" Then
            gsMsg = "會計版別已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboAccTypeCode.SetFocus
            Exit Function
            End If
        
        End If
    
    Chk_cboAccTypeCode = True
    
End Function

Private Sub cmdOpen()
    Dim newForm As New frmAT001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboAccTypeCode, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_AT001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, IIf(wiAction = AddRec, txtAccTypeCode.Text, cboAccTypeCode.Text))
    Call SetSPPara(adcmdSave, 3, txtAccTypeDesc)
    Call SetSPPara(adcmdSave, 4, cboAccTypeCOSML)
    Call SetSPPara(adcmdSave, 5, cboAccTypeINVML)
    Call SetSPPara(adcmdSave, 6, cboAccTypeSALML)
    Call SetSPPara(adcmdSave, 7, gsUserID)
    Call SetSPPara(adcmdSave, 8, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 9)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - AT001!"
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
    
    ReDim vFilterAry(5, 2)
    vFilterAry(1, 1) = "會計版別編碼"
    vFilterAry(1, 2) = "AccTypeCode"
    
    vFilterAry(2, 1) = "註解"
    vFilterAry(2, 2) = "AccTypeDesc"
    
    vFilterAry(3, 1) = "銷售成本會計級別"
    vFilterAry(3, 2) = "AccTypeCOSML"
    
    vFilterAry(4, 1) = "存貨會計級別"
    vFilterAry(4, 2) = "AccTypeINVML"
    
    vFilterAry(5, 1) = "銷售會計級別"
    vFilterAry(5, 2) = "AccTypeSALML"
    
    ReDim vAry(5, 3)
    vAry(1, 1) = "會計版別編碼"
    vAry(1, 2) = "AccTypeCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "註解"
    vAry(2, 2) = "AccTypeDesc"
    vAry(2, 3) = "4000"
    
    vAry(3, 1) = "銷售成本會計級別"
    vAry(3, 2) = "AccTypeCOSML"
    vAry(3, 3) = "1500"
    
    vAry(4, 1) = "存貨會計級別"
    vAry(4, 2) = "AccTypeINVML"
    vAry(4, 3) = "1500"
    
    vAry(5, 1) = "銷售會計級別"
    vAry(5, 2) = "AccTypeSALML"
    vAry(5, 3) = "1500"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT MstAccountType.AccTypeCode, MstAccountType.AccTypeDesc "
        wsSQL = wsSQL + "FROM MstAccountType "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE MstAccountType.AccTypeStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstAccountType.AccTypeCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboAccTypeCode Then
        cboAccTypeCode = Trim(frmShareSearch.Tag)
        cboAccTypeCode.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
    
End Sub

Private Sub txtAccTypeCode_GotFocus()
    FocusMe txtAccTypeCode
End Sub

Private Sub txtAccTypeCode_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLenA(txtAccTypeCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        If Chk_txtAccTypeCode() = True Then
            Call Ini_Scr_AfrKey
        End If
        
        
    End If
End Sub

Private Sub txtAccTypeCode_LostFocus()
    FocusMe txtAccTypeCode, True
End Sub

Private Sub txtAccTypeDesc_LostFocus()
    FocusMe txtAccTypeDesc, True
End Sub

Private Sub txtAccTypeDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtAccTypeDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtAccTypeDesc() = True Then
            cboAccTypeCOSML.SetFocus
        End If
    End If
End Sub

Private Sub txtAccTypeDesc_GotFocus()
    FocusMe txtAccTypeDesc
End Sub

Private Function Chk_txtAccTypeDesc() As Boolean
    
    Chk_txtAccTypeDesc = False
    
    If Trim(txtAccTypeDesc.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtAccTypeDesc.SetFocus
        Exit Function
    End If
    
    Chk_txtAccTypeDesc = True
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

Private Sub cboAccTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboAccTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboAccTypeCode() = True Then
            Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub cboAccTypeCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboAccTypeCode
    
    wsSQL = "SELECT AccTypeCode, AccTypeDesc FROM MstAccountType WHERE AccTypeStatus = '1'"
    wsSQL = wsSQL & " AND AccTypeCode LIKE '%" & IIf(cboAccTypeCode.SelLength > 0, "", Set_Quote(cboAccTypeCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY AccTypeCode "
    Call Ini_Combo(2, wsSQL, cboAccTypeCode.Left, cboAccTypeCode.Top + cboAccTypeCode.Height, tblCommon, "AT001", "TBLACCTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboAccTypeCode_GotFocus()
    FocusMe cboAccTypeCode
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT AccTypeStatus FROM MstAccountType WHERE AccTypeCode = '" & Set_Quote(txtAccTypeCode) & "'"
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
        .TableKey = "AccTypeCode"
        .KeyLen = 10
        Set .ctlKey = txtAccTypeCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Function Chk_cboAccTypeCOSML() As Boolean
    Dim wsStatus As String
 
    Chk_cboAccTypeCOSML = False
    
    If Trim(cboAccTypeCOSML.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeCOSML.SetFocus
        Exit Function
    End If

    If Chk_AccTypeML(cboAccTypeCOSML.Text, wsStatus) = False Then
                  
        gsMsg = "銷售成本會計級別不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeCOSML.SetFocus
        Exit Function
        
    Else
        
        If wsStatus = "2" Then
        gsMsg = "銷售成本會計級別已存在但已無效!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeCOSML.SetFocus
        Exit Function
        End If
    
    End If
    
    Chk_cboAccTypeCOSML = True
    
End Function

Private Function Chk_cboAccTypeINVML() As Boolean
    Dim wsStatus As String
 
    Chk_cboAccTypeINVML = False
    
    If Trim(cboAccTypeINVML.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeINVML.SetFocus
        Exit Function
    End If

    If Chk_AccTypeML(cboAccTypeINVML.Text, wsStatus) = False Then
        gsMsg = "存貨會計級別不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeINVML.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "存貨會計級別已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboAccTypeINVML.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboAccTypeINVML = True
End Function

Private Function Chk_cboAccTypeSALML() As Boolean
    Dim wsStatus As String
 
    Chk_cboAccTypeSALML = False
    
    If Trim(cboAccTypeSALML.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeSALML.SetFocus
        Exit Function
    End If

    If Chk_AccTypeML(cboAccTypeSALML.Text, wsStatus) = False Then
        gsMsg = "銷售會計級別不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboAccTypeSALML.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "銷售會計級別已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboAccTypeSALML.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboAccTypeSALML = True
End Function


Private Function Chk_AccTypeML(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_AccTypeML = False
    
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
    
    Chk_AccTypeML = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

