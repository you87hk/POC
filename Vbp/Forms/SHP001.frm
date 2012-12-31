VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSHP001 
   BackColor       =   &H8000000A&
   Caption         =   "SHP001"
   ClientHeight    =   5505
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8895
   Icon            =   "SHP001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8895
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "SHP001.frx":08CA
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboShipCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   1320
      Width           =   2730
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   5055
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   8715
      Begin VB.ComboBox cboCardCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   2730
      End
      Begin VB.OptionButton optCard 
         Caption         =   "CARDVDR"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optCard 
         Caption         =   "CARDCUS"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtShipRemark 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   13
         Top             =   3960
         Width           =   6930
      End
      Begin VB.TextBox txtShipper 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   3600
         Width           =   2730
      End
      Begin VB.TextBox txtShipFaxNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5880
         TabIndex        =   11
         Top             =   3240
         Width           =   2730
      End
      Begin VB.TextBox txtShipTelNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   10
         Top             =   3240
         Width           =   2730
      End
      Begin VB.TextBox txtShipCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   2730
      End
      Begin VB.TextBox txtShipName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   2730
      End
      Begin VB.PictureBox picShipAdr 
         BackColor       =   &H80000009&
         Height          =   1455
         Left            =   1680
         ScaleHeight     =   1395
         ScaleWidth      =   6855
         TabIndex        =   24
         Top             =   1680
         Width           =   6915
         Begin VB.TextBox txtShipAdr 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   6735
         End
         Begin VB.TextBox txtShipAdr 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   7
            Top             =   345
            Width           =   6735
         End
         Begin VB.TextBox txtShipAdr 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   3
            Left            =   0
            TabIndex        =   8
            Top             =   690
            Width           =   6735
         End
         Begin VB.TextBox txtShipAdr 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   4
            Left            =   0
            TabIndex        =   9
            Top             =   1035
            Width           =   6735
         End
      End
      Begin VB.Label lblCardCode 
         Caption         =   "CARDCODE"
         Height          =   240
         Left            =   360
         TabIndex        =   29
         Top             =   670
         Width           =   1260
      End
      Begin VB.Label lblShipRemark 
         Caption         =   "SHIPREMARK"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Label lblShipper 
         Caption         =   "SHIPPER"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label lblShipFaxNo 
         Caption         =   "SHIPFAXNO"
         Height          =   255
         Left            =   4560
         TabIndex        =   26
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label lblShipTelNo 
         Caption         =   "SHIPTELNO"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label lblShipLastUpd 
         Caption         =   "SHIPLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   22
         Top             =   4605
         Width           =   1500
      End
      Begin VB.Label lblShipLastUpdDate 
         Caption         =   "SHIPLASTUPDDATE"
         Height          =   240
         Left            =   4320
         TabIndex        =   21
         Top             =   4605
         Width           =   1500
      End
      Begin VB.Label lblDspShipLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1920
         TabIndex        =   20
         Top             =   4560
         Width           =   2265
      End
      Begin VB.Label lblDspShipLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5880
         TabIndex        =   19
         Top             =   4560
         Width           =   2265
      End
      Begin VB.Label lblShipCode 
         Caption         =   "SHIPCODE"
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
         TabIndex        =   18
         Top             =   1020
         Width           =   1260
      End
      Begin VB.Label lblShipAdr 
         Caption         =   "SHIPADR"
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   1740
         Width           =   1380
      End
      Begin VB.Label lblShipName 
         Caption         =   "SHIPNAME"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1380
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
            Picture         =   "SHP001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SHP001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   15
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
Attribute VB_Name = "frmSHP001"
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
Private wlCardID As Long

Private wcCombo As Control

Private wsActNam(4) As String
'Row Lock Variable

Private Const wsKeyType = "MstShip"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private Sub cboCardCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCardCode
    
    If optCard(0).Value = True Then
        wsSQL = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrStatus = '1'"
        wsSQL = wsSQL & " AND VdrInactive = 'N' "
        wsSQL = wsSQL & " AND VdrCode LIKE '%" & IIf(cboCardCode.SelLength > 0, "", Set_Quote(cboCardCode.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY VdrCode "
    ElseIf optCard(1).Value = True Then
        wsSQL = "SELECT CusCode, CusName FROM MstCustomer WHERE CusStatus = '1'"
        wsSQL = wsSQL & " AND CusInactive = 'N' "
        wsSQL = wsSQL & " AND CusCode LIKE '%" & IIf(cboCardCode.SelLength > 0, "", Set_Quote(cboCardCode.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY CusCode "
    End If
    
    Call Ini_Combo(2, wsSQL, cboCardCode.Left, cboCardCode.Top + cboCardCode.Height, tblCommon, "SHP001", "TBLC", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCardCode_GotFocus()
    FocusMe cboCardCode
End Sub

Private Sub cboCardCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCardCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboCardCode() = True Then
            If wiAction = AddRec Then
                txtShipCode.SetFocus
            Else
                cboShipCode.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cboCardCode_LostFocus()
    FocusMe cboCardCode, True
End Sub

Private Sub cboShipCode_LostFocus()
    FocusMe cboShipCode, True
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
            txtShipCode.Enabled = False
            cboShipCode.Enabled = False
        
            txtShipName.Enabled = False
            picShipAdr.Enabled = False
            
            txtShipTelNo.Enabled = False
            txtShipFaxNo.Enabled = False
            txtShipPer.Enabled = False
            txtShipRemark.Enabled = False
            
            Me.cboShipCode.Visible = False
            Me.txtShipCode.Visible = True
            
            optCard(1).Enabled = False
            optCard(0).Enabled = False
            
            cboCardCode.Enabled = False
            
        Case "AfrActAdd"
            cboShipCode.Enabled = False
            cboShipCode.Visible = False
            
            txtShipCode.Enabled = True
            txtShipCode.Visible = True
            
            optCard(1).Enabled = True
            optCard(0).Enabled = True
            
            cboCardCode.Enabled = True
            
        Case "AfrActEdit"
            cboShipCode.Enabled = True
            cboShipCode.Visible = True
            
            txtShipCode.Enabled = False
            txtShipCode.Visible = False
            
            optCard(1).Enabled = True
            optCard(0).Enabled = True
            
            cboCardCode.Enabled = True
            
        Case "AfrKey"
            cboShipCode.Enabled = False
            txtShipCode.Enabled = False
            
            txtShipName.Enabled = True
            picShipAdr.Enabled = True
            
            txtShipTelNo.Enabled = True
            txtShipFaxNo.Enabled = True
            txtShipPer.Enabled = True
            txtShipRemark.Enabled = True
            
            If optCard(0).Value = True Then
                optCard(1).Enabled = False
                optCard(0).Enabled = False
            Else
                optCard(0).Enabled = False
                optCard(1).Enabled = False
            End If
            
            cboCardCode.Enabled = False
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    If Chk_txtShipName = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim iCounter As Integer
    Dim wiCardClass As Integer
    Dim wlCardID As Long
        
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From MstShip "
    wsSQL = wsSQL + "WHERE (((MstShip.ShipCode)='" + Set_Quote(cboShipCode.Text) + "' AND ShipStatus = '1'))"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False

    Else
        cboShipCode = ReadRs(rsRcd, "ShipCode")
        txtShipName = ReadRs(rsRcd, "ShipName")
        
        For iCounter = 1 To 4
            txtShipAdr(iCounter) = ReadRs(rsRcd, "ShipAdr" & iCounter)
        Next iCounter
        
        txtShipTelNo = ReadRs(rsRcd, "ShipTelNo")
        txtShipFaxNo = ReadRs(rsRcd, "ShipFaxNo")
        txtShipPer = ReadRs(rsRcd, "Shipper")
        txtShipRemark = ReadRs(rsRcd, "ShipRemark")
        
        lblDspShipLastUpd = ReadRs(rsRcd, "ShipLastUpd")
        lblDspShipLastUpdDate = ReadRs(rsRcd, "ShipLastUpdDate")
        
        wiCardClass = ReadRs(rsRcd, "ShipCardClass")
        wlCardID = ReadRs(rsRcd, "ShipCardID")
        
        If wiCardClass = 1 Then
            optCard(0).Value = True
            
            wsSQL = "SELECT  VdrCode CardCode "
            wsSQL = wsSQL + "FROM MstVendor "
            wsSQL = wsSQL + "WHERE (MstVendor.VdrID)=" + CStr(wlCardID)
        Else
            optCard(1).Value = True
            
            wsSQL = "SELECT  CusCode CardCode "
            wsSQL = wsSQL + "From MstCustomer "
            wsSQL = wsSQL + "WHERE (MstCustomer.CusID)=" + CStr(wlCardID)
        End If
        
        rsRcd.Close
        Set rsRcd = Nothing
        
        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
        If rsRcd.RecordCount <> 0 Then
            cboCardCode = ReadRs(rsRcd, "CardCode")
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
    Set frmSHP001 = Nothing
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
 '   Me.Left = 0
 '   Me.Top = 0
 '   Me.Width = Screen.Width
 '   Me.Height = Screen.Height
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "SHP001"
    wsTrnCd = ""
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblShipCode.Caption = Get_Caption(waScrItm, "SHIPCODE")
    lblShipName.Caption = Get_Caption(waScrItm, "SHIPNAME")
    lblShipAdr.Caption = Get_Caption(waScrItm, "SHIPADR")
    lblShipTelNo.Caption = Get_Caption(waScrItm, "SHIPTELNO")
    lblShipFaxNo.Caption = Get_Caption(waScrItm, "SHIPFAXNO")
    lblShipPer.Caption = Get_Caption(waScrItm, "SHIPPER")
    lblShipRemark.Caption = Get_Caption(waScrItm, "SHIPREMARK")
    lblShipLastUpd.Caption = Get_Caption(waScrItm, "SHIPLASTUPD")
    lblShipLastUpdDate.Caption = Get_Caption(waScrItm, "SHIPLASTUPDDATE")
    lblCardCode.Caption = Get_Caption(waScrItm, "CARDCODE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    optCard(0).Caption = Get_Caption(waScrItm, "OPTCARD0")
    optCard(1).Caption = Get_Caption(waScrItm, "OPTCARD1")
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "SHPADD")
    wsActNam(2) = Get_Caption(waScrItm, "SHPEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SHPDELETE")
    
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
        cboCardCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCardCode.SetFocus
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboShipCode.SetFocus
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
                If RowLock(wsConnTime, wsKeyType, cboShipCode, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
    End Select
    
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtShipName.SetFocus
End Sub

Private Function Chk_ShipCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_ShipCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT ShipStatus "
    wsSQL = wsSQL & " FROM MstShip WHERE ShipCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "ShipStatus")
    
    Chk_ShipCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_txtShipCode() As Boolean
    Dim wsStatus As String

    Chk_txtShipCode = False
    
    If Trim(cboCardCode.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCardCode.SetFocus
        Exit Function
    End If
    
    Call GetCardID(cboCardCode, wlCardID)
    
    If wlCardID = 0 Then
        gsMsg = "沒有輸入正確之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCardCode.SetFocus
        Exit Function
    End If
    
    If Trim(txtShipCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtShipCode.SetFocus
        Exit Function
    End If

    If Chk_ShipCode(txtShipCode.Text, wsStatus) = True Then
        
        If wsStatus = "2" Then
            gsMsg = "貨運編碼已存在但已無效!"
        Else
            gsMsg = "貨運編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtShipCode.SetFocus
        Exit Function
        
    End If
    
    Chk_txtShipCode = True
End Function

Private Function Chk_cboShipCode() As Boolean
    Dim wsStatus As String
 
    Chk_cboShipCode = False
    
    If Trim(cboShipCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboShipCode.SetFocus
        Exit Function
    End If

    If Chk_ShipCode(cboShipCode.Text, wsStatus) = False Then
                  
        gsMsg = "貨運編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboShipCode.SetFocus
        Exit Function
        
    Else
        
        If wsStatus = "2" Then
        gsMsg = "貨運編碼已存在但已無效!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboShipCode.SetFocus
        Exit Function
        End If
    
    End If
    
    Chk_cboShipCode = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmSHP001
    
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboShipCode, wsFormID) Then
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
    
    GetCardID cboCardCode, wlCardID
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SHP001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlCardID)
    Call SetSPPara(adcmdSave, 3, IIf(optCard(0).Value = True, 1, 2))
    Call SetSPPara(adcmdSave, 4, IIf(wiAction = AddRec, txtShipCode.Text, cboShipCode.Text))
    Call SetSPPara(adcmdSave, 5, txtShipName)
    Call SetSPPara(adcmdSave, 6, txtShipAdr(1))
    Call SetSPPara(adcmdSave, 7, txtShipAdr(2))
    Call SetSPPara(adcmdSave, 8, txtShipAdr(3))
    Call SetSPPara(adcmdSave, 9, txtShipAdr(4))
    Call SetSPPara(adcmdSave, 10, txtShipTelNo)
    Call SetSPPara(adcmdSave, 11, txtShipFaxNo)
    Call SetSPPara(adcmdSave, 12, txtShipPer)
    Call SetSPPara(adcmdSave, 13, txtShipRemark)
    Call SetSPPara(adcmdSave, 14, gsUserID)
    Call SetSPPara(adcmdSave, 15, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 16)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - SHP001!"
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
    vFilterAry(1, 1) = "貨運編碼"
    vFilterAry(1, 2) = "ShipCode"
    
    vFilterAry(2, 1) = "名稱"
    vFilterAry(2, 2) = "ShipName"
    
    ReDim vAry(2, 3)
    vAry(1, 1) = "貨運編碼"
    vAry(1, 2) = "ShipCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "名稱"
    vAry(2, 2) = "ShipName"
    vAry(2, 3) = "5000"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT MstShip.ShipCode, MstShip.ShipName "
        wsSQL = wsSQL + "FROM MstShip "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE MstShip.ShipStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstShip.ShipCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboShipCode Then
        cboShipCode = Trim(frmShareSearch.Tag)
        cboShipCode.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
    
End Sub

Private Sub txtShipAdr_GotFocus(Index As Integer)
    FocusMe txtShipAdr(Index)
End Sub

Private Sub txtShipAdr_KeyPress(Index As Integer, KeyAscii As Integer)
    Call chk_InpLen(txtShipAdr(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Index = 4 Then
            txtShipTelNo.SetFocus
        Else
            txtShipAdr(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtShipAdr_LostFocus(Index As Integer)
    FocusMe txtShipAdr(Index), True
End Sub

Private Sub txtShipCode_GotFocus()
    FocusMe txtShipCode
End Sub

Private Sub txtShipCode_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtShipCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtShipCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtShipCode_LostFocus()
    FocusMe txtShipCode, True
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

Private Sub cboShipCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboShipCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboShipCode() = True Then
            Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub cboShipCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboShipCode
    
    
    If optCard(0).Value = True Then
    
    wsSQL = "SELECT ShipCode, VdrCode, ShipName  FROM MstShip, MstVendor"
    wsSQL = wsSQL & " WHERE ShipStatus = '1' "
    wsSQL = wsSQL & " AND ShipCardClass = 1 "
    wsSQL = wsSQL & " AND ShipCardID = VdrID "
    wsSQL = wsSQL & " AND VdrCode LIKE '%" & IIf(cboCardCode.SelLength > 0, "", Set_Quote(cboCardCode.Text)) & "%' "
    wsSQL = wsSQL & " AND ShipCode LIKE '%" & IIf(cboShipCode.SelLength > 0, "", Set_Quote(cboShipCode.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ShipCode "
    
    Else
    
    wsSQL = "SELECT ShipCode, CusCode, ShipName  FROM MstShip, MstCustomer"
    wsSQL = wsSQL & " WHERE ShipStatus = '1' "
    wsSQL = wsSQL & " AND ShipCardClass = 2 "
    wsSQL = wsSQL & " AND ShipCardID = CusID "
    wsSQL = wsSQL & " AND CusCode LIKE '%" & IIf(cboCardCode.SelLength > 0, "", Set_Quote(cboCardCode.Text)) & "%' "
    wsSQL = wsSQL & " AND ShipCode LIKE '%" & IIf(cboShipCode.SelLength > 0, "", Set_Quote(cboShipCode.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ShipCode "
    
    End If
    
    Call Ini_Combo(3, wsSQL, cboShipCode.Left, cboShipCode.Top + cboShipCode.Height, tblCommon, "SHP001", "TBLSHP", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboShipCode_GotFocus()
    FocusMe cboShipCode
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT ShipStatus FROM MstShip WHERE ShipCode = '" & Set_Quote(txtShipCode) & "'"
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
        .TableKey = "ShipCode"
        .KeyLen = 10
        Set .ctlKey = txtShipCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub txtShipName_GotFocus()
    FocusMe txtShipName
End Sub

Private Sub txtShipName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtShipName() = True Then
            txtShipAdr(1).SetFocus
        End If
    End If
End Sub

Private Sub txtShipName_LostFocus()
    FocusMe txtShipName, True
End Sub

Private Function Chk_txtShipName() As Boolean
    Chk_txtShipName = False
    
    If Len(Trim(txtShipName)) <= 0 Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtShipName.SetFocus
        Exit Function
    End If
    
    Chk_txtShipName = True
End Function

Private Sub txtShipTelNo_GotFocus()
    FocusMe txtShipTelNo
End Sub

Private Sub txtShipTelNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipTelNo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtShipFaxNo.SetFocus
    End If
End Sub

Private Sub txtShipTelNo_LostFocus()
    FocusMe txtShipTelNo, True
End Sub

Private Sub txtShipFaxNo_GotFocus()
    FocusMe txtShipFaxNo
End Sub

Private Sub txtShipFaxNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipFaxNo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtShipPer.SetFocus
    End If
End Sub

Private Sub txtShipFaxNo_LostFocus()
    FocusMe txtShipFaxNo, True
End Sub

Private Sub txtShipPer_GotFocus()
    FocusMe txtShipPer
End Sub

Private Sub txtShipPer_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipPer, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtShipRemark.SetFocus
    End If
End Sub

Private Sub txtShipPer_LostFocus()
    FocusMe txtShipPer, True
End Sub

Private Sub txtShipRemark_GotFocus()
    FocusMe txtShipRemark
End Sub

Private Sub txtShipRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipRemark, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtShipName.SetFocus
    End If
End Sub

Private Sub txtShipRemark_LostFocus()
    FocusMe txtShipRemark, True
End Sub

Private Function Chk_cboCardCode() As Boolean
    Dim wsStatus As String

    Chk_cboCardCode = False
    
        If Trim(cboCardCode.Text) = "" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCardCode.SetFocus
            Exit Function
        End If
    
        If optCard(0).Value = True Then
            If Chk_VdrCode(cboCardCode) = False Then
                gsMsg = "商戶不存在!"
                
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                cboCardCode.SetFocus
                Exit Function
            End If
        ElseIf optCard(1).Value = True Then
            If Chk_CusCode(cboCardCode) = False Then
                gsMsg = "商戶不存在!"
                
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                cboCardCode.SetFocus
                Exit Function
            End If
        End If
    
    Chk_cboCardCode = True
End Function

Private Function Chk_VdrCode(ByVal inCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_VdrCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT VdrStatus "
    wsSQL = wsSQL & " FROM MstVendor WHERE VdrCode = '" & Set_Quote(inCode) & "'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_VdrCode = True
    
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

Private Sub GetCardID(ByVal inCode As String, outCode As Long)
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    If Trim(inCode) = "" Then
        Exit Sub
    End If
    
    If optCard(0).Value = True Then
        wsSQL = "SELECT VdrID CardID"
        wsSQL = wsSQL & " FROM MstVendor WHERE VdrCode = '" & Set_Quote(inCode) & "'"
    ElseIf optCard(1).Value = True Then
        wsSQL = "SELECT CusID CardID"
        wsSQL = wsSQL & " FROM MstCustomer WHERE CusCode = '" & Set_Quote(inCode) & "'"
    End If
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Sub
    End If
    
    outCode = ReadRs(rsRcd, "CardID")
    
    rsRcd.Close
    Set rsRcd = Nothing
End Sub
