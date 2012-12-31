VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSD001 
   BackColor       =   &H8000000A&
   Caption         =   "銷售折扣"
   ClientHeight    =   3825
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "frmSD001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "frmSD001.frx":08CA
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   8355
      Begin VB.TextBox txtSDID 
         Height          =   270
         Left            =   7800
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSDDiscount 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         TabIndex        =   3
         Top             =   2205
         Width           =   855
      End
      Begin VB.ComboBox cboSDCDISCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   1485
         Width           =   1170
      End
      Begin VB.ComboBox cboSDNatureCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   1005
         Width           =   1170
      End
      Begin VB.ComboBox cboSDMethodCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1170
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   8055
         Begin VB.Label lblDspSDCDisDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   2640
            TabIndex        =   20
            Top             =   1240
            Width           =   5265
         End
         Begin VB.Label lblDspSDNatureDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   2640
            TabIndex        =   19
            Top             =   760
            Width           =   5265
         End
         Begin VB.Label lblDspSDMethodDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   2640
            TabIndex        =   18
            Top             =   240
            Width           =   5265
         End
         Begin VB.Label lblSDMethodCode 
            Caption         =   "SDMETHODCODE"
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
            TabIndex        =   16
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label lblSDNatureCode 
            Caption         =   "SDNATURECODE"
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
            TabIndex        =   15
            Top             =   840
            Width           =   1380
         End
         Begin VB.Label lblSDCDisCode 
            Caption         =   "SDCDISCODE"
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
            TabIndex        =   14
            Top             =   1300
            Width           =   1380
         End
      End
      Begin VB.Label lblPercent 
         Caption         =   "%"
         Height          =   240
         Left            =   2760
         TabIndex        =   21
         Top             =   2280
         Width           =   300
      End
      Begin VB.Label lblArrow 
         Caption         =   "=>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2205
         Width           =   375
      End
      Begin VB.Label lblSDDiscount 
         Caption         =   "SDDISCOUNT"
         Height          =   240
         Left            =   600
         TabIndex        =   11
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label lblSaleDiscountLastUpd 
         Caption         =   "SALEDISCOUNTLASTUPD"
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   2925
         Width           =   1140
      End
      Begin VB.Label lblSaleDiscountLastUpdDate 
         Caption         =   "SALEDISCOUNTLASTUPDDATE"
         Height          =   240
         Left            =   4200
         TabIndex        =   7
         Top             =   2925
         Width           =   1260
      End
      Begin VB.Label lblDspSaleDiscountLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         Top             =   2880
         Width           =   2505
      End
      Begin VB.Label lblDspSaleDiscountLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5520
         TabIndex        =   5
         Top             =   2880
         Width           =   2505
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
            Picture         =   "frmSD001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSD001.frx":6945
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
Attribute VB_Name = "frmSD001"
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
Private wlKey As Long
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control

Private Const wsKeyType = "MstSaleDiscount"
Private wsUsrId As String
Private wsTrnCd As String

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
        
            If tbrProcess.Buttons(tcFind).Enabled = True Then Call cmdFind
            
        Case vbKeyF10
        
            If tbrProcess.Buttons(tcSave).Enabled = True Then Call cmdSave
            
            
            
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
    
    IniForm
    Ini_Caption
    Ini_Scr
    
    MousePointer = vbDefault
  
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 4230
        Me.Width = 8700
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
            Me.cboSDMethodCode.Enabled = False
            Me.cboSDNatureCode.Enabled = False
            Me.cboSDCDISCode.Enabled = False
            Me.txtSDDiscount.Enabled = False
            
        Case "AfrActAdd"
            Me.cboSDMethodCode.Enabled = True
            Me.cboSDNatureCode.Enabled = True
            Me.cboSDCDISCode.Enabled = True
            Me.txtSDDiscount.Enabled = False
            
        Case "AfrActEdit"
            Me.cboSDMethodCode.Enabled = True
            Me.cboSDNatureCode.Enabled = True
            Me.cboSDCDISCode.Enabled = True
            Me.txtSDDiscount.Enabled = False
            
        Case "AfrKey"
            Me.cboSDMethodCode.Enabled = False
            Me.cboSDNatureCode.Enabled = False
            Me.cboSDCDISCode.Enabled = False
            Me.txtSDDiscount.Enabled = True
            
    End Select
End Sub

Private Function Chk_KeyFld() As Boolean
    
        
    Chk_KeyFld = False
    
    If Chk_cboSDMethodCode() = False Then
        Exit Function
    End If
    
    If Chk_cboSDNatureCode() = False Then
        Exit Function
    End If
    
    If Chk_cboSDCDISCode() = False Then
        Exit Function
    End If
    
    If Chk_SaleDiscount(cboSDNatureCode, cboSDMethodCode, cboSDCDISCode) = False Then
    
    If wiAction = CorRec Then
        gsMsg = "銷售折扣不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDCDISCode.SetFocus
        Exit Function
    End If
    
    Else
    
    If wiAction = AddRec Then
        gsMsg = "銷售折扣已存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDCDISCode.SetFocus
        Exit Function
    End If
    
    End If
    
    Chk_KeyFld = True
    
End Function


'-- Input validation checking.
Private Function InputValidation() As Boolean
    Dim sMsg As String
        
    InputValidation = False
    
    If Chk_txtSDDiscount() = False Then
        Exit Function
    End If
    
    
    InputValidation = True
End Function


Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstSaleDiscount.* "
    wsSql = wsSql + "From MstSaleDiscount "
    wsSql = wsSql + "WHERE (((MstSaleDiscount.SDMethodCode)='" + Set_Quote(cboSDMethodCode) + "') "
    wsSql = wsSql + "AND ((MstSaleDiscount.SDNatureCode)='" + Set_Quote(cboSDNatureCode) + "') "
    wsSql = wsSql + "AND ((MstSaleDiscount.SDCDisCode)='" + Set_Quote(cboSDCDISCode) + "') "
    wsSql = wsSql + "AND ((SaleDiscountStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "SDID")
        Me.txtSDDiscount = Format(To_Value(ReadRs(rsRcd, "SDDiscount")), gsAmtFmt)
        Me.lblDspSaleDiscountLastUpd = ReadRs(rsRcd, "SaleDiscountLastUpd")
        Me.lblDspSaleDiscountLastUpdDate = ReadRs(rsRcd, "SaleDiscountLastUpdDate")
        
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
    Set frmSD001 = Nothing
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
 '   Me.Left = 0
 '   Me.Top = 0
 '   Me.Width = Screen.Width
 '   Me.Height = Screen.Height
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "SD001"
    wsTrnCd = ""
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblSDNatureCode.Caption = Get_Caption(waScrItm, "SDNATURECODE")
    lblSDMethodCode.Caption = Get_Caption(waScrItm, "SDMETHODCODE")
    lblSDCDisCode.Caption = Get_Caption(waScrItm, "SDCDISCODE")
    lblSDDiscount.Caption = Get_Caption(waScrItm, "SDDISCOUNT")
    lblPercent.Caption = Get_Caption(waScrItm, "PERCENT")
    lblSaleDiscountLastUpd.Caption = Get_Caption(waScrItm, "SALEDISCOUNTLASTUPD")
    lblSaleDiscountLastUpdDate.Caption = Get_Caption(waScrItm, "SALEDISCOUNTLASTUPDDATE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
   
    wsActNam(1) = Get_Caption(waScrItm, "SDADD")
    wsActNam(2) = Get_Caption(waScrItm, "SDEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SDDELETE")
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
    wlKey = 0
    
    Me.txtSDDiscount = Format(0, gsAmtFmt)

    
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
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
       
    
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
    End Select
    
    cboSDMethodCode.SetFocus
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Dim Ctrl As Control
    
    Select Case wiAction
    
    Case CorRec, DelRec

        If LoadRecord() = False Then
            gsMsg = "沒有要存取之折扣!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Sub
        Else
            If RowLock(wsConnTime, wsKeyType, txtSDID, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtSDDiscount.SetFocus
End Sub

Private Function Chk_SDNatureCode(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_SDNatureCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT NatureCode, NatureDesc "
    wsSql = wsSql & " FROM MstNature WHERE NatureCode = '" & Set_Quote(inCode) & "' And NatureStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        OutDesc = ""
        Exit Function
    Else
        OutDesc = ReadRs(rsRcd, "NatureDesc")
    End If
    
    Chk_SDNatureCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_SDCDisCode(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_SDCDisCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT CDisCode, CDisDesc "
    wsSql = wsSql & " FROM MstCategoryDiscount WHERE CDisCode = '" & Set_Quote(inCode) & "' And CDisStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        OutDesc = ""
        Exit Function
    Else
        OutDesc = ReadRs(rsRcd, "CDisDesc")
    End If
    
    Chk_SDCDisCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_SDMethodCode(ByVal inCode As String, ByRef OutDesc As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_SDMethodCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT MethodCode, MethodDesc "
    wsSql = wsSql & " FROM MstMethod WHERE MethodCode = '" & Set_Quote(inCode) & "' And MethodStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        OutDesc = ""
        Exit Function
    Else
        OutDesc = ReadRs(rsRcd, "MethodDesc")
    End If
    
    Chk_SDMethodCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Public Function Chk_SaleDiscount(inNatureCode, InMethodCode, inCDISCode) As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    Chk_SaleDiscount = False
    
    wsSql = "SELECT MstSaleDiscount.* "
    wsSql = wsSql + "From MstSaleDiscount "
    wsSql = wsSql + "WHERE (((MstSaleDiscount.SDMethodCode)='" + Set_Quote(InMethodCode) + "') "
    wsSql = wsSql + "AND ((MstSaleDiscount.SDNatureCode)='" + Set_Quote(inNatureCode) + "') "
    wsSql = wsSql + "AND ((MstSaleDiscount.SDCDISCode)='" + Set_Quote(inCDISCode) + "') "
    wsSql = wsSql + "AND ((SaleDiscountStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount <= 0 Then
        
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    Chk_SaleDiscount = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Function Chk_txtSDDiscount() As Boolean
    Chk_txtSDDiscount = False
    
    If txtSDDiscount < 0 Or txtSDDiscount > 100 Then
        gsMsg = "折扣範圍錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtSDDiscount.SetFocus
        Exit Function
    End If
    
    Chk_txtSDDiscount = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmSD001
    
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
Private Sub cmdFind()
   Call OpenPromptForm
End Sub
Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim wsNo As String
    Dim adcmdSave As New ADODB.Command
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, txtSDID, wsFormID) Then
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
    

    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SD001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, cboSDNatureCode)
    Call SetSPPara(adcmdSave, 4, cboSDMethodCode)
    Call SetSPPara(adcmdSave, 5, cboSDCDISCode)
    Call SetSPPara(adcmdSave, 6, txtSDDiscount)
    Call SetSPPara(adcmdSave, 7, gsUserID)
    Call SetSPPara(adcmdSave, 8, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 9)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - SD001!"
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
    Dim sTmpSQL As String
    
    ReDim vFilterAry(7, 2)
    vFilterAry(1, 1) = "客戶銷售渠道編碼"
    vFilterAry(1, 2) = "MethodCode"
    
    vFilterAry(2, 1) = "渠道註解"
    vFilterAry(2, 2) = "MethodDesc"
    
    vFilterAry(3, 1) = "客戶性質編碼"
    vFilterAry(3, 2) = "SDNatureCode"
    
    vFilterAry(4, 1) = "性質註解"
    vFilterAry(4, 2) = "NatureDesc"
    
    vFilterAry(5, 1) = "圖書折扣分類編碼"
    vFilterAry(5, 2) = "CDISCode"
    
    vFilterAry(6, 1) = "分類註解"
    vFilterAry(6, 2) = "CDisDesc"
    
    vFilterAry(7, 1) = "折扣"
    vFilterAry(7, 2) = "SDDiscount"
    
    ReDim vAry(8, 3)
    vAry(1, 1) = ""
    vAry(1, 2) = "SDID"
    vAry(1, 3) = "0"
    
    vAry(2, 1) = "銷售渠道編碼"
    vAry(2, 2) = "SDMethodCode"
    vAry(2, 3) = "1300"
    
    vAry(3, 1) = "渠道註解"
    vAry(3, 2) = "MethodDesc"
    vAry(3, 3) = "1300"
    
    vAry(4, 1) = "客戶性質編碼"
    vAry(4, 2) = "SDNatureCode"
    vAry(4, 3) = "1300"
    
    vAry(5, 1) = "性質註解"
    vAry(5, 2) = "NatureDesc"
    vAry(5, 3) = "1300"
    
    vAry(6, 1) = "圖書折扣分類編碼"
    vAry(6, 2) = "SDCDISCode"
    vAry(6, 3) = "1300"
    
    vAry(7, 1) = "分類註解"
    vAry(7, 2) = "CDisDesc"
    vAry(7, 3) = "1300"
    
    vAry(8, 1) = "折扣"
    vAry(8, 2) = "SaleDiscount"
    vAry(8, 3) = "800"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstSaleDiscount.SDID, MstSaleDiscount.SDNatureCode, MstNature.NatureDesc, MstSaleDiscount.SDMethodCode, MstMethod.MethodDesc, MstSaleDiscount.SDCDisCode, MstCategoryDiscount.CDisDesc, MstSaleDiscount.SDDiscount "
        sSQL = sSQL + "FROM MstSaleDiscount, MstCategoryDiscount, MstNature, MstMethod "
        .sBindSQL = sSQL
        sTmpSQL = "WHERE MstSaleDiscount.SaleDiscountStatus = '1' "
        sTmpSQL = sTmpSQL + "AND MstNature.NatureCode = MstSaleDiscount.SDNatureCode "
        sTmpSQL = sTmpSQL + "AND MstMethod.MethodCode = MstSaleDiscount.SDMethodCode "
        sTmpSQL = sTmpSQL + "AND MstCategoryDiscount.CDisCode = MstSaleDiscount.SDCDisCode "
        .sBindWhereSQL = sTmpSQL
        .sBindOrderSQL = "ORDER BY MstSaleDiscount.SDNatureCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> txtSDID Then
        txtSDID = Trim(frmShareSearch.Tag)
        cboSDCDISCode.SetFocus
        SendKeys "{ENTER}"
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

Private Sub cboSDMethodCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboSDMethodCode, 10, KeyAscii)
    
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboSDMethodCode() = False Then
            Exit Sub
        End If
        
        cboSDNatureCode.SetFocus
     
    End If
End Sub

Private Sub cboSDMethodCode_DropDown()
    
    Dim wsSql As String
  
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSDMethodCode
    
    wsSql = "SELECT MethodCode, MethodDesc FROM MstMethod WHERE MethodStatus = '1'"
    wsSql = wsSql & " AND MethodCode LIKE '%" & IIf(cboSDMethodCode.SelLength > 0, "", Set_Quote(cboSDMethodCode.Text)) & "%' "
   
    wsSql = wsSql & "ORDER BY MethodCode "
    Call Ini_Combo(2, wsSql, cboSDMethodCode.Left, cboSDMethodCode.Top + cboSDMethodCode.Height, tblCommon, "SD001", "TBLM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboSDMethodCode_GotFocus()
    FocusMe cboSDMethodCode
End Sub

Private Sub cboSDMethodCode_LostFocus()
    FocusMe cboSDMethodCode, True
End Sub

Private Sub cboSDNatureCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboSDNatureCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboSDNatureCode() = False Then
            Exit Sub
        End If
        
        cboSDCDISCode.SetFocus
    End If
End Sub

Private Function Chk_cboSDNatureCode() As Boolean
    Dim wsDesc As String
    
    Chk_cboSDNatureCode = False
    
    If Trim(cboSDNatureCode) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDNatureCode.SetFocus
        Exit Function
    End If

    If Chk_SDNatureCode(cboSDNatureCode.Text, wsDesc) = False Then
        gsMsg = "客戶性質編碼不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDNatureCode.SetFocus
        Exit Function
    End If

    lblDspSDNatureDesc = wsDesc
    
    Chk_cboSDNatureCode = True
End Function

Private Function Chk_cboSDCDISCode() As Boolean
    Dim wsDesc As String
    
    Chk_cboSDCDISCode = False
    
    If Trim(cboSDCDISCode) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDCDISCode.SetFocus
        Exit Function
    End If

    If Chk_SDCDisCode(cboSDCDISCode.Text, wsDesc) = False Then
        gsMsg = "圖書分類編碼不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDCDISCode.SetFocus
        Exit Function
    End If
    
    lblDspSDCDisDesc = wsDesc
    
    Chk_cboSDCDISCode = True
End Function

Private Function Chk_cboSDMethodCode() As Boolean
    Dim wsDesc As String

    Chk_cboSDMethodCode = False
    
    If Trim(cboSDMethodCode) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDMethodCode.SetFocus
        Exit Function
    End If

    If Chk_SDMethodCode(cboSDMethodCode.Text, wsDesc) = False Then
        gsMsg = "銷售渠道編碼不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSDMethodCode.SetFocus
        Exit Function
    End If

    lblDspSDMethodDesc = wsDesc
    
    Chk_cboSDMethodCode = True
End Function

Private Sub cboSDNatureCode_DropDown()
    
    Dim wsSql As String
 
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSDNatureCode
    
    wsSql = "SELECT NatureCode, NatureDesc FROM MstNature WHERE NatureStatus = '1'"
    wsSql = wsSql & " AND NatureCode LIKE '%" & IIf(cboSDNatureCode.SelLength > 0, "", Set_Quote(cboSDNatureCode.Text)) & "%' "
    wsSql = wsSql & "ORDER BY NatureCode "
    Call Ini_Combo(2, wsSql, cboSDNatureCode.Left, cboSDNatureCode.Top + cboSDNatureCode.Height, tblCommon, "SD001", "TBLN", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboSDNatureCode_GotFocus()
    FocusMe cboSDNatureCode
End Sub

Private Sub cboSDNatureCode_LostFocus()
    FocusMe cboSDNatureCode, True
End Sub

Private Sub cboSDCDISCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboSDCDISCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
      
    If Chk_cboSDCDISCode() = False Then
        Exit Sub
    End If
      
          
    If Chk_KeyFld = True Then
           Call Ini_Scr_AfrKey
     End If
     
    End If
End Sub

Private Sub cboSDCDISCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboSDCDISCode
    
    wsSql = "SELECT CDisCode, CDisDesc FROM MstCategoryDiscount WHERE CDisStatus = '1'"
    wsSql = wsSql & " AND CDisCode LIKE '%" & IIf(cboSDCDISCode.SelLength > 0, "", Set_Quote(cboSDCDISCode.Text)) & "%' "
   
    wsSql = wsSql & "ORDER BY CDisCode "
    Call Ini_Combo(2, wsSql, cboSDCDISCode.Left, cboSDCDISCode.Top + cboSDCDISCode.Height, tblCommon, "SD001", "TBLCDIS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboSDCDISCode_GotFocus()
    FocusMe cboSDCDISCode
End Sub

Private Sub cboSDCDISCode_LostFocus()
    FocusMe cboSDCDISCode, True
End Sub

Private Sub txtSDDiscount_GotFocus()
    FocusMe txtSDDiscount
End Sub

Private Sub txtSDDiscount_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtSDDiscount, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
      
    If Chk_txtSDDiscount() = False Then
        Exit Sub
    End If
      
    End If
End Sub

Private Sub txtSDDiscount_LostFocus()
FocusMe txtSDDiscount, True
End Sub

Private Sub txtSDID_Change()
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstSaleDiscount.* "
    wsSql = wsSql + "From MstSaleDiscount "
    wsSql = wsSql + "WHERE (((MstSaleDiscount.SDID)='" + Set_Quote(txtSDID) + "') "
    wsSql = wsSql + "AND ((SaleDiscountStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount > 0 Then
        Me.cboSDMethodCode = ReadRs(rsRcd, "SDMethodCode")
        Me.cboSDNatureCode = ReadRs(rsRcd, "SDNatureCode")
        Me.cboSDCDISCode = ReadRs(rsRcd, "SDCDISCode")
    End If
    
    rsRcd.Close
    
    Set rsRcd = Nothing
End Sub
