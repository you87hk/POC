VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAPR001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAPR001.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9360
      OleObjectBlob   =   "frmAPR001.frx":0442
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5280
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5280
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   11775
      Begin VB.Frame fraSelect 
         Height          =   525
         Left            =   7320
         TabIndex        =   19
         Top             =   120
         Width           =   3975
         Begin VB.OptionButton optDocType 
            Caption         =   "SO"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   20
            Top             =   200
            Width           =   1095
         End
         Begin VB.OptionButton optDocType 
            Caption         =   "SO"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   7
            Top             =   200
            Width           =   1095
         End
         Begin VB.OptionButton optDocType 
            Caption         =   "SN"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   200
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin MSMask.MaskEdBox medPrdTo 
         Height          =   285
         Left            =   5280
         TabIndex        =   5
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "####/##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPrdFr 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   7
         Mask            =   "####/##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCusNoFr 
         Caption         =   "Customer Code From"
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblPrdFr 
         Caption         =   "Period From"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   990
         Width           =   1890
      End
      Begin VB.Label lblCusNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   14
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblPrdTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   13
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label lblDocNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   12
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   1890
      End
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6375
      Left            =   120
      OleObjectBlob   =   "frmAPR001.frx":2B45
      TabIndex        =   8
      Top             =   1800
      Width           =   11775
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   11400
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":B018
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":B8F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":C1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":C61E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":CA70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":CD8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":D1DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":D62E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":D948
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":DC62
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":E0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":E990
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":ECB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":F10C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":F428
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":F744
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":FB98
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":FEB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":101D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":104F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":10948
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":10C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPR001.frx":10F7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Convert"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Can"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Finish"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Complete"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Lock"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   120
      TabIndex        =   17
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmAPR001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
'Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Private wbErr As Boolean

Private wiSort As Integer
Private wsSortBy As String

Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wiActFlg As Integer
Private wsMark As String
Private wsTrnCd As String
Private wsTitle As String

Private Const tcConvert = "Convert"
Private Const tcCan = "Can"
Private Const tcCopy = "Copy"
Private Const tcFinish = "Finish"
Private Const tcComplete = "Complete"
Private Const tcPrint = "Print"
Private Const tcLock = "Lock"
Private Const tcExport = "Export"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"


Private Const SSEL = 0
Private Const SDOCDATE = 1
Private Const SDOCNO = 2
Private Const SCUSCODE = 3
Private Const SCUSNAME = 4
Private Const SFROMDATE = 5
Private Const STODATE = 6
Private Const SQTY = 7
Private Const SNET = 8
Private Const SORI = 9
Private Const SREADY = 10
Private Const SID = 11



Private Sub cboCusNoFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    wsSQL = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND CusStatus <> '2' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSQL, cboCusNoFr.Left, cboCusNoFr.Top + cboCusNoFr.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoFr_GotFocus()
        FocusMe cboCusNoFr
    Set wcCombo = cboCusNoFr
End Sub

Private Sub cboCusNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoFr, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboCusNoFr.Text) <> "" And _
            Trim(cboCusNoTo.Text) = "" Then
            cboCusNoTo.Text = cboCusNoFr.Text
        End If
        cboCusNoTo.SetFocus
    End If
End Sub


Private Sub cboCusNoFr_LostFocus()
    FocusMe cboCusNoFr, True
End Sub

Private Sub cboCusNoTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    wsSQL = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND CusStatus <> '2' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & " ORDER BY Cuscode "
    Call Ini_Combo(2, wsSQL, cboCusNoTo.Left, cboCusNoTo.Top + cboCusNoTo.Height, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusNoTo_GotFocus()
    FocusMe cboCusNoTo
    Set wcCombo = cboCusNoTo
End Sub

Private Sub cboCusNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusNoTo = False Then
            Exit Sub
        End If
        
        medPrdFr.SetFocus
    End If
End Sub



Private Sub cboCusNoTo_LostFocus()
FocusMe cboCusNoTo, True
End Sub




Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
End Sub

Private Sub medPrdFr_GotFocus()
    FocusMe medPrdFr
End Sub


Private Sub medPrdFr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_medPrdFr = False Then
            Exit Sub
        End If
        
        If Trim(medPrdFr) <> "/" And _
            Trim(medPrdTo) = "/" Then
            medPrdTo.Text = medPrdFr.Text
        End If
        medPrdTo.SetFocus
    End If
End Sub

Private Sub medPrdFr_LostFocus()
    FocusMe medPrdFr, True
End Sub

Private Sub medPrdTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_medPrdTo = False Then
            Exit Sub
        End If
        
        If LoadRecord = True Then
            tblDetail.SetFocus
        End If
       
    End If
End Sub

Private Sub medPrdTo_GotFocus()
    FocusMe medPrdTo
End Sub
Private Sub medPrdTo_LostFocus()
    FocusMe medPrdTo, True
End Sub

Private Sub cboDocNoFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoFr
  
    Select Case wsTrnCd
    Case "SN"
    
    wsSQL = "SELECT SNHDDOCNO, CUSCODE, SNHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSNHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SNHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SNHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SNHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SNHDDOCNO "
    
    Case "SO"
    
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
    
    Case "SP"
    
    wsSQL = "SELECT SPHDDOCNO, CUSCODE, SPHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSPHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SPHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SPHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SPHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SPHDDOCNO "
    
    Case "SW"
    
    wsSQL = "SELECT SWHDDOCNO, CUSCODE, SWHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSWHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SWHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SWHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SWHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SWHDDOCNO "
    
    Case "SD"
    
    wsSQL = "SELECT SDHDDOCNO, CUSCODE, SDHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSDHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SDHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SDHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SDHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SDHDDOCNO "
    
    Case "IV"
    
    wsSQL = "SELECT IVHDDOCNO, CUSCODE, IVHDDOCDATE "
    wsSQL = wsSQL & " FROM soaIVHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND IVHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO "
        
    Case "VQ"
    
    wsSQL = "SELECT VQHDDOCNO, CUSCODE, VQHDDOCDATE "
    wsSQL = wsSQL & " FROM soaVQHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE VQHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND VQHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND VQHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY VQHDDOCNO "
    
    
    End Select
    Call Ini_Combo(3, wsSQL, cboDocNoFr.Left, cboDocNoFr.Top + cboDocNoFr.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoFr_GotFocus()
    FocusMe cboDocNoFr
    Set wcCombo = cboDocNoFr
End Sub

Private Sub cboDocNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboDocNoFr.Text) <> "" And _
            Trim(cboDocNoTo.Text) = "" Then
            cboDocNoTo.Text = cboDocNoFr.Text
        End If
        cboDocNoTo.SetFocus
    End If
End Sub

Private Sub cboDocNoFr_LostFocus()
    FocusMe cboDocNoFr, True
End Sub

Private Sub cboDocNoTo_DropDown()
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNoTo
  
    Select Case wsTrnCd
    Case "SN"
    
    wsSQL = "SELECT SNHDDOCNO, CUSCODE, SNHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSNHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SNHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SNHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SNHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SNHDDOCNO "
    
    Case "SO"
    
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
    
    Case "SP"
    
    wsSQL = "SELECT SPHDDOCNO, CUSCODE, SPHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSPHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SPHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SPHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SPHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SPHDDOCNO "
    
    Case "SW"
    
    wsSQL = "SELECT SWHDDOCNO, CUSCODE, SWHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSWHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SWHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SWHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SWHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SWHDDOCNO "
    
    Case "SD"
    
    wsSQL = "SELECT SDHDDOCNO, CUSCODE, SDHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSDHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SDHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SDHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SDHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SDHDDOCNO "
    
    Case "IV"
    
    wsSQL = "SELECT IVHDDOCNO, CUSCODE, IVHDDOCDATE "
    wsSQL = wsSQL & " FROM soaIVHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND IVHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO "
    
   Case "VQ"
    
    wsSQL = "SELECT VQHDDOCNO, CUSCODE, VQHDDOCDATE "
    wsSQL = wsSQL & " FROM soaVQHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE VQHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND VQHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND VQHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY VQHDDOCNO "
    
    
    End Select
    
    Call Ini_Combo(3, wsSQL, cboDocNoTo.Left, cboDocNoTo.Top + cboDocNoTo.Height, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNoTo_GotFocus()
    FocusMe cboDocNoTo
    Set wcCombo = cboDocNoTo
End Sub

Private Sub cboDocNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboDocNoTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboDocNoTo = False Then
            Call cboDocNoTo_GotFocus
            Exit Sub
        End If
        
       cboCusNoFr.SetFocus
        
        
    End If
End Sub

Private Sub cboDocNoTo_LostFocus()
    FocusMe cboDocNoTo, True
End Sub
Private Function chk_cboDocNoTo() As Boolean
    chk_cboDocNoTo = False
    
    If UCase(cboDocNoFr.Text) > UCase(cboDocNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboDocNoTo = True
End Function

Private Function chk_cboCusNoTo() As Boolean
    chk_cboCusNoTo = False
    
    If UCase(cboCusNoFr.Text) > UCase(cboCusNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboCusNoTo = True
End Function
Private Function chk_medPrdFr() As Boolean
    chk_medPrdFr = False
    
    If Trim(medPrdFr) = "/" Then
        chk_medPrdFr = True
        Exit Function
    End If
    
    If Chk_Period(medPrdFr) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdFr.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdFr = True
End Function

Private Function chk_medPrdTo() As Boolean
    chk_medPrdTo = False
    
    If UCase(medPrdFr.Text) > UCase(medPrdTo.Text) Then
        gsMsg = "To must > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    End If
    
    If Trim(medPrdTo) = "/" Then
        chk_medPrdTo = True
        Exit Function
    End If

    If Chk_Period(medPrdTo) = False Then
        gsMsg = "Wrong Period!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medPrdTo.SetFocus
        Exit Function
    
    End If
    
    chk_medPrdTo = True
End Function
Private Sub Chk_Sel(inRow As Long)
    
    Dim wlCtr As Long
     
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If waResult(wlCtr, SSEL) = "-1" Then
                  waResult(wlCtr, SSEL) = "0"
                  Exit Sub
               End If
            End If
        Next

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF2
        If tbrProcess.Buttons(tcConvert).Enabled = False Then Exit Sub
           Call cmdSave(2)
        
        Case vbKeyF3
        If tbrProcess.Buttons(tcCan).Enabled = False Then Exit Sub
           Call cmdSave(3)
           
        Case vbKeyF8
        If tbrProcess.Buttons(tcCopy).Enabled = False Then Exit Sub
           Call cmdSave(4)
              
        Case vbKeyF10
        If tbrProcess.Buttons(tcFinish).Enabled = False Then Exit Sub
            Call cmdSave(5)
        
        Case vbKeyF9
        If tbrProcess.Buttons(tcComplete).Enabled = False Then Exit Sub
            Call cmdSave(8)
            
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
             
        Case vbKeyF5
           Call cmdSelect(1)
           
        Case vbKeyF6
           Call cmdSelect(0)
        
        Case vbKeyF7
            Call LoadRecord
            
    End Select
    
    
                
       
End Sub



Private Sub optDocType_Click(Index As Integer)
    Call cmdRefresh
End Sub

Private Sub optDocType_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call cmdRefresh
        tblDetail.SetFocus
        
    End If
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    
   If tbrProcess.Buttons(Button.Key).Enabled = False Then Exit Sub
        
    
    Select Case Button.Key
        Case tcConvert
            Call cmdSave(2)
            
        Case tcCan
            Call cmdSave(3)
            
        Case tcCopy
           ' Call cmdSave(4)
           Call cmdCopy(4)
                
        Case tcFinish
            Call cmdSave(5)
            
        Case tcExport
            Call cmdExport
            
        
        Case tcComplete
            Call cmdSave(8)
            
        Case tcPrint
            Call cmdPrint
            
        Case tcLock
            Call cmdReady
            
        Case tcCancel
        
           Call cmdCancel
            
        
        Case tcSAll
        
           Call cmdSelect(1)
        
        Case tcDAll
        
           Call cmdSelect(0)
           
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call cmdRefresh
            
            
    End Select
End Sub

Private Sub Form_Load()
    
    
  MousePointer = vbHourglass
  
    IniForm
    Ini_Caption
    Ini_Grid
    Ini_Scr

    
   MousePointer = vbDefault
    
    
End Sub

Private Sub cmdCancel()
    
    
  MousePointer = vbHourglass
  
    Ini_Scr
    
   MousePointer = vbDefault
    
    
End Sub



Private Sub cmdRefresh()
    
    
  MousePointer = vbHourglass
  
    If wsSortBy = "ASC" Then
    wsSortBy = "DESC"
    Else
    wsSortBy = "ASC"
    End If
    
  
    Call Set_tbrProcess
    Call LoadRecord
    
    
   MousePointer = vbDefault
    
    
End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, SSEL, SID
    
    
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
   '         Case "ComboBox"
   '             MyControl.Clear
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

    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    wiExit = False
    
    Call SetPeriodMask(medPrdFr)
    Call SetPeriodMask(medPrdTo)
    
    
    medPrdFr.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))
    medPrdTo.Text = Dsp_PeriodDate(Left(gsSystemDate, 7))
    
    cboDocNoFr.Text = ""
    cboDocNoTo.Text = ""
    cboCusNoFr.Text = ""
    cboCusNoTo.Text = ""
    
    wiSort = 0
    wsSortBy = "ASC"

     Call cmdRefresh
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
    Set waScrItm = Nothing
 '   Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmAPR001 = Nothing
 
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    optDocType(0).Value = True
    

   
End Sub


Private Sub Set_tbrProcess()

With tbrProcess
    
    Select Case wsFormID
    
    Case "APR001"
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    Case 1
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    Case 2
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    End Select
    
    ''Tom 20090203
    .Buttons(tcExport).Enabled = False
    .Buttons(tcLock).Enabled = False
    .Buttons(tcPrint).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
        
    Case "APR002"
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = True
    .Buttons(tcComplete).Enabled = True
    ''Tom 20090203
    .Buttons(tcLock).Enabled = True
    Case 1
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    ''Tom 20090203
    .Buttons(tcLock).Enabled = False
    
    Case 2
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
''Tom 20090203
    .Buttons(tcLock).Enabled = False
        
    End Select
    
    .Buttons(tcExport).Enabled = True
    .Buttons(tcPrint).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
     Case "APR003"
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = True
    .Buttons(tcComplete).Enabled = False
    Case 1
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    Case 2
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    End Select
    
    ''Tom 20090203
    .Buttons(tcExport).Enabled = False
    .Buttons(tcLock).Enabled = False
    .Buttons(tcPrint).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
     Case "APR004"
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
    .Buttons(tcConvert).Enabled = True
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    Case 1
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    Case 2
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    End Select
    
    ''Tom 20090203
    .Buttons(tcExport).Enabled = False
    .Buttons(tcLock).Enabled = False
    .Buttons(tcPrint).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
     Case "APR005"
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
     Case 1

    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    Case 2
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    End Select
     
    ''Tom 20090203
    .Buttons(tcExport).Enabled = False
    .Buttons(tcLock).Enabled = False
    .Buttons(tcPrint).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
    Case "APR006"
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = True
    .Buttons(tcComplete).Enabled = False
    .Buttons(tcPrint).Enabled = True
     Case 1

    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    .Buttons(tcPrint).Enabled = True
    Case 2
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = True
    .Buttons(tcCopy).Enabled = True
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    .Buttons(tcPrint).Enabled = False
    End Select
     
    ''Tom 20090203
    .Buttons(tcExport).Enabled = False
    .Buttons(tcLock).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
    Case "APR007"
    
    .Buttons(tcConvert).Enabled = False
    .Buttons(tcCan).Enabled = False
    .Buttons(tcCopy).Enabled = False
    .Buttons(tcFinish).Enabled = False
    .Buttons(tcComplete).Enabled = False
    .Buttons(tcPrint).Enabled = False
     
    ''Tom 20090203
    .Buttons(tcExport).Enabled = False
    .Buttons(tcLock).Enabled = False
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
    End Select
    

    
End With

End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
  '  Call Get_Scr_Item("TOOLTIP_A", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblPrdFr.Caption = Get_Caption(waScrItm, "PRDFR")
    lblPrdTo.Caption = Get_Caption(waScrItm, "PRDTO")
    optDocType(0).Caption = Get_Caption(waScrItm, "OPT1")
    optDocType(1).Caption = Get_Caption(waScrItm, "OPT2")
    optDocType(2).Caption = Get_Caption(waScrItm, "OPT3")

    
                
    wsTitle = Get_Caption(waScrItm, "TITLE")
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SCUSCODE).Caption = Get_Caption(waScrItm, "SCUSCODE")
        .Columns(SCUSNAME).Caption = Get_Caption(waScrItm, "SCUSNAME")
        .Columns(SDOCDATE).Caption = Get_Caption(waScrItm, "SDOCDATE")
        .Columns(SFROMDATE).Caption = Get_Caption(waScrItm, "SFROMDATE")
        .Columns(STODATE).Caption = Get_Caption(waScrItm, "STODATE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SNET).Caption = Get_Caption(waScrItm, "SNET")
        .Columns(SORI).Caption = Get_Caption(waScrItm, "SORI")
        'Tom 20090203
        If wsTrnCd = "SO" Then
            .Columns(SREADY).Caption = "Lock"
        End If
        
    End With
    
    With tbrProcess
    .Buttons(tcConvert).ToolTipText = Get_Caption(waScrItm, tcConvert) & "(F2)"
    .Buttons(tcCan).ToolTipText = Get_Caption(waScrItm, tcCan) & "(F3)"
    .Buttons(tcCopy).ToolTipText = Get_Caption(waScrItm, tcCopy) & "(F8)"
    .Buttons(tcFinish).ToolTipText = Get_Caption(waScrItm, tcFinish) & "(F10)"
    .Buttons(tcComplete).ToolTipText = Get_Caption(waScrItm, tcComplete) & "(F9)"
    .Buttons(tcPrint).ToolTipText = Get_Caption(waScrItm, tcPrint)
    .Buttons(tcLock).ToolTipText = "Ready / Unlock"
    .Buttons(tcExport).ToolTipText = Get_Caption(waScrItm, tcExport)
    
    
    .Buttons(tcRefresh).ToolTipText = Get_Caption(waScrItm, tcRefresh) & "(F7)"
    .Buttons(tcCancel).ToolTipText = Get_Caption(waScrItm, tcCancel) & "(F11)"
    .Buttons(tcSAll).ToolTipText = Get_Caption(waScrItm, tcSAll) & "(F5)"
    .Buttons(tcDAll).ToolTipText = Get_Caption(waScrItm, tcDAll) & "(F6)"
    .Buttons(tcExit).ToolTipText = Get_Caption(waScrItm, tcExit) & "(F12)"
   End With

End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .UPDATE
    End With

   If ColIndex = SSEL Then
   
 '   tblDetail.ReBind
 '   tblDetail.Bookmark = 0
         
   End If
   
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsBookID As String
Dim wsBookCode As String
Dim wsBarCode As String
Dim wsBookName As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double
Dim wsLotNo As String


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case SSEL
            
           '   If .Columns(ColIndex).Text = "-1" Then
           '       Call Chk_Sel(.Row + To_Value(.FirstRow))
           '    End If
                
               ' If Chk_grdSoNo(.Columns(ColIndex).Text) = False Then
               '    GoTo Tbl_BeforeColUpdate_Err
               ' End If
                
            End Select
    End With
    
    Exit Sub
    
Tbl_BeforeColUpdate_Err:
    tblDetail.Columns(ColIndex).Text = OldValue
    Cancel = True
    Exit Sub

tblDetail_BeforeColUpdate_Err:
    
    MsgBox "Check tblDeiail BeforeColUpdate!"
    tblDetail.Columns(ColIndex).Text = OldValue
    Cancel = True
End Sub



Private Sub tblDetail_ButtonClick(ByVal ColIndex As Integer)

    
    On Error GoTo tblDetail_ButtonClick_Err
    
    
    With tblDetail
        Select Case ColIndex
            Case SDOCNO
                
                 If .Columns(SDOCNO).Text <> "" Then
                    
                    
                  Select Case wsFormID
                    Case "APR001", "APR002", "APR003", "APR006"
                    
                    frmAPR0011.InDocID = .Columns(SID).Text
                    frmAPR0011.InCusNo = .Columns(SCUSCODE).Text
                    frmAPR0011.TrnCd = wsTrnCd
                    frmAPR0011.FormID = wsFormID & "1"
                    frmAPR0011.UpdFlg = False
                    frmAPR0011.Show vbModal
                    
                    Case "APR004", "APR005", "APR007"
                    
                    frmAPR0012.InDocID = .Columns(SID).Text
                    frmAPR0012.InCusNo = .Columns(SCUSCODE).Text
                    frmAPR0012.TrnCd = wsTrnCd
                    frmAPR0012.FormID = wsFormID & "1"
                    frmAPR0012.Show vbModal
                  
                  End Select
  
                    
                    
                    
                 End If
        
            Case SDOCDATE
            
            If .Columns(SDOCNO).Text <> "" And wsFormID = "APR002" Then
            
                    frmSOLOG.InDocID = .Columns(SID).Text
                    frmSOLOG.InCusNo = .Columns(SCUSCODE).Text
                    frmSOLOG.TrnCd = wsTrnCd
                    frmSOLOG.FormID = "SOLOG"
                    frmSOLOG.Show vbModal
                    
            End If
            
                
           End Select
    End With
    
    Exit Sub
    
tblDetail_ButtonClick_Err:
     MsgBox "Check tblDeiail ButtonClick!"
 
End Sub

Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode
        Case vbKeyF4        ' CALL COMBO BOX
            KeyCode = vbDefault
            Call tblDetail_ButtonClick(.Col)
            
        Case vbKeyReturn
            Select Case .Col
            Case SORI
                 KeyCode = vbKeyDown
                 .Col = SSEL
            Case Else
                 KeyCode = vbDefault
                 .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SSEL Then
                .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SORI
                    KeyCode = vbKeyDown
                    .Col = SSEL
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
                
            End Select
        
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub






Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        
        
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                 
                Case SCUSNAME
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = .Columns(SCUSNAME).Text
                    
                Case SFROMDATE
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = .Columns(SFROMDATE).Text
                  
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
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


Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = False
        .AllowUpdate = True
        .AllowDelete = False
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = SSEL To SID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SSEL
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).Locked = False
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 25
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                Case SCUSCODE
                   .Columns(wiCtr).Width = 800
                   .Columns(wiCtr).DataWidth = 10
                Case SCUSNAME
                   .Columns(wiCtr).Width = 1500
                   .Columns(wiCtr).DataWidth = 50
                Case SDOCDATE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Button = True
                Case SFROMDATE
                    .Columns(wiCtr).Width = 2500
                    .Columns(wiCtr).DataWidth = 10
                Case STODATE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Visible = False
                Case SQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case SNET
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SORI
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = False
                Case SREADY
                'Tom 20090203
                    If wsTrnCd = "SO" Then
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).ValueItems.Presentation = dbgCheckBox
                    Else
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
                    End If
                Case SID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 15
                End Select
                
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wdCreLmt As Double
    Dim wdCreLft As Double
    Dim wsStatus As String
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    
    Select Case Opt_Getfocus(optDocType, 3, 0)
    Case 0
     wsStatus = "1"
    Case 1
     wsStatus = "4"
    Case 2
     wsStatus = "2"
    End Select
    
    Select Case wsTrnCd
    Case "SN"
    
    wsSQL = "SELECT CUSNAME, SNHDDOCID DOCID, SNHDDOCNO DOCNO, SNHDDOCDATE DOCDATE, SNHDSHIPFROM +' '+ SNHDSHIPTO +' '+ SNHDSHIPVIA JOBREF, SNHDCUSID, CUSCODE, SNHDREFNO REFNO, SUM(SNPTJQTY) QTY, "
    wsSQL = wsSQL & " SUM(SNPTJNET) NET "
    wsSQL = wsSQL & " FROM  SOASNHD, SOASNPTJ, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SNHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SNHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND SNHDDOCID = SNPTJDOCID "
    wsSQL = wsSQL & " AND SNHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SNHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " GROUP BY CUSNAME, SNHDDOCID, SNHDDOCNO, SNHDDOCDATE, SNHDSHIPFROM, SNHDSHIPTO, SNHDSHIPVIA, SNHDCUSID, CUSCODE, SNHDREFNO "
    'wsSQL = wsSQL & " ORDER BY SNHDDOCDATE, SNHDDOCNO "
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY SNHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY SNHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY SNHDREFNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY SNHDDOCDATE, SNHDDOCNO " & wsSortBy
    End If
    
    Case "SO"
    
    wsSQL = "SELECT CUSNAME, SOHDDOCID DOCID, SOHDDOCNO DOCNO, SOHDDOCDATE DOCDATE, SOHDSHIPFROM JOBREF, SOHDCUSID, CUSCODE, SOHDCORRNO REFNO, SOHDREADY RDY, SUM(SOPTJQTY) QTY, "
    wsSQL = wsSQL & " SUM(SOPTJNET) NET "
    wsSQL = wsSQL & " FROM  SOASOHD, SOASOPTJ, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SOHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND SOHDDOCID = SOPTJDOCID "
    wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " GROUP BY CUSNAME, SOHDDOCID, SOHDDOCNO, SOHDDOCDATE, SOHDSHIPFROM, SOHDCUSID, CUSCODE, SOHDCORRNO, SOHDREADY "
    'wsSQL = wsSQL & " ORDER BY SOHDDOCDATE, SOHDDOCNO "
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY SOHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY SOHDCORRNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY SOHDDOCDATE, SOHDDOCNO " & wsSortBy
    End If
    
    Case "SP"
    
    wsSQL = "SELECT CUSNAME, SPHDDOCID DOCID, SPHDDOCNO DOCNO, SPHDDOCDATE DOCDATE, SPHDSHIPFROM JOBREF, SPHDCUSID, CUSCODE, SPHDREFNO REFNO, SUM(SPDTQTY) QTY, "
    wsSQL = wsSQL & " SUM(SPDTMERCST) NET "
    wsSQL = wsSQL & " FROM  SOASPHD, SOASPDT, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SPHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SPHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND SPHDDOCID = SPDTDOCID "
    wsSQL = wsSQL & " AND SPHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SPHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " GROUP BY CUSNAME, SPHDDOCID, SPHDDOCNO, SPHDDOCDATE, SPHDSHIPFROM,  SPHDCUSID, CUSCODE, SPHDREFNO "
    'wsSQL = wsSQL & " ORDER BY SPHDDOCDATE, SPHDDOCNO "
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY SPHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY SPHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY SPHDREFNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY SPHDDOCDATE, SPHDDOCNO " & wsSortBy
    End If
    
    Case "SW"
    
    wsSQL = "SELECT CUSNAME, SWHDDOCID DOCID, SWHDDOCNO DOCNO, SWHDDOCDATE DOCDATE, SWHDSHIPFROM JOBREF, SWHDCUSID, CUSCODE, SWHDREFNO REFNO, SUM(SWDTQTY) QTY, "
    wsSQL = wsSQL & " SUM(SWDTMERCST) NET "
    wsSQL = wsSQL & " FROM  SOASWHD, SOASWDT, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SWHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SWHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND SWHDDOCID = SWDTDOCID "
    wsSQL = wsSQL & " AND SWHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SWHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " AND SWHDUPDFLG = 'N' "
    wsSQL = wsSQL & " GROUP BY CUSNAME, SWHDDOCID, SWHDDOCNO, SWHDDOCDATE, SWHDSHIPFROM, SWHDCUSID, CUSCODE, SWHDREFNO "
    
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY SWHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY SWHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY SWHDREFNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY SWHDDOCDATE, SWHDDOCNO " & wsSortBy
    End If
    
    Case "SD"
    
    wsSQL = "SELECT CUSNAME, SDHDDOCID DOCID, SDHDDOCNO DOCNO, SDHDDOCDATE DOCDATE, SDHDSHIPFROM JOBREF, SDHDCUSID, CUSCODE, SDHDREFNO REFNO, SUM(SDPTJQTY) QTY, "
    wsSQL = wsSQL & " SUM(SDPTJNET) NET "
    wsSQL = wsSQL & " FROM  SOASDHD, SOASDPTJ, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE SDHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SDHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND SDHDDOCID = SDPTJDOCID "
    wsSQL = wsSQL & " AND SDHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SDHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " GROUP BY CUSNAME, SDHDDOCID, SDHDDOCNO, SDHDDOCDATE, SDHDSHIPFROM, SDHDCUSID, CUSCODE, SDHDREFNO "
    'wsSQL = wsSQL & " ORDER BY SDHDDOCDATE, SDHDDOCNO "
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY SDHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY SDHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY SDHDREFNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY SDHDDOCDATE, SDHDDOCNO " & wsSortBy
    End If
    
    
     Case "IV"
    
    wsSQL = "SELECT CUSNAME, IVHDDOCID DOCID,  LEFT(IVHDDOCNO,4) + '/' + IVHDREFNO + '/' + RIGHT(IVHDDOCNO,LEN(IVHDDOCNO)-4) DOCNO, IVHDDOCDATE DOCDATE, IVHDSHIPFROM JOBREF, IVHDCUSID, CUSCODE, IVHDCORRNO REFNO, SUM(IVDTQTY) QTY, "
    wsSQL = wsSQL & " SUM(IVDTNET) NET "
    wsSQL = wsSQL & " FROM  SOAIVHD, SOAIVDT, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND IVHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND IVHDDOCID = IVDTDOCID "
    wsSQL = wsSQL & " AND IVHDCUSID = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " GROUP BY CUSNAME, IVHDDOCID, IVHDDOCNO, IVHDDOCDATE, IVHDSHIPFROM, IVHDCUSID, CUSCODE, IVHDCORRNO, IVHDREFNO "
   ' wsSQL = wsSQL & " ORDER BY IVHDDOCDATE "
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY IVHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY IVHDCORRNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY IVHDDOCDATE, IVHDDOCNO " & wsSortBy
    End If
    
    
    
    Case "VQ"
    
    wsSQL = "SELECT CUSNAME, VQHDDOCID DOCID, VQHDDOCNO DOCNO, VQHDDOCDATE DOCDATE, VQHDSHIPFROM JOBREF, VQHDCUSID, CUSCODE, SOHDDOCNO REFNO, VQHDREADY RDY, SUM(VQPTJQTY) QTY, "
    wsSQL = wsSQL & " SUM(VQPTJNET) NET "
    wsSQL = wsSQL & " FROM  SOAVQHD, SOAVQPTJ, MSTCUSTOMER, SOASOHD "
    wsSQL = wsSQL & " WHERE VQHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND VQHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND VQHDDOCID = VQPTJDOCID "
    wsSQL = wsSQL & " AND VQHDCUSID = CUSID "
    wsSQL = wsSQL & " AND VQHDREFDOCID = SOHDDOCID "
    wsSQL = wsSQL & " AND VQHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " GROUP BY CUSNAME, VQHDDOCID, VQHDDOCNO, VQHDDOCDATE, VQHDSHIPFROM, VQHDCUSID, CUSCODE, SOHDDOCNO, VQHDREADY "
    'wsSQL = wsSQL & " ORDER BY VQHDDOCDATE, VQHDDOCNO "
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY VQHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY VQHDDOCDATE " & wsSortBy
    ElseIf wiSort = 2 Then
    wsSQL = wsSQL & " ORDER BY CUSCODE " & wsSortBy
    ElseIf wiSort = 3 Then
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY VQHDDOCDATE, VQHDDOCNO " & wsSortBy
    End If
    
    End Select
    
     rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, SSEL, SID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
    
     
    With waResult
    .ReDim 0, -1, SSEL, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
  '    wdCreLft = Get_CreditLimit(ReadRs(rsRcd, "SNHDCUSID"), gsSystemDate)
       wdCreLft = 0

     .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "DOCNO")
        waResult(.UpperBound(1), SCUSCODE) = ReadRs(rsRcd, "CUSCODE")
        waResult(.UpperBound(1), SCUSNAME) = ReadRs(rsRcd, "CUSNAME")
        waResult(.UpperBound(1), SDOCDATE) = ReadRs(rsRcd, "DOCDATE")
        waResult(.UpperBound(1), SFROMDATE) = ReadRs(rsRcd, "JOBREF")
        waResult(.UpperBound(1), STODATE) = ""
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
        waResult(.UpperBound(1), SNET) = Format(To_Value(ReadRs(rsRcd, "NET")), gsAmtFmt)
        waResult(.UpperBound(1), SORI) = ReadRs(rsRcd, "REFNO")
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "DOCID")
        
        If wsTrnCd = "SO" Then
            waResult(.UpperBound(1), SREADY) = IIf(ReadRs(rsRcd, "RDY") = "Y", "-1", "0")
        End If
        
    rsRcd.MoveNext
    Loop
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    LoadRecord = True
    Me.MousePointer = vbNormal
    
End Function


Private Function Chk_GrdRow(ByVal LastRow As Long) As Boolean

    Dim wlCtr As Long
    Dim wsDes As String
    Dim wsExcRat As String
    
    Chk_GrdRow = False
    
    On Error GoTo Chk_GrdRow_Err
    
    With tblDetail
        
        If To_Value(LastRow) > waResult.UpperBound(1) Then
           Chk_GrdRow = True
           Exit Function
        End If
        
        
        
        If wiActFlg = 2 And wsFormID = "APR001" Then
        If Chk_grdSoNo("SO", waResult(LastRow, SORI)) = False Then
                .Col = SORI
                Exit Function
        End If
        End If
        
        If wiActFlg = 3 And wsFormID = "APR002" Then
        If DelValidation(To_Value(waResult(LastRow, SID))) = False Then
                .Col = SID
                Exit Function
        End If
        End If
    
    
    
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function


Private Sub cmdSave(ByVal inActFlg As Integer)

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wsCusNo As String
    Dim wsStorep As String
    Dim wsItemNo As String
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    wiActFlg = inActFlg
    
    

    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    '' Last Check when Add
   
    Select Case wiActFlg
    Case 2
    gsMsg = "你是否確認要轉換?"
    Case 3
    gsMsg = "你是否取消此文件?"
    If Opt_Getfocus(optDocType, 3, 0) = 2 Then
        inActFlg = 6
        gsMsg = "你是否完全刪除此文件?"
    End If
    Case 4
    gsMsg = "你是否拷貝此文件?"
    Case 5
    If wsFormID = "APR002" Then
        inActFlg = 7
        gsMsg = "你是否轉換送貨單?"
    Else
        gsMsg = "你是否批核完成此文件?"
    End If
    Case 8
       gsMsg = "你是否完成此文件?"
    End Select
    
    
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

    wsMark = "0"
    
   
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_APR001A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, inActFlg)
                Call SetSPPara(adcmdSave, 2, wsTrnCd)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SID))
                Call SetSPPara(adcmdSave, 4, IIf(wiActFlg = 2 And wsFormID = "APR001", waResult(wiCtr, SORI), ""))
                Call SetSPPara(adcmdSave, 5, wsFormID)
                Call SetSPPara(adcmdSave, 6, gsUserID)
                Call SetSPPara(adcmdSave, 7, wsGenDte)
 
                wsMark = waResult(wiCtr, SID)
                wsCusNo = waResult(wiCtr, SCUSCODE)
                adcmdSave.Execute
                wsDocNo = GetSPPara(adcmdSave, 8)
                wsItemNo = GetSPPara(adcmdSave, 9)
            End If
        Next
    End If
    
    
    If wiActFlg = 2 And wsDocNo = "-1" Then
        gsMsg = "Item ： " & wsItemNo & " length too long!Please convert it to normal item!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        GoTo cmdSave_ItemErr
        
    End If
    
    cnCon.CommitTrans
    
  
    Select Case wiActFlg
    Case 1
        gsMsg = "已完成!"
    Case 2, 4
        gsMsg = "文件 ： " & wsDocNo & " 已完成!"
    Case 3
        gsMsg = "文件已取消完成!"
    Case 5
    If wsFormID = "APR002" Then
        gsMsg = "文件 ： " & wsDocNo & " 已完成!"
    Else
        gsMsg = "已完成!"
    End If
        
    End Select
    MsgBox gsMsg, vbOKOnly, gsTitle
        

    
    Set adcmdSave = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    
    Exit Sub
    
cmdSave_ItemErr:
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub

Private Function InputValidation() As Boolean
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, SSEL)) = "-1" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
        tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function





Private Sub cmdSelect(ByVal wiSelect As Integer)
    Dim wiCtr As Long
    
    Me.MousePointer = vbHourglass
    
    
     
    With waResult
    For wiCtr = 0 To .UpperBound(1)
        waResult(wiCtr, SSEL) = IIf(wiSelect = 1, "-1", "0")
    Next wiCtr
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    Me.MousePointer = vbNormal
    
End Sub


Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property


Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property


Private Function Chk_grdSoNo(ByVal InTrnCd As String, ByVal inAccNo As String) As Boolean
Dim wsStatus As String
    
     Chk_grdSoNo = False
    
    
    If Trim(inAccNo) = "" Then
        Chk_grdSoNo = True
        Exit Function
    End If
    
    
   If Chk_TrnHdDocNo(InTrnCd, inAccNo, wsStatus) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "工程編號已入數, 不能啟用!"
            MsgBox gsMsg, vbOKOnly, gsTitle
        End If
        
        If wsStatus = "2" Then
            gsMsg = "工程編號已刪除, 不能啟用!"
            MsgBox gsMsg, vbOKOnly, gsTitle
        End If
    
        If wsStatus = "3" Then
            gsMsg = "工程編號已無效, 不能啟用!"
            MsgBox gsMsg, vbOKOnly, gsTitle
        End If
        
        If wsStatus = "1" Then
            gsMsg = "工程編號已用, 不能啟用!"
            MsgBox gsMsg, vbOKOnly, gsTitle
        End If
        
        Exit Function
    End If
    
    Chk_grdSoNo = True
    
    
End Function


Private Function DelValidation(ByVal InDocID As Long) As Boolean
Dim OutTrnCd As String
Dim OutDocNo As String

    
    
    DelValidation = False
    
    On Error GoTo DelValidation_Err
    
    
    
 '   If Not chk_txtRevNo Then Exit Function
    If Chk_SoRefDoc(InDocID, OutTrnCd, OutDocNo) = True Then
        
        Select Case OutTrnCd
        Case "SP"
        gsMsg = "配貨單 : " & OutDocNo & " 是以此銷售轉為!不能刪除"
        Case "SW"
        gsMsg = "發貨單 : " & OutDocNo & " 是以此銷售轉為!不能刪除"
        Case "SR"
        gsMsg = "退貨單 : " & OutDocNo & " 是以此銷售轉為!不能刪除"
        Case "IV"
        gsMsg = "發票 : " & OutDocNo & " 是以此銷售轉為!不能刪除"
        Case "PO"
        gsMsg = "採購單 : " & OutDocNo & " 是以此銷售轉為!不能刪除"
        End Select
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    
    End If
    
    DelValidation = True
    
    Exit Function
    
DelValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function


Private Sub tblDetail_HeadClick(ByVal ColIndex As Integer)

    
    On Error GoTo tblDetail_HeadClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case SDOCNO
                wiSort = 0
                cmdRefresh
            Case SDOCDATE
                wiSort = 1
                cmdRefresh
            Case SCUSCODE
                wiSort = 2
                cmdRefresh
            Case SORI
                wiSort = 3
                cmdRefresh
           End Select
    End With

    
    Exit Sub
    
tblDetail_HeadClick_Err:
     MsgBox "Check tblDeiail HeadClick!"

End Sub
Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    'If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = ""
    wsSelection(2) = ""
    wsSelection(3) = ""
    wsSelection(4) = ""
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTJOB005 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wsTitle & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboCusNoFr.Text) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdFr.Text) = "/", "000000", Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(medPrdTo.Text) = "/", "999999", Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "', "
    wsSQL = wsSQL & "'" & Opt_Getfocus(optDocType, 3, 0) & "',"
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTJOB005"
    Else
    wsRptName = "RPTJOB005"
    End If
    
    NewfrmPrint.ReportID = "JOB005"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "JOB005"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdCopy(ByVal inActFlg As Integer)
' Tom 20090203
    Dim wsGenDte As String
    Dim adcmdCopy As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wsCusNo As String
    Dim wsStorep As String
    Dim wsItemNo As String
    
    Dim wsCmpCd As String
    
     
    On Error GoTo cmdCopy_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    wiActFlg = inActFlg
    
    

    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    '' Last Check when Add
    
    wsCmpCd = GetCompanySelect
    
    If wsCmpCd = "" Then
       gsMsg = "No company Code!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       MousePointer = vbDefault
       Exit Sub
    End If

   
    Select Case wsCmpCd
    Case "FR"
    gsMsg = "你是否拷貝此文件到輝域(FR)"
    Case "CF"
    gsMsg = "你是否拷貝此文件到忠輝(CF)"
    Case "CY"
    gsMsg = "你是否拷貝此文件到盈福(CY)"
    Case Else
    gsMsg = "你是否拷貝此文件到輝域(FR)"
    wsCmpCd = "FR"
    End Select
    
    
    'gsMsg = "你是否拷貝此文件?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

    wsMark = "0"
   
    
    cnCon.BeginTrans
    Set adcmdCopy.ActiveConnection = cnCon
 
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdCopy.CommandText = "usp_APR001A_CPY"
        adcmdCopy.CommandType = adCmdStoredProc
        adcmdCopy.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdCopy, 1, inActFlg)
                Call SetSPPara(adcmdCopy, 2, wsCmpCd)
                Call SetSPPara(adcmdCopy, 3, wsTrnCd)
                Call SetSPPara(adcmdCopy, 4, waResult(wiCtr, SID))
                Call SetSPPara(adcmdCopy, 5, IIf(wiActFlg = 2 And wsFormID = "APR001", waResult(wiCtr, SORI), ""))
                Call SetSPPara(adcmdCopy, 6, wsFormID)
                Call SetSPPara(adcmdCopy, 7, gsUserID)
                Call SetSPPara(adcmdCopy, 8, wsGenDte)
 
                wsMark = waResult(wiCtr, SID)
                wsCusNo = waResult(wiCtr, SCUSCODE)
                adcmdCopy.Execute
                wsDocNo = GetSPPara(adcmdCopy, 9)
                wsItemNo = GetSPPara(adcmdCopy, 10)
            End If
        Next
    End If
    
    cnCon.CommitTrans
    
    gsMsg = "文件 ： " & wsDocNo & " 已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    Set adcmdCopy = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    
    Exit Sub
    
cmdCopy_ItemErr:
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdCopy = Nothing
    
    Exit Sub
    
cmdCopy_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdCopy = Nothing
    
End Sub

Private Function GetCompanySelect() As String

    Dim Newfrm As New frmSelectComp
    
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
      '  Set .ctlKey = GetCompanySelect
        .Show vbModal
    GetCompanySelect = .ctlKey
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Function



Private Sub cmdReady()
' Tom 20090203
    Dim wsGenDte As String
    Dim adcmdReady As New ADODB.Command
    Dim wiCtr As Integer
    Dim bPass As Boolean
    
    
     
    On Error GoTo cmdReady_Err
    
    If wsTrnCd <> "SO" Then
        Exit Sub
    End If
    
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    '' Last Check when Add
    
   
    gsMsg = "你是否 Ready/unlock 此文件"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

    
    
    cnCon.BeginTrans
    Set adcmdReady.ActiveConnection = cnCon
 
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdReady.CommandText = "usp_APR001A_RDY"
        adcmdReady.CommandType = adCmdStoredProc
        adcmdReady.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                bPass = True
                
                If Trim(waResult(wiCtr, SREADY)) = "-1" Then
                    bPass = chkSpecialPassword(waResult(wiCtr, SDOCNO))
                End If
                
                If (bPass) Then
                Call SetSPPara(adcmdReady, 1, wsTrnCd)
                Call SetSPPara(adcmdReady, 2, waResult(wiCtr, SID))
                Call SetSPPara(adcmdReady, 3, IIf(waResult(wiCtr, SREADY) = "-1", "N", "Y"))
                Call SetSPPara(adcmdReady, 4, wsFormID)
                Call SetSPPara(adcmdReady, 5, gsUserID)
                Call SetSPPara(adcmdReady, 6, wsGenDte)
                adcmdReady.Execute
                End If
                
            End If
        Next wiCtr
    End If
    
    cnCon.CommitTrans
    
    gsMsg = "已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    Set adcmdReady = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    
    Exit Sub
    
cmdReady_ItemErr:
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdReady = Nothing
    
    Exit Sub
    
cmdReady_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdReady = Nothing
    
End Sub


Private Function chkSpecialPassword(ByVal inDoc As String) As Boolean

    Dim Newfrm As New frmSpecialPassword
    Dim outResult As Boolean
    
    Me.MousePointer = vbHourglass
    outResult = False
    
    With Newfrm
        .DOCNO = inDoc
        .Show vbModal
        outResult = .Result
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
    
    chkSpecialPassword = outResult
    
End Function

Private Sub cmdExport()

    Dim wsGenDte As String
    Dim wiCtr As Integer
    Dim wsTrnCode As String
    Dim wsStorep As String
    Dim wiMod As Integer
    Dim wsPath As String
    Dim wsDoc As String
    
    
    
     
    On Error GoTo cmdExport_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate

    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    '' Last Check when Add
   
    gsMsg = "你是否確認要匯出文件？"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If
    
    wsTrnCode = "DN"
    

    If Trim(gsHHPath) <> "" Then
        wsPath = gsHHPath + "send\HHTORDER.TXT"
    Else
        wsPath = App.Path + "send\HHTORDER.TXT"
    End If
    
'    cnCon.BeginTrans
'    Set adcmdExport.ActiveConnection = cnCon
 
    wiMod = 1
    wsDoc = ""
    If waResult.UpperBound(1) >= 0 Then
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
            
            
            If ExportToHHFile(wsPath, wsTrnCode, waResult(wiCtr, SID), wiMod, "") = False Then
                gsMsg = waResult(wiCtr, SDOCNO) & " 匯出Error!"
                MsgBox gsMsg, vbOKOnly, gsTitle
            Else
            wiMod = 2
            wsDoc = wsDoc & IIf(wsDoc = "", waResult(wiCtr, SID), "," & waResult(wiCtr, SID))
            
            End If
            
            End If
        Next wiCtr
    End If
    
    
 '   cnCon.CommitTrans
    Sleep (500)
    If SendToHH(wsPath) = True Then
  
    gsMsg = "匯出文件已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    End If
        

    
   ' Set adcmdExport = Nothing
    
    Call LoadRecord
    
    MousePointer = vbDefault
    
    Exit Sub
    
cmdExport_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
  '  cnCon.RollbackTrans
  '  Set adcmdExport = Nothing
    
End Sub

