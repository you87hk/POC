VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmINQ001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmINQ001.frx":0000
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
      OleObjectBlob   =   "frmINQ001.frx":0442
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboItemNoFr 
      Height          =   300
      Left            =   2040
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1812
   End
   Begin VB.ComboBox cboItemNoTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoFr 
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboCusNoTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoFr 
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1812
   End
   Begin VB.ComboBox cboDocNoTo 
      Height          =   300
      Left            =   5520
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   11775
      Begin VB.Frame fraSelect 
         Height          =   525
         Left            =   7320
         TabIndex        =   15
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
            Index           =   1
            Left            =   2040
            TabIndex        =   5
            Top             =   200
            Width           =   1335
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
            TabIndex        =   4
            Top             =   200
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Label lblItemNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   17
         Top             =   975
         Width           =   1095
      End
      Begin VB.Label lblItemNoFr 
         Caption         =   "Customer Code From"
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label lblCusNoFr 
         Caption         =   "Customer Code From"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblCusNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   11
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblDocNoTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4080
         TabIndex        =   10
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblDocNoFr 
         Caption         =   "Document # From"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   1890
      End
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6255
      Left            =   120
      OleObjectBlob   =   "frmINQ001.frx":2B45
      TabIndex        =   6
      Top             =   1920
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
            Picture         =   "frmINQ001.frx":B5D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":BEB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":C78C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":CBDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":D030
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":D34A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":D79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":DBEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":DF08
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":E222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":E674
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":EF50
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":F278
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":F6CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":F9E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":FD04
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":10158
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":10474
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":10794
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":10AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":11390
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":116AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ001.frx":119C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   14
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
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Print"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "重新整理 (F5)"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblDspNetAmtOrg 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "9.999.999.999.99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   10440
      TabIndex        =   22
      Top             =   8280
      Width           =   1410
   End
   Begin VB.Label lblDspGrsAmtOrg 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "9.999.999.999.99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9120
      TabIndex        =   21
      Top             =   8280
      Width           =   1290
   End
   Begin VB.Label lblDspTotalQty 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   "9.999.999.999.99"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   8280
      TabIndex        =   20
      Top             =   8280
      Width           =   810
   End
   Begin VB.Label lblDspItmDesc 
      BorderStyle     =   1  '單線固定
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   8280
      Width           =   8145
   End
End
Attribute VB_Name = "frmINQ001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Dim waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private wcCombo As Control
Private wbErr As Boolean
Private wsMark As String



Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wiActFlg As Integer
Private wsTrnCd As String


Private Const tcPrint = "Print"


Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"


Private Const SDOCDATE = 0
Private Const SDOCNO = 1
Private Const SCUSCODE = 2
Private Const SDOCLINE = 3
Private Const SITMCODE = 4
Private Const SWHSCODE = 5
Private Const SUPRICE = 6
Private Const SQTY = 7
Private Const SDISPER = 8
Private Const SAMT = 9
Private Const SNET = 10
Private Const SID = 11
Private Const SDUMMY = 12


Private Sub cboCusNoFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    Select Case wsTrnCd
        Case "SO", "IV", "SN"
            wsSQL = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
            wsSQL = wsSQL & " AND CusStatus <> '2' "
            wsSQL = wsSQL & " AND CusInactive = 'N' "
            wsSQL = wsSQL & " ORDER BY Cuscode "
    
        Case "PO", "PV"
            wsSQL = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
            wsSQL = wsSQL & " AND VdrStatus <> '2' "
            wsSQL = wsSQL & " AND VdrInactive = 'N' "
            wsSQL = wsSQL & " ORDER BY Vdrcode "
    
    End Select
    
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
    Select Case wsTrnCd
        Case "SO", "IV", "SN"
            wsSQL = "SELECT CusCode, CusName FROM mstCustomer WHERE CusCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
            wsSQL = wsSQL & " AND CusStatus <> '2' "
            wsSQL = wsSQL & " AND CusInactive = 'N' "
            wsSQL = wsSQL & " ORDER BY Cuscode "
    
        Case "PO", "PV"
            wsSQL = "SELECT VdrCode, VdrName FROM mstVendor WHERE VdrCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
            wsSQL = wsSQL & " AND VdrStatus <> '2' "
            wsSQL = wsSQL & " AND VdrInactive = 'N' "
            wsSQL = wsSQL & " ORDER BY Vdrcode "
    
    End Select
    
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
        
        cboItemNoFr.SetFocus
    End If
End Sub



Private Sub cboCusNoTo_LostFocus()
FocusMe cboCusNoTo, True
End Sub

'''
Private Sub cboItemNoFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    wsSQL = "SELECT ItmCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & " FROM mstItem WHERE ItmCode LIKE '%" & IIf(cboItemNoFr.SelLength > 0, "", Set_Quote(cboItemNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmStatus <> '2' "
    wsSQL = wsSQL & " ORDER BY Itmcode "
    Call Ini_Combo(2, wsSQL, cboItemNoFr.Left, cboItemNoFr.Top + cboItemNoFr.Height, tblCommon, wsFormID, "TBLItemNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItemNoFr_GotFocus()
        FocusMe cboItemNoFr
    Set wcCombo = cboItemNoFr
End Sub

Private Sub cboItemNoFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItemNoFr, 30, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItemNoFr.Text) <> "" And _
            Trim(cboItemNoTo.Text) = "" Then
            cboItemNoTo.Text = cboItemNoFr.Text
        End If
        cboItemNoTo.SetFocus
    End If
End Sub


Private Sub cboItemNoFr_LostFocus()
    FocusMe cboItemNoFr, True
End Sub

Private Sub cboItemNoTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    wsSQL = "SELECT ItmCode, " & IIf(gsLangID = "1", "ItmEngName", "ItmChiName") & " FROM mstItem WHERE ItmCode LIKE '%" & IIf(cboItemNoTo.SelLength > 0, "", Set_Quote(cboItemNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmStatus <> '2' "
    wsSQL = wsSQL & " ORDER BY Itmcode "
    
    Call Ini_Combo(2, wsSQL, cboItemNoTo.Left, cboItemNoTo.Top + cboItemNoTo.Height, tblCommon, wsFormID, "TBLItemNo", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItemNoTo_GotFocus()
    FocusMe cboItemNoTo
    Set wcCombo = cboItemNoTo
End Sub

Private Sub cboItemNoTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItemNoTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItemNoTo = False Then
            Exit Sub
        End If
        
        If LoadRecord = True Then
            tblDetail.SetFocus
        End If
       
    End If
End Sub

Private Sub cboItemNoTo_LostFocus()
FocusMe cboItemNoTo, True
End Sub


Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
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
    Case "IV"
    wsSQL = "SELECT IVHDDOCNO, CUSCODE, IVHDDOCDATE "
    wsSQL = wsSQL & " FROM soaIVHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND IVHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO "
    Case "PO"
    wsSQL = "SELECT POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSQL = wsSQL & " FROM popPOHD, mstVENDOR "
    wsSQL = wsSQL & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND POHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO "
    Case "PV"
    wsSQL = "SELECT PVHDDOCNO, VDRCODE, PVHDDOCDATE "
    wsSQL = wsSQL & " FROM popPVHD, mstVENDOR "
    wsSQL = wsSQL & " WHERE PVHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND PVHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND PVHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY PVHDDOCNO "
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
    Case "IV"
    wsSQL = "SELECT IVHDDOCNO, CUSCODE, IVHDDOCDATE "
    wsSQL = wsSQL & " FROM soaIVHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND IVHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY IVHDDOCNO "
     Case "PO"
    wsSQL = "SELECT POHDDOCNO, VDRCODE, POHDDOCDATE "
    wsSQL = wsSQL & " FROM popPOHD, mstVENDOR "
    wsSQL = wsSQL & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND POHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO "
    Case "PV"
    wsSQL = "SELECT PVHDDOCNO, VDRCODE, PVHDDOCDATE "
    wsSQL = wsSQL & " FROM popPVHD, mstVENDOR "
    wsSQL = wsSQL & " WHERE PVHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND PVHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND PVHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY PVHDDOCNO "
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
Private Function chk_cboItemNoTo() As Boolean
    chk_cboItemNoTo = False
    
    If UCase(cboItemNoFr.Text) > UCase(cboItemNoTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItemNoTo.SetFocus
        Exit Function
    End If
    
    chk_cboItemNoTo = True
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode

        
  '      Case vbKeyF3
           
        Case vbKeyF11
           Call cmdCancel
          
            
        Case vbKeyF12
            Unload Me
             
        
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
     
    
        Case tcPrint
        
        
        Case tcCancel
        
           Call cmdCancel
           
              
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
  
   ' Ini_Scr
    Call LoadRecord
    
   MousePointer = vbDefault
    
    
End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, SDOCDATE, SID
    
    
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
    wsMark = "0"
    

    cboDocNoFr.Text = ""
    cboDocNoTo.Text = ""
    cboCusNoFr.Text = ""
    cboCusNoTo.Text = ""
    cboItemNoFr.Text = ""
    cboItemNoTo.Text = ""


    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   

    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmINQ001 = Nothing
 
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    optDocType(0).Value = True

End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblDocNoFr.Caption = Get_Caption(waScrItm, "DOCNOFR")
    lblDocNoTo.Caption = Get_Caption(waScrItm, "DOCNOTO")
    lblCusNoFr.Caption = Get_Caption(waScrItm, "CUSNOFR")
    lblCusNoTo.Caption = Get_Caption(waScrItm, "CUSNOTO")
    lblItemNoFr.Caption = Get_Caption(waScrItm, "ITEMNOFR")
    lblItemNoTo.Caption = Get_Caption(waScrItm, "ITEMNOTO")
    optDocType(0).Caption = Get_Caption(waScrItm, "OPT1")
    optDocType(1).Caption = Get_Caption(waScrItm, "OPT2")
                
     
    With tblDetail
        .Columns(SDOCDATE).Caption = Get_Caption(waScrItm, "SDOCDATE")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SCUSCODE).Caption = Get_Caption(waScrItm, "SCUSCODE")
        .Columns(SDOCLINE).Caption = Get_Caption(waScrItm, "SDOCLINE")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SWHSCODE).Caption = Get_Caption(waScrItm, "SWHSCODE")
        .Columns(SUPRICE).Caption = Get_Caption(waScrItm, "SUPRICE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(SDISPER).Caption = Get_Caption(waScrItm, "SDISPER")
        .Columns(SAMT).Caption = Get_Caption(waScrItm, "SAMT")
        .Columns(SNET).Caption = Get_Caption(waScrItm, "SNET")
        
    End With
    

    'tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F11)"
    
    
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    

End Sub










Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .Update
    End With


   
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    
    
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
                
             '    If .Columns(SDOCNO).Text <> "" Then
                    
             '       frmINQ0011.InDocID = .Columns(SID).Text
             '       frmINQ0011.InCusNo = .Columns(SCUSCODE).Text
             '       frmINQ0011.Show vbModal
                 
             '   End If
                
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
            Case SNET
                 KeyCode = vbKeyDown
                 .Col = SDOCDATE
            Case Else
                 KeyCode = vbDefault
                 .Col = .Col + 1
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SDOCDATE Then
                .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SNET
                    KeyCode = vbKeyDown
                    .Col = SDOCDATE
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
   
                Case SCUSCODE
                
                     lblDspItmDesc.Caption = ""
                     
                     Select Case wsTrnCd
                     Case "SO", "IV", "SN"
                        lblDspItmDesc.Caption = Get_TableInfo("MSTCUSTOMER", "CUSCODE = '" & Set_Quote(.Columns(SCUSCODE).Text) & "'", "CUSNAME")
                     Case "PO", "PV"
                        lblDspItmDesc.Caption = Get_TableInfo("MSTVENDOR", "VDRCODE = '" & Set_Quote(.Columns(SCUSCODE).Text) & "'", "VDRNAME")
                     End Select
                     
                Case SITMCODE
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = Get_TableInfo("MSTITEM", "ITMCODE = '" & Set_Quote(.Columns(SITMCODE).Text) & "'", IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME"))
                  
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
        
        For wiCtr = SDOCDATE To SDUMMY
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SDOCDATE
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 1000
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 1500
                Case SCUSCODE
                   .Columns(wiCtr).Width = 800
                   .Columns(wiCtr).DataWidth = 10
                Case SDOCLINE
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Width = 500
                Case SITMCODE
                   .Columns(wiCtr).Width = 2500
                   .Columns(wiCtr).DataWidth = 50
                Case SWHSCODE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SUPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                Case SQTY
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SAMT
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SNET
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SDISPER
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case SDUMMY
                    .Columns(wiCtr).Width = 100
                    .Columns(wiCtr).DataWidth = 0
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
    Dim wdOutQty As Double
    Dim wsStatus As String
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
        wsStatus = "1"
     Else
        wsStatus = "4"
    End If
    
    Select Case wsTrnCd
    Case "SN"
    wsSQL = "SELECT SNHDDOCID DOCID, SNHDDOCNO DOCNO, SNHDDOCDATE DOCDATE, CUSCODE CCODE, SNDTDOCLINE DOCLINE, ITMCODE, SNDTWHSCODE WHSCODE, "
    wsSQL = wsSQL & " SNDTUPRICE UPRICE, SNDTDISPER DISPER, SNDTTOTQTY QTY, SNDTTOTAMT AMT, SNDTTOTNET NET "
    wsSQL = wsSQL & " FROM  SOASNHD, SOASNDT, MSTCUSTOMER, MSTITEM "
    wsSQL = wsSQL & " WHERE SNHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND ITMCODE BETWEEN '" & cboItemNoFr & "' AND '" & IIf(Trim(cboItemNoTo.Text) = "", String(13, "z"), Set_Quote(cboItemNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SNHDDOCID = SNDTDOCID "
    wsSQL = wsSQL & " AND SNHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SNDTITEMID = ITMID "
    wsSQL = wsSQL & " AND SNHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " ORDER BY SNHDDOCDATE, SNHDDOCNO, SNDTDOCLINE "
    Case "SO"
    wsSQL = "SELECT SOHDDOCID DOCID, SOHDDOCNO DOCNO, SOHDDOCDATE DOCDATE, CUSCODE CCODE, SODTDOCLINE DOCLINE, ITMCODE, SODTWHSCODE WHSCODE, "
    wsSQL = wsSQL & " SODTUPRICE UPRICE, SODTDISPER DISPER, SODTTOTQTY QTY, SODTTOTAMT AMT, SODTTOTNET NET "
    wsSQL = wsSQL & " FROM  SOASOHD, SOASODT, MSTCUSTOMER, MSTITEM "
    wsSQL = wsSQL & " WHERE SOHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND ITMCODE BETWEEN '" & cboItemNoFr & "' AND '" & IIf(Trim(cboItemNoTo.Text) = "", String(13, "z"), Set_Quote(cboItemNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SOHDDOCID = SODTDOCID "
    wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
    wsSQL = wsSQL & " AND SOHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " ORDER BY SOHDDOCDATE, SOHDDOCNO, SODTDOCLINE "
    Case "IV"
    wsSQL = "SELECT IVHDDOCID DOCID, IVHDDOCNO DOCNO, IVHDDOCDATE DOCDATE, CUSCODE CCODE, IVDTDOCLINE DOCLINE, IVDTDESC1 ITMCODE, IVDTDESC2 WHSCODE, "
    wsSQL = wsSQL & " IVDTUPRICE UPRICE, IVDTDISPER DISPER, IVDTQTY QTY, IVDTAMT AMT, IVDTNET NET "
    wsSQL = wsSQL & " FROM  SOAIVHD, SOAIVDT, MSTCUSTOMER "
    wsSQL = wsSQL & " WHERE IVHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND CUSCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND IVHDDOCID = IVDTDOCID "
    wsSQL = wsSQL & " AND IVHDCUSID = CUSID "
    wsSQL = wsSQL & " AND IVHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " ORDER BY IVHDDOCDATE, IVHDDOCNO, IVDTDOCLINE "
    Case "PO"
    wsSQL = "SELECT POHDDOCID DOCID, POHDDOCNO DOCNO, POHDDOCDATE DOCDATE, VDRCODE CCODE, PODTDOCLINE DOCLINE, ITMCODE, PODTWHSCODE WHSCODE, "
    wsSQL = wsSQL & " PODTUPRICE UPRICE, PODTDISPER DISPER, PODTQTY QTY, PODTAMT AMT, PODTNET NET "
    wsSQL = wsSQL & " FROM  POPPOHD, POPPODT, MSTVENDOR, MSTITEM "
    wsSQL = wsSQL & " WHERE POHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND VDRCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND ITMCODE BETWEEN '" & cboItemNoFr & "' AND '" & IIf(Trim(cboItemNoTo.Text) = "", String(13, "z"), Set_Quote(cboItemNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND POHDVDRID = VDRID "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND POHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " ORDER BY POHDDOCDATE, POHDDOCNO, PODTDOCLINE "
    Case "PV"
    wsSQL = "SELECT PVHDDOCID DOCID, PVHDDOCNO DOCNO, PVHDDOCDATE DOCDATE, VDRCODE CCODE, PVDTDOCLINE DOCLINE, ITMCODE, PVDTWHSCODE WHSCODE, "
    wsSQL = wsSQL & " PVDTUPRICE UPRICE, PVDTDISPER DISPER, PVDTQTY QTY, PVDTAMT AMT, PVDTNET NET "
    wsSQL = wsSQL & " FROM  POPPVHD, POPPVDT, MSTVENDOR, MSTITEM "
    wsSQL = wsSQL & " WHERE PVHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND VDRCODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND ITMCODE BETWEEN '" & cboItemNoFr & "' AND '" & IIf(Trim(cboItemNoTo.Text) = "", String(13, "z"), Set_Quote(cboItemNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND PVHDDOCID = PVDTDOCID "
    wsSQL = wsSQL & " AND PVHDVDRID = VDRID "
    wsSQL = wsSQL & " AND PVDTITEMID = ITMID "
    wsSQL = wsSQL & " AND PVHDSTATUS = '" & wsStatus & "'"
    wsSQL = wsSQL & " ORDER BY PVHDDOCDATE, PVHDDOCNO, PVDTDOCLINE "
    
    
    End Select
    
     rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, SDOCDATE, SID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
    
     
    With waResult
    .ReDim 0, -1, SDOCDATE, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
     
     .AppendRows
        waResult(.UpperBound(1), SDOCDATE) = ReadRs(rsRcd, "DOCDATE")
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "DOCNO")
        waResult(.UpperBound(1), SCUSCODE) = ReadRs(rsRcd, "CCODE")
        waResult(.UpperBound(1), SDOCLINE) = Format(To_Value(ReadRs(rsRcd, "DOCLINE")), "000")
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waResult(.UpperBound(1), SWHSCODE) = ReadRs(rsRcd, "WHSCODE")
        waResult(.UpperBound(1), SUPRICE) = Format(To_Value(ReadRs(rsRcd, "UPRICE")), gsUprFmt)
        waResult(.UpperBound(1), SQTY) = Format(To_Value(ReadRs(rsRcd, "QTY")), gsQtyFmt)
        waResult(.UpperBound(1), SDISPER) = Format(To_Value(ReadRs(rsRcd, "DISPER")), gsQtyFmt)
        waResult(.UpperBound(1), SAMT) = Format(To_Value(ReadRs(rsRcd, "AMT")), gsAmtFmt)
        waResult(.UpperBound(1), SNET) = Format(To_Value(ReadRs(rsRcd, "NET")), gsAmtFmt)
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "DOCID")
    rsRcd.MoveNext
    Loop
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    
    Call Calc_Total
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    LoadRecord = True
    Me.MousePointer = vbNormal
    
End Function


Private Function Calc_Total(Optional ByVal LastRow As Variant) As Boolean
    
    Dim wiTotalGrs As Double
    Dim wiTotalNet As Double
    Dim wiTotalQty As Double
    
    Dim wiRowCtr As Integer
    
    
    Calc_Total = False
    
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotalGrs = wiTotalGrs + To_Value(waResult(wiRowCtr, SAMT))
        wiTotalNet = wiTotalNet + To_Value(waResult(wiRowCtr, SNET))
        wiTotalQty = wiTotalQty + To_Value(waResult(wiRowCtr, SQTY))
    Next
    
    lblDspGrsAmtOrg.Caption = Format(CStr(wiTotalGrs), gsAmtFmt)
    lblDspNetAmtOrg.Caption = Format(CStr(wiTotalNet), gsAmtFmt)
    lblDspTotalQty.Caption = Format(CStr(wiTotalQty), gsQtyFmt)
    
    Calc_Total = True

End Function
Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property
Public Property Let TrnCd(InTrnCd As String)
    wsTrnCd = InTrnCd
End Property
