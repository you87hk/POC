VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmSTKCNT 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmSTKCNT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   8620.47
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   11923.82
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   11280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":1A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":1E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":21B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":2606
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":2A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":2D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":308C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":34DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":3DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":40E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":4536
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":4852
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":4B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":4FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":52DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":55FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":5A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSTKCNT.frx":5D6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9360
      OleObjectBlob   =   "frmSTKCNT.frx":608E
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboWhsCodeFr 
      Height          =   300
      Left            =   2520
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1812
   End
   Begin VB.ComboBox cboWhsCodeTo 
      Height          =   300
      Left            =   5400
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1812
   End
   Begin VB.ComboBox cboItmBarCodeFr 
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboItmBarCodeTo 
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboItmCodeFr 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1812
   End
   Begin VB.ComboBox cboItmCodeTo 
      Height          =   300
      Left            =   5400
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   11775
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   7320
         TabIndex        =   17
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton optShow 
            Caption         =   "SN"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optShow 
            Caption         =   "SN"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Label lblWhsCodeTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblWhsCodeFr 
         Caption         =   "ItmTypeCode From"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label lblItmBarCodeFr 
         Caption         =   "Itm Barcode From"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblItmBarCodeTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4560
         TabIndex        =   11
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblItmCodeTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4560
         TabIndex        =   10
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label lblItmCodeFr 
         Caption         =   "Itm # From"
         Height          =   225
         Left            =   120
         TabIndex        =   9
         Top             =   255
         Width           =   1890
      End
   End
   Begin TrueDBGrid60.TDBGrid tblDetail 
      Height          =   6015
      Left            =   120
      OleObjectBlob   =   "frmSTKCNT.frx":8791
      TabIndex        =   6
      Top             =   2160
      Width           =   11775
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
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "OK"
            Object.ToolTipText     =   "選取 (F2)"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   13
      Top             =   8280
      Width           =   11655
   End
End
Attribute VB_Name = "frmSTKCNT"
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



Private wiExit As Boolean
Private wsFormCaption As String
Private wsFormID As String
Private wiActFlg As Integer


Private Const tcOK = "OK"
Private Const tcPrint = "Print"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"

Private Const SSEL = 0
Private Const SITMCODE = 1
Private Const SITMNAME = 2
Private Const SWHSCODE = 3
Private Const SLOTNO = 4
Private Const SSOH = 5
Private Const SCOUNTED = 6
Private Const SQTYDIFF = 7
Private Const SDUMMY = 8
Private Const SID = 9



Private Sub cboItmBarCodeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    
    wsSQL = "SELECT SCHDDOCNO, SCHDDOCDATE "
    wsSQL = wsSQL & " FROM ICSTKCNT "
    wsSQL = wsSQL & " WHERE SCHDDOCNO LIKE '%" & IIf(cboItmBarCodeFr.SelLength > 0, "", Set_Quote(cboItmBarCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SCHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND SCHDTRNCODE  = 'SC' "
    wsSQL = wsSQL & " ORDER BY SCHDDOCNO "
    
    
    Call Ini_Combo(2, wsSQL, cboItmBarCodeFr.Left, cboItmBarCodeFr.Top + cboItmBarCodeFr.Height, tblCommon, wsFormID, "TBLItmBarCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmBarCodeFr_GotFocus()
        FocusMe cboItmBarCodeFr
    Set wcCombo = cboItmBarCodeFr
End Sub

Private Sub cboItmBarCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmBarCodeFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItmBarCodeFr.Text) <> "" And _
            Trim(cboItmBarCodeTo.Text) = "" Then
            cboItmBarCodeTo.Text = cboItmBarCodeFr.Text
        End If
        cboItmBarCodeTo.SetFocus
    End If
End Sub


Private Sub cboItmBarCodeFr_LostFocus()
    FocusMe cboItmBarCodeFr, True
End Sub

Private Sub cboItmBarCodeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass


    wsSQL = "SELECT SCHDDOCNO, SCHDDOCDATE "
    wsSQL = wsSQL & " FROM ICSTKCNT "
    wsSQL = wsSQL & " WHERE SCHDDOCNO LIKE '%" & IIf(cboItmBarCodeTo.SelLength > 0, "", Set_Quote(cboItmBarCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SCHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND SCHDTRNCODE  = 'SC' "
    wsSQL = wsSQL & " ORDER BY SCHDDOCNO "

    Call Ini_Combo(2, wsSQL, cboItmBarCodeTo.Left, cboItmBarCodeTo.Top + cboItmBarCodeTo.Height, tblCommon, wsFormID, "TBLItmBarCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmBarCodeTo_GotFocus()
    FocusMe cboItmBarCodeTo
    Set wcCombo = cboItmBarCodeTo
End Sub

Private Sub cboItmBarCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmBarCodeTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmBarCodeTo = False Then
            Exit Sub
        End If
        
        cboWhsCodeFr.SetFocus
    End If
End Sub



Private Sub cboItmBarCodeTo_LostFocus()
FocusMe cboItmBarCodeTo, True
End Sub








Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
End Sub



Private Sub cboItmCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeFr
  
    wsSQL = "SELECT ItmCode, ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmCode LIKE '%" & IIf(cboItmCodeFr.SelLength > 0, "", Set_Quote(cboItmCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    Call Ini_Combo(3, wsSQL, cboItmCodeFr.Left, cboItmCodeFr.Top + cboItmCodeFr.Height, tblCommon, wsFormID, "TBLItmCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeFr_GotFocus()
    FocusMe cboItmCodeFr
    Set wcCombo = cboItmCodeFr
End Sub

Private Sub cboItmCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeFr, 30, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboItmCodeFr.Text) <> "" And _
            Trim(cboItmCodeTo.Text) = "" Then
            cboItmCodeTo.Text = cboItmCodeFr.Text
        End If
        cboItmCodeTo.SetFocus
    End If
End Sub

Private Sub cboItmCodeFr_LostFocus()
    FocusMe cboItmCodeFr, True
End Sub

Private Sub cboItmCodeTo_DropDown()
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmCodeTo
  
    wsSQL = "SELECT ItmCode, ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmCode LIKE '%" & IIf(cboItmCodeTo.SelLength > 0, "", Set_Quote(cboItmCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    Call Ini_Combo(3, wsSQL, cboItmCodeTo.Left, cboItmCodeTo.Top + cboItmCodeTo.Height, tblCommon, wsFormID, "TBLItmCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCodeTo_GotFocus()
    FocusMe cboItmCodeTo
    Set wcCombo = cboItmCodeTo
End Sub

Private Sub cboItmCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCodeTo, 30, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmCodeTo = False Then
            Call cboItmCodeTo_GotFocus
            Exit Sub
        End If
        
       cboItmBarCodeFr.SetFocus
        
        
    End If
End Sub

Private Sub cboItmCodeTo_LostFocus()
    FocusMe cboItmCodeTo, True
End Sub
Private Function chk_cboItmCodeTo() As Boolean
    chk_cboItmCodeTo = False
    
    If UCase(cboItmCodeFr.Text) > UCase(cboItmCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboItmCodeTo = True
End Function

Private Function chk_cboItmBarCodeTo() As Boolean
    chk_cboItmBarCodeTo = False
    
    If UCase(cboItmBarCodeFr.Text) > UCase(cboItmBarCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItmBarCodeTo.SetFocus
        Exit Function
    End If
    
    chk_cboItmBarCodeTo = True
End Function



Private Sub cboWhsCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWhsCodeFr
  
    wsSQL = "SELECT WhsCode, WhsDesc FROM MstWarehouse "
    wsSQL = wsSQL & " WHERE WhsStatus = '1'"
    wsSQL = wsSQL & " AND WhsCode LIKE '%" & IIf(cboWhsCodeFr.SelLength > 0, "", Set_Quote(cboWhsCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY WhsCode "
    
    Call Ini_Combo(2, wsSQL, cboWhsCodeFr.Left, cboWhsCodeFr.Top + cboWhsCodeFr.Height, tblCommon, wsFormID, "TBLWhsCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboWhsCodeFr_GotFocus()
    FocusMe cboWhsCodeFr
    Set wcCombo = cboWhsCodeFr
End Sub

Private Sub cboWhsCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWhsCodeFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboWhsCodeFr.Text) <> "" And _
            Trim(cboWhsCodeTo.Text) = "" Then
            cboWhsCodeTo.Text = cboWhsCodeFr.Text
        End If
        cboWhsCodeTo.SetFocus
    End If
End Sub

Private Sub cboWhsCodeFr_LostFocus()
    FocusMe cboWhsCodeFr, True
End Sub

Private Sub cboWhsCodeTo_DropDown()
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboWhsCodeTo
  
    wsSQL = "SELECT WhsCode, WhsDesc FROM MstWarehouse "
    wsSQL = wsSQL & " WHERE WhsStatus = '1'"
    wsSQL = wsSQL & " AND WhsCode LIKE '%" & IIf(cboWhsCodeTo.SelLength > 0, "", Set_Quote(cboWhsCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY WhsCode "
    Call Ini_Combo(2, wsSQL, cboWhsCodeTo.Left, cboWhsCodeTo.Top + cboWhsCodeTo.Height, tblCommon, wsFormID, "TBLWhsCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboWhsCodeTo_GotFocus()
    FocusMe cboWhsCodeTo
    Set wcCombo = cboWhsCodeTo
End Sub

Private Sub cboWhsCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboWhsCodeTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboWhsCodeTo = False Then
            Call cboWhsCodeTo_GotFocus
            Exit Sub
        End If
        
       Call Opt_Setfocus(optShow, 2, 0)
        
        
    End If
End Sub

Private Sub cboWhsCodeTo_LostFocus()
    FocusMe cboWhsCodeTo, True
End Sub
Private Function chk_cboWhsCodeTo() As Boolean
    chk_cboWhsCodeTo = False
    
    If UCase(cboWhsCodeFr.Text) > UCase(cboWhsCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboWhsCodeTo = True
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF6
           Call cmdSave(1)
            
        Case vbKeyF11
           Call cmdPrint
           
        Case vbKeyF3
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
            
        Case vbKeyF9
           Call cmdSelect(1)
           
        Case vbKeyF10
           Call cmdSelect(0)
        
        Case vbKeyF5
            Call LoadRecord
        
      
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


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, SSEL, SID
  
    
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
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

    Me.Caption = wsFormCaption
    
    tblCommon.Visible = False
    wiExit = False
    
    optShow(0).Value = True
      
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   

    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmSTKCNT = Nothing

    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "STKCNT"
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblItmCodeFr.Caption = Get_Caption(waScrItm, "ItmCodeFR")
    lblItmCodeTo.Caption = Get_Caption(waScrItm, "ItmCodeTO")
    lblItmBarCodeFr.Caption = Get_Caption(waScrItm, "ItmBarCodeFR")
    lblItmBarCodeTo.Caption = Get_Caption(waScrItm, "ItmBarCodeTO")
    lblWhsCodeFr.Caption = Get_Caption(waScrItm, "WhsCodeFR")
    lblWhsCodeTo.Caption = Get_Caption(waScrItm, "WhsCodeTO")
        

    optShow(0).Caption = Get_Caption(waScrItm, "SHOW0")
    optShow(1).Caption = Get_Caption(waScrItm, "SHOW1")
    
    
    
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SItmCode")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SItmNAME")
        .Columns(SWHSCODE).Caption = Get_Caption(waScrItm, "SWHSCODE")
        .Columns(SCOUNTED).Caption = Get_Caption(waScrItm, "SCOUNTED")
        .Columns(SSOH).Caption = Get_Caption(waScrItm, "SSOH")
        .Columns(SLOTNO).Caption = Get_Caption(waScrItm, "SLOTNO")
        .Columns(SQTYDIFF).Caption = Get_Caption(waScrItm, "SQTYDIFF")
        
    End With
    
    
    
    
    tbrProcess.Buttons(tcOK).ToolTipText = Get_Caption(waScrToolTip, tcOK) & "(F6)"
    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F11)"
    
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F5)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F3)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrToolTip, tcSAll) & "(F9)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrToolTip, tcDAll) & "(F10)"
    
    

End Sub













Private Sub optShow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
                
      Call LoadRecord
       
    End If
End Sub
Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    With tblDetail
        .UPDATE
    End With
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
          
            
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





Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode

            
        Case vbKeyReturn
            Select Case .Col
            Case SQTYDIFF
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
                Case SQTYDIFF
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


Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    Select Case tblDetail.Col
        

       
    End Select

End Sub
Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        
        
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                
                Case SITMCODE
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = .Columns(SITMNAME).Text
                
                Case SITMNAME
                    lblDspItmDesc.Caption = ""
                    lblDspItmDesc.Caption = .Columns(SITMNAME).Text
                    
    
              
                 
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case tcOK
            Call cmdSave(1)
            
        Case tcPrint
        
            Call cmdPrint
            
        Case tcCancel
        
           Call cmdCancel
           
        Case tcSAll
           Call cmdSelect(1)
           
        Case tcDAll
           Call cmdSelect(0)
            
        Case tcExit
            Unload Me
            
        Case tcRefresh
            Call LoadRecord
            
    End Select
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
                    .Columns(wiCtr).Visible = False
                Case SITMCODE
                    .Columns(wiCtr).DataWidth = 30
                    .Columns(wiCtr).Width = 1500
                Case SITMNAME
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Width = 3500
                Case SWHSCODE
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Width = 1200
                Case SSOH
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SCOUNTED
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SLOTNO
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 20
                Case SQTYDIFF
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
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
    Dim adcmdSave As New ADODB.Command
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsDteTim As String
    
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
            
    wsDteTim = Now
    wsDteTim = Change_SQLDate(wsDteTim)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
        adcmdSave.CommandText = "USP_RPTSTKCNT"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        Call SetSPPara(adcmdSave, 1, gsUserID)
        Call SetSPPara(adcmdSave, 2, Change_SQLDate(wsDteTim))
        Call SetSPPara(adcmdSave, 3, "盤點數目報表")
        Call SetSPPara(adcmdSave, 4, Opt_Getfocus(optShow, 2, 0))
        Call SetSPPara(adcmdSave, 5, cboItmCodeFr)
        Call SetSPPara(adcmdSave, 6, IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), cboItmCodeTo.Text))
        Call SetSPPara(adcmdSave, 7, cboItmBarCodeFr)
        Call SetSPPara(adcmdSave, 8, IIf(Trim(cboItmBarCodeTo.Text) = "", String(13, "z"), cboItmBarCodeTo.Text))
        Call SetSPPara(adcmdSave, 9, cboWhsCodeFr)
        Call SetSPPara(adcmdSave, 10, IIf(Trim(cboWhsCodeTo.Text) = "", String(10, "z"), cboWhsCodeTo.Text))
        Call SetSPPara(adcmdSave, 11, gsLangID)
        
        adcmdSave.Execute
        
    
    cnCon.CommitTrans
    
    wsSQL = "SELECT RPTITMID, RPTITMCODE, RPTITMNAM , RPTBARCODE, RPTSOH, RPTCOUNTED, RPTWHSCODE, RPTLOTNO, "
    wsSQL = wsSQL & " RPTQTYDIFF "
    wsSQL = wsSQL & " FROM RPTSTKCNT "
    wsSQL = wsSQL & " WHERE RPTDTETIM = '" & wsDteTim & "' "
    wsSQL = wsSQL & " AND RPTUSRID = '" & gsUserID & "' "
    
    If optShow(1).Value = True Then
    wsSQL = wsSQL & " AND RPTSTKFLG = 'Y' "
    End If
    wsSQL = wsSQL & " ORDER BY RPTITMCODE "
    
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

         .AppendRows
        waResult(.UpperBound(1), SSEL) = "0"
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "RPTITMCODE")
        waResult(.UpperBound(1), SITMNAME) = ReadRs(rsRcd, "RPTITMNAM")
        waResult(.UpperBound(1), SWHSCODE) = ReadRs(rsRcd, "RPTWHSCODE")
        waResult(.UpperBound(1), SLOTNO) = ReadRs(rsRcd, "RPTLOTNO")
        waResult(.UpperBound(1), SSOH) = Format(ReadRs(rsRcd, "RPTSOH"), gsQtyFmt)
        waResult(.UpperBound(1), SCOUNTED) = Format(ReadRs(rsRcd, "RPTCOUNTED"), gsQtyFmt)
        waResult(.UpperBound(1), SQTYDIFF) = Format(ReadRs(rsRcd, "RPTQTYDIFF"), gsQtyFmt)
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "RPTITMID")
        
   
      
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
        
       
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function




Private Sub cmdSave(ByVal inActFlg As Integer)

 
    
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
                'If Chk_GrdRow(wlCtr) = False Then
                '    tblDetail.SetFocus
                '    Exit Function
                'End If
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
Private Function Chk_grdQty(inCode As String) As Boolean
    
    Chk_grdQty = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入數量!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdQty = False
        Exit Function
    End If

    If To_Value(inCode) = 0 Then
        gsMsg = "數量必需大於零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdQty = False
        Exit Function
    End If
    
End Function


Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    
    
    
    Me.MousePointer = vbHourglass
    
    
   ' If InputValidation() = False Then
    '   MousePointer = vbDefault
    '   Exit Sub
    'End If
    
    
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = ""
    wsSelection(2) = ""
    wsSelection(3) = ""
    wsSelection(4) = ""
    
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTSTKCNT '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'盤點數目報表', "
    wsSQL = wsSQL & "'" & Opt_Getfocus(optShow, 2, 0) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmCodeFr) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), Set_Quote(cboItmCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmBarCodeFr) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmBarCodeTo.Text) = "", String(15, "z"), Set_Quote(cboItmBarCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboWhsCodeFr) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboWhsCodeTo.Text) = "", String(10, "z"), Set_Quote(cboWhsCodeTo.Text)) & "', "
    'wsSql = wsSql & "'" & IIf(Trim(medPrdFr.Text) = "/", "0000/00", medPrdFr.Text) & "', "
    'wsSql = wsSql & "'" & IIf(Trim(medPrdTo.Text) = "/", "9999/99", medPrdTo.Text) & "', "
    'wsSql = wsSql & "" & To_Value(txtQtyLvlFr) & ", "
    'wsSql = wsSql & "" & IIf(To_Value(txtQtyLvlTo.Text) = 0, "999999999.99", To_Value(txtQtyLvlTo.Text)) & ", "
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTSTKCNT"
    Else
    wsRptName = "RPTSTKCNT"
    End If
    
    NewfrmPrint.ReportID = "STKCNT"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "STKCNT"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
        
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
    
    
    
End Sub
