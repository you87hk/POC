VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmMRP001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmMRP001.frx":0000
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
            Picture         =   "frmMRP001.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":1A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":1E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":21B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":2606
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":2A58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":2D72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":308C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":34DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":3DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":40E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":4536
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":4852
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":4B6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":4FC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":52DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":55FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":5A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMRP001.frx":5D6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9360
      OleObjectBlob   =   "frmMRP001.frx":608E
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboItmTypeCodeFr 
      Height          =   300
      Left            =   2520
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1812
   End
   Begin VB.ComboBox cboItmTypeCodeTo 
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
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   11775
      Begin VB.Frame fraSelect 
         Height          =   1245
         Left            =   7320
         TabIndex        =   17
         Top             =   120
         Width           =   4335
         Begin VB.TextBox txtQtyLvlTo 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   21
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtQtyLvlFr 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   2160
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
         Begin VB.OptionButton optType 
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
            Width           =   3135
         End
         Begin VB.OptionButton optType 
            Caption         =   "SO"
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
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblQtyLvlTo 
            Caption         =   "To"
            Height          =   225
            Left            =   3000
            TabIndex        =   22
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Label lblItmTypeCodeTo 
         Caption         =   "To"
         Height          =   225
         Left            =   4560
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblItmTypeCodeFr 
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
      Height          =   6255
      Left            =   120
      OleObjectBlob   =   "frmMRP001.frx":8791
      TabIndex        =   6
      Top             =   1920
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
Attribute VB_Name = "frmMRP001"
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
Private wsDte As String


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
Private Const SITMTYPE = 3
Private Const SUPRICE = 4
Private Const SSTKUOM = 5
Private Const SSTKQTY = 6
Private Const SRDRQTY = 7
Private Const SVDRCODE = 8
Private Const SCURR = 9
Private Const SPOPRICE = 10
Private Const SPRCUOM = 11
Private Const SPOQTY = 12
Private Const SDUMMY = 13
Private Const SID = 14



Private Sub cboItmBarCodeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    wsSQL = "SELECT ITMBARCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME "
    wsSQL = wsSQL & " FROM MSTITEM "
    wsSQL = wsSQL & " WHERE ITMBARCODE LIKE '%" & IIf(cboItmBarCodeFr.SelLength > 0, "", Set_Quote(cboItmBarCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
    
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
    Call chk_InpLen(cboItmBarCodeFr, 10, KeyAscii)
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
    wsSQL = "SELECT ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmBarCode LIKE '%" & IIf(cboItmBarCodeTo.SelLength > 0, "", Set_Quote(cboItmBarCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " AND ItmSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmCode "
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
    Call chk_InpLen(cboItmBarCodeTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmBarCodeTo = False Then
            Exit Sub
        End If
        
        cboItmTypeCodeFr.SetFocus
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
  
    wsSQL = "SELECT ItmCode, ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME "
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
  
    wsSQL = "SELECT ItmCode, ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME "
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



Private Sub cboItmTypeCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeFr
  
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", " ItmTypeEngDesc", " ItmTypeChiDesc") & " FROM MstItemType "
    wsSQL = wsSQL & " WHERE ItmTypeStatus = '1'"
    wsSQL = wsSQL & " AND ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeFr.SelLength > 0, "", Set_Quote(cboItmTypeCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    
    Call Ini_Combo(2, wsSQL, cboItmTypeCodeFr.Left, cboItmTypeCodeFr.Top + cboItmTypeCodeFr.Height, tblCommon, wsFormID, "TBLItmTypeCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmTypeCodeFr_GotFocus()
    FocusMe cboItmTypeCodeFr
    Set wcCombo = cboItmTypeCodeFr
End Sub

Private Sub cboItmTypeCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCodeFr, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Trim(cboItmTypeCodeFr.Text) <> "" And _
            Trim(cboItmTypeCodeTo.Text) = "" Then
            cboItmTypeCodeTo.Text = cboItmTypeCodeFr.Text
        End If
        cboItmTypeCodeTo.SetFocus
    End If
End Sub

Private Sub cboItmTypeCodeFr_LostFocus()
    FocusMe cboItmTypeCodeFr, True
End Sub

Private Sub cboItmTypeCodeTo_DropDown()
Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeTo
  
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", " ItmTypeEngDesc", " ItmTypeChiDesc") & " FROM MstItemType "
    wsSQL = wsSQL & " WHERE ItmTypeStatus = '1'"
    wsSQL = wsSQL & " AND ItmTypeCode LIKE '%" & IIf(cboItmTypeCodeTo.SelLength > 0, "", Set_Quote(cboItmTypeCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY ItmTypeCode "
    Call Ini_Combo(2, wsSQL, cboItmTypeCodeTo.Left, cboItmTypeCodeTo.Top + cboItmTypeCodeTo.Height, tblCommon, wsFormID, "TBLItmTypeCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmTypeCodeTo_GotFocus()
    FocusMe cboItmTypeCodeTo
    Set wcCombo = cboItmTypeCodeTo
End Sub

Private Sub cboItmTypeCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCodeTo, 15, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmTypeCodeTo = False Then
            Call cboItmTypeCodeTo_GotFocus
            Exit Sub
        End If
        
        Call Opt_Setfocus(optType, 2, 0)
        
        
    End If
End Sub

Private Sub cboItmTypeCodeTo_LostFocus()
    FocusMe cboItmTypeCodeTo, True
End Sub
Private Function chk_cboItmTypeCodeTo() As Boolean
    chk_cboItmTypeCodeTo = False
    
    If UCase(cboItmTypeCodeFr.Text) > UCase(cboItmTypeCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    End If
    
    chk_cboItmTypeCodeTo = True
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF10
        If tbrProcess.Buttons(tcOK).Enabled = False Then Exit Sub
           Call cmdSave(1)
            
        Case vbKeyF9
           Call cmdPrint
           
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

    
    optType(0).Value = True
    
    txtQtyLvlFr.Enabled = False
    txtQtyLvlTo.Enabled = False
 '   cboItmCodeFr.Text = ""
 '   cboItmCodeTo.Text = ""
 '   cboItmBarCodeFr.Text = ""
 '   cboItmBarCodeTo.Text = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   

    
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmMRP001 = Nothing

    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "MRP001"
    wsDte = Change_SQLDate(Now)
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblItmCodeFr.Caption = Get_Caption(waScrItm, "ItmCodeFR")
    lblItmCodeTo.Caption = Get_Caption(waScrItm, "ItmCodeTO")
    lblItmBarCodeFr.Caption = Get_Caption(waScrItm, "ItmBarCodeFR")
    lblItmBarCodeTo.Caption = Get_Caption(waScrItm, "ItmBarCodeTO")
    lblItmTypeCodeFr.Caption = Get_Caption(waScrItm, "ItmTypeCodeFR")
    lblItmTypeCodeTo.Caption = Get_Caption(waScrItm, "ItmTypeCodeTO")
        

    
    optType(0).Caption = Get_Caption(waScrItm, "OPT0")
    optType(1).Caption = Get_Caption(waScrItm, "OPT2")
    
    lblQtyLvlTo.Caption = Get_Caption(waScrItm, "QTYLVLTO")
    
    
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SItmCode")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SItmNAME")
        .Columns(SITMTYPE).Caption = Get_Caption(waScrItm, "SITMTYPE")
        .Columns(SUPRICE).Caption = Get_Caption(waScrItm, "SUPRICE")
        .Columns(SSTKUOM).Caption = Get_Caption(waScrItm, "SSTKUOM")
        .Columns(SSTKQTY).Caption = Get_Caption(waScrItm, "SSTKQTY")
        .Columns(SRDRQTY).Caption = Get_Caption(waScrItm, "SRDRQTY")
        .Columns(SPOQTY).Caption = Get_Caption(waScrItm, "SPOQTY")
        .Columns(SVDRCODE).Caption = Get_Caption(waScrItm, "SVDRCODE")
        .Columns(SCURR).Caption = Get_Caption(waScrItm, "SCURR")
        .Columns(SPOPRICE).Caption = Get_Caption(waScrItm, "SPOPRICE")
        .Columns(SPRCUOM).Caption = Get_Caption(waScrItm, "SPRCUOM")
        
    End With
    
    
    
    
    tbrProcess.Buttons(tcOK).ToolTipText = Get_Caption(waScrToolTip, tcOK) & "(F10)"
    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F9)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcSAll).ToolTipText = Get_Caption(waScrToolTip, tcSAll) & "(F5)"
    tbrProcess.Buttons(tcDAll).ToolTipText = Get_Caption(waScrToolTip, tcDAll) & "(F6)"
    
    

End Sub










Private Sub optType_Click(Index As Integer)
If optType(1).Value = True Then
    txtQtyLvlFr.Enabled = True
    txtQtyLvlTo.Enabled = True
Else
    txtQtyLvlFr.Enabled = False
    txtQtyLvlTo.Enabled = False
End If

End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
                
       If Index <> 1 Then
            If LoadRecord = True Then
                tblDetail.SetFocus
            End If
       Else
            txtQtyLvlFr.SetFocus
       End If
       
    End If
End Sub

Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    With tblDetail
        .Update
    End With
End Sub




Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim RetCurr As String
Dim RetUom As String
Dim RetCost As Double


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case SPOQTY
            
              If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
              End If

            Case SVDRCODE
                
              If Chk_grdVdrCode(.Columns(ColIndex).Text, .Columns(SITMCODE).Text, RetCurr, RetUom, RetCost) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
              End If
              
              .Columns(SCURR).Text = RetCurr
              .Columns(SPRCUOM).Text = RetUom
              .Columns(SPOPRICE).Text = RetCost
              
                
            Case SPOPRICE
            
             If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
              End If

            
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
    
    Dim wsSQL As String
    Dim wiTop As Long
    Dim wiCtr As Integer
    
    On Error GoTo tblDetail_ButtonClick_Err
    
    With tblDetail
        Select Case ColIndex
            Case SVDRCODE
                
            wsSQL = "SELECT VDRCODE, VDRNAME "
            wsSQL = wsSQL & " FROM mstVENDOR, mstVDRITEM, MSTITEM "
            wsSQL = wsSQL & " WHERE ITMCODE = '" & Set_Quote(.Columns(SITMCODE).Text) & "' "
            wsSQL = wsSQL & " AND VDRID = VDRITEMVDRID "
            wsSQL = wsSQL & " AND ITMID = VDRITEMITMID "
            wsSQL = wsSQL & " AND VDRITEMSTATUS <> '2' "
            wsSQL = wsSQL & " ORDER BY VDRCODE "
            
                
            Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLVDRCODE", Me.Width, Me.Height)
            tblCommon.Visible = True
            tblCommon.SetFocus
            Set wcCombo = tblDetail
                
                
           End Select
    End With
    
    Exit Sub
    
tblDetail_ButtonClick_Err:
     MsgBox "Check tblDetail ButtonClick!"
 
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
            Case SSEL
                 KeyCode = vbDefault
                 .Col = SPOQTY
            Case SPOQTY
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
                Case SPOQTY
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
        
        Case SPOQTY
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        Case SPOPRICE
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
            
       
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
                    
    
                Case SPOQTY
                 
               '  Call Chk_grdQty(waResult(LastRow, SPOQTY))
                 
                Case SVDRCODE
                 
                 Call Chk_grdVdrCode(waResult(LastRow, SVDRCODE), waResult(LastRow, SITMCODE), "", "", 0)
                 
                Case SPOPRICE
                 
                 Call Chk_grdQty(waResult(LastRow, SPOPRICE))
                 
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
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case SVDRCODE
               wcCombo.Text = tblCommon.Columns(0).Text
          Case Else
               wcCombo.Text = tblCommon.Columns(0).Text
       End Select
    Else
       wcCombo.Text = tblCommon.Columns(0).Text
    End If
    
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
        If wcCombo.Name = tblDetail.Name Then
            tblDetail.EditActive = True
            Select Case wcCombo.Col
              Case SVDRCODE
                   wcCombo.Text = tblCommon.Columns(0).Text
              Case Else
                   wcCombo.Text = tblCommon.Columns(0).Text
           End Select
        Else
           wcCombo.Text = tblCommon.Columns(0).Text
        End If
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
                Case SITMCODE
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 2000
                Case SITMNAME
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Width = 1500
                Case SUPRICE
                   .Columns(wiCtr).Width = 1000
                   .Columns(wiCtr).DataWidth = 15
                   .Columns(wiCtr).NumberFormat = gsUprFmt
                Case SITMTYPE
                    .Columns(wiCtr).Width = 1300
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Visible = False
                Case SSTKUOM
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 10
                 Case SSTKQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SRDRQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case SPOQTY
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                    .Columns(wiCtr).Locked = False
                Case SVDRCODE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).Button = True
                Case SPOPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = False
                Case SPRCUOM
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 10
                Case SCURR
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 3
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
    Dim wdStkQty As Double
    Dim wsMth As String
    Dim wdRdrQty As Double
    
    wsMth = Mid(gsSystemDate, 6, 2)
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    
    Call cmdCrtRecord
    
    
    wsSQL = "SELECT RPTITMID, RPTITMCODE, RPTITMNAM, RPTITMTYPE, RPTUPRICE, RPTVDRID, RPTSTKUOM, "
    wsSQL = wsSQL & " RPTPOQTY,  RPTRDRLVL, RPTNEEDQTY, RPTVDRCODE, RPTCURR, RPTPOPRICE, RPTPRCUOM "
    wsSQL = wsSQL & " FROM  RPTMRP001 "
    wsSQL = wsSQL & " WHERE RPTUSRID  = '" & gsUserID & "' "
    wsSQL = wsSQL & " AND RPTDTETIM  = '" & wsDte & "' "
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
        waResult(.UpperBound(1), SITMTYPE) = ReadRs(rsRcd, "RPTITMTYPE")
        waResult(.UpperBound(1), SUPRICE) = Format(To_Value(ReadRs(rsRcd, "RPTUPRICE")), gsUprFmt)
        waResult(.UpperBound(1), SSTKUOM) = ReadRs(rsRcd, "RPTSTKUOM")
        waResult(.UpperBound(1), SSTKQTY) = Format(ReadRs(rsRcd, "RPTNEEDQTY"), gsQtyFmt)
        waResult(.UpperBound(1), SRDRQTY) = Format(ReadRs(rsRcd, "RPTRDRLVL"), gsQtyFmt)
        waResult(.UpperBound(1), SPOQTY) = Format(ReadRs(rsRcd, "RPTPOQTY"), gsQtyFmt)
        waResult(.UpperBound(1), SVDRCODE) = ReadRs(rsRcd, "RPTVDRCODE")
        waResult(.UpperBound(1), SCURR) = ReadRs(rsRcd, "RPTCURR")
        waResult(.UpperBound(1), SPOPRICE) = Format(To_Value(ReadRs(rsRcd, "RPTPOPRICE")), gsAmtFmt)
        waResult(.UpperBound(1), SPRCUOM) = ReadRs(rsRcd, "RPTPRCUOM")
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
        
        If Chk_grdQty(waResult(LastRow, SPOQTY)) = False Then
            .Col = SPOQTY
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdVdrCode(waResult(LastRow, SVDRCODE), waResult(LastRow, SITMCODE), "", "", 0) = False Then
            .Col = SVDRCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdQty(waResult(LastRow, SPOPRICE)) = False Then
            .Col = SPOPRICE
            .Row = LastRow
            Exit Function
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
    Dim wsFirstFlg As String
    Dim wiLine As Integer
    Dim wsCtlPrd As String
    Dim wsDteTim As String
    Dim wiflg As Integer
     
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
    Case 1
    gsMsg = "你是否確認此文件?"
    Case 2
    gsMsg = "你是否取消此文件?"
    End Select
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If
   
    wsDteTim = Change_SQLDate(Now)
    wsCtlPrd = Left(gsSystemDate, 4) & Mid(gsSystemDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_MRP001A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                
                    Call SetSPPara(adcmdSave, 1, wsDteTim)
                    Call SetSPPara(adcmdSave, 2, wsCtlPrd)
                    Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SID))
                    Call SetSPPara(adcmdSave, 4, waResult(wiCtr, SVDRCODE))
                    Call SetSPPara(adcmdSave, 5, waResult(wiCtr, SCURR))
                    Call SetSPPara(adcmdSave, 6, waResult(wiCtr, SPOQTY))
                    Call SetSPPara(adcmdSave, 7, waResult(wiCtr, SPOPRICE))
                    Call SetSPPara(adcmdSave, 8, wsFormID)
                    Call SetSPPara(adcmdSave, 9, gsUserID)
                    Call SetSPPara(adcmdSave, 10, wsGenDte)
                    adcmdSave.Execute
                    
            End If
        Next wiCtr
        
        adcmdSave.CommandText = "USP_MRP001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
        Call SetSPPara(adcmdSave, 1, gsUserID)
        Call SetSPPara(adcmdSave, 2, wsDteTim)
        Call SetSPPara(adcmdSave, 3, gsLangID)
        Call SetSPPara(adcmdSave, 4, wsFormID)
        Call SetSPPara(adcmdSave, 5, wsGenDte)
        adcmdSave.Execute
        
    End If
    
    
    
    cnCon.CommitTrans
    
    gsMsg = "已完成!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Call LoadRecord
    Set adcmdSave = Nothing
    
    
    MousePointer = vbDefault
    
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
    
    
    If Not chk_txtQtyLvlTo Then
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

Private Function Chk_grdVdrCode(ByVal inCode As String, ByVal inItmCode As String, ByRef OutCurr As String, ByRef OutUOM As String, ByRef OutCost As Double) As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset

    
    If Trim(inCode) = "" Then
        gsMsg = "沒有輸入供應商!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If


    wsSQL = "SELECT VdrItemCurr, VdrItemCost, VdrItemUomCode FROM MstVendor, mstVdrItem, MstItem "
    wsSQL = wsSQL & " WHERE VdrCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & " And ItmCode = '" & Set_Quote(inItmCode) & "' "
    wsSQL = wsSQL & " And VdrItemVdrID = VdrID "
    wsSQL = wsSQL & " And VdrItemItmID = ItmID "
    wsSQL = wsSQL & " And VdrItemStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutCurr = ReadRs(rsRcd, "VdrItemCurr")
        OutUOM = ReadRs(rsRcd, "VdrItemUomCode")
        OutCost = Format(ReadRs(rsRcd, "VdrItemCost"), gsAmtFmt)
        
        Chk_grdVdrCode = True
    Else
        OutCurr = ""
        OutUOM = ""
        OutCost = 0
    
        gsMsg = "沒有此供應商價格!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdVdrCode = False
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    rsRcd.Close
    Set rsRcd = Nothing
    
   
    
End Function
Private Sub cmdPrint()

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

    wsSQL = "EXEC usp_RPTMRP001 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & wsDte & "', "
    wsSQL = wsSQL & "'物料再訂購告示', "
    wsSQL = wsSQL & "'" & Opt_Getfocus(optType, 2, 0) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmCodeFr) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), Set_Quote(cboItmCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmBarCodeFr) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmBarCodeTo.Text) = "", String(13, "z"), Set_Quote(cboItmBarCodeTo.Text)) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboItmTypeCodeFr) & "', "
    wsSQL = wsSQL & "'" & IIf(Trim(cboItmTypeCodeTo.Text) = "", String(10, "z"), Set_Quote(cboItmTypeCodeTo.Text)) & "', "
    wsSQL = wsSQL & "" & To_Value(txtQtyLvlFr) & ", "
    wsSQL = wsSQL & "" & IIf(To_Value(txtQtyLvlTo.Text) = 0, "999999999.99", To_Value(txtQtyLvlTo.Text)) & ", "
    wsSQL = wsSQL & gsLangID
    
       
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTMRP001"
    Else
    wsRptName = "RPTMRP001"
    End If
    
    NewfrmPrint.ReportID = "MRP001"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "MRP001"
    NewfrmPrint.RptDteTim = wsDte
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
        
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
    
    
    
End Sub

Private Function chk_txtQtyLvlTo() As Boolean
    chk_txtQtyLvlTo = False
    
    If To_Value(txtQtyLvlFr.Text) > To_Value(txtQtyLvlTo.Text) Then
        gsMsg = "範圍錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    chk_txtQtyLvlTo = True
End Function
Private Sub txtQtyLvlFr_GotFocus()
    FocusMe txtQtyLvlFr
End Sub

Private Sub txtQtyLvlFr_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtQtyLvlFr, True, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
      

        txtQtyLvlTo.SetFocus
      
    End If
End Sub

Private Sub txtQtyLvlFr_LostFocus()
FocusMe txtQtyLvlFr, True
End Sub

Private Sub txtQtyLvlTo_GotFocus()
    FocusMe txtQtyLvlTo
End Sub

Private Sub txtQtyLvlTo_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtQtyLvlTo, True, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
      
    If chk_txtQtyLvlTo() = False Then
        Exit Sub
    End If
    
    If LoadRecord = True Then
        tblDetail.SetFocus
    End If
       
      
    End If
End Sub

Private Sub txtQtyLvlTo_LostFocus()
FocusMe txtQtyLvlTo, True
End Sub

Private Sub cmdCrtRecord()

    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
     
    On Error GoTo cmdCrtRecord_Err
    
    MousePointer = vbHourglass

    
    
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    adcmdSave.CommandText = "USP_RPTMRP001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, wsDte)
    Call SetSPPara(adcmdSave, 3, "")
    Call SetSPPara(adcmdSave, 4, Opt_Getfocus(optType, 2, 0))
    Call SetSPPara(adcmdSave, 5, cboItmCodeFr)
    Call SetSPPara(adcmdSave, 6, IIf(Trim(cboItmCodeTo.Text) = "", String(30, "z"), cboItmCodeTo.Text))
    Call SetSPPara(adcmdSave, 7, cboItmBarCodeFr)
    Call SetSPPara(adcmdSave, 8, IIf(Trim(cboItmBarCodeTo.Text) = "", String(30, "z"), cboItmBarCodeTo.Text))
    Call SetSPPara(adcmdSave, 9, cboItmTypeCodeFr)
    Call SetSPPara(adcmdSave, 10, IIf(Trim(cboItmTypeCodeTo.Text) = "", String(10, "z"), cboItmTypeCodeTo.Text))
    Call SetSPPara(adcmdSave, 11, txtQtyLvlFr.Text)
    Call SetSPPara(adcmdSave, 12, IIf(To_Value(txtQtyLvlTo.Text) = 0, "999999999.99", txtQtyLvlTo.Text))
    Call SetSPPara(adcmdSave, 13, gsLangID)
    adcmdSave.Execute
    
    
    cnCon.CommitTrans
    

    
    
    'Call UnLockAll(wsConnTime, wsFormID)

    Set adcmdSave = Nothing
    
    
    MousePointer = vbDefault
    
    Exit Sub
    
    
cmdCrtRecord_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub

