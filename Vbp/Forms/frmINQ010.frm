VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmINQ010 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmINQ010.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      OleObjectBlob   =   "frmINQ010.frx":0442
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   4575
   End
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
            Picture         =   "frmINQ010.frx":2B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":341F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":3CF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":414B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":459D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":48B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":4D09
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":515B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":5475
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":578F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":5BE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":64BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":67E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":6C39
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":6F55
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":7271
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":76C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":79E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":7CFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":8151
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmINQ010.frx":846D
            Key             =   ""
         EndProperty
      EndProperty
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
   Begin VB.ComboBox cboItmAccTypeCodeFr 
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   960
      Width           =   1812
   End
   Begin VB.ComboBox cboItmAccTypeCodeTo 
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
      Begin VB.Label lblItmAccTypeCodeFr 
         Caption         =   "Itm Barcode From"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label lblItmAccTypeCodeTo 
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
      OleObjectBlob   =   "frmINQ010.frx":8791
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
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Print"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "frmINQ010"
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
Private wsDteTim As String

Private Const tcPrint = "Print"
Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExit = "Exit"

Private Const SITMCODE = 0
Private Const SITMNAME = 1
Private Const SITMTYPE = 2
Private Const STRNCODE = 3
Private Const SDOCDATE = 4
Private Const SDOCNO = 5
Private Const SCURR = 6
Private Const SUPRICE = 7
Private Const SQTY = 8
Private Const SID = 9
Private Const SDUMMY = 10




Private Sub cboItmAccTypeCodeFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    wsSQL = "SELECT ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmBarCode LIKE '%" & IIf(cboItmAccTypeCodeFr.SelLength > 0, "", Set_Quote(cboItmAccTypeCodeFr.Text)) & "%' "
    wsSQL = wsSQL & " AND ITMSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmBarCode "
    
    Call Ini_Combo(2, wsSQL, cboItmAccTypeCodeFr.Left, cboItmAccTypeCodeFr.Top + cboItmAccTypeCodeFr.Height, tblCommon, wsFormID, "TBLItmAccTypeCode", Me.Width, Me.Height)
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmAccTypeCodeFr_GotFocus()
        FocusMe cboItmAccTypeCodeFr
    Set wcCombo = cboItmAccTypeCodeFr
End Sub

Private Sub cboItmAccTypeCodeFr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmAccTypeCodeFr, 13, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Trim(cboItmAccTypeCodeFr.Text) <> "" And _
            Trim(cboItmAccTypeCodeTo.Text) = "" Then
            cboItmAccTypeCodeTo.Text = cboItmAccTypeCodeFr.Text
        End If
        cboItmAccTypeCodeTo.SetFocus
    End If
End Sub


Private Sub cboItmAccTypeCodeFr_LostFocus()
    FocusMe cboItmAccTypeCodeFr, True
End Sub

Private Sub cboItmAccTypeCodeTo_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    wsSQL = "SELECT ItmBarCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " "
    wsSQL = wsSQL & " FROM mstItem "
    wsSQL = wsSQL & " WHERE ItmBarCode LIKE '%" & IIf(cboItmAccTypeCodeTo.SelLength > 0, "", Set_Quote(cboItmAccTypeCodeTo.Text)) & "%' "
    wsSQL = wsSQL & " AND ITMSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY ItmBarCode "
    
    Call Ini_Combo(2, wsSQL, cboItmAccTypeCodeTo.Left, cboItmAccTypeCodeTo.Top + cboItmAccTypeCodeTo.Height, tblCommon, wsFormID, "TBLItmAccTypeCode", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmAccTypeCodeTo_GotFocus()
    FocusMe cboItmAccTypeCodeTo
    Set wcCombo = cboItmAccTypeCodeTo
End Sub

Private Sub cboItmAccTypeCodeTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmAccTypeCodeTo, 13, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmAccTypeCodeTo = False Then
            Exit Sub
        End If
        
        cboItmTypeCodeFr.SetFocus
    End If
End Sub



Private Sub cboItmAccTypeCodeTo_LostFocus()
FocusMe cboItmAccTypeCodeTo, True
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
        
       cboItmAccTypeCodeFr.SetFocus
        
        
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

Private Function chk_cboItmAccTypeCodeTo() As Boolean
    chk_cboItmAccTypeCodeTo = False
    
    If UCase(cboItmAccTypeCodeFr.Text) > UCase(cboItmAccTypeCodeTo.Text) Then
        gsMsg = "To > From!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboItmAccTypeCodeTo.SetFocus
        Exit Function
    End If
    
    chk_cboItmAccTypeCodeTo = True
End Function



Private Sub cboItmTypeCodeFr_DropDown()
   Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItmTypeCodeFr
  
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ItmTypeEngDesc", "ItmTypeChiDesc") & " FROM MstItemType "
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
  
    wsSQL = "SELECT ItmTypeCode, " & IIf(gsLangID = "1", "ItmTypeEngDesc", "ItmTypeChiDesc") & " FROM MstItemType "
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
    Call chk_InpLen(cboItmTypeCodeTo, 10, KeyAscii)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboItmTypeCodeTo = False Then
            Call cboItmTypeCodeTo_GotFocus
            Exit Sub
        End If
        
        If LoadRecord Then
            tblDetail.SetFocus
        End If
        
        
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
       
            
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
            
     
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
    
    waResult.ReDim 0, -1, SITMCODE, SID
  
    
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
    
    
    
    cboItmCodeFr.Text = ""
    cboItmCodeTo.Text = ""
    cboItmAccTypeCodeFr.Text = ""
    cboItmAccTypeCodeTo.Text = ""
    cboItmTypeCodeFr.Text = ""
    cboItmTypeCodeTo.Text = ""

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   
    cnCon.Execute "DELETE FROM RPTINQ010 WHERE RPTUSRID = '" & gsUserID & "' AND RPTDTETIM = '" & wsDteTim & "' "

    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waResult = Nothing
    Set frmINQ010 = Nothing


    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsFormID = "INQ010"

    
End Sub

Private Sub Ini_Caption()
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblItmCodeFr.Caption = Get_Caption(waScrItm, "ItmCodeFR")
    lblItmCodeTo.Caption = Get_Caption(waScrItm, "ItmCodeTO")
    lblItmAccTypeCodeFr.Caption = Get_Caption(waScrItm, "ItmAccTypeCodeFR")
    lblItmAccTypeCodeTo.Caption = Get_Caption(waScrItm, "ItmAccTypeCodeTO")
    lblItmTypeCodeFr.Caption = Get_Caption(waScrItm, "ItmTypeCodeFR")
    lblItmTypeCodeTo.Caption = Get_Caption(waScrItm, "ItmTypeCodeTO")

    
    
    
    
    With tblDetail
        .Columns(SITMCODE).Caption = Get_Caption(waScrItm, "SITMCODE")
        .Columns(SITMNAME).Caption = Get_Caption(waScrItm, "SITMNAME")
        .Columns(SUPRICE).Caption = Get_Caption(waScrItm, "SUPRICE")
        .Columns(SITMTYPE).Caption = Get_Caption(waScrItm, "SITMTYPE")
        .Columns(SQTY).Caption = Get_Caption(waScrItm, "SQTY")
        .Columns(STRNCODE).Caption = Get_Caption(waScrItm, "STRNCODE")
        .Columns(SDOCDATE).Caption = Get_Caption(waScrItm, "SDOCDATE")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SCURR).Caption = Get_Caption(waScrItm, "SCURR")
        
    End With
    
    
    
    
'    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint) & "(F11)"
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





Private Sub tblDetail_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim wlRet As Integer
    Dim wlRow As Long
    
    On Error GoTo tblDetail_KeyDown_Err
    
    With tblDetail
        Select Case KeyCode

            
        Case vbKeyReturn
            Select Case .Col
            Case SQTY
                 KeyCode = vbKeyDown
                 .Col = SITMCODE
            Case Else
                 KeyCode = vbDefault
                 .Col = .Col + 1
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> SITMCODE Then
                .Col = .Col - 1
            End If
        Case vbKeyRight
            Select Case .Col
                Case SQTY
                KeyCode = vbKeyDown
                    .Col = SITMCODE
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
    
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
       
        
    
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
            
        Case tcPrint
        
            Call cmdPrint
            
        Case tcCancel
        
           Call cmdCancel

            
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
        
        For wiCtr = SITMCODE To SDUMMY
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = True
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case SITMCODE
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 2000
                Case SITMNAME
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Width = 2000
                Case SUPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                Case SITMTYPE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 10
                Case SQTY
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsQtyFmt
                Case STRNCODE
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 3
                Case SDOCDATE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
                Case SDOCNO
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 20
                Case SCURR
                    .Columns(wiCtr).Width = 1000
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
    
    
    wsDteTim = Change_SQLDate(Now)
    
    Call cmdSave
    
    wsSQL = "SELECT RPTITMID, RPTITMCODE, RPTITMNAME, RPTUPRICE, RPTITMTYPECODE, RPTQTY, RPTTRNCODE, RPTDOCDATE, RPTDOCNO, RPTCURR "
    wsSQL = wsSQL & " From RPTINQ010 "
    wsSQL = wsSQL & " WHERE RPTUSRID = '" & gsUserID & "' "
    wsSQL = wsSQL & " AND RPTDTETIM = '" & wsDteTim & "' "
    wsSQL = wsSQL & " ORDER BY RPTDOCDATE, RPTITMCODE "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, SITMCODE, SID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
    
     
    With waResult
    .ReDim 0, -1, SITMCODE, SID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF

       .AppendRows
        waResult(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "RPTITMCODE")
        waResult(.UpperBound(1), SITMNAME) = ReadRs(rsRcd, "RPTITMNAME")
        waResult(.UpperBound(1), SITMTYPE) = ReadRs(rsRcd, "RPTITMTYPECODE")
        waResult(.UpperBound(1), SQTY) = Format(ReadRs(rsRcd, "RPTQTY"), gsQtyFmt)
        waResult(.UpperBound(1), SUPRICE) = Format(ReadRs(rsRcd, "RPTUPRICE"), gsUprFmt)
        waResult(.UpperBound(1), STRNCODE) = ReadRs(rsRcd, "RPTTRNCODE")
        waResult(.UpperBound(1), SDOCDATE) = ReadRs(rsRcd, "RPTDOCDATE")
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "RPTDOCNO")
        waResult(.UpperBound(1), SCURR) = ReadRs(rsRcd, "RPTCURR")
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "RPTITMID")
        
        'End If

      
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

Private Sub cmdPrint()
    
    
End Sub

Private Sub cmdSave()
    Dim adcmdSave As New ADODB.Command

     
    On Error GoTo cmdSave_Err
    
    'MousePointer = vbHourglass
    
    
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    adcmdSave.CommandText = "USP_RPTINQ010"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
     
    Call SetSPPara(adcmdSave, 1, gsUserID)
    Call SetSPPara(adcmdSave, 2, wsDteTim)
    Call SetSPPara(adcmdSave, 3, cboItmCodeFr)
    Call SetSPPara(adcmdSave, 4, cboItmCodeTo)
    Call SetSPPara(adcmdSave, 5, cboItmAccTypeCodeFr)
    Call SetSPPara(adcmdSave, 6, cboItmAccTypeCodeTo)
    Call SetSPPara(adcmdSave, 7, cboItmTypeCodeFr)
    Call SetSPPara(adcmdSave, 8, cboItmTypeCodeTo)
    Call SetSPPara(adcmdSave, 9, gsLangID)
    
    adcmdSave.Execute
        
    cnCon.CommitTrans
    
    
    
    'Call UnLockAll(wsConnTime, wsFormID)
    Set adcmdSave = Nothing
    
    
  '  MousePointer = vbDefault
    
    Exit Sub
    
cmdSave_Err:
    MsgBox Err.Description
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing
    
End Sub

