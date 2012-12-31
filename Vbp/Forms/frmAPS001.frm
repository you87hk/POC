VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAPS001 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "Stock Reserve"
   ClientHeight    =   8625
   ClientLeft      =   75
   ClientTop       =   1005
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmAPS001.frx":0000
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
      OleObjectBlob   =   "frmAPS001.frx":0442
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
            Index           =   1
            Left            =   2040
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   200
            Value           =   -1  'True
            Width           =   1455
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
         Top             =   960
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
      OleObjectBlob   =   "frmAPS001.frx":2B45
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
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":B278
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":BB52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":C42C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":C87E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":CCD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":CFEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":D43C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":D88E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":DBA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":DEC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":E314
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":EBF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":EF18
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":F36C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":F688
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":F9A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":FDF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":10114
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":10434
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":10754
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAPS001.frx":10BA8
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
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Can"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Finish"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "取消 (F3)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAll"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DAll"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "frmAPS001"
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

Private Const tcCan = "Can"
Private Const tcFinish = "Finish"

Private Const tcRefresh = "Refresh"
Private Const tcCancel = "Cancel"
Private Const tcExport = "Export"

Private Const tcExit = "Exit"
Private Const tcSAll = "SAll"
Private Const tcDAll = "DAll"


Private Const SSEL = 0
Private Const STRNCODE = 1
Private Const STRNNAME = 2
Private Const SDOCNO = 3
Private Const SDOCDATE = 4
Private Const SCUSCODE = 5
Private Const SPRTFLG = 6
Private Const SAPRFLG = 7
Private Const SUPDUSR = 8
Private Const SUPDDATE = 9
Private Const SDUMMY = 10
Private Const SID = 11



Private Sub cboCusNoFr_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    wsSQL = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboCusNoFr.SelLength > 0, "", Set_Quote(cboCusNoFr.Text)) & "%' "
    wsSQL = wsSQL & "AND SaleStatus = '1' "
    wsSQL = wsSQL & "AND SaleType = 'W' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
    
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
    
  
    wsSQL = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboCusNoTo.SelLength > 0, "", Set_Quote(cboCusNoTo.Text)) & "%' "
    wsSQL = wsSQL & "AND SaleStatus = '1' "
    wsSQL = wsSQL & "AND SaleType = 'W' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
    
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
  
    wsSQL = "SELECT SJHDDOCNO, SALECODE, SJHDDOCDATE "
    wsSQL = wsSQL & " FROM ICSTKADJ, mstSALESMAN "
    wsSQL = wsSQL & " WHERE SJHDDOCNO LIKE '%" & IIf(cboDocNoFr.SelLength > 0, "", Set_Quote(cboDocNoFr.Text)) & "%' "
    wsSQL = wsSQL & " AND SJHDSTAFFID  *= SALEID "
    wsSQL = wsSQL & " AND SJHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SJHDDOCNO "
    
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
  
    wsSQL = "SELECT SJHDDOCNO, SALECODE, SJHDDOCDATE "
    wsSQL = wsSQL & " FROM ICSTKADJ, mstSALESMAN "
    wsSQL = wsSQL & " WHERE SJHDDOCNO LIKE '%" & IIf(cboDocNoTo.SelLength > 0, "", Set_Quote(cboDocNoTo.Text)) & "%' "
    wsSQL = wsSQL & " AND SJHDSTAFFID  *= SALEID "
    wsSQL = wsSQL & " AND SJHDSTATUS = '1' "
    wsSQL = wsSQL & " ORDER BY SJHDDOCNO "
    
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
        
        Case vbKeyF3
          If tbrProcess.Buttons(tcCan).Enabled = False Then Exit Sub
     
           Call cmdSave(3)
       
        Case vbKeyF10
        
        If tbrProcess.Buttons(tcFinish).Enabled = False Then Exit Sub
       
           Call cmdSave(5)
           
        
        Case vbKeyF8
        
        If tbrProcess.Buttons(tcExport).Enabled = False Then Exit Sub
       
           Call cmdExport
           
        Case vbKeyF11
           Call cmdCancel
            
        Case vbKeyF12
            Unload Me
             
        Case vbKeyF5
           Call cmdSelect(1)
           
        Case vbKeyF6
           Call cmdSelect(0)
        
        Case vbKeyF5
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
            
        Case tcCan
            Call cmdSave(3)
            
                
        Case tcFinish
            Call cmdSave(5)
            
        Case tcExport
            Call cmdExport
            
        
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
    Set frmAPS001 = Nothing
    
End Sub



Private Sub IniForm()
    Me.KeyPreview = True
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    optDocType(0).Value = True
 '   wsFormID = "APS001"
End Sub


Private Sub Set_tbrProcess()

With tbrProcess
    
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
    .Buttons(tcCan).Enabled = False
    .Buttons(tcFinish).Enabled = True
    Else
    .Buttons(tcCan).Enabled = False
    .Buttons(tcFinish).Enabled = False
    End If
    
    .Buttons(tcExport).Enabled = True
    .Buttons(tcRefresh).Enabled = True
    .Buttons(tcCancel).Enabled = True
    .Buttons(tcSAll).Enabled = True
    .Buttons(tcDAll).Enabled = True
    .Buttons(tcExit).Enabled = True
    
    
    
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
                
    
    
    With tblDetail
        .Columns(SSEL).Caption = Get_Caption(waScrItm, "SSEL")
        .Columns(STRNCODE).Caption = Get_Caption(waScrItm, "STRNCODE")
        .Columns(STRNNAME).Caption = Get_Caption(waScrItm, "STRNNAME")
        .Columns(SDOCNO).Caption = Get_Caption(waScrItm, "SDOCNO")
        .Columns(SCUSCODE).Caption = Get_Caption(waScrItm, "SCUSCODE")
        .Columns(SDOCDATE).Caption = Get_Caption(waScrItm, "SDOCDATE")
        .Columns(SAPRFLG).Caption = Get_Caption(waScrItm, "SAPRFLG")
        .Columns(SPRTFLG).Caption = Get_Caption(waScrItm, "SPRTFLG")
        .Columns(SUPDUSR).Caption = Get_Caption(waScrItm, "SUPDUSR")
        .Columns(SUPDDATE).Caption = Get_Caption(waScrItm, "SUPDDATE")
    End With
    
    With tbrProcess
    .Buttons(tcCan).ToolTipText = Get_Caption(waScrItm, tcCan) & "(F3)"
    .Buttons(tcFinish).ToolTipText = Get_Caption(waScrItm, tcFinish) & "(F10)"
    .Buttons(tcExport).ToolTipText = Get_Caption(waScrItm, tcExport) & "(F8)"
    .Buttons(tcRefresh).ToolTipText = Get_Caption(waScrItm, tcRefresh) & "(F7)"
    .Buttons(tcCancel).ToolTipText = Get_Caption(waScrItm, tcCancel) & "(F11)"
    .Buttons(tcSAll).ToolTipText = Get_Caption(waScrItm, tcSAll) & "(F5)"
    .Buttons(tcDAll).ToolTipText = Get_Caption(waScrItm, tcDAll) & "(F6)"
    .Buttons(tcExit).ToolTipText = Get_Caption(waScrItm, tcExit) & "(F12)"
    
   End With

End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .Update
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
                    
                    frmAPS0011.InDocID = .Columns(SID).Text
                    frmAPS0011.InCusNo = .Columns(SCUSCODE).Text
                    frmAPS0011.FormID = wsFormID & "1"
                    frmAPS0011.Show vbModal
                    
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
            Case SUPDDATE
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
                Case SUPDDATE
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
                Case STRNCODE
                    .Columns(wiCtr).DataWidth = 3
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).Button = True
                Case STRNNAME
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                Case SDOCNO
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                Case SCUSCODE
                   .Columns(wiCtr).Width = 1500
                   .Columns(wiCtr).DataWidth = 10
                Case SDOCDATE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 10
                Case SAPRFLG
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).DataWidth = 5
                Case SPRTFLG
                    .Columns(wiCtr).Width = 700
                    .Columns(wiCtr).DataWidth = 5
                Case SUPDUSR
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 20
                Case SUPDDATE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 10
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
    Dim wdCreLmt As Double
    Dim wdCreLft As Double
    Dim wsStatus As String
    
    Me.MousePointer = vbHourglass
    LoadRecord = False
    
    
    If Opt_Getfocus(optDocType, 2, 0) = 0 Then
     wsStatus = "1"
    Else
     wsStatus = "4"
    End If
    
    
    wsSQL = "SELECT SJHDTRNCODE TRNCODE, SCDDESC TRNNAME, SJHDDOCID DOCID, SJHDDOCNO DOCNO, SJHDDOCDATE DOCDATE, SALECODE STAFFCODE, SJHDPRTFLG PRTFLG, SJHDAPRFLG APRFLG, SJHDUPDUSR UPDUSR, SJHDUPDDATE UPDDATE "
    wsSQL = wsSQL & " FROM  ICSTKADJ, MSTSALESMAN, SYSCODEDESC "
    wsSQL = wsSQL & " WHERE SJHDDOCNO BETWEEN '" & cboDocNoFr & "' AND '" & IIf(Trim(cboDocNoTo.Text) = "", String(15, "z"), Set_Quote(cboDocNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SALECODE BETWEEN '" & cboCusNoFr & "' AND '" & IIf(Trim(cboCusNoTo.Text) = "", String(10, "z"), Set_Quote(cboCusNoTo.Text)) & "'"
    wsSQL = wsSQL & " AND SJHDCTLPRD BETWEEN '" & IIf(Trim(medPrdFr.Text) = "/", String(6, "000000"), Left(medPrdFr.Text, 4) & Right(medPrdFr.Text, 2)) & "'"
    wsSQL = wsSQL & " AND '" & IIf(Trim(medPrdTo.Text) = "/", String(6, "999999"), Left(medPrdTo.Text, 4) & Right(medPrdTo.Text, 2)) & "'"
    wsSQL = wsSQL & " AND SJHDTRNCODE = SCDCODE "
    wsSQL = wsSQL & " AND SJHDSTAFFID *= SALEID "
    wsSQL = wsSQL & " AND SCDLANGID = '" & gsLangID & "' "
    wsSQL = wsSQL & " AND SJHDSTATUS = '" & wsStatus & "'"
    'wsSQL = wsSQL & " ORDER BY SJHDTRNCODE, SJHDDOCDATE, SJHDDOCNO "
    
    
    If wiSort = 0 Then
    wsSQL = wsSQL & " ORDER BY SJHDDOCNO " & wsSortBy
    ElseIf wiSort = 1 Then
    wsSQL = wsSQL & " ORDER BY SJHDDOCDATE " & wsSortBy
    Else
    wsSQL = wsSQL & " ORDER BY SJHDDOCDATE, SJHDDOCNO " & wsSortBy
    End If
    
    
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
        waResult(.UpperBound(1), STRNCODE) = ReadRs(rsRcd, "TRNCODE")
        waResult(.UpperBound(1), STRNNAME) = ReadRs(rsRcd, "TRNNAME")
        waResult(.UpperBound(1), SDOCNO) = ReadRs(rsRcd, "DOCNO")
        waResult(.UpperBound(1), SCUSCODE) = ReadRs(rsRcd, "STAFFCODE")
        waResult(.UpperBound(1), SDOCDATE) = ReadRs(rsRcd, "DOCDATE")
        waResult(.UpperBound(1), SAPRFLG) = ReadRs(rsRcd, "APRFLG")
        waResult(.UpperBound(1), SPRTFLG) = ReadRs(rsRcd, "PRTFLG")
        waResult(.UpperBound(1), SUPDUSR) = ReadRs(rsRcd, "UPDUSR")
        waResult(.UpperBound(1), SUPDDATE) = ReadRs(rsRcd, "UPDDATE")
        waResult(.UpperBound(1), SID) = ReadRs(rsRcd, "DOCID")
        
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

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wsCusNo As String
    Dim wsStorep As String
     
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
    Case 3
    gsMsg = "你是否取消此文件?"
    Case 5
    gsMsg = "你是否確認完成此文件?"
    End Select
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If

    wsMark = "0"
    
   
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
 
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_APS001A"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SSEL)) = "-1" Then
                Call SetSPPara(adcmdSave, 1, inActFlg)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, STRNCODE))
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SID))
                Call SetSPPara(adcmdSave, 4, "")
                Call SetSPPara(adcmdSave, 5, wsFormID)
                Call SetSPPara(adcmdSave, 6, gsUserID)
                Call SetSPPara(adcmdSave, 7, wsGenDte)
 
                wsMark = waResult(wiCtr, SID)
                wsCusNo = waResult(wiCtr, SCUSCODE)
                adcmdSave.Execute
                wsDocNo = GetSPPara(adcmdSave, 8)
            End If
        Next
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
        If wsDocNo <> "" Then
        gsMsg = "物料:" & wsDocNo & "不足!不能退貨!"
        Else
        gsMsg = "文件已完成!"
        End If
    End Select
    MsgBox gsMsg, vbOKOnly, gsTitle
        

    
    Set adcmdSave = Nothing
    
    Call LoadRecord
    
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
           End Select
    End With

    
    Exit Sub
    
tblDetail_HeadClick_Err:
     MsgBox "Check tblDeiail HeadClick!"

End Sub



Private Sub cmdExport()

    Dim wsGenDte As String
    'Dim adcmdExport As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsTrnCode As String
    Dim wsVdrNo As String
    Dim wsStorep As String
    Dim wiMod As Integer
    Dim wsPath As String
    Dim wsDoc As String
    
    Dim wsWhsCd As String
    
     
    On Error GoTo cmdExport_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate

    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Sub
    End If

    '' Last Check when Add
    
    wsWhsCd = GetWhsSelect
    
    If wsWhsCd = "" Then
       gsMsg = "No Warehouse Code!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       MousePointer = vbDefault
       Exit Sub
    End If


    Select Case wsWhsCd
    Case "CF-HK"
    gsMsg = "你是否在(香港倉)匯出文件"
    Case "CF-HZ"
    gsMsg = "你是否在(鶴山倉)匯出文件"
    Case Else
    gsMsg = "你是否在(香港倉)匯出文件"
    wsWhsCd = "CF-HK"
    End Select
        
    
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       MousePointer = vbDefault
       Exit Sub
    End If
    
    wsTrnCode = "TR"
    

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
            
            
            
            If ExportToHHFile(wsPath, wsTrnCode, waResult(wiCtr, SID), wiMod, wsWhsCd) = False Then
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

Private Function GetWhsSelect() As String

    Dim Newfrm As New frmSelectWhs
    
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
      '  Set .ctlKey = GetCompanySelect
        .Show vbModal
    GetWhsSelect = .ctlKey
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Function
