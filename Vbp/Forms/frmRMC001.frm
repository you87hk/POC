VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmRMC001 
   BackColor       =   &H8000000A&
   Caption         =   "對換率輸入"
   ClientHeight    =   5535
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   8580
   Icon            =   "frmRMC001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8580
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboExcCurr 
      Height          =   300
      Left            =   5520
      TabIndex        =   55
      Top             =   800
      Width           =   2730
   End
   Begin VB.ComboBox cboExcYr 
      Height          =   300
      Left            =   1440
      TabIndex        =   54
      Top             =   800
      Width           =   2730
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10080
      OleObjectBlob   =   "frmRMC001.frx":08CA
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame fraDetailInfo 
      Caption         =   "FRADETAILINFO"
      Height          =   5055
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Width           =   8355
      Begin VB.Frame fraRates 
         Caption         =   "FRARATES"
         Height          =   2895
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   7935
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   7
            Left            =   5400
            TabIndex        =   13
            Top             =   705
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   7
            Left            =   6435
            TabIndex        =   14
            Top             =   705
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   8
            Left            =   6435
            TabIndex        =   16
            Top             =   1005
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   8
            Left            =   5400
            TabIndex        =   15
            Top             =   1005
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   9
            Left            =   6435
            TabIndex        =   18
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   9
            Left            =   5400
            TabIndex        =   17
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   10
            Left            =   6435
            TabIndex        =   20
            Top             =   1680
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   10
            Left            =   5400
            TabIndex        =   19
            Top             =   1680
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   11
            Left            =   6435
            TabIndex        =   22
            Top             =   1980
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   11
            Left            =   5400
            TabIndex        =   21
            Top             =   1980
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   12
            Left            =   6435
            TabIndex        =   24
            Top             =   2295
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   12
            Left            =   5400
            TabIndex        =   23
            Top             =   2295
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1320
            TabIndex        =   1
            Top             =   705
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2355
            TabIndex        =   2
            Top             =   705
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   2355
            TabIndex        =   4
            Top             =   1005
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   1320
            TabIndex        =   3
            Top             =   1005
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   2355
            TabIndex        =   6
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   1320
            TabIndex        =   5
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   2355
            TabIndex        =   8
            Top             =   1680
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   1320
            TabIndex        =   7
            Top             =   1680
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   5
            Left            =   2355
            TabIndex        =   10
            Top             =   1980
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   5
            Left            =   1320
            TabIndex        =   9
            Top             =   1980
            Width           =   1000
         End
         Begin VB.TextBox txtExcBRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   6
            Left            =   2355
            TabIndex        =   12
            Top             =   2295
            Width           =   1000
         End
         Begin VB.TextBox txtExcRate 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Index           =   6
            Left            =   1320
            TabIndex        =   11
            Top             =   2295
            Width           =   1000
         End
         Begin VB.Label lblExcRate1 
            Caption         =   "EXCRATE"
            Height          =   240
            Left            =   5520
            TabIndex        =   53
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblExcBRate1 
            Caption         =   "EXCBRATE"
            Height          =   240
            Left            =   6480
            TabIndex        =   52
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblJul 
            Caption         =   "JUL"
            Height          =   240
            Left            =   4440
            TabIndex        =   51
            Top             =   750
            Width           =   1020
         End
         Begin VB.Label lblAug 
            Caption         =   "AUG"
            Height          =   240
            Left            =   4440
            TabIndex        =   50
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblSep 
            Caption         =   "SEP"
            Height          =   240
            Left            =   4440
            TabIndex        =   49
            Top             =   1395
            Width           =   1020
         End
         Begin VB.Label lblOct 
            Caption         =   "OCT"
            Height          =   240
            Left            =   4440
            TabIndex        =   48
            Top             =   1725
            Width           =   1020
         End
         Begin VB.Label lblNov 
            Caption         =   "NOV"
            Height          =   240
            Left            =   4440
            TabIndex        =   47
            Top             =   2055
            Width           =   1020
         End
         Begin VB.Label lblDec 
            Caption         =   "DEC"
            Height          =   240
            Left            =   4440
            TabIndex        =   46
            Top             =   2370
            Width           =   1020
         End
         Begin VB.Label lblExcRate 
            Caption         =   "EXCRATE"
            Height          =   240
            Left            =   1440
            TabIndex        =   45
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblExcBRate 
            Caption         =   "EXCBRATE"
            Height          =   240
            Left            =   2400
            TabIndex        =   44
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label lblJan 
            Caption         =   "JAN"
            Height          =   240
            Left            =   360
            TabIndex        =   43
            Top             =   750
            Width           =   1020
         End
         Begin VB.Label lblFeb 
            Caption         =   "FEB"
            Height          =   240
            Left            =   360
            TabIndex        =   42
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label lblMar 
            Caption         =   "MAR"
            Height          =   240
            Left            =   360
            TabIndex        =   41
            Top             =   1395
            Width           =   1020
         End
         Begin VB.Label lblApr 
            Caption         =   "APR"
            Height          =   240
            Left            =   360
            TabIndex        =   40
            Top             =   1725
            Width           =   1020
         End
         Begin VB.Label lblMay 
            Caption         =   "MAY"
            Height          =   240
            Left            =   360
            TabIndex        =   39
            Top             =   2055
            Width           =   1020
         End
         Begin VB.Label lblJun 
            Caption         =   "JUN"
            Height          =   240
            Left            =   360
            TabIndex        =   38
            Top             =   2370
            Width           =   1020
         End
      End
      Begin VB.TextBox txtExcID 
         Height          =   270
         Left            =   7800
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtExcDesc 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   0
         Top             =   990
         Width           =   6780
      End
      Begin VB.Frame FraKeyField 
         Height          =   615
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   8055
         Begin VB.Label lblExcYr 
            Caption         =   "EXCYR"
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
            TabIndex        =   36
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label lblExcCurr 
            Caption         =   "EXCCURR"
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
            Left            =   3960
            TabIndex        =   35
            Top             =   270
            Width           =   1260
         End
      End
      Begin VB.Label lblExcDesc 
         Caption         =   "EXCDESC"
         Height          =   240
         Left            =   360
         TabIndex        =   32
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label lblExcLastUpd 
         Caption         =   "EXCLASTUPD"
         Height          =   240
         Left            =   360
         TabIndex        =   29
         Top             =   4605
         Width           =   1140
      End
      Begin VB.Label lblExcLastUpdDate 
         Caption         =   "EXCLASTUPDDATE"
         Height          =   240
         Left            =   4200
         TabIndex        =   28
         Top             =   4605
         Width           =   1260
      End
      Begin VB.Label lblDspExcLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1500
         TabIndex        =   27
         Top             =   4560
         Width           =   2505
      End
      Begin VB.Label lblDspExcLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   5595
         TabIndex        =   26
         Top             =   4560
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
            Picture         =   "frmRMC001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRMC001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   30
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
Attribute VB_Name = "frmRMC001"
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

Private Const wsKeyType = "MstExchangeRate"
Private wsUsrId As String
Private wsTrnCd As String

Private Sub cboExcYr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboExcYr
    
    wsSql = "SELECT DISTINCT ExcYr FROM MstExchangeRate WHERE ExcStatus <> '2' "
    wsSql = wsSql & "ORDER BY ExcYr"
    Call Ini_Combo(1, wsSql, cboExcYr.Left, cboExcYr.Top + cboExcYr.Height, tblCommon, "EXC001", "TBLYR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboExcYr_GotFocus()
    FocusMe cboExcYr
End Sub

Private Sub cboExcYr_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, cboExcYr, False, False)
    Call chk_InpLen(cboExcYr, 4, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboExcCurr.SetFocus
    End If
End Sub

Private Sub cboExcYr_LostFocus()
    FocusMe cboExcYr, True
End Sub

Private Sub cboExcCurr_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboExcCurr
    
    wsSql = "SELECT DISTINCT ExcCurr FROM MstExchangeRate WHERE ExcStatus <> '2' "
    wsSql = wsSql & "ORDER BY ExcCurr"
    Call Ini_Combo(1, wsSql, cboExcCurr.Left, cboExcCurr.Top + cboExcCurr.Height, tblCommon, "EXC001", "TBLCURR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboExcCurr_GotFocus()
    FocusMe cboExcCurr
End Sub

Private Sub cboExcCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboExcCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboExcCurr = True Then
            Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboExcCurr_LostFocus()
    FocusMe cboExcCurr, True
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
    MousePointer = vbHourglass
    
    IniForm
    Ini_Caption
    Ini_Scr
    
    MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 5940
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
    Dim iCounter As Integer

    Select Case sStatus
        Case "Default"
            cboExcYr.Enabled = False
            cboExcCurr.Enabled = False
            txtExcDesc.Enabled = False
            
            For iCounter = 1 To 12
                txtExcRate(iCounter).Enabled = False
                txtExcBRate(iCounter).Enabled = False
                txtExcRate(iCounter) = 0
                txtExcBRate(iCounter) = 0
            Next
            
        Case "AfrActAdd"
            cboExcYr.Enabled = True
            cboExcCurr.Enabled = True
            
            txtExcDesc.Enabled = False
            
            For iCounter = 1 To 12
                txtExcRate(iCounter).Enabled = False
                txtExcBRate(iCounter).Enabled = False
                txtExcRate(iCounter) = 0
                txtExcBRate(iCounter) = 0
            Next
            
        Case "AfrActEdit"
            cboExcYr.Enabled = True
            cboExcCurr.Enabled = True
            
            txtExcDesc.Enabled = False
            
            For iCounter = 1 To 12
                txtExcRate(iCounter).Enabled = False
                txtExcBRate(iCounter).Enabled = False
            Next
            
        Case "AfrKey"
            cboExcYr.Enabled = False
            cboExcCurr.Enabled = False
            
            txtExcDesc.Enabled = True
            
            For iCounter = 1 To 12
                txtExcRate(iCounter).Enabled = True
                txtExcBRate(iCounter).Enabled = True
            Next
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    InputValidation = False
    
    If Chk_cboExcCurr = False Then
        Exit Function
    End If
    
    If Chk_AlltxtExcRate = False Then
        Exit Function
    End If
    
    If Chk_AlltxtExcBRate = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    Dim iExcMn As Integer
    
    wsSql = "SELECT MstExchangeRate.* "
    wsSql = wsSql + "From MstExchangeRate "
    wsSql = wsSql + "WHERE (((MstExchangeRate.ExcYr)='" & Set_Quote(cboExcYr) & "') "
    wsSql = wsSql + "AND ((MstExchangeRate.ExcCurr)='" & Set_Quote(cboExcCurr) & "') "
    wsSql = wsSql + "AND ((MstExchangeRate.ExcStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
    Else
        Me.txtExcDesc = ReadRs(rsRcd, "ExcDesc")
        Me.lblDspExcLastUpd = ReadRs(rsRcd, "ExcLastUpd")
        Me.lblDspExcLastUpdDate = ReadRs(rsRcd, "ExcLastUpdDate")
        
        rsRcd.MoveFirst
        
        Do While Not rsRcd.EOF
            iExcMn = ReadRs(rsRcd, "ExcMn")
            Me.txtExcRate(iExcMn) = Format(ReadRs(rsRcd, "ExcRate"), gsExrFmt)
            Me.txtExcBRate(iExcMn) = Format(ReadRs(rsRcd, "ExcBRate"), gsExrFmt)
            rsRcd.MoveNext
        Loop
        
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
    
    Set frmEXC001 = Nothing
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
  '  Me.Left = 0
  '  Me.Top = 0
  '  Me.Width = Screen.Width
  '  Me.Height = Screen.Height
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "EXC001"
    wsTrnCd = ""
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    lblExcYr.Caption = Get_Caption(waScrItm, "EXCYR")
    lblExcCurr.Caption = Get_Caption(waScrItm, "EXCCURR")
    lblExcDesc.Caption = Get_Caption(waScrItm, "EXCDESC")
    lblExcRate.Caption = Get_Caption(waScrItm, "EXCRATE")
    lblExcBRate.Caption = Get_Caption(waScrItm, "EXCBRATE")
    
    lblExcRate1.Caption = lblExcRate.Caption
    lblExcBRate1.Caption = lblExcBRate.Caption
    
    lblExcLastUpd.Caption = Get_Caption(waScrItm, "EXCLASTUPD")
    lblExcLastUpdDate.Caption = Get_Caption(waScrItm, "EXCLASTUPDDATE")
    
    lblJan.Caption = LoadMthDesc(1)
    lblFeb.Caption = LoadMthDesc(2)
    lblMar.Caption = LoadMthDesc(3)
    
    lblApr.Caption = LoadMthDesc(4)
    lblMay.Caption = LoadMthDesc(5)
    lblJun.Caption = LoadMthDesc(6)
    
    lblJul.Caption = LoadMthDesc(7)
    lblAug.Caption = LoadMthDesc(8)
    lblSep.Caption = LoadMthDesc(9)
    
    lblOct.Caption = LoadMthDesc(10)
    lblNov.Caption = LoadMthDesc(11)
    lblDec.Caption = LoadMthDesc(12)
    
    fraDetailInfo.Caption = Get_Caption(waScrItm, "FRADETAILINFO")
    fraRates.Caption = Get_Caption(waScrItm, "FRARATES")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"

   
    wsActNam(1) = Get_Caption(waScrItm, "EXCADD")
    wsActNam(2) = Get_Caption(waScrItm, "EXCEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "EXCDELETE")
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
    
    cboExcYr.SetFocus
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
            If RowLock(wsConnTime, wsKeyType, cboExcYr & cboExcCurr, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtExcDesc.SetFocus
End Sub

Public Function Chk_ExcRate(InExcYr, InExcCurr) As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    Chk_ExcRate = False
    
    wsSql = "SELECT MstExchangeRate.* "
    wsSql = wsSql + "From MstExchangeRate "
    wsSql = wsSql + "WHERE (((MstExchangeRate.ExcYr)='" + Set_Quote(InExcYr) + "') "
    wsSql = wsSql + "AND ((MstExchangeRate.ExcCurr)='" + Set_Quote(InExcCurr) + "') "
    wsSql = wsSql + "AND ((MstExchangeRate.ExcStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount <= 0 Then
        
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    Chk_ExcRate = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Function Chk_cboExcYr() As Boolean
    Chk_cboExcYr = False
    
    If Len(Trim(cboExcYr)) <> 4 Then
        gsMsg = "年份錯誤! 請輸入四位數字之年份!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboExcYr.SetFocus
        Exit Function
    End If
    
    Chk_cboExcYr = True
End Function

Private Function Chk_cboExcCurr() As Boolean
    Dim wsStatus As String
    
    Chk_cboExcCurr = False
    
    If Len(Trim(cboExcCurr)) = 0 And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入需要資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboExcCurr.SetFocus
        Exit Function
    End If
    
    If Chk_ExcCurr(cboExcYr, cboExcCurr, wsStatus) = False Then
        If wiAction <> AddRec Then
            gsMsg = "貨幣不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboExcCurr.SetFocus
            Exit Function
        End If
    Else
        If wiAction = AddRec Then
            If wsStatus = "2" Then
                gsMsg = "貨幣已存在但已無效!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                cboExcCurr.SetFocus
                Exit Function
            Else
                gsMsg = "貨幣已存在!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                cboExcCurr.SetFocus
                Exit Function
            End If
        End If
                
        If wsStatus = "2" Then
            gsMsg = "貨幣已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboExcCurr.SetFocus
            Exit Function
        End If
    End If
    
    Chk_cboExcCurr = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmEXC001
    
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
    Dim iCounter As Integer
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboExcYr & cboExcCurr, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_EXC001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, cboExcYr)
    Call SetSPPara(adcmdSave, 4, cboExcCurr)
    Call SetSPPara(adcmdSave, 5, txtExcDesc)
   
    For iCounter = 1 To 12
        Call SetSPPara(adcmdSave, 5 + iCounter, txtExcRate(iCounter))
        Call SetSPPara(adcmdSave, 17 + iCounter, txtExcBRate(iCounter))
    Next
    
    
    Call SetSPPara(adcmdSave, 30, gsUserID)
    Call SetSPPara(adcmdSave, 31, wsGenDte)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 32)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - EXC001!"
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
    
    ReDim vFilterAry(4, 2)
    vFilterAry(1, 1) = "年份"
    vFilterAry(1, 2) = "ExcYr"
    
    vFilterAry(2, 1) = "月份"
    vFilterAry(2, 2) = "ExcMn"
    
    vFilterAry(3, 1) = "貨幣"
    vFilterAry(3, 2) = "ExcCurr"
    
    vFilterAry(4, 1) = "對換率註解"
    vFilterAry(4, 2) = "ExcDesc"
    
    ReDim vAry(7, 3)
    vAry(1, 1) = ""
    vAry(1, 2) = "ExcID"
    vAry(1, 3) = "0"
    
    vAry(2, 1) = "年份"
    vAry(2, 2) = "ExcYr"
    vAry(2, 3) = "700"
    
    vAry(3, 1) = "月份"
    vAry(3, 2) = "ExcMn"
    vAry(3, 3) = "900"
    
    vAry(4, 1) = "貨幣"
    vAry(4, 2) = "ExcCurr"
    vAry(4, 3) = "1100"
    
    vAry(5, 1) = "對換率註解"
    vAry(5, 2) = "ExcDesc"
    vAry(5, 3) = "2000"
    
    vAry(6, 1) = "對換率"
    vAry(6, 2) = "ExcRate"
    vAry(6, 3) = "1000"
    
    vAry(7, 1) = "購貨對換率"
    vAry(7, 2) = "ExcBRate"
    vAry(7, 3) = "1200"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstExchangeRate.ExcID, MstExchangeRate.ExcYr, MstExchangeRate.ExcMn, MstExchangeRate.ExcCurr, MstExchangeRate.ExcDesc, "
        sSQL = sSQL + "MstExchangeRate.ExcRate, MstExchangeRate.ExcBRate "
        sSQL = sSQL + "FROM MstExchangeRate "
        .sBindSQL = sSQL
        sTmpSQL = "WHERE MstExchangeRate.ExcStatus = '1' "
        .sBindWhereSQL = sTmpSQL
        .sBindOrderSQL = "ORDER BY MstExchangeRate.ExcYr"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> txtExcID Then
        txtExcID = Trim(frmShareSearch.Tag)
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

Private Sub txtExcDesc_GotFocus()
    FocusMe txtExcDesc
End Sub

Private Sub txtExcDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtExcDesc, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
      
        txtExcRate(1).SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtExcDesc_LostFocus()
    FocusMe txtExcDesc, True
End Sub

Private Sub txtExcID_Change()
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstExchangeRate.* "
    wsSql = wsSql + "From MstExchangeRate "
    wsSql = wsSql + "WHERE (((MstExchangeRate.ExcID)='" + Set_Quote(txtExcID) + "') "
    wsSql = wsSql + "AND ((MstExchangeRate.ExcStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount > 0 Then
        Me.cboExcYr = ReadRs(rsRcd, "ExcYr")
        Me.cboExcCurr = ReadRs(rsRcd, "ExcCurr")
    End If
    
    If LoadRecord = True Then
        cboExcCurr.SetFocus
        SendKeys "{ENTER}"
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Sub

Private Function Chk_txtExcRate(Index As Integer) As Boolean
    Chk_txtExcRate = False
    
    If Not IsNumeric(txtExcRate(Index)) Then
        gsMsg = "對換率錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtExcRate(Index).SetFocus
        Exit Function
    End If
    
    Chk_txtExcRate = True
End Function

Private Function Chk_AlltxtExcRate() As Boolean
    Dim iCounter As Integer
    
    Chk_AlltxtExcRate = False
    
    For iCounter = 1 To 12
        If Chk_txtExcRate(iCounter) = False Then
            Exit Function
        End If
    Next
    
    Chk_AlltxtExcRate = True
End Function

Private Function Chk_txtExcBRate(Index As Integer) As Boolean
    Chk_txtExcBRate = False
    
    If Not IsNumeric(txtExcBRate(Index)) Then
        gsMsg = "購貨對換率錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtExcBRate(Index).SetFocus
        Exit Function
    End If
    
    Chk_txtExcBRate = True
End Function

Private Function Chk_AlltxtExcBRate() As Boolean
    Dim iCounter As Integer
    
    Chk_AlltxtExcBRate = False
    
    For iCounter = 1 To 12
        If Chk_txtExcBRate(iCounter) = False Then
            Exit Function
        End If
    Next
    
    Chk_AlltxtExcBRate = True
End Function

Private Function LoadMthID(inCode) As String
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstMonth.MthId "
    wsSql = wsSql + "From MstMonth "
    wsSql = wsSql + "WHERE ((MstMonth.MthLongDesc)='" + Set_Quote(inCode) + "') "
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount <= 0 Then
        LoadMthID = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    LoadMthID = ReadRs(rsRcd, "MthID")
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function LoadMthDesc(inCode) As String
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSql = "SELECT MstMonth.MthLongDesc "
    wsSql = wsSql + "From MstMonth "
    wsSql = wsSql + "WHERE ((MstMonth.MthID)=" + Set_Quote(inCode) + ") AND MstMonth.MthLang ='" & Set_Quote(gsLangID) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount <= 0 Then
        LoadMthDesc = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    LoadMthDesc = ReadRs(rsRcd, "MthLongDesc")
    rsRcd.Close
    Set rsRcd = Nothing
End Function


Private Sub txtExcRate_GotFocus(Index As Integer)
    FocusMe txtExcRate(Index)
End Sub

Private Sub txtExcRate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtExcRate(Index), False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtExcRate(Index) = True Then
            txtExcBRate(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtExcRate_LostFocus(Index As Integer)
    txtExcRate(Index) = Format(txtExcRate(Index), gsExrFmt)
    FocusMe txtExcRate(Index), True
End Sub

Private Sub txtExcBRate_GotFocus(Index As Integer)
    FocusMe txtExcBRate(Index)
End Sub

Private Sub txtExcBRate_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtExcBRate(Index), False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtExcBRate(Index) = True Then
            If Index = 12 Then
                txtExcDesc.SetFocus
            Else
                txtExcRate(Index + 1).SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtExcBRate_LostFocus(Index As Integer)
    txtExcBRate(Index) = Format(txtExcBRate(Index), gsExrFmt)
    FocusMe txtExcBRate(Index), True
End Sub

Private Function Chk_ExcCurr(ByVal inCode As String, ByVal inCode1 As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    Chk_ExcCurr = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSql = "SELECT ExcStatus "
    wsSql = wsSql & " FROM MstExchangeRate WHERE ExcYr = '" & Set_Quote(inCode) & "' AND ExcCurr = '" & Set_Quote(inCode1) & "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "ExcStatus")
    
    Chk_ExcCurr = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function
