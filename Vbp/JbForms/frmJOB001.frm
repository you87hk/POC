VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmJOB001 
   Caption         =   "JOB001"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   11880
   Icon            =   "frmJOB001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   12120
      OleObjectBlob   =   "frmJOB001.frx":030A
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtJobType 
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   8655
   End
   Begin VB.CheckBox chkInActive 
      Alignment       =   1  '靠右對齊
      Caption         =   "INACTIVE"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10560
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkDetail 
      Alignment       =   1  '靠右對齊
      Caption         =   "DETAIL"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10560
      TabIndex        =   3
      Top             =   1160
      Width           =   1215
   End
   Begin VB.TextBox txtJobName 
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   8655
   End
   Begin VB.ComboBox cboJobCode 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   6615
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmJOB001.frx":2A0D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboCusCode"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmJOB001.frx":2A29
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tblDetail"
      Tab(1).Control(1)=   "fraEst2"
      Tab(1).Control(2)=   "fraEst1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmJOB001.frx":2A45
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tblActual"
      Tab(2).ControlCount=   1
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   4935
         Left            =   -74520
         OleObjectBlob   =   "frmJOB001.frx":2A61
         TabIndex        =   41
         Top             =   1200
         Width           =   11175
      End
      Begin VB.ComboBox cboCusCode 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Frame fra1 
         Height          =   5895
         Left            =   240
         TabIndex        =   16
         Top             =   420
         Width           =   11415
         Begin VB.TextBox txtJobRemark 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   11
            Top             =   4200
            Width           =   9975
         End
         Begin VB.TextBox txtJobComplete 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   10
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txtCusPo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   9
            Top             =   3000
            Width           =   4215
         End
         Begin VB.TextBox txtCusContactPerson 
            BackColor       =   &H80000014&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1320
            TabIndex        =   6
            Top             =   1800
            Width           =   9975
         End
         Begin MSMask.MaskEdBox medFromDate 
            Height          =   285
            Left            =   1320
            TabIndex        =   7
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medToDate 
            Height          =   285
            Left            =   4200
            TabIndex        =   8
            Top             =   2400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblJobRemark 
            Caption         =   "JOBREMARK"
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
            Left            =   120
            TabIndex        =   40
            Top             =   4260
            Width           =   1215
         End
         Begin VB.Label lblPercent 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2760
            TabIndex        =   39
            Top             =   3660
            Width           =   300
         End
         Begin VB.Label lblJobComplete 
            Caption         =   "JOBCOMPLETE"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   3660
            Width           =   2100
         End
         Begin VB.Label lblCusPo 
            Caption         =   "CUSPO"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   26
            Top             =   3060
            Width           =   2100
         End
         Begin VB.Label lblToDate 
            Caption         =   "TODATE"
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
            Left            =   3000
            TabIndex        =   25
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label lblFromDate 
            Caption         =   "FROMDATE"
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
            Left            =   120
            TabIndex        =   24
            Top             =   2460
            Width           =   1200
         End
         Begin VB.Label lblCusName 
            Caption         =   "CUSNAME"
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
            Left            =   120
            TabIndex        =   23
            Top             =   1245
            Width           =   1215
         End
         Begin VB.Label lblDspCusName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1320
            TabIndex        =   19
            Top             =   1200
            Width           =   9975
         End
         Begin VB.Label lblCusCode 
            Caption         =   "CUSCODE"
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
            Left            =   120
            TabIndex        =   18
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label lblCusContact 
            Caption         =   "CUSCONTACT"
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
            Left            =   120
            TabIndex        =   17
            Top             =   1860
            Width           =   1215
         End
      End
      Begin TrueDBGrid60.TDBGrid tblActual 
         Height          =   5655
         Left            =   -74760
         OleObjectBlob   =   "frmJOB001.frx":89C8
         TabIndex        =   27
         Top             =   720
         Width           =   11415
      End
      Begin VB.Frame fraEst2 
         Height          =   5895
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   11415
         Begin VB.Frame Frame1 
            Height          =   450
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   6135
            Begin VB.Label lblDeleteLine 
               Caption         =   "REMARK"
               Height          =   225
               Left            =   4800
               TabIndex        =   33
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lblInsertLine 
               Caption         =   "REMARK"
               Height          =   225
               Left            =   3360
               TabIndex        =   32
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lblComboPrompt 
               Caption         =   "REMARK"
               Height          =   225
               Left            =   1920
               TabIndex        =   31
               Top             =   180
               Width           =   1215
            End
            Begin VB.Label lblKeyDesc 
               Caption         =   "REMARK"
               Height          =   225
               Left            =   360
               TabIndex        =   30
               Top             =   180
               Width           =   1215
            End
         End
      End
      Begin VB.Frame fraEst1 
         Height          =   5895
         Left            =   -74760
         TabIndex        =   34
         Top             =   360
         Width           =   11415
         Begin VB.TextBox txtExpense 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   300
            Left            =   1560
            TabIndex        =   13
            Top             =   2160
            Width           =   2715
         End
         Begin VB.TextBox txtIncome 
            Alignment       =   1  '靠右對齊
            Enabled         =   0   'False
            Height          =   300
            Left            =   1560
            TabIndex        =   12
            Top             =   1440
            Width           =   2715
         End
         Begin VB.Label lblExpense 
            Caption         =   "EXPENSE"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   600
            TabIndex        =   37
            Top             =   2220
            Width           =   1020
         End
         Begin VB.Label lblIncome 
            Caption         =   "INCOME"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   600
            TabIndex        =   35
            Top             =   1500
            Width           =   1020
         End
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   7920
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":E92F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":F209
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":FAE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":FF35
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":10387
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":106A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":10AF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":10F45
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":1125F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":11579
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":119CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":122A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJOB001.frx":125CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Height          =   360
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   17160
      _ExtentX        =   30268
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open (F6)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add (F2)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit (F5)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Find (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblJobType 
      Caption         =   "JOBTYPE"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   22
      Top             =   1500
      Width           =   1380
   End
   Begin VB.Label lblJobName 
      Caption         =   "JOBNAME"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   21
      Top             =   1140
      Width           =   1380
   End
   Begin VB.Label lblJobCode 
      Caption         =   "JOBCODE"
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
      Left            =   360
      TabIndex        =   20
      Top             =   780
      Width           =   1215
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop Up"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmJOB001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Private waActual As New XArrayDB
Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB

Private waPopUpSub As New XArrayDB
Private wcCombo As Control
Private wbReadOnly As Boolean

Private wsOldCusNo As String

Private wgsTitle As String

Private Const COSTCODE = 0
Private Const COSTDESC = 1
Private Const UNIT = 2
Private Const INCOME = 3
Private Const EXPENSE = 4
Private Const COSTID = 5

Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"
Private Const tcRefresh = "Refresh"
Private Const tcPrint = "Print"

Private wiOpenDoc As Integer
Private wiAction As Integer
Private wlCusID As Long
Private wlLineNo As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "MstJob"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, COSTCODE, COSTID
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    waActual.ReDim 0, -1, COSTCODE, COSTID
    Set tblActual.Array = waActual
    tblActual.ReBind
    tblActual.Bookmark = 0
    
    wiAction = DefaultPage
    
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


    wbReadOnly = False
    
    Call SetButtonStatus("Default")
    Call SetFieldStatus("Default")
    Call SetFieldStatus("AfrActEdit")
    
    Call SetDateMask(medFromDate)
    Call SetDateMask(medToDate)
    
    wsOldCusNo = ""
    
    wlKey = 0
    wlCusID = 0
    wlLineNo = 1
    chkDetail.Value = 0
    chkInActive.Value = 0
    txtJobComplete = 0
    txtIncome = 0
    txtExpense = 0
    wsTrnCd = "JB"
        
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    tabDetailInfo.Tab = 0
    FocusMe cboJobCode
End Sub

Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
End Sub

Private Sub Ini_Scr_AfrKey()
    If LoadRecord() = False Then
        wiAction = AddRec
        medFromDate.Text = Dsp_Date(Now)
        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboJobCode.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        wsOldCusNo = cboCusCode.Text
      
         Call SetButtonStatus("AfrKeyEdit")
    End If
    
     Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    
    Call SetFieldStatus("AfrKey")
    
    txtJobName.SetFocus
    tabDetailInfo.Tab = 0
    
  '      wiAction = AddRec
  '      Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
  '      Call SetButtonStatus("AfrKeyAdd")
  '      Call SetFieldStatus("AfrKey")
        
  '      cboSaleCode.SetFocus
End Sub

Private Sub cboJobCode_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboJobCode
  
    wsSql = "SELECT JOBCODE, JOBNAME, JOBACTIVE "
    wsSql = wsSql & " FROM MSTJOB "
    wsSql = wsSql & " WHERE JOBCODE LIKE '%" & IIf(cboJobCode.SelLength > 0, "", Set_Quote(cboJobCode.Text)) & "%' "
    wsSql = wsSql & " AND JOBSTATUS  = '1' "
    wsSql = wsSql & " ORDER BY JOBCODE DESC "
    Call Ini_Combo(3, wsSql, cboJobCode.Left, cboJobCode.Top + cboJobCode.Height, tblCommon, wsFormID, "TBLJOBCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboJobCode_GotFocus()
    FocusMe cboJobCode
End Sub

Private Sub cboJobCode_LostFocus()
    FocusMe cboJobCode, True
End Sub

Private Sub cboJobCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboJobCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboJobCode() = False Then Exit Sub
        
        Call Ini_Scr_AfrKey
        
    End If

End Sub

Private Function Chk_cboJobCode() As Boolean
    Dim wsStatus As String
    
    Chk_cboJobCode = False
    
    If Trim(cboJobCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboJobCode.SetFocus
   
        Exit Function
    End If
        
    If Chk_TrnHdDocNo(wsTrnCd, cboJobCode, wsStatus) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已刪除, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
        End If
    
        If wsStatus = "3" Then
            gsMsg = "文件已無效, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
        End If
    End If
    
    Chk_cboJobCode = True
End Function

Private Sub chkDetail_Click()
        If chkDetail.Value = 0 Then
            fraEst1.Visible = True
            fraEst2.Visible = False
            tblDetail.Visible = False
        Else
            fraEst1.Visible = False
            fraEst2.Visible = True
            tblDetail.Visible = True
        End If
End Sub

Private Sub chkDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chkDetail.Value = 0 Then
            fraEst1.Visible = True
            fraEst2.Visible = False
            tblDetail.Visible = False
        Else
            fraEst1.Visible = False
            fraEst2.Visible = True
            tblDetail.Visible = True
        End If
        
        chkInActive.SetFocus
    End If
End Sub

Private Sub chkDetail_LostFocus()
    'If chkDetail.Value = 0 Then
    '    fraEst1.Visible = True
    '    fraEst2.Visible = False
    '    tblDetail.Visible = False
    'Else
    '    fraEst1.Visible = False
    '    fraEst2.Visible = True
    '    tblDetail.Visible = True
    'End If
End Sub

Private Sub chkInActive_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        cboCusCode.SetFocus
    End If

End Sub

Private Sub Form_Activate()
    If OpenDoc = True Then
        OpenDoc = False
        Set wcCombo = cboCusCode
        Call cboCusCode_DropDown
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
   Select Case KeyCode
         Case vbKeyPageDown
            KeyCode = 0
            If tabDetailInfo.Tab < tabDetailInfo.Tabs - 1 Then
                tabDetailInfo.Tab = tabDetailInfo.Tab + 1
                Exit Sub
            End If
        Case vbKeyPageUp
            KeyCode = 0
            If tabDetailInfo.Tab > 0 Then
                tabDetailInfo.Tab = tabDetailInfo.Tab - 1
                Exit Sub
            End If
       
       Case vbKeyF6
            Call cmdOpen
        
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
            
        
        'Case vbKeyF5
        '    If wiAction = DefaultPage Then Call cmdEdit
       
        
        'Case vbKeyF3
        '    If wiAction = DefaultPage Then Call cmdDel
        
        Case vbKeyF7
        
            If tbrProcess.Buttons(tcRefresh).Enabled = True Then Call cmdRefresh
            
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
    
    Call Ini_Form
    Call Ini_Grid
    Call Ini_Caption
    Call Ini_Scr
  
    MousePointer = vbDefault
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    lblJobCode.Caption = Get_Caption(waScrItm, "JOBCODE")
    lblJobName.Caption = Get_Caption(waScrItm, "JOBNAME")
    lblJobType.Caption = Get_Caption(waScrItm, "JOBTYPE")
  
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    lblCusContact.Caption = Get_Caption(waScrItm, "CUSCONTACTPERSON")
    
    chkDetail.Caption = Get_Caption(waScrItm, "DETAIL")
    chkInActive.Caption = Get_Caption(waScrItm, "INACTIVE")

    lblFromDate.Caption = Get_Caption(waScrItm, "FROMDATE")
    lblToDate.Caption = Get_Caption(waScrItm, "TODATE")
    lblCusPo.Caption = Get_Caption(waScrItm, "CUSPO")
    lblJobComplete.Caption = Get_Caption(waScrItm, "JOBCOMPLETE")
    lblPercent.Caption = Get_Caption(waScrItm, "PERCENT")
    lblJobRemark.Caption = Get_Caption(waScrItm, "JOBREMARK")
    
    With tblDetail
        .Columns(COSTCODE).Caption = Get_Caption(waScrItm, "COSTCODE")
        .Columns(COSTDESC).Caption = Get_Caption(waScrItm, "COSTDESC")
        .Columns(UNIT).Caption = Get_Caption(waScrItm, "UNIT")
        .Columns(INCOME).Caption = Get_Caption(waScrItm, "INCOME")
        .Columns(EXPENSE).Caption = Get_Caption(waScrItm, "EXPENSE")
    End With
    
    With tblActual
        .Columns(COSTCODE).Caption = Get_Caption(waScrItm, "COSTCODE")
        .Columns(COSTDESC).Caption = Get_Caption(waScrItm, "COSTDESC")
        .Columns(UNIT).Caption = Get_Caption(waScrItm, "UNIT")
        .Columns(INCOME).Caption = Get_Caption(waScrItm, "INCOME")
        .Columns(EXPENSE).Caption = Get_Caption(waScrItm, "EXPENSE")
    End With
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcRefresh).ToolTipText = Get_Caption(waScrToolTip, tcRefresh) & "(F7)"
    tbrProcess.Buttons(tcPrint).ToolTipText = Get_Caption(waScrToolTip, tcPrint)
   
    lblKeyDesc = Get_Caption(waScrToolTip, "KEYDESC")
    lblComboPrompt = Get_Caption(waScrToolTip, "COMBOPROMPT")
    lblInsertLine = Get_Caption(waScrToolTip, "INSERTLINE")
    lblDeleteLine = Get_Caption(waScrToolTip, "DELETELINE")
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO03")

    wsActNam(1) = Get_Caption(waScrItm, "JOBADD")
    wsActNam(2) = Get_Caption(waScrItm, "JOBEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "JOBDELETE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'    If Button = 2 Then
'        PopupMenu mnuMaster
'    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    Call UnLockAll(wsConnTime, wsFormID)
    Set waResult = Nothing
    Set waActual = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmJOB001 = Nothing
End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
         If cboCusCode.Enabled Then
            cboCusCode.SetFocus
         End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
        If chkDetail.Value = 0 Then
            fraEst1.Visible = True
            fraEst2.Visible = False
            tblDetail.Visible = False
            If txtIncome.Enabled = True Then
                txtIncome.SetFocus
            End If
        Else
            fraEst1.Visible = False
            fraEst2.Visible = True
            tblDetail.Visible = True
            
            If tblDetail.Enabled Then
                tblDetail.Col = COSTCODE
                tblDetail.SetFocus
            End If
        End If
    ElseIf tabDetailInfo.Tab = 2 Then
    End If
End Sub

Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case COSTCODE
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
              Case COSTCODE
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
    
    tblCommon.Visible = False
    If wcCombo.Enabled = True Then
        wcCombo.SetFocus
    Else
        Set wcCombo = Nothing
    End If

End Sub

Private Function cmdSave() As Boolean

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
    Dim wsDocNo As String
    Dim wlRowCtr As Long
    Dim wsCtlPrd As String
    Dim wsSts As String
    Dim i As Integer
     
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboJobCode.Text, wsFormID) Or wbReadOnly Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
   
    If InputValidation() = False Then
       MousePointer = vbDefault
       Exit Function
    End If
    
    '' Last Check when Add
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
         
   ' If lblDspNetAmtOrg.Caption > Get_CreditLimit(wlCusID, wlKey, Trim(medDocDate.Text)) Then
   '    gsMsg = "已超過信貸額!"
   '    MsgBox gsMsg, vbOKOnly, gsTitle
   '    MousePointer = vbDefault
   '    Exit Function
   ' End If
    
    wlRowCtr = waResult.UpperBound(1)
    'wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_JOB001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboJobCode.Text))
    Call SetSPPara(adcmdSave, 5, txtJobName)
    Call SetSPPara(adcmdSave, 6, txtCusContactPerson.Text)
    Call SetSPPara(adcmdSave, 7, wlCusID)
    Call SetSPPara(adcmdSave, 8, txtJobType)
    Call SetSPPara(adcmdSave, 9, txtCusPo)
    Call SetSPPara(adcmdSave, 10, To_Value(txtJobComplete))
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medFromDate.Text))
    'Call SetSPPara(adcmdSave, 12, Set_MedDate(medExpiryDate.Text))
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medToDate.Text))
    
    Call SetSPPara(adcmdSave, 13, IIf(chkDetail.Value = 0, "N", "Y"))
    
    Call SetSPPara(adcmdSave, 14, IIf(chkInActive.Value = 0, "Y", "N"))
    Call SetSPPara(adcmdSave, 15, txtJobRemark)
    
    Call SetSPPara(adcmdSave, 16, wsFormID)
    Call SetSPPara(adcmdSave, 17, gsWorkStationID)
    Call SetSPPara(adcmdSave, 18, gsUserID)
    Call SetSPPara(adcmdSave, 19, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 20)
    wsDocNo = GetSPPara(adcmdSave, 21)
    
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_JOB001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, COSTCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, wiCtr + 1)
                Call SetSPPara(adcmdSave, 4, IIf(chkDetail.Value = 1, "D", "S"))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, COSTCODE))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, UNIT))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, INCOME))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, EXPENSE))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, COSTDESC))
                Call SetSPPara(adcmdSave, 10, IIf(wlRowCtr = wiCtr, "Y", "N"))
                adcmdSave.Execute
            End If
        Next
    End If
    cnCon.CommitTrans
    
    If wiAction = AddRec Then
    If Trim(wsDocNo) <> "" Then
      '  Call cmdPrint(wsDocNo)
        gsMsg = "文件號 : " & wsDocNo & " 已製成!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    Else
        
        gsMsg = "文件儲存失敗!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    End If
    End If
    
    If wiAction = CorRec Then
        gsMsg = "文件已儲存!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    End If
    
    'Call UnLockAll(wsConnTime, wsFormID)
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
Private Function cmdDel() As Boolean

    Dim wsGenDte As String
    Dim adcmdDelete As New ADODB.Command
    Dim wsDocNo As String
    Dim i As Integer
    
    cmdDel = False
    
    MousePointer = vbHourglass
    
    On Error GoTo cmdDelete_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboJobCode.Text, wsFormID) Or wbReadOnly Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
    End If
    
    gsMsg = "你是否確認要刪除此檔案?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       wiAction = CorRec
       MousePointer = vbDefault
       Exit Function
    End If
    
    wiAction = DelRec
    
      cnCon.BeginTrans
    Set adcmdDelete.ActiveConnection = cnCon
        
    adcmdDelete.CommandText = "USP_JOB001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboJobCode.Text))
    Call SetSPPara(adcmdDelete, 5, txtJobName)
    Call SetSPPara(adcmdDelete, 6, txtCusContactPerson.Text)
    Call SetSPPara(adcmdDelete, 7, wlCusID)
    Call SetSPPara(adcmdDelete, 8, txtJobType)
    Call SetSPPara(adcmdDelete, 9, txtCusPo)
    Call SetSPPara(adcmdDelete, 10, To_Value(txtJobComplete))
    
    Call SetSPPara(adcmdDelete, 11, Set_MedDate(medFromDate.Text))
    'Call SetSPPara(adcmdDelete, 12, Set_MedDate(medExpiryDate.Text))
    Call SetSPPara(adcmdDelete, 12, Set_MedDate(medToDate.Text))
    
    Call SetSPPara(adcmdDelete, 13, IIf(chkDetail.Value = 0, "N", "Y"))
    
    Call SetSPPara(adcmdDelete, 14, IIf(chkInActive.Value = 0, "Y", "N"))
    Call SetSPPara(adcmdDelete, 15, txtJobRemark)
    
    Call SetSPPara(adcmdDelete, 16, wsFormID)
    Call SetSPPara(adcmdDelete, 17, gsWorkStationID)
    Call SetSPPara(adcmdDelete, 18, gsUserID)
    Call SetSPPara(adcmdDelete, 19, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 20)
    wsDocNo = GetSPPara(adcmdDelete, 21)
    
    cnCon.CommitTrans
    
    gsMsg = wsDocNo & " 檔案已刪除!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    Call cmdCancel
    MousePointer = vbDefault
    
    Set adcmdDelete = Nothing
    cmdDel = True
    
    Exit Function
    
cmdDelete_Err:
    MsgBox "Check cmdDel"
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdDelete = Nothing

End Function

Private Function InputValidation() As Boolean
    
    Dim wsExcRate As String
    Dim wsExcDesc As String

    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    If chk_txtJobName = False Then Exit Function
    If Chk_medToDate = False Then Exit Function
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, COSTCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.Col = COSTCODE
                    tblDetail.SetFocus
                    tabDetailInfo.Tab = 1
   
                    Exit Function
                End If
            End If
            'For wlCtr1 = 0 To .UpperBound(1)
            '    If wlCtr <> wlCtr1 Then
            '        If waResult(wlCtr, BOOKCODE) = waResult(wlCtr1, BOOKCODE) Then
            '          gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
            '          MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            '          tblDetail.Col = BOOKCODE
            '          tblDetail.SetFocus
            '          Exit Function
            '        End If
            '    End If
            'Next
        Next
    End With
    
    If wiEmptyGrid = True And chkDetail.Value = 0 Then
        gsMsg = "訂購單沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tblDetail.Col = COSTCODE
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    'If Chk_NoDup(To_Value(tblDetail.Bookmark)) = False Then
    '    tblDetail.FirstRow = tblDetail.Row
    '    tblDetail.Col = BOOKCODE
    '    tblDetail.SetFocus
    '    Exit Function
    'End If
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
    


Private Sub cmdNew()

    Dim newForm As New frmJOB001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmJOB001
    
    newForm.OpenDoc = True
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show

End Sub

Private Sub Ini_Form()

    Me.KeyPreview = True
'    Me.Left = (Screen.Width - Me.Width) / 2
'    Me.Top = (Screen.Height - Me.Height) / 2
    
    Me.WindowState = 2
    'Me.tblDetail.Height = Me.Height - Me.tbrProcess.Height - Me.fra1.Height
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "JOB001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")

    
    Call LoadWSINFO
    

End Sub



Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("Default")
    
End Sub



Public Property Get OpenDoc() As Integer
    OpenDoc = wiOpenDoc
End Property

Public Property Let OpenDoc(SearchDoc As Integer)
    wiOpenDoc = SearchDoc
End Property



Private Sub tblDetail_BeforeRowColChange(Cancel As Integer)

    On Error GoTo tblDetail_BeforeRowColChange_Err
    With tblDetail
       ' If .Bookmark <> .DestinationRow Then
            If Chk_GrdRow(To_Value(.Bookmark)) = False Then
                Cancel = True
                Exit Sub
            End If
       ' End If
    End With
    
    Exit Sub
    
tblDetail_BeforeRowColChange_Err:
    
    MsgBox "Check tblDeiail BeforeRowColChange!"
    Cancel = True

End Sub


Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub tbrProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
 
 Select Case Button.Key
        Case tcOpen
            Call cmdOpen
        Case tcAdd
            Call cmdNew
    '    Case tcEdit
     '       Call cmdEdit
        Case tcDelete
            Call cmdDel
        Case tcSave
            Call cmdSave
        Case tcCancel
           If tbrProcess.Buttons(tcSave).Enabled = True Then
           If MsgBox("你是否確定要放棄現時之作業?", vbYesNo, gsTitle) = vbYes Then
                Call cmdCancel
           End If
           Else
                Call cmdCancel
           End If
       Case tcRefresh
            Call cmdRefresh
       Case tcPrint
            Call cmdPrint
       Case tcExit
            Unload Me
    End Select
    
End Sub



Private Sub cboCusCode_DropDown()
   
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusCode
    
    If gsLangID = "1" Then
        wsSql = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSql = wsSql & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
        wsSql = wsSql & "AND CUSSTATUS = '1' "
        wsSql = wsSql & "ORDER BY CUSCODE "
    Else
        wsSql = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSql = wsSql & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
        wsSql = wsSql & "AND CUSSTATUS = '1' "
        wsSql = wsSql & "ORDER BY CUSCODE "
    End If
    Call Ini_Combo(2, wsSql, cboCusCode.Left + tabDetailInfo.Left, cboCusCode.Top + cboCusCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCUSCODE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
   
End Sub

Private Sub cboCusCode_GotFocus()
    
    Set wcCombo = cboCusCode
    'TREtoolsbar1.ButtonEnabled(tcCusSrh) = True
    FocusMe cboCusCode
    
End Sub

Private Sub cboCusCode_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboCusCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_cboCusCode() = False Then Exit Sub
        
        txtCusContactPerson = Get_TableInfo("MstCustomer", "CusCode ='" & Set_Quote(cboCusCode.Text) & "'", "CusContactPerson")
        
        'If wiAction = AddRec Or wsOldCusNo <> cboCusCode.Text Then Call Get_DefVal
           
        txtCusContactPerson.SetFocus
            
    End If
    
End Sub

Private Function chk_cboCusCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    Dim wsTel As String
    Dim wsFax As String
    
    chk_cboCusCode = False
    
    If Trim(cboCusCode) = "" Then
        gsMsg = "請輸入客戶編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    End If
    
    If Chk_CusCode(cboCusCode, wlID, wsName, wsTel, wsFax) Then
        wlCusID = wlID
        lblDspCusName.Caption = wsName
    Else
        wlCusID = 0
        lblDspCusName.Caption = ""
        gsMsg = "客戶不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    End If
    
    chk_cboCusCode = True

End Function

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
      '  .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = COSTCODE To COSTID
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case COSTCODE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Button = True
                Case COSTDESC
                    .Columns(wiCtr).Width = 6000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case UNIT
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case INCOME
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case EXPENSE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case COSTID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
    End With
    
    With tblActual
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
      '  .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = COSTCODE To COSTID
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case COSTCODE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Locked = True
                Case COSTDESC
                    .Columns(wiCtr).Width = 6000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case UNIT
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                Case INCOME
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case EXPENSE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case COSTID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    Dim sTemp As String
   
    With tblDetail
        sTemp = .Columns(ColIndex)
        .Update
    End With

    'If ColIndex = COSTCODE Then
    '    Call LoadJobCost(sTemp)
    'End If
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wsCostCode As String
    Dim wsCostDesc As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case COSTCODE
                'If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                '    GoTo Tbl_BeforeColUpdate_Err
                'End If
                
                If Chk_grdCostCode(.Columns(ColIndex).Text, wsCostCode, wsCostDesc) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If

                .Columns(COSTCODE).Text = wsCostCode
                .Columns(COSTDESC).Text = wsCostDesc
                .Columns(UNIT).Text = 0
                .Columns(INCOME).Text = 0
                .Columns(EXPENSE).Text = 0
                
                If Trim(.Columns(ColIndex).Text) <> wsCostCode Then
                    .Columns(ColIndex).Text = wsCostCode
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
    Dim wsSql As String
    Dim wiTop As Long
    Dim wiCtr As Integer
    
    On Error GoTo tblDetail_ButtonClick_Err

    With tblDetail
        Select Case ColIndex
            Case COSTCODE
                
                If gsLangID = 1 Then
                    wsSql = "SELECT COSTCODE, COSTNAME FROM MSTCOST "
                    wsSql = wsSql & " WHERE COSTSTATUS = '1' AND COSTACTIVE = 'Y' AND COSTCODE LIKE '%" & Set_Quote(.Columns(COSTCODE).Text) & "%' "
                    If waResult.UpperBound(1) > -1 Then
                          wsSql = wsSql & " AND COSTCODE NOT IN ( "
                          For wiCtr = 0 To waResult.UpperBound(1)
                                wsSql = wsSql & " '" & waResult(wiCtr, COSTCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                          Next
                    End If
                    wsSql = wsSql & " ORDER BY COSTCODE "
                Else
                    wsSql = "SELECT COSTCODE, COSTNAME FROM MSTCOST "
                    wsSql = wsSql & " WHERE COSTSTATUS = '1' AND COSTACTIVE = 'Y' AND COSTCODE LIKE '%" & Set_Quote(.Columns(COSTCODE).Text) & "%' "
                    If waResult.UpperBound(1) > -1 Then
                          wsSql = wsSql & " AND COSTCODE NOT IN ( "
                          For wiCtr = 0 To waResult.UpperBound(1)
                                wsSql = wsSql & " '" & waResult(wiCtr, COSTCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                          Next
                    End If
                    wsSql = wsSql & " ORDER BY COSTCODE "
                End If
                
                Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCOSTCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
          '  Case WhsCode
                
          '      wsSql = "SELECT WHSCODE, WHSDESC FROM mstWareHouse "
          '      wsSql = wsSql & " WHERE WHSSTATUS <> '2' AND WHSCODE LIKE '%" & Set_Quote(.Columns(WhsCode).Text) & "%' "
          '      wsSql = wsSql & " ORDER BY WHSCODE "
                
          '      Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
          '      tblCommon.Visible = True
          '      tblCommon.SetFocus
          '      Set wcCombo = tblDetail
                
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
        
        Case vbKeyF5        ' INSERT LINE
            KeyCode = vbDefault
            If .Bookmark = waResult.UpperBound(2) Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case vbKeyF8        ' DELETE LINE
            KeyCode = vbDefault
            If IsNull(.Bookmark) Then Exit Sub
            If .EditActive = True Then Exit Sub
            gsMsg = "你是否確定要刪除此列?"
            If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then Exit Sub
            .Delete
            .Update
            If .Row = -1 Then
                .Row = 0
            End If
            .Refresh
            .SetFocus

        Case vbKeyReturn
            Select Case .Col
                Case COSTCODE, COSTDESC, UNIT, INCOME
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case EXPENSE
                    KeyCode = vbKeyDown
                    .Col = COSTCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
              Select Case .Col
                Case EXPENSE, INCOME, UNIT, COSTDESC
                    .Col = .Col - 1
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case COSTCODE, COSTDESC, UNIT, INCOME
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
        
        'Case Qty
        '    Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        'Case Price, DisPer
        '    Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
            
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = COSTCODE
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case COSTCODE
                    Call Chk_grdCostCode(.Columns(COSTCODE).Text, "", "")
                'Case UNIT
                '    Call Chk_grdQty(.Columns(UNIT).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdCostCode(inAccNo As String, outAccNo As String, outAccDesc As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入物料號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdCostCode = False
        Exit Function
    End If
    
    If gsLangID = "1" Then
        wsSql = "SELECT COSTCODE, COSTNAME FROM MstCost"
        wsSql = wsSql & " WHERE (COSTCODE = '" & Set_Quote(inAccNo) & "') AND COSTACTIVE = 'Y' "
    Else
        wsSql = "SELECT COSTCODE, COSTNAME FROM MstCost"
        wsSql = wsSql & " WHERE (COSTCODE = '" & Set_Quote(inAccNo) & "') AND COSTACTIVE = 'Y' "
    End If
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccNo = ReadRs(rsDes, "COSTCODE")
        outAccDesc = ReadRs(rsDes, "COSTNAME")
       
        Chk_grdCostCode = True
    Else
        outAccDesc = ""
        
        gsMsg = "沒有此成本!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdCostCode = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

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

Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(COSTCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, COSTCODE)) = "" And _
                   Trim(waResult(inRow, COSTDESC)) = "" And _
                   Trim(waResult(inRow, UNIT)) = "" And _
                   Trim(waResult(inRow, INCOME)) = "" And _
                   Trim(waResult(inRow, EXPENSE)) = "" And _
                   Trim(waResult(inRow, COSTID)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyRow = False
    
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
        
        If IsEmptyRow(To_Value(LastRow)) = True Then
            .Delete
            .Refresh
            .SetFocus
            Chk_GrdRow = False
            Exit Function
        End If
        
        If Chk_grdCostCode(waResult(LastRow, COSTCODE), "", "") = False Then
            .Col = COSTCODE
            .Row = LastRow
            Exit Function
        End If
        
        'If Chk_grdWhsCode(waResult(LastRow, WHSCODE)) = False Then
        '        .Col = WHSCODE
        '        .Row = LastRow
        '        Exit Function
        'End If
        
        'If Chk_grdWantedDate(waResult(LastRow, WANTED)) = False Then
        '        .Col = WANTED
        '        .Row = LastRow
        '        Exit Function
        'End If
        
        'If Chk_grdQty(waResult(LastRow, Qty)) = False Then
        '        .Col = Qty
        '        .Row = LastRow
        '        Exit Function
        'End If
        
        'If Chk_grdDisPer(waResult(LastRow, DisPer)) = False Then
        '        .Col = DisPer
        '        .Row = LastRow
        '        Exit Function
        'End If
        
        'If Chk_Amount(waResult(LastRow, Amt)) = False Then
        '    .Col = Amt
        '    .Row = LastRow
        '    Exit Function
        'End If
        
    
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function

Private Function SaveData() As Boolean

    Dim wiRet As Long
    
    SaveData = False
    
     If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) And _
        tbrProcess.Buttons(tcSave).Enabled = True Then
        
        gsMsg = "你是否確定不儲存現時之變更而離開?"
        If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbOK Then
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

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            Me.cboJobCode.Enabled = False
            Me.txtJobName.Enabled = False
            Me.txtJobType.Enabled = False
            Me.chkDetail.Enabled = False
            Me.chkInActive.Enabled = False
            
            Me.cboCusCode.Enabled = False
            Me.txtCusContactPerson.Enabled = False
            Me.medFromDate.Enabled = False
            Me.medToDate.Enabled = False
            Me.txtCusPo.Enabled = False
            Me.txtJobComplete.Enabled = False
            Me.txtJobRemark.Enabled = False
            
            Me.txtIncome.Enabled = False
            Me.txtExpense.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboJobCode.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboJobCode.Enabled = True
        
        Case "AfrKey"
            Me.cboJobCode.Enabled = False
            Me.txtJobName.Enabled = True
            Me.txtJobType.Enabled = True
            Me.chkDetail.Enabled = True
            Me.chkInActive.Enabled = True
            
            Me.cboCusCode.Enabled = True
            Me.txtCusContactPerson.Enabled = True
            Me.medFromDate.Enabled = True
            Me.medToDate.Enabled = True
            Me.txtCusPo.Enabled = True
            Me.txtJobComplete.Enabled = True
            Me.txtJobRemark.Enabled = True
            
            Me.txtIncome.Enabled = True
            Me.txtExpense.Enabled = True
            
            Me.tblDetail.Enabled = True
    End Select
End Sub

Private Sub LoadWSINFO()
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT * FROM sysWSINFO WHERE WSID ='" + gsWorkStationID + "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        If gsLangID = "2" Then
            wgsTitle = ReadRs(rsRcd, "WSCTITLE")
        Else
            wgsTitle = ReadRs(rsRcd, "WSTITLE")
        End If
    Else
        wgsTitle = ""
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    
End Sub

Private Function LoadRecord() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False

    wsSql = "SELECT JOBID, JOBCODE, JOBNAME, JOBCONTACTPERSON, JOBCUSID, "
    wsSql = wsSql & "JOBTYPE, JOBCUSPO, JOBCOMPLETE, JOBFROMDATE, JOBTODATE, "
    wsSql = wsSql & "JOBDETAIL, JOBACTIVE, JOBREMARK, JOBSTATUS, JOBLASTUPD, "
    wsSql = wsSql & "JOBLASTUPDDATE, "
    wsSql = wsSql & "CUSCODE, CUSNAME "
    wsSql = wsSql & "FROM  MstJob, MstCustomer "
    wsSql = wsSql & "WHERE JOBCODE = '" & Set_Quote(cboJobCode) & "' "
    wsSql = wsSql & "AND CUSID = JOBCUSID "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wlKey = ReadRs(rsRcd, "JOBID")
    txtJobName = ReadRs(rsRcd, "JOBNAME")
    Me.txtCusContactPerson = ReadRs(rsRcd, "JOBCONTACTPERSON")
    wlCusID = ReadRs(rsRcd, "JOBCUSID")
    txtJobType = ReadRs(rsRcd, "JOBTYPE")
    txtCusPo = ReadRs(rsRcd, "JOBCUSPO")
    txtJobComplete = ReadRs(rsRcd, "JOBCOMPLETE")
    medFromDate = ReadRs(rsRcd, "JOBFROMDATE")
    medToDate = ReadRs(rsRcd, "JOBTODATE")
    txtJobRemark = ReadRs(rsRcd, "JOBREMARK")
    
    If ReadRs(rsRcd, "JOBDETAIL") = "Y" Then
        chkDetail.Value = 1
    Else
        chkDetail.Value = 0
    End If
    
    If ReadRs(rsRcd, "JOBACTIVE") = "Y" Then
        chkInActive.Value = 0
    Else
        chkInActive.Value = 1
    End If
    
    cboCusCode = ReadRs(rsRcd, "CUSCODE")
    lblDspCusName = ReadRs(rsRcd, "CUSNAME")
    
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    wsSql = "SELECT JBEID, JBEJOBID, JBEDOCLINE, JBELNTYPE, JBECOSTCODE, "
    wsSql = wsSql & "JBEQTY, JBEINCOME, JBEEXPENSE, JBEREMARK "
    wsSql = wsSql & "FROM MstJobEstimation "
    wsSql = wsSql & "WHERE JBEJOBID = " & wlKey & " "
    wsSql = wsSql & "ORDER BY JBEDOCLINE "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
    Else
        rsRcd.MoveFirst
        With waResult
             .ReDim 0, -1, COSTCODE, COSTID
             Do While Not rsRcd.EOF
                 wiCtr = wiCtr + 1
                 .AppendRows
                 waResult(.UpperBound(1), COSTCODE) = ReadRs(rsRcd, "JBECOSTCODE")
                 waResult(.UpperBound(1), COSTDESC) = ReadRs(rsRcd, "JBEREMARK")
                 waResult(.UpperBound(1), UNIT) = ReadRs(rsRcd, "JBEQTY")
                 waResult(.UpperBound(1), INCOME) = ReadRs(rsRcd, "JBEINCOME")
                 waResult(.UpperBound(1), EXPENSE) = ReadRs(rsRcd, "JBEEXPENSE")
                 waResult(.UpperBound(1), COSTID) = ReadRs(rsRcd, "JBEID")
                 rsRcd.MoveNext
             Loop
             'wlLineNo = waResult(.UpperBound(1), LINENO) + 1
        End With
        tblDetail.ReBind
        tblDetail.FirstRow = 0
        
        rsRcd.Close
        Set rsRcd = Nothing
    End If
    
    wsSql = "SELECT JBCID, JBCJOBID, JBCLNTYPE, COAACCCODE, "
    wsSql = wsSql & "COADESC, JBCQTY, JBCINCOME, JBCEXPENSE "
    wsSql = wsSql & "FROM MstJobCost, MstCOA "
    wsSql = wsSql & "WHERE JBCJOBID = " & wlKey & " "
    wsSql = wsSql & "AND JBCACCID = COAACCID "
    wsSql = wsSql & "ORDER BY COAACCCODE "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
    Else
        rsRcd.MoveFirst
        With waActual
             .ReDim 0, -1, COSTCODE, COSTID
             Do While Not rsRcd.EOF
                 wiCtr = wiCtr + 1
                 .AppendRows
                 waActual(.UpperBound(1), COSTCODE) = ReadRs(rsRcd, "COAACCCODE")
                 waActual(.UpperBound(1), COSTDESC) = ReadRs(rsRcd, "COADESC")
                 waActual(.UpperBound(1), UNIT) = ReadRs(rsRcd, "JBCQTY")
                 waActual(.UpperBound(1), INCOME) = ReadRs(rsRcd, "JBCINCOME")
                 waActual(.UpperBound(1), EXPENSE) = ReadRs(rsRcd, "JBCEXPENSE")
                 waActual(.UpperBound(1), COSTID) = ReadRs(rsRcd, "JBCID")
                 rsRcd.MoveNext
             Loop
             'wlLineNo = waResult(.UpperBound(1), LINENO) + 1
        End With
        tblActual.ReBind
        tblActual.FirstRow = 0
        rsRcd.Close
        Set rsRcd = Nothing
        
    End If
    
    
    
    'Call Calc_Total
    
    LoadRecord = True
    
End Function

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String

    wsSql = "SELECT JOBSTATUS FROM MSTJOB WHERE JOBCODE = '" & Set_Quote(cboJobCode) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
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
        .TableKey = "JOBCODE"
        .KeyLen = 15
        Set .ctlKey = cboJobCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuPopUpSub_Click(Index As Integer)
    Call Call_PopUpMenu(waPopUpSub, Index)
End Sub

Private Sub Call_PopUpMenu(ByVal inArray As XArrayDB, inMnuIdx As Integer)

    Dim wsAct As String
    
    wsAct = inArray(inMnuIdx, 0)
    
    With tblDetail
    Select Case wsAct
        Case "DELETE"
           
           If IsNull(.Bookmark) Then Exit Sub
            If .EditActive = True Then Exit Sub
            gsMsg = "你是否確定要刪除此列?"
            If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then Exit Sub
            .Delete
            .Update
            If .Row = -1 Then
                .Row = 0
            End If
            .Refresh
            .SetFocus
            
        
        Case "INSERT"
            
            If .Bookmark = waResult.UpperBound(2) Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case Else
            Exit Sub
                    
            
    End Select
    
    End With
             
    
End Sub
Public Sub SetButtonStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = False
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
                .Buttons(tcRefresh).Enabled = False
                .Buttons(tcPrint).Enabled = False
            End With
            
            
        
            
        
        Case "AfrKeyAdd"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
                .Buttons(tcRefresh).Enabled = False
                .Buttons(tcPrint).Enabled = False
            End With
        
        Case "AfrKeyEdit"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = True
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
                .Buttons(tcRefresh).Enabled = True
                .Buttons(tcPrint).Enabled = True
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
                .Buttons(tcRefresh).Enabled = False
                .Buttons(tcPrint).Enabled = False
            
            End With
            
       
    
    End Select
End Sub

Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSql As String
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
    wsSql = "EXEC usp_RPTSN002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & Set_Quote(cboJobCode.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboJobCode.Text) & "', "
    wsSql = wsSql & "'" & "" & "', "
    wsSql = wsSql & "'" & String(10, "z") & "', "
    wsSql = wsSql & "'" & "0000/00/00" & "', "
    wsSql = wsSql & "'" & "9999/99/99" & "', "
    wsSql = wsSql & "'" & "%" & "', "
    wsSql = wsSql & gsLangID
    
    
    If gsLangID = "2" Then wsRptName = "C" + "RPTSN002"
    
    NewfrmPrint.ReportID = "SN002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SN002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub txtCusContactPerson_GotFocus()
    FocusMe txtCusContactPerson
End Sub

Private Sub txtCusContactPerson_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusContactPerson, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        medFromDate.SetFocus
    End If
End Sub

Private Sub txtCusContactPerson_LostFocus()
    FocusMe txtCusContactPerson, True
End Sub

Private Sub txtCusPo_GotFocus()
    FocusMe txtCusPo
End Sub

Private Sub txtCusPo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusPo, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtJobComplete.SetFocus
    End If
End Sub

Private Sub txtCusPo_LostFocus()
    FocusMe txtCusPo, True
End Sub

Private Sub txtExpense_GotFocus()
    FocusMe txtExpense
End Sub

Private Sub txtExpense_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtExpense, 11, KeyAscii)
    Call Chk_InpNum(KeyAscii, txtExpense, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If chk_txtExpense Then
            tblActual.SetFocus
        End If
    End If
End Sub

Private Sub txtExpense_LostFocus()
    FocusMe txtExpense, True
End Sub

Private Sub txtIncome_GotFocus()
    FocusMe txtIncome
End Sub

Private Sub txtIncome_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtIncome, 11, KeyAscii)
    Call Chk_InpNum(KeyAscii, txtIncome, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If chk_txtIncome Then
            txtExpense.SetFocus
        End If
    End If
End Sub

Private Sub txtIncome_LostFocus()
    FocusMe txtIncome, True
End Sub

Private Sub txtJobComplete_GotFocus()
    FocusMe txtJobComplete
End Sub

Private Sub txtJobComplete_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtJobComplete, 3, KeyAscii)
    Call Chk_InpNum(KeyAscii, txtJobComplete, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If chk_txtJobComplete Then
            txtJobRemark.SetFocus
        End If
    End If
End Sub

Private Sub txtJobComplete_LostFocus()
    FocusMe txtJobComplete, True
End Sub

Private Sub txtJobName_GotFocus()
    FocusMe txtJobName
End Sub

Private Sub txtJobName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtJobName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If chk_txtJobName Then
            txtJobType.SetFocus
        End If
    End If
End Sub

Private Sub txtJobName_LostFocus()
    FocusMe txtJobName, True
End Sub

Private Sub txtJobRemark_GotFocus()
    FocusMe txtJobRemark
End Sub

Private Sub txtJobRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtJobRemark, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        tabDetailInfo.Tab = 1
        If chkDetail.Value = 0 Then
            fraEst1.Visible = True
            fraEst2.Visible = False
            tblDetail.Visible = False
            txtIncome.SetFocus
        Else
            fraEst1.Visible = False
            fraEst2.Visible = True
            tblDetail.Visible = True
            tblDetail.SetFocus
        End If
    End If
End Sub

Private Sub txtJobRemark_LostFocus()
    FocusMe txtJobRemark, True
End Sub

Private Sub txtJobType_GotFocus()
    FocusMe txtJobType
End Sub

Private Sub txtJobType_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtJobType, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        chkDetail.SetFocus
    End If
End Sub

Private Sub txtJobType_LostFocus()
    FocusMe txtJobType
End Sub

Private Sub medFromDate_GotFocus()
    FocusMe medFromDate
End Sub

Private Sub medFromDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medFromDate Then
            If Trim(medToDate.Text) = "/  /" Then
                medToDate.Text = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(medFromDate.Text))), "yyyy/mm/dd")
            End If
            
            tabDetailInfo.Tab = 0
            medToDate.SetFocus
        End If
    End If
End Sub

Private Sub medFromDate_LostFocus()
    FocusMe medFromDate, True
End Sub

Private Sub medToDate_GotFocus()
    FocusMe medToDate
End Sub

Private Sub medToDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medToDate Then
            tabDetailInfo.Tab = 0
            txtCusPo.SetFocus
        End If
    End If
End Sub

Private Sub medToDate_LostFocus()
    FocusMe medToDate, True
End Sub

Private Function Chk_medFromDate() As Boolean
    
    Chk_medFromDate = False
    
    If Trim(medFromDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medFromDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medFromDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medFromDate.SetFocus
        Exit Function
    End If
    
    Chk_medFromDate = True

End Function

Private Function Chk_medToDate() As Boolean
    Chk_medToDate = False
    
    If Trim(medToDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medToDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medToDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medToDate.SetFocus
        Exit Function
    End If
    
    Chk_medToDate = True
End Function

Private Function chk_txtJobComplete() As Boolean
    chk_txtJobComplete = False
    
    If Trim(txtJobComplete.Text) = "" Then
        gsMsg = "完成比例錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtJobComplete.SetFocus
        Exit Function
    End If
    
    If CInt(txtJobComplete) > 100 Or CInt(txtJobComplete) < 0 Then
        gsMsg = "完成比例錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtJobComplete.SetFocus
        Exit Function
    End If
    
    chk_txtJobComplete = True
End Function

Private Function chk_txtIncome() As Boolean
    chk_txtIncome = False
    
    If Trim(txtIncome.Text) = "" Then
        gsMsg = "收入錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtIncome.SetFocus
        Exit Function
    End If
    
    If CInt(txtIncome) < 0 Then
        gsMsg = "收入錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtIncome.SetFocus
        Exit Function
    End If
    
    chk_txtIncome = True
End Function

Private Function chk_txtExpense() As Boolean
    chk_txtExpense = False
    
    If Trim(txtExpense.Text) = "" Then
        gsMsg = "支出錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtExpense.SetFocus
        Exit Function
    End If
    
    If CInt(txtExpense) < 0 Then
        gsMsg = "支出錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtExpense.SetFocus
        Exit Function
    End If
    
    chk_txtExpense = True
End Function

Private Function chk_txtJobName() As Boolean
    chk_txtJobName = False
    
    If Trim(txtJobName.Text) = "" Then
        gsMsg = "必須輸入項目名稱!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtJobName.SetFocus
        Exit Function
    End If
    
    chk_txtJobName = True
End Function

Private Function cmdRefresh() As Boolean
    cmdRefresh = False
    cmdRefresh = True
End Function
