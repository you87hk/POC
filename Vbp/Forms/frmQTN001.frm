VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmQTN001 
   Caption         =   "Quotations"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   -15
   ClientWidth     =   9795
   Icon            =   "frmQTN001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11040
      OleObjectBlob   =   "frmQTN001.frx":030A
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   4575
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   8055
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmQTN001.frx":2A0D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboSaleCode"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboCusCode"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboDocNo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboMethodCode"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmQTN001.frx":2A29
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTotalQty"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblGrsAmtOrg"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblDisAmtOrg"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblNetAmtOrg"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblDspTotalQty"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblDspGrsAmtOrg"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblDspDisAmtOrg"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblDspNetAmtOrg"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "tblDetail"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.ComboBox cboSaleCode 
         Height          =   300
         Left            =   -72960
         TabIndex        =   4
         Top             =   3840
         Width           =   1890
      End
      Begin VB.ComboBox cboCusCode 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72960
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cboDocNo 
         Height          =   300
         Left            =   -72960
         TabIndex        =   0
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cboMethodCode 
         Height          =   300
         Left            =   -72960
         TabIndex        =   5
         Top             =   4200
         Width           =   1890
      End
      Begin VB.Frame fra1 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   10
         Top             =   240
         Width           =   11295
         Begin VB.PictureBox picRmk 
            BackColor       =   &H80000009&
            Height          =   1695
            Left            =   1920
            ScaleHeight     =   1635
            ScaleWidth      =   8715
            TabIndex        =   37
            Top             =   4320
            Width           =   8775
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   42
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   345
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   41
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   40
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   690
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   39
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1035
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   38
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1395
               Width           =   7545
            End
         End
         Begin VB.TextBox txtRevNo 
            Height          =   324
            Left            =   5880
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "12345678901234567890"
            Top             =   240
            Width           =   408
         End
         Begin MSMask.MaskEdBox medDocDate 
            Height          =   285
            Left            =   5880
            TabIndex        =   3
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblRmk 
            Caption         =   "RMK"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   4440
            Width           =   1620
         End
         Begin VB.Label lblDspMethodDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   3840
            TabIndex        =   28
            Top             =   3960
            Width           =   6855
         End
         Begin VB.Label lblMethodCode 
            Caption         =   "METHODCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   27
            Top             =   4020
            Width           =   1665
         End
         Begin VB.Label lblDspSaleDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   3840
            TabIndex        =   26
            Top             =   3600
            Width           =   6855
         End
         Begin VB.Label lblSaleCode 
            Caption         =   "SALECODE"
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
            TabIndex        =   25
            Top             =   3600
            Width           =   1185
         End
         Begin VB.Label lblCusTel 
            Caption         =   "CUSTEL"
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
            Top             =   2580
            Width           =   1215
         End
         Begin VB.Label lblCusFax 
            Caption         =   "CUSFAX"
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
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblDspCusFax 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1920
            TabIndex        =   22
            Top             =   2880
            Width           =   3135
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
            TabIndex        =   21
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label lblDspCusTel 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1920
            TabIndex        =   20
            Top             =   2520
            Width           =   3135
         End
         Begin VB.Label lblDspCusName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1920
            TabIndex        =   19
            Top             =   960
            Width           =   8775
         End
         Begin VB.Label lblCusCode 
            Caption         =   "CUSCODE"
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
            Left            =   120
            TabIndex        =   18
            Top             =   660
            Width           =   1815
         End
         Begin VB.Label lblDocDate 
            Caption         =   "DOCDATE"
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
            Left            =   4200
            TabIndex        =   17
            Top             =   660
            Width           =   1560
         End
         Begin VB.Label lblRevNo 
            Caption         =   "REVNO"
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
            Left            =   4200
            TabIndex        =   16
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label lblDocNo 
            Caption         =   "DOCNO"
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
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblDspCusAddress 
            BorderStyle     =   1  '單線固定
            Height          =   780
            Left            =   1920
            TabIndex        =   14
            Top             =   1680
            Width           =   8775
         End
         Begin VB.Label lblCusAddress 
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
            TabIndex        =   13
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblDspCusContact 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1920
            TabIndex        =   12
            Top             =   1320
            Width           =   8775
         End
         Begin VB.Label lblCusContact 
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
            TabIndex        =   11
            Top             =   1380
            Width           =   1215
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "frmQTN001.frx":2A45
         TabIndex        =   6
         Top             =   840
         Width           =   11655
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
         Left            =   6960
         TabIndex        =   36
         Top             =   480
         Width           =   4770
      End
      Begin VB.Label lblDspDisAmtOrg 
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
         Left            =   4560
         TabIndex        =   35
         Top             =   480
         Width           =   2370
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
         Left            =   2040
         TabIndex        =   34
         Top             =   480
         Width           =   2490
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
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label lblNetAmtOrg 
         Alignment       =   2  '置中對齊
         Caption         =   "NETAMTORG"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   32
         Top             =   120
         Width           =   4755
      End
      Begin VB.Label lblDisAmtOrg 
         Alignment       =   2  '置中對齊
         Caption         =   "DISAMTORG"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label lblGrsAmtOrg 
         Alignment       =   2  '置中對齊
         Caption         =   "GRSAMTORG"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   120
         Width           =   2475
      End
      Begin VB.Label lblTotalQty 
         Alignment       =   2  '置中對齊
         Caption         =   "NETAMTORG"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   120
         Width           =   1755
      End
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   10080
      Top             =   2760
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
            Picture         =   "frmQTN001.frx":B87C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":C156
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":CA30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":CE82
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":D2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":D5EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":DA40
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":DE92
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":E1AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":E4C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":E918
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQTN001.frx":F1F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
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
Attribute VB_Name = "frmQTN001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private waResult As New XArrayDB
Private waItem As New XArrayDB
Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB

Private waPopUpSub As New XArrayDB
Private wcCombo As Control




Private wsOldCusNo As String
Private wsDueDate As String
Private wsPayCode As String
Private wsWhsCode As String
Private wsMethodCode As String
Private wsCurCode As String
Private wdExcr As Double
Private wgsTitle As String
Private wsShipName As String
Private wsShipAdr1 As String
Private wsShipAdr2 As String
Private wsShipAdr3 As String
Private wsShipAdr4 As String


Private Const BOOKCODE = 0
Private Const BARCODE = 1
Private Const WHSCODE = 2
Private Const BOOKNAME = 3
Private Const WANTED = 4
Private Const PUBLISHER = 5
Private Const Qty = 6
Private Const Price = 7
Private Const DisPer = 8
Private Const Dis = 9
Private Const Amt = 10
Private Const Net = 11
Private Const Netl = 12
Private Const Disl = 13
Private Const Amtl = 14
Private Const BOOKID = 15
Private Const BOM = 16


Private Const STYPE = 0
Private Const SITMCODE = 1
Private Const SITMDESC = 2
Private Const SUPRICE = 3
Private Const SQTY = 4
Private Const SAMT = 5
Private Const SJBCODE = 6
Private Const SSTATUS = 7

Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"


Private wiOpenDoc As Integer
Private wiAction As Integer
Private wiRevNo As Integer
Private wlCusID As Long
Private wlSaleID As Long
Private wlCusTyp As Long

Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "soaQTHD"
Private wsFormID As String
Private wsUsrID As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String





Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, BOOKCODE, BOM
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    
    frmQTN002.InvDoc.ReDim 0, -1, STYPE, SSTATUS
    waItem.ReDim 0, -1, STYPE, SSTATUS
    
 
    
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

    Call SetButtonStatus("AfrActEdit")
    Call SetFieldStatus("Default")
    Call SetFieldStatus("AfrActEdit")
    
    Call SetDateMask(medDocDate)
      
    
    wsOldCusNo = ""
     wiRevNo = Format(0, "##0")
  
    
    wlKey = 0
    wlCusID = 0
    wlSaleID = 0
    medDocDate.Text = Dsp_Date(Now)
    
    wsPayCode = "'"
    wsDueDate = ""
    wsShipName = ""
    wsShipAdr1 = ""
    wsShipAdr2 = ""
    wsShipAdr3 = ""
    wsShipAdr4 = ""
    
    cboMethodCode.Text = wsMethodCode
    
    tblCommon.Visible = False

    
    Me.Caption = wsFormCaption
    tabDetailInfo.Tab = 0
    FocusMe cboDocNo
    
    
   
    

End Sub



Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
End Sub






Private Sub Ini_Scr_AfrKey()
   
    If LoadRecord() = False Then
        wiAction = AddRec
        txtRevNo.Text = Format(0, "##0")
        txtRevNo.Enabled = False
        medDocDate.Text = Dsp_Date(Now)
        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrID) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrID
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        txtRevNo.Enabled = True
        wsOldCusNo = cboCusCode.Text
      
         Call SetButtonStatus("AfrKeyEdit")
    End If
    
     Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    
    Call SetFieldStatus("AfrKey")
    Call LoadQTItem
    cboCusCode.SetFocus
    tabDetailInfo.Tab = 0
   
    
  '      wiAction = AddRec
  '      Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
  '      Call SetButtonStatus("AfrKeyAdd")
  '      Call SetFieldStatus("AfrKey")
        
  '      cboSaleCode.SetFocus
End Sub








Private Sub cboDocNo_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSql = "SELECT QTHDDOCNO, CUSCODE, QTHDDOCDATE "
    wsSql = wsSql & " FROM soaQTHD, mstCUSTOMER "
    wsSql = wsSql & " WHERE QTHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSql = wsSql & " AND QTHDCUSID  = CUSID "
    wsSql = wsSql & " AND QTHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY QTHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNo.Left + tabDetailInfo.Left, cboDocNo.Top + cboDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub



Private Sub cboDocNo_GotFocus()
FocusMe cboDocNo
End Sub

Private Sub cboDocNo_LostFocus()
FocusMe cboDocNo, True
End Sub

Private Sub cboDocNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboDocNo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If Chk_cboDocNo() = False Then Exit Sub
        
        Call Ini_Scr_AfrKey
        
    End If

End Sub

Private Function Chk_cboDocNo() As Boolean
    
Dim wsStatus As String
    
    Chk_cboDocNo = False
    
    If Trim(cboDocNo.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboDocNo.SetFocus
        tabDetailInfo.Tab = 0
   
        Exit Function
    End If
    
        
   If Chk_DocNo(cboDocNo, wsStatus) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            tabDetailInfo.Tab = 0
   
            Exit Function
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            tabDetailInfo.Tab = 0
   
            Exit Function
        End If
    
    
    End If
    
    
    Chk_cboDocNo = True

End Function

Private Sub cboMethodCode_GotFocus()
    FocusMe cboMethodCode
End Sub

Private Sub cboMethodCode_LostFocus()
    FocusMe cboMethodCode, True
End Sub


Private Sub cboMethodCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboMethodCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboMethodCode = False Then
                Exit Sub
        End If
        
        txtRmk(1).SetFocus
        
        'tblDetail.SetFocus
        'tabDetailInfo.Tab = 1
   
       
    End If
    
End Sub

Private Sub cboMethodCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboMethodCode
    
    wsSql = "SELECT MethodCode, MethodDESC FROM mstMethod WHERE MethodCode LIKE '%" & IIf(cboMethodCode.SelLength > 0, "", Set_Quote(cboMethodCode.Text)) & "%' "
    wsSql = wsSql & "AND METHODSTATUS = '1' "
    wsSql = wsSql & "ORDER BY MethodCode "
    Call Ini_Combo(2, wsSql, cboMethodCode.Left + tabDetailInfo.Left, cboMethodCode.Top + cboMethodCode.Height + tabDetailInfo.Top, tblCommon, "QTN001", "TBLMETHODCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboMethodCode() As Boolean
Dim wsDesc As String

    Chk_cboMethodCode = False
     
    If Trim(cboMethodCode.Text) = "" Then
        gsMsg = "必需輸入銷售渠道!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboMethodCode.SetFocus
        tabDetailInfo.Tab = 0
   
        Exit Function
    End If
    
    
    If Chk_Method(cboMethodCode, wsDesc) = False Then
        gsMsg = "沒有此銷售渠道!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboMethodCode.SetFocus
        tabDetailInfo.Tab = 0
        lblDspMethodDesc = ""
       Exit Function
    End If
    
    lblDspMethodDesc = wsDesc
    
    Chk_cboMethodCode = True
    
End Function
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
        
         'Case vbKeyF9
        
         '   If tbrProcess.Buttons(tcFind).Enabled = True Then Call cmdFind
            
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
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblRevNo.Caption = Get_Caption(waScrItm, "REVNO")
  
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    lblCusTel.Caption = Get_Caption(waScrItm, "CUSTEL")
    lblCusFax.Caption = Get_Caption(waScrItm, "CUSFAX")
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblCusAddress.Caption = Get_Caption(waScrItm, "CUSADDRESS")
    lblCusContact.Caption = Get_Caption(waScrItm, "CONTACT")
    lblMethodCode.Caption = Get_Caption(waScrItm, "METHODCODE")
    
    
    
    lblGrsAmtOrg.Caption = Get_Caption(waScrItm, "GRSAMTORG")
    lblNetAmtOrg.Caption = Get_Caption(waScrItm, "NETAMTORG")
    lblDisAmtOrg.Caption = Get_Caption(waScrItm, "DISAMTORG")
    lblTotalQty.Caption = Get_Caption(waScrItm, "TOTALQTY")
    
    With tblDetail
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(PUBLISHER).Caption = Get_Caption(waScrItm, "PUBLISHER")
        .Columns(Qty).Caption = Get_Caption(waScrItm, "QTY")
        .Columns(Price).Caption = Get_Caption(waScrItm, "PRICE")
        .Columns(DisPer).Caption = Get_Caption(waScrItm, "DISPER")
        .Columns(Net).Caption = Get_Caption(waScrItm, "NET")
        .Columns(BOM).Caption = Get_Caption(waScrItm, "BOM")
    End With
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")

    wsActNam(1) = Get_Caption(waScrItm, "SOADD")
    wsActNam(2) = Get_Caption(waScrItm, "SOEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SODELETE")
    
    
     Call Ini_PgmMenu(mnuPopUpSub, "POPUP", waPopUpSub)
  
    
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
    Set waItem = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmQTN001 = Nothing

End Sub
















Private Sub medDocDate_GotFocus()
    
  FocusMe medDocDate
    
End Sub


Private Sub medDocDate_LostFocus()

    FocusMe medDocDate, True
    
End Sub


Private Sub medDocDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medDocDate Then
        cboSaleCode.SetFocus
        tabDetailInfo.Tab = 0
        End If
    End If
End Sub

Private Function Chk_medDocDate() As Boolean

    
    Chk_medDocDate = False
    
    If Trim(medDocDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocDate.SetFocus
        tabDetailInfo.Tab = 0
   
        Exit Function
    End If
    
    If Chk_Date(medDocDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocDate.SetFocus
        tabDetailInfo.Tab = 0
   
        Exit Function
    End If
    
    
    Chk_medDocDate = True

End Function




















Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
         If cboCusCode.Enabled Then
            cboCusCode.SetFocus
         End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
    
        If tblDetail.Enabled Then
            tblDetail.SetFocus
        End If
    
    End If
End Sub



Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case BOOKCODE
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
              Case BOOKCODE
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





Private Function Chk_KeyFld() As Boolean
    
        
    Chk_KeyFld = False
    
    If chk_cboCusCode = False Then
        Exit Function
    End If
    
    If Chk_medDocDate = False Then
        Exit Function
    End If
    
    
    tblDetail.Enabled = True
    Chk_KeyFld = True

End Function
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Then
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
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_QTN001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlCusID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, txtRevNo.Text)
    Call SetSPPara(adcmdSave, 8, wsCurCode)
    Call SetSPPara(adcmdSave, 9, wdExcr)
    Call SetSPPara(adcmdSave, 10, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 11, wsDueDate)
    Call SetSPPara(adcmdSave, 12, "")
    Call SetSPPara(adcmdSave, 13, "")
    
    Call SetSPPara(adcmdSave, 14, wlSaleID)
    Call SetSPPara(adcmdSave, 15, wlCusTyp)
    
    Call SetSPPara(adcmdSave, 16, wsPayCode)
    Call SetSPPara(adcmdSave, 17, "")
    Call SetSPPara(adcmdSave, 18, "")
    Call SetSPPara(adcmdSave, 19, cboMethodCode)
    Call SetSPPara(adcmdSave, 20, "")
    Call SetSPPara(adcmdSave, 21, "")
    
    Call SetSPPara(adcmdSave, 22, "")
    Call SetSPPara(adcmdSave, 23, "")
    Call SetSPPara(adcmdSave, 24, "")
    
    Call SetSPPara(adcmdSave, 25, "")
    Call SetSPPara(adcmdSave, 26, "")
    Call SetSPPara(adcmdSave, 27, "")
    Call SetSPPara(adcmdSave, 28, "")
    Call SetSPPara(adcmdSave, 29, wsShipName)
    Call SetSPPara(adcmdSave, 30, wsShipAdr1)
    Call SetSPPara(adcmdSave, 31, wsShipAdr2)
    Call SetSPPara(adcmdSave, 32, wsShipAdr3)
    Call SetSPPara(adcmdSave, 33, wsShipAdr4)
    
    For i = 1 To 5
    Call SetSPPara(adcmdSave, 34 + i - 1, txtRmk(i))
    Next
    For i = 1 To 5
    Call SetSPPara(adcmdSave, 39 + i - 1, "")
    Next
    
    
    Call SetSPPara(adcmdSave, 44, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 45, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 46, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 47, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 48, lblDspNetAmtOrg)
    Call SetSPPara(adcmdSave, 49, lblDspNetAmtOrg)
    
    Call SetSPPara(adcmdSave, 50, wsFormID)
    Call SetSPPara(adcmdSave, 51, gsWorkStationID)
    Call SetSPPara(adcmdSave, 52, gsUserID)
    Call SetSPPara(adcmdSave, 53, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 54)
    wsDocNo = GetSPPara(adcmdSave, 55)
    
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_QTN001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, BOOKCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, BOOKCODE))
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, BARCODE))
                Call SetSPPara(adcmdSave, 5, wiCtr + 1)
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, BOOKNAME))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, Qty))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, Price))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, DisPer))
                Call SetSPPara(adcmdSave, 10, Set_MedDate(waResult(wiCtr, WANTED)))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, WHSCODE))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, Amt))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, Amtl))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, Dis))
                Call SetSPPara(adcmdSave, 15, waResult(wiCtr, Disl))
                Call SetSPPara(adcmdSave, 16, waResult(wiCtr, Net))
                Call SetSPPara(adcmdSave, 17, waResult(wiCtr, Netl))
                Call SetSPPara(adcmdSave, 18, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 19, waResult(wiCtr, BOM))
                adcmdSave.Execute
            End If
        Next
    End If
    
     If waItem.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_QTN001C"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waItem.UpperBound(1)
            If Trim(waItem(wiCtr, SJBCODE)) <> "" And Trim(waItem(wiCtr, SITMCODE)) <> "" And Trim(waItem(wiCtr, SSTATUS)) = "1" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waItem(wiCtr, SJBCODE))
                Call SetSPPara(adcmdSave, 4, waItem(wiCtr, SITMCODE))
                Call SetSPPara(adcmdSave, 5, wiCtr + 1)
                Call SetSPPara(adcmdSave, 6, waItem(wiCtr, SITMDESC))
                Call SetSPPara(adcmdSave, 7, waItem(wiCtr, SQTY))
                Call SetSPPara(adcmdSave, 8, waItem(wiCtr, SUPRICE))
                Call SetSPPara(adcmdSave, 9, waItem(wiCtr, SAMT))
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
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Then
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
        
    adcmdDelete.CommandText = "USP_QTN001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, wlCusID)
    Call SetSPPara(adcmdDelete, 6, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 7, txtRevNo.Text)
    Call SetSPPara(adcmdDelete, 8, wsCurCode)
    Call SetSPPara(adcmdDelete, 9, wdExcr)
    Call SetSPPara(adcmdDelete, 10, "")
    
    Call SetSPPara(adcmdDelete, 11, wsDueDate)
    Call SetSPPara(adcmdDelete, 12, "")
    Call SetSPPara(adcmdDelete, 13, "")
    
    Call SetSPPara(adcmdDelete, 14, wlSaleID)
    Call SetSPPara(adcmdDelete, 15, wlCusTyp)
    
    Call SetSPPara(adcmdDelete, 16, wsPayCode)
    Call SetSPPara(adcmdDelete, 17, "")
    Call SetSPPara(adcmdDelete, 18, "")
    Call SetSPPara(adcmdDelete, 19, cboMethodCode)
    Call SetSPPara(adcmdDelete, 20, "")
    Call SetSPPara(adcmdDelete, 21, "")
    
    Call SetSPPara(adcmdDelete, 22, "")
    Call SetSPPara(adcmdDelete, 23, "")
    Call SetSPPara(adcmdDelete, 24, "")
    
    Call SetSPPara(adcmdDelete, 25, "")
    Call SetSPPara(adcmdDelete, 26, "")
    Call SetSPPara(adcmdDelete, 27, "")
    Call SetSPPara(adcmdDelete, 28, "")
    Call SetSPPara(adcmdDelete, 29, "")
    Call SetSPPara(adcmdDelete, 30, "")
    Call SetSPPara(adcmdDelete, 31, "")
    Call SetSPPara(adcmdDelete, 32, "")
    Call SetSPPara(adcmdDelete, 33, "")
    
    For i = 1 To 10
    Call SetSPPara(adcmdDelete, 34 + i - 1, "")
    Next
    
    Call SetSPPara(adcmdDelete, 44, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdDelete, 45, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdDelete, 46, lblDspDisAmtOrg)
    Call SetSPPara(adcmdDelete, 47, lblDspDisAmtOrg)
    Call SetSPPara(adcmdDelete, 48, lblDspNetAmtOrg)
    Call SetSPPara(adcmdDelete, 49, lblDspNetAmtOrg)
    
    Call SetSPPara(adcmdDelete, 50, wsFormID)
    Call SetSPPara(adcmdDelete, 51, gsWorkStationID)
    Call SetSPPara(adcmdDelete, 52, gsUserID)
    Call SetSPPara(adcmdDelete, 53, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 54)
    wsDocNo = GetSPPara(adcmdDelete, 55)
   
    
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
    
    
    If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboCusCode() Then Exit Function
    If Not Chk_cboSaleCode Then Exit Function
    If Not Chk_cboMethodCode Then Exit Function
    
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    Dim wlCtr1 As Long
    
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, BOOKCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    tabDetailInfo.Tab = 1
   
                    Exit Function
                End If
            End If
            For wlCtr1 = 0 To .UpperBound(1)
                If wlCtr <> wlCtr1 Then
                    If waResult(wlCtr, BOOKCODE) = waResult(wlCtr1, BOOKCODE) Then
                      gsMsg = "重覆書本!"
                      MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                      tblDetail.SetFocus
                      Exit Function
                    End If
                End If
            Next
        Next
    
    
    
    
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "訂購單沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
        tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    If Chk_NoDup(To_Value(tblDetail.Bookmark)) = False Then
        tblDetail.FirstRow = tblDetail.Row
        tblDetail.Col = BOOKCODE
        tblDetail.SetFocus
        Exit Function
    End If
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
    


Private Sub cmdNew()

    Dim newForm As New frmQTN001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmQTN001
    
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
    wsFormID = "QTN001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "QT"
    
    Call LoadWSINFO
    

End Sub



Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
    cboDocNo.SetFocus
    
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
     '   Case tcFind
      '      Call cmdFind
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
    Call Ini_Combo(2, wsSql, cboCusCode.Left + tabDetailInfo.Left, cboCusCode.Top + cboCusCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
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
        If wiAction = AddRec Or wsOldCusNo <> cboCusCode.Text Then Call Get_DefVal
           
            If Chk_KeyFld Then
                cboSaleCode.SetFocus
            End If
            
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
        lblDspCusTel.Caption = wsTel
        lblDspCusFax.Caption = wsFax
    Else
        wlCusID = 0
        lblDspCusName.Caption = ""
        lblDspCusTel.Caption = ""
        lblDspCusFax.Caption = ""
        gsMsg = "客戶不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    End If
    
    chk_cboCusCode = True

End Function

Private Sub Get_DefVal()
    
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSql As String
    Dim wsExcDesc As String
    Dim wsExcRate As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSql = "SELECT * "
    wsSql = wsSql & "FROM  mstCUSTOMER "
    wsSql = wsSql & "WHERE CUSID = " & wlCusID
    rsDefVal.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        wlSaleID = ReadRs(rsDefVal, "CUSSALEID")
        wlCusTyp = To_Value(ReadRs(rsDefVal, "CUSTYPID"))
        wsPayCode = ReadRs(rsDefVal, "CUSPAYCODE")
        lblDspCusContact = ReadRs(rsDefVal, "CUSCONTACTPERSON")
        lblDspCusAddress = ReadRs(rsDefVal, "CUSADDRESS1") & Chr(13) & Chr(10) & _
                           ReadRs(rsDefVal, "CUSADDRESS2") & Chr(13) & Chr(10) & _
                           ReadRs(rsDefVal, "CUSADDRESS3") & Chr(13) & Chr(10) & _
                           ReadRs(rsDefVal, "CUSADDRESS4")
        wsShipName = ReadRs(rsDefVal, "CUSNAME")
        wsShipAdr1 = ReadRs(rsDefVal, "CUSADDRESS1")
        wsShipAdr2 = ReadRs(rsDefVal, "CUSADDRESS2")
        wsShipAdr3 = ReadRs(rsDefVal, "CUSADDRESS3")
        wsShipAdr4 = ReadRs(rsDefVal, "CUSADDRESS4")
        
          Else
        wlSaleID = 0
        wlCusTyp = 0
        wsPayCode = ""
        wsShipName = ""
        wsShipAdr1 = ""
        wsShipAdr2 = ""
        wsShipAdr3 = ""
        wsShipAdr4 = ""
        
        
        
    End If
    rsDefVal.Close
    Set rsDefVal = Nothing
    
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    
    'get Due Date Payment Term
    wsDueDate = Dsp_Date(Get_DueDte(wsPayCode, medDocDate))

End Sub



Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 1
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = BOOKCODE To BOM
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case BOOKCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 15
                Case BARCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case WHSCODE
                   .Columns(wiCtr).Visible = False
                   .Columns(wiCtr).DataWidth = 10
                Case BOOKNAME
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = False
                Case WANTED
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 10
                    
                Case PUBLISHER
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Visible = False
                Case Qty
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case Price
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case DisPer
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Net
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Dis
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
                Case Amt
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                    
                Case Netl
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                Case Disl
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                Case Amtl
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                Case BOOKID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case BOM
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).Width = 500
            End Select
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
Dim sTemp As String
   
    With tblDetail
  '      sTemp = .Columns(ColIndex)
        .Update
    End With


  '  If ColIndex = BOOKCODE Then
  '      Call LoadBookGroup(sTemp)
  '  End If
     
     
End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsBookID As String
Dim wsBookCode As String
Dim wsBarCode As String
Dim wsBookName As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case BOOKCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdBookCode(.Columns(ColIndex).Text, wsBookID, wsBookCode, wsBarCode, wsBookName, wsPub, wdPrice, wdDisPer) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If

                
                .Columns(BOOKID).Text = wsBookID
                .Columns(BARCODE).Text = wsBarCode
                .Columns(BOOKNAME).Text = wsBookName
                .Columns(PUBLISHER).Text = wsPub
                .Columns(Price).Text = Format(wdPrice, gsAmtFmt)
                .Columns(Qty).Text = "1"
                .Columns(DisPer).Text = Format(wdDisPer, "#,##0")
                .Columns(WANTED).Text = medDocDate
                .Columns(WHSCODE).Text = wsWhsCode
                .Columns(BOM).Text = "N"
                
                
                If Trim(.Columns(ColIndex).Text) <> wsBookCode Then
                    .Columns(ColIndex).Text = wsBookCode
                End If
                If Trim(.Columns(Price).Text) <> "" Then
                .Columns(Amt).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text), gsAmtFmt)
                .Columns(Amtl).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text) * wdExcr, gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Dis).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                .Columns(Disl).Text = Format(To_Value(.Columns(Amtl).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(Dis).Text) <> "" Then
                .Columns(Net).Text = Format(To_Value(.Columns(Amt).Text) - To_Value(.Columns(Dis).Text), gsAmtFmt)
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) - To_Value(.Columns(Disl).Text), gsAmtFmt)
                
                End If
                
        '     Case WhsCode
        '        If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
        '            GoTo Tbl_BeforeColUpdate_Err
        '        End If
                
        '        If Chk_grdWhsCode(.Columns(ColIndex).Text) = False Then
        '                GoTo Tbl_BeforeColUpdate_Err
        '        End If
        '    Case WANTED
        '        If Chk_grdWantedDate(.Columns(ColIndex).Text) = False Then
        '                GoTo Tbl_BeforeColUpdate_Err
         '       End If
            Case Qty, Price, DisPer
            
                If ColIndex = Qty Then
                        If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                        End If
                ElseIf ColIndex = DisPer Then
                        If Chk_grdDisPer(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                        End If
                End If
                    
                If Trim(.Columns(Price).Text) <> "" Then
                .Columns(Amt).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text), gsAmtFmt)
                .Columns(Amtl).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Dis).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                .Columns(Disl).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(Dis).Text) <> "" Then
                .Columns(Net).Text = Format(To_Value(.Columns(Amt).Text) - To_Value(.Columns(Dis).Text), gsAmtFmt)
                .Columns(Netl).Text = Format(To_Value(.Columns(Amt).Text) - To_Value(.Columns(Dis).Text), gsAmtFmt)
                
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
            Case BOOKCODE
                

                wsSql = "SELECT JBCODE, JBTYPE, JBDESC FROM mstJOB "
                wsSql = wsSql & " WHERE JBSTATUS <> '2' AND JBCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                If waResult.UpperBound(1) > -1 Then
                      wsSql = wsSql & " AND JBCODE NOT IN ( "
                      For wiCtr = 0 To waResult.UpperBound(1)
                            wsSql = wsSql & " '" & waResult(wiCtr, BOOKCODE) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                      Next
                End If
                wsSql = wsSql & " ORDER BY JBCODE "
                
                Call Ini_Combo(3, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLBOOKCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
          Case BOM
              
                
                If wiAction = DelRec Or Trim(.Columns(BOOKCODE).Text) = "" Then Exit Sub
                    
                If waItem.UpperBound(1) >= 0 Then
                    frmQTN002.InvDoc.ReDim 0, waItem.UpperBound(1), STYPE, SSTATUS
                 End If
                    
                    frmQTN002.InJob = .Columns(BOOKCODE).Text
                    frmQTN002.InvDoc = waItem
                    frmQTN002.Show vbModal
                    waItem.ReDim 0, frmQTN002.InvDoc.UpperBound(1), STYPE, SSTATUS
                    Set waItem = frmQTN002.InvDoc
                    Unload frmQTN002
                    .Columns(BOM).Text = "Y"
                
                
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
                Case Net, BOM
                    KeyCode = vbKeyDown
                    .Col = BOOKCODE
                Case BOOKCODE, Qty
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case BARCODE
                    KeyCode = vbDefault
                    .Col = BOOKNAME
                Case BOOKNAME
                    KeyCode = vbDefault
                    .Col = Qty
                Case DisPer
                    KeyCode = vbDefault
                    .Col = Net
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            Select Case .Col
            Case DisPer, Price, BARCODE
                   .Col = .Col - 1
            Case BOM
                   .Col = Net
            Case Net
                   .Col = DisPer
            Case Qty
                   .Col = BOOKNAME
            Case BOOKNAME
                   .Col = BARCODE
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
            Case BOOKCODE, Qty, Price
                    .Col = .Col + 1
            Case BARCODE
                   .Col = BOOKNAME
            Case BOOKNAME
                   .Col = Qty
            Case DisPer
                   .Col = Net
            Case Net
                   .Col = BOM
            End Select
            
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    Select Case tblDetail.Col
        
        Case Qty
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        Case Price, DisPer
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
        
              
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = BOOKCODE
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case BOOKCODE
                    Call Chk_grdBookCode(.Columns(BOOKCODE).Text, "", "", "", "", "", 0, 0)
                Case WHSCODE
                    Call Chk_grdWhsCode(.Columns(WHSCODE).Text)
                Case WANTED
                    Call Chk_grdWantedDate(.Columns(WANTED).Text)
                Case Qty
                    Call Chk_grdQty(.Columns(Qty).Text)
                Case DisPer
                    Call Chk_grdDisPer(.Columns(DisPer).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdBookCode(inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String, outPub As String, outPrice As Double, outDisPer As Double) As Boolean
    
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    
    If Trim(inAccNo) = "" Then
       gsMsg = "Must Input a Code!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       Chk_grdBookCode = False
       Exit Function
    End If
    
    wsSql = "SELECT JBCODE, JBTYPE, JBDESC, JBUNITPRICE  FROM mstJOB "
    wsSql = wsSql & " WHERE JBCODE = '" & Set_Quote(inAccNo) & "' "
    
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       outAccID = 0
       outAccNo = ReadRs(rsDes, "JBCODE")
       OutName = ReadRs(rsDes, "JBDESC")
       OutBarCode = ReadRs(rsDes, "JBTYPE")
       outPub = ""
       outPrice = To_Value(ReadRs(rsDes, "JBUNITPRICE"))
       wsCurr = "HKD"
       
       'wdPrice = getCusItemPrice(wlCusID, outAccID, wsCurCode)
       wdPrice = 0
       
     '  If Chk_NoDup2(outAccNo, wsWhsCode) = False Then
     '     tblDetail.SetFocus
     '     tblDetail.Col = BOOKCODE
     '     Exit Function
     '  End If
    
       
       If wdPrice = 0 Then
       If wsCurCode <> wsCurr Then
       If getExcRate(wsCurr, medDocDate, wsExcr, "") = True Then
       outPrice = NBRnd(outPrice * To_Value(wsExcr) / wdExcr, giExrDp)
       End If
       End If
       Else
        outPrice = wdPrice
       End If
       
        'outDisPer = Get_SaleDiscount(cboMethodCode, To_Value(wlCusID), To_Value(outAccID))
       outDisPer = 0
       
       Chk_grdBookCode = True
    Else
        outAccID = ""
        outAccNo = inAccNo
        OutName = ""
        OutBarCode = ""
        outPub = ""
        outPrice = 0
        outDisPer = 0
        Chk_grdBookCode = True
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdWhsCode(inNo As String) As Boolean
    
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdWhsCode = False
    
    If Trim(inNo) = "" Then
        gsMsg = "必需輸入貨倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    wsSql = "SELECT *  FROM mstWareHouse"
    wsSql = wsSql & " WHERE WHSCODE = '" & Set_Quote(inNo) & "' "
    wsSql = wsSql & " AND WHSSTATUS = '1' "
       
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
       Chk_grdWhsCode = True
    Else
        gsMsg = "沒有此貨倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdWhsCode = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing

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

Private Function Chk_grdDisPer(inCode As String) As Boolean
    
    Chk_grdDisPer = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入折扣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdDisPer = False
        Exit Function
    End If

    If To_Value(inCode) < 0 Or To_Value(inCode) > 100 Then
        gsMsg = "折扣必需為零至一百!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdDisPer = False
        Exit Function
    End If
    
End Function
Private Function Chk_grdWantedDate(inCode As String) As Boolean
    
    Chk_grdWantedDate = False
    
    If Trim(inCode) = "/  /" Or Trim(inCode) = "" Then
        Chk_grdWantedDate = True
        Exit Function
    End If

   
    
    If Chk_MedDate(inCode) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    Chk_grdWantedDate = True
    
    
End Function





Private Function Chk_Amount(inAmt As String) As Integer
    
    Chk_Amount = False
    
    If Trim(inAmt) = "" Then
        gsMsg = "必需輸入金額!"
        MsgBox gsMsg, vbOKOnly, gsTitle
       Exit Function
    End If
    
  '  If To_Value(inAmt) = 0 Then
  '     gsMsg = "Amount Must not zero!"
  '      MsgBox gsMsg, vbOKOnly, gsTitle
  '     Exit Function
  '  End If
    Chk_Amount = True

End Function

Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(BOOKCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, BOOKCODE)) = "" And _
                   Trim(waResult(inRow, BOOKNAME)) = "" And _
                   Trim(waResult(inRow, PUBLISHER)) = "" And _
                   Trim(waResult(inRow, Qty)) = "" And _
                   Trim(waResult(inRow, Price)) = "" And _
                   Trim(waResult(inRow, DisPer)) = "" And _
                   Trim(waResult(inRow, Amt)) = "" And _
                   Trim(waResult(inRow, Amtl)) = "" And _
                   Trim(waResult(inRow, Dis)) = "" And _
                   Trim(waResult(inRow, Disl)) = "" And _
                   Trim(waResult(inRow, Net)) = "" And _
                   Trim(waResult(inRow, Netl)) = "" And _
                   Trim(waResult(inRow, BOOKID)) = "" Then
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
        
        If Chk_grdBookCode(waResult(LastRow, BOOKCODE), "", "", "", "", "", 0, 0) = False Then
            .Col = BOOKCODE
            Exit Function
        End If
        
        If Chk_grdWhsCode(waResult(LastRow, WHSCODE)) = False Then
                .Col = WHSCODE
                Exit Function
        End If
        
        If Chk_grdWantedDate(waResult(LastRow, WANTED)) = False Then
                .Col = WANTED
                Exit Function
        End If
        
        If Chk_grdQty(waResult(LastRow, Qty)) = False Then
                .Col = Qty
                Exit Function
        End If
        
        If Chk_grdDisPer(waResult(LastRow, DisPer)) = False Then
                .Col = DisPer
                Exit Function
        End If
        
        If Chk_Amount(waResult(LastRow, Amt)) = False Then
            .Col = Amt
            Exit Function
        End If
        
    
        
    End With
        
    Chk_GrdRow = True

    Exit Function
    
Chk_GrdRow_Err:
    MsgBox "Check Chk_GrdRow"
    
End Function

Private Function Calc_Total(Optional ByVal LastRow As Variant) As Boolean
    
    Dim wiTotalGrs As Double
    Dim wiTotalDis As Double
    Dim wiTotalNet As Double
    Dim wiTotalQty As Integer
    
    Dim wiRowCtr As Integer
    
    
    Calc_Total = False
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotalGrs = wiTotalGrs + To_Value(waResult(wiRowCtr, Amt))
        wiTotalDis = wiTotalDis + To_Value(waResult(wiRowCtr, Dis))
        wiTotalNet = wiTotalNet + To_Value(waResult(wiRowCtr, Net))
        wiTotalQty = wiTotalQty + To_Value(waResult(wiRowCtr, Qty))
    Next
    
    lblDspGrsAmtOrg.Caption = Format(CStr(wiTotalGrs), gsAmtFmt)
    lblDspDisAmtOrg.Caption = Format(CStr(wiTotalDis), gsAmtFmt)
    lblDspNetAmtOrg.Caption = Format(CStr(wiTotalNet), gsAmtFmt)
    lblDspTotalQty.Caption = Format(CStr(wiTotalQty), gsQtyFmt)
    
    Calc_Total = True

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
                .Buttons(tcFind).Enabled = True
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
            Me.cboDocNo.Enabled = False
            Me.txtRevNo.Enabled = False
            Me.cboCusCode.Enabled = False
            Me.medDocDate.Enabled = False
            
            Me.cboSaleCode.Enabled = False
            Me.cboMethodCode.Enabled = False
            Me.picRmk.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
           Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
           Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            Me.txtRevNo.Enabled = True
            Me.cboCusCode.Enabled = True
            Me.medDocDate.Enabled = True
            Me.cboSaleCode.Enabled = True
            Me.cboMethodCode.Enabled = True
            Me.picRmk.Enabled = True
            Me.tblDetail.Enabled = True
            
       
            
    End Select
End Sub




Private Sub cboSaleCode_GotFocus()
    FocusMe cboSaleCode
End Sub

Private Sub cboSaleCode_LostFocus()
    FocusMe cboSaleCode, True
End Sub


Private Sub cboSaleCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboSaleCode = False Then
                Exit Sub
        End If
        
        cboMethodCode.SetFocus
        tabDetailInfo.Tab = 0
   
       
    End If
    
End Sub

Private Sub cboSaleCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSaleCode
    
    wsSql = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' "
    wsSql = wsSql & "AND SaleStatus = '1' "
    wsSql = wsSql & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSql, cboSaleCode.Left + tabDetailInfo.Left, cboSaleCode.Top + cboSaleCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSALECOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboSaleCode() As Boolean
Dim wsDesc As String

    Chk_cboSaleCode = False
     
    If Trim(cboSaleCode.Text) = "" Then
        gsMsg = "必需輸入營業員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSaleCode.SetFocus
        tabDetailInfo.Tab = 0
   
        Exit Function
    End If
    
    
    If Chk_Salesman(cboSaleCode, wlSaleID, wsDesc) = False Then
        gsMsg = "沒有此營業員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboSaleCode.SetFocus
        lblDspSaleDesc = ""
       Exit Function
    End If
    
    lblDspSaleDesc = wsDesc
    
    Chk_cboSaleCode = True
    
End Function








Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Chk_NoDup = False
    
        wsCurRec = tblDetail.Columns(BOOKCODE)
 '       wsCurRecLn = tblDetail.Columns(wsWhsCode)
 
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, BOOKCODE) Then
                  gsMsg = "重覆書本!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function

Private Function Chk_NoDup2(inItmCode As String, inWhsCode As String) As Boolean
' CHECK NEW ENTRY FRO DUPLICATES
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    
    Chk_NoDup2 = False
    
    If waResult.UpperBound(1) = -1 Then
       Chk_NoDup2 = True
       Exit Function
    End If
    
    If Trim(inItmCode) = "" Then Exit Function
    
   ' If optStlMtd(0).Value = True Then
   '     For wlCtr = 0 To waInvoice.UpperBound(1)
   '
   '         If inInvNo = waInvoice(wlCtr, Tab1InvNo) And _
   '            inInvLn = waInvoice(wlCtr, Tab1InvLn) Then
   '            Call Dsp_Err("E0014", "", "E", Me.Caption)
   '           Exit Function
   '        End If
   '    Next
   ' Else
        For wlCtr = 0 To waResult.UpperBound(1)
            If inItmCode = waResult(wlCtr, BOOKCODE) Then
                gsMsg = "重覆書本!"
                MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
               Exit Function
            End If
        Next
    
    'End If
    Chk_NoDup2 = True

End Function

Private Sub cmdPrint(InDocNo As String)
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
    wsSql = "EXEC usp_RPTQTN001 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & InDocNo & "', "
    wsSql = wsSql & "'" & InDocNo & "', "
    wsSql = wsSql & "'" & "" & "', "
    wsSql = wsSql & "'" & String(10, "z") & "', "
    wsSql = wsSql & "'" & String(6, "0") & "', "
    wsSql = wsSql & "'" & String(6, "9") & "', "
    wsSql = wsSql & "'" & "N" & "', "
    wsSql = wsSql & gsLangID
    
    
    If gsLangID = "2" Then wsRptName = "C" + "RPTQTN001"
    
    NewfrmPrint.ReportID = "QTN001"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "QTN001"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub LoadWSINFO()
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT * FROM sysWSINFO WHERE WSID ='" + gsWorkStationID + "'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
    
    
    wsWhsCode = ReadRs(rsRcd, "WSWHSCODE")
    wsMethodCode = ReadRs(rsRcd, "WSMETHODCODE")
    wsCurCode = ReadRs(rsRcd, "WSCURR")
    wdExcr = To_Value(ReadRs(rsRcd, "WSEXCR"))
    If gsLangID = "2" Then
    wgsTitle = ReadRs(rsRcd, "WSCTITLE")
    Else
    wgsTitle = ReadRs(rsRcd, "WSTITLE")
    End If
    
    
    Else
    
    wsWhsCode = ""
    wsMethodCode = ""
    wsCurCode = wsBaseCurCd
    wdExcr = 1
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
    
        wsSql = "SELECT QTHDDOCID, QTHDDOCNO, QTHDCUSID, CUSID, CUSCODE, CUSNAME, CUSCONTACTPERSON, CUSADDRESS1, CUSADDRESS2, CUSADDRESS3, CUSADDRESS4, CUSTEL, CUSFAX, "
        wsSql = wsSql & "QTHDDOCDATE, QTHDREVNO, QTHDCURR, QTHDEXCR, "
        wsSql = wsSql & "QTHDDUEDATE, QTHDONDATE, QTHDETADATE, QTHDPAYCODE, QTHDPRCCODE, QTHDSALEID, QTHDCUSTYP, QTHDMLCODE, QTHDMETHODCODE, "
        wsSql = wsSql & "QTHDCUSPO, QTHDLCNO, QTHDPORTNO, QTHDSHIPPER, QTHDSHIPFROM, QTHDSHIPTO, QTHDSHIPVIA, QTHDSHIPNAME, "
        wsSql = wsSql & "QTHDSHIPCODE, QTHDSHIPADR1,  QTHDSHIPADR2,  QTHDSHIPADR3,  QTHDSHIPADR4, "
        wsSql = wsSql & "QTHDRMKCODE, QTHDRMK1,  QTHDRMK2,  QTHDRMK3,  QTHDRMK4, QTHDRMK5, "
        wsSql = wsSql & "QTHDRMK6,  QTHDRMK7,  QTHDRMK8,  QTHDRMK9, QTHDRMK10, "
        wsSql = wsSql & "QTHDGRSAMT , QTHDGRSAMTL, QTHDDISAMT, QTHDDISAMTL, QTHDNETAMT, QTHDNETAMTL, "
        wsSql = wsSql & "QTDTJBCODE, QTDTJBTYPE, QTDTWHSCODE, QTDTJBDESC, QTDTWANTED, QTDTQTY, QTDTUPRICE, QTDTDISPER, QTDTAMT, QTDTAMTL, QTDTDIS, QTDTDISL, QTDTNET, QTDTNETL "
        wsSql = wsSql & "FROM  soaQTHD, soaQTDT, mstCUSTOMER "
        wsSql = wsSql & "WHERE QTHDDOCNO = '" & cboDocNo & "' "
        wsSql = wsSql & "AND QTHDDOCID = QTDTDOCID "
        wsSql = wsSql & "AND QTHDCUSID = CUSID "
        wsSql = wsSql & "ORDER BY QTDTDOCLINE "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsRcd, "QTHDDOCID")
    txtRevNo.Text = Format(ReadRs(rsRcd, "QTHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsRcd, "QTHDREVNO"))
    medDocDate.Text = ReadRs(rsRcd, "QTHDDOCDATE")
    wlCusID = ReadRs(rsRcd, "CUSID")
    cboCusCode.Text = ReadRs(rsRcd, "CUSCODE")
    lblDspCusName.Caption = ReadRs(rsRcd, "CUSNAME")
    lblDspCusTel.Caption = ReadRs(rsRcd, "CUSTEL")
    lblDspCusFax.Caption = ReadRs(rsRcd, "CUSFAX")
    lblDspCusContact = ReadRs(rsRcd, "CUSCONTACTPERSON")
    lblDspCusAddress = ReadRs(rsRcd, "CUSADDRESS1") & Chr(13) & Chr(10) & _
                           ReadRs(rsRcd, "CUSADDRESS2") & Chr(13) & Chr(10) & _
                           ReadRs(rsRcd, "CUSADDRESS3") & Chr(13) & Chr(10) & _
                           ReadRs(rsRcd, "CUSADDRESS4")
       
    wlSaleID = To_Value(ReadRs(rsRcd, "QTHDSALEID"))
    wlCusTyp = To_Value(ReadRs(rsRcd, "QTHDCUSTYP"))
    
    wsPayCode = ReadRs(rsRcd, "QTHDPAYCODE")
    
 
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, BOOKCODE, BOM
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsRcd, "QTDTJBCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "QTDTJBTYPE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "QTDTJBDESC")
             waResult(.UpperBound(1), WHSCODE) = ReadRs(rsRcd, "QTDTWHSCODE")
             waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsRcd, "")
             waResult(.UpperBound(1), WANTED) = Dsp_MedDate(ReadRs(rsRcd, "QTDTWANTED"))
             waResult(.UpperBound(1), Qty) = Format(ReadRs(rsRcd, "QTDTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), Price) = Format(ReadRs(rsRcd, "QTDTUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsRcd, "QTDTDISPER"), "0.0")
             waResult(.UpperBound(1), Amt) = Format(ReadRs(rsRcd, "QTDTAMT"), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsRcd, "QTDTAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(ReadRs(rsRcd, "QTDTDIS"), gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(ReadRs(rsRcd, "QTDTDISL"), gsAmtFmt)
             waResult(.UpperBound(1), Net) = Format(ReadRs(rsRcd, "QTDTNET"), gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(ReadRs(rsRcd, "QTDTNETL"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "QTDTITEMID")
             waResult(.UpperBound(1), BOM) = "Y"
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    Call Calc_Total
    
    LoadRecord = True
    
End Function
Private Function Chk_KeyExist() As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT QTHDSTATUS FROM soaQTHD WHERE QTHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
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
        .TableKey = "QTHDDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub txtRevNo_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtRevNo.Text, False, False)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtRevNo Then
            medDocDate.SetFocus
            tabDetailInfo.Tab = 0
   
        End If
    End If

End Sub

Private Function chk_txtRevNo() As Boolean
    
    chk_txtRevNo = False
    
    If Trim(txtRevNo) = "" Then
        gsMsg = "對換率超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtRevNo.SetFocus
        Exit Function
    End If
    
    If To_Value(txtRevNo) > wiRevNo + 1 Or _
        To_Value(txtRevNo) < wiRevNo Then
        gsMsg = "修改號錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtRevNo.SetFocus
        Exit Function
    End If
    
    chk_txtRevNo = True

End Function

Private Function LoadBookGroup(ByVal ISBN As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiCtr As Long
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    Dim wsItmID As String
    Dim wsISBN As String
    Dim wsName As String
    Dim wsBarCode As String
    Dim wsPub As String
    Dim wsPrice As Double
    Dim wsDisPer As Double
    Dim wdAmt As Double
    Dim wdAmtl As Double
    Dim wdDis As Double
    Dim wdDisl As Double
    Dim wdNet As Double
    Dim wdNetl As Double
    Dim wsSeries As String
    
    Dim wsMtd As String
    
    LoadBookGroup = False
    wsMtd = ""
    
    If Trim(ISBN) = "" Then
        Exit Function
    End If
    
        wsSql = "SELECT ITMSERIESNO "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ITMSERIESNO = '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
       
        rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

        If rsRcd.RecordCount > 0 Then
            wsMtd = "1"
            wsSeries = ISBN
        End If
        rsRcd.Close
        Set rsRcd = Nothing
        
        If wsMtd <> "1" Then
    
        wsSql = "SELECT ITMSERIESNO "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmCode = '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "AND ITMSERIESNO <> '" & Set_Quote(ISBN) & "' "
       
        rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

        If rsRcd.RecordCount <= 0 Then
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        End If
    
        If IsNull(ReadRs(rsRcd, "ITMSERIESNO")) Or Trim(ReadRs(rsRcd, "ITMSERIESNO")) = "" Then
            rsRcd.Close
            Set rsRcd = Nothing
            Exit Function
        Else
            wsSeries = ReadRs(rsRcd, "ITMSERIESNO")
        End If
    
         rsRcd.Close
         Set rsRcd = Nothing
       
         End If
         
    If gsLangID = "1" Then
        
        wsSql = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "ORDER BY ItmCode "
    Else
    
        wsSql = "SELECT ITMID, ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMPUBLISHER, ITMCURR, ITMDEFAULTPRICE "
        wsSql = wsSql & "FROM  mstITEM "
        wsSql = wsSql & "WHERE ItmSeriesNo = '" & Set_Quote(wsSeries) & "' "
        wsSql = wsSql & "AND ITMCODE <> '" & Set_Quote(ISBN) & "' "
        wsSql = wsSql & "ORDER BY ItmCode "
        
    End If
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    If wsMtd = "" Then
    gsMsg = "此書為套裝書之一, 你是否要於此單表選擇全套書?"
    Else
    gsMsg = "此書為套裝書, 你是否要於此單表選擇全套書?"
    End If
    
    If MsgBox(gsMsg, vbInformation + vbYesNo, gsTitle) = vbNo Then
        tblDetail.ReBind
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    If wsMtd = "1" Then
         With tblDetail
            .Delete
            .Update
            If .Row = -1 Then
                .Row = 0
            End If
         End With
    End If
    
    rsRcd.MoveFirst
    Do While Not rsRcd.EOF
  
       wsItmID = ReadRs(rsRcd, "ITMID")
       wsISBN = ReadRs(rsRcd, "ITMCODE")
       wsName = ReadRs(rsRcd, "ITNAME")
       wsBarCode = ReadRs(rsRcd, "ITMBARCODE")
       wsPub = ReadRs(rsRcd, "ITMPUBLISHER")
       wsPrice = To_Value(ReadRs(rsRcd, "ITMDEFAULTPRICE"))
       wsCurr = ReadRs(rsRcd, "ITMCURR")
       
       wdPrice = getCusItemPrice(wlCusID, wsItmID, wsCurCode)
       
       If wdPrice = 0 Then
       If wsCurCode <> wsCurr Then
       If getExcRate(wsCurr, medDocDate, wsExcr, "") = True Then
       wsPrice = NBRnd(wsPrice * To_Value(wsExcr) / wdExcr, giExrDp)
       End If
       End If
       Else
        wsPrice = wdPrice
       End If
       
       wsDisPer = Get_SaleDiscount(cboMethodCode, To_Value(wlCusID), To_Value(wsItmID))
       wdAmt = wsPrice
       wdAmtl = wsPrice * wdExcr
       wdDis = wsPrice * wsDisPer / 100
       wdDisl = wdDis * wdExcr
       wdNet = wdAmt - wdDis
       wdNetl = wdAmtl - wdDisl
       
       With waResult
             .AppendRows
             waResult(.UpperBound(1), BOOKCODE) = wsISBN
             waResult(.UpperBound(1), BARCODE) = wsBarCode
             waResult(.UpperBound(1), BOOKNAME) = wsName
             waResult(.UpperBound(1), WHSCODE) = wsWhsCode
             waResult(.UpperBound(1), PUBLISHER) = wsPub
             waResult(.UpperBound(1), WANTED) = medDocDate
             waResult(.UpperBound(1), Qty) = "1"
             waResult(.UpperBound(1), Price) = Format(wsPrice, gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(wsDisPer, "0.0")
             waResult(.UpperBound(1), Amt) = Format(wdAmt, gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(wdAmtl, gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(wdDis, gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(wdDisl, gsAmtFmt)
             waResult(.UpperBound(1), Net) = Format(wdNet, gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(wdNetl, gsAmtFmt)
             waResult(.UpperBound(1), BOOKID) = wsItmID
        End With
    
     rsRcd.MoveNext
     Loop

    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    Call Calc_Total
    
    LoadBookGroup = True
    
End Function



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

Private Sub txtRmk_GotFocus(Index As Integer)

        FocusMe txtRmk(Index)
            
End Sub

Private Sub txtRmk_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtRmk(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        If Index = 5 Then
        tabDetailInfo.Tab = 1
        tblDetail.SetFocus
        Else
        tabDetailInfo.Tab = 0
        txtRmk(Index + 1).SetFocus
        End If
        
    End If
End Sub

Private Sub txtRmk_LostFocus(Index As Integer)
        
        FocusMe txtRmk(Index), True

End Sub

Public Function Chk_DocNo(ByVal InDocNo As String, ByRef OutStatus As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT QTHDSTATUS FROM soaQTHD WHERE QTHDDOCNO = '" & Set_Quote(InDocNo) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
        OutStatus = ReadRs(rsRcd, "QTHDSTATUS")
        Chk_DocNo = True
    Else
        OutStatus = ""
        Chk_DocNo = False
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    

End Function


Private Sub LoadQTItem()

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    Dim wiRow As Long
    
    On Error GoTo LoadQTItem_Err
    
    wsSql = " SELECT QTTITMTYPECODE, ITMCODE, QTTITMDESC, QTTUPRICE, QTTQTY, QTTAMT, QTTJBCODE "
    wsSql = wsSql & " FROM soaQTItem, MstItem "
    wsSql = wsSql & " WHERE QttDocID = " & wlKey
    wsSql = wsSql & " AND ItmID = QttItmID "
    
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    wiRow = 0
    waItem.ReDim 0, rsRcd.RecordCount - 1, STYPE, SSTATUS
    
    If rsRcd.RecordCount = 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Sub
    End If
    
    Do Until rsRcd.EOF
        waItem(wiRow, STYPE) = ReadRs(rsRcd, "QTTITMTYPECODE")
        waItem(wiRow, SITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waItem(wiRow, SITMDESC) = ReadRs(rsRcd, "QTTITMDESC")
        waItem(wiRow, SUPRICE) = Format(ReadRs(rsRcd, "QTTUPRICE"), gsAmtFmt)
        waItem(wiRow, SQTY) = Format(ReadRs(rsRcd, "QTTQTY"), gsQtyFmt)
        waItem(wiRow, SAMT) = Format(ReadRs(rsRcd, "QTTAMT"), gsAmtFmt)
        waItem(wiRow, SJBCODE) = ReadRs(rsRcd, "QTTJBCODE")
        waItem(wiRow, SSTATUS) = "1"
        wiRow = wiRow + 1
        rsRcd.MoveNext
    Loop

    rsRcd.Close
    Set rsRcd = Nothing
    
    Exit Sub
        
LoadQTItem_Err:
    MsgBox "LoadQTItem Err!"
    rsRcd.Close
    Set rsRcd = Nothing
    
End Sub

