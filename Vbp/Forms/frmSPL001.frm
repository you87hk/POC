VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSPL001 
   Caption         =   "送貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmSPL001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11400
      OleObjectBlob   =   "frmSPL001.frx":030A
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   10800
      Top             =   120
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
            Picture         =   "frmSPL001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSPL001.frx":66AD
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
      Width           =   11880
      _ExtentX        =   20955
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
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   8055
      Left            =   0
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Header Information"
      TabPicture(0)   =   "frmSPL001.frx":69C9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraKey"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraInfo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboCusCode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboDocNo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraShip"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboShipCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboRefDocNo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmSPL001.frx":69E5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDspTotalQty"
      Tab(1).Control(1)=   "lblTotalQty"
      Tab(1).Control(2)=   "tblDetail"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Item Information"
      TabPicture(2)   =   "frmSPL001.frx":6A01
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboRmkCode"
      Tab(2).Control(1)=   "fraRmk"
      Tab(2).ControlCount=   2
      Begin VB.ComboBox cboRefDocNo 
         Height          =   300
         Left            =   1800
         TabIndex        =   2
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Left            =   1800
         TabIndex        =   10
         Top             =   4200
         Width           =   2010
      End
      Begin VB.Frame fraShip 
         Height          =   3135
         Left            =   120
         TabIndex        =   54
         Top             =   3840
         Width           =   11535
         Begin VB.TextBox txtShipPer 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   11
            Text            =   "01234567890123457890"
            Top             =   720
            Width           =   4305
         End
         Begin VB.TextBox txtShipName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   12
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1080
            Width           =   4305
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1455
            Left            =   1680
            ScaleHeight     =   1395
            ScaleWidth      =   9555
            TabIndex        =   55
            Top             =   1440
            Width           =   9615
            Begin VB.TextBox txtShipAdr1 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   13
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   9465
            End
            Begin VB.TextBox txtShipAdr2 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   14
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   9465
            End
            Begin VB.TextBox txtShipAdr3 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   15
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   720
               Width           =   9465
            End
            Begin VB.TextBox txtShipAdr4 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   16
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1080
               Width           =   9465
            End
         End
         Begin VB.Label lblShipAdr 
            Caption         =   "SHIPADR"
            Height          =   240
            Left            =   120
            TabIndex        =   59
            Top             =   1480
            Width           =   1500
         End
         Begin VB.Label lblShipPer 
            Caption         =   "SHIPPER"
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   760
            Width           =   1500
         End
         Begin VB.Label lblShipName 
            Caption         =   "SHIPNAME"
            Height          =   240
            Left            =   120
            TabIndex        =   57
            Top             =   1120
            Width           =   1380
         End
         Begin VB.Label lblShipCode 
            Caption         =   "SHIPCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   56
            Top             =   400
            Width           =   1500
         End
      End
      Begin VB.ComboBox cboRmkCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   17
         Top             =   480
         Width           =   1890
      End
      Begin VB.ComboBox cboDocNo 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   420
         Width           =   1935
      End
      Begin VB.ComboBox cboCusCode 
         Height          =   300
         Left            =   5280
         TabIndex        =   3
         Top             =   780
         Width           =   1935
      End
      Begin VB.Frame fraInfo 
         Height          =   1815
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   11535
         Begin VB.TextBox txtCusPo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   6
            Text            =   "0123456789012345789"
            Top             =   240
            Width           =   5265
         End
         Begin VB.TextBox txtShipTo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   8
            Text            =   "0123456789012345789"
            Top             =   960
            Width           =   5265
         End
         Begin VB.TextBox txtShipVia 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   9
            Text            =   "0123456789012345789"
            Top             =   1320
            Width           =   5265
         End
         Begin VB.TextBox txtShipFrom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   7
            Text            =   "0123456789012345789"
            Top             =   600
            Width           =   5265
         End
         Begin VB.Label lblCusPo 
            Caption         =   "CUSPO"
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   280
            Width           =   2100
         End
         Begin VB.Label lblShipTo 
            Caption         =   "SHIPTO"
            Height          =   240
            Left            =   120
            TabIndex        =   35
            Top             =   1000
            Width           =   2100
         End
         Begin VB.Label lblShipVia 
            Caption         =   "SHIPVIA"
            Height          =   240
            Left            =   120
            TabIndex        =   34
            Top             =   1360
            Width           =   2100
         End
         Begin VB.Label lblShipFrom 
            Caption         =   "SHIPFROM"
            Height          =   240
            Left            =   120
            TabIndex        =   33
            Top             =   640
            Width           =   2100
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   6855
         Left            =   -74880
         OleObjectBlob   =   "frmSPL001.frx":6A1D
         TabIndex        =   28
         Top             =   840
         Width           =   11535
      End
      Begin VB.Frame fraRmk 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   39
         Top             =   240
         Width           =   11535
         Begin VB.PictureBox picRmk 
            BackColor       =   &H80000009&
            Height          =   3495
            Left            =   1680
            ScaleHeight     =   3435
            ScaleWidth      =   9555
            TabIndex        =   40
            Top             =   600
            Width           =   9615
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   19
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   18
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   20
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   690
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   23
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1740
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   21
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1035
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   22
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1395
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   24
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2085
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   8
               Left            =   0
               TabIndex        =   25
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2430
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   9
               Left            =   0
               TabIndex        =   26
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2775
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   10
               Left            =   0
               TabIndex        =   27
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   3120
               Width           =   7545
            End
         End
         Begin VB.Label lblRmkCode 
            Caption         =   "RMKCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   280
            Width           =   1500
         End
         Begin VB.Label lblRmk 
            Caption         =   "RMK"
            Height          =   240
            Left            =   120
            TabIndex        =   41
            Top             =   650
            Width           =   1500
         End
      End
      Begin VB.Frame fraKey 
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   120
         Width           =   11535
         Begin VB.TextBox txtRevNo 
            Height          =   324
            Left            =   5160
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "12345678901234567890"
            Top             =   300
            Width           =   408
         End
         Begin MSMask.MaskEdBox medDocDate 
            Height          =   285
            Left            =   9360
            TabIndex        =   4
            Top             =   1020
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medETADate 
            Height          =   285
            Left            =   9360
            TabIndex        =   5
            Top             =   1380
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblRefDocNo 
            Caption         =   "CUSCODE"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblETADate 
            Caption         =   "DOCDATE"
            Height          =   255
            Left            =   7365
            TabIndex        =   60
            Top             =   1440
            Width           =   1680
         End
         Begin VB.Label lblCusCode 
            Caption         =   "CUSCODE"
            Height          =   255
            Left            =   3840
            TabIndex        =   53
            Top             =   705
            Width           =   1575
         End
         Begin VB.Label lblDocNo 
            Caption         =   "DOCNO"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblRevNo 
            Caption         =   "REVNO"
            Height          =   255
            Left            =   3840
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblDocDate 
            Caption         =   "DOCDATE"
            Height          =   255
            Left            =   7365
            TabIndex        =   50
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lblDspCusName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   49
            Top             =   1020
            Width           =   5535
         End
         Begin VB.Label lblDspCusTel 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   48
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lblCusName 
            Caption         =   "CUSNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblDspCusFax 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   5160
            TabIndex        =   46
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lblCusFax 
            Caption         =   "CUSFAX"
            Height          =   255
            Left            =   3840
            TabIndex        =   45
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCusTel 
            Caption         =   "CUSTEL"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1440
            Width           =   1575
         End
      End
      Begin VB.Label lblTotalQty 
         Alignment       =   2  '置中對齊
         Caption         =   "TOTALBOOK"
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
         Left            =   -74760
         TabIndex        =   38
         Top             =   60
         Width           =   1755
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
         Left            =   -74880
         TabIndex        =   37
         Top             =   420
         Width           =   1890
      End
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
Attribute VB_Name = "frmSPL001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private waResult As New XArrayDB
Private waScrItm As New XArrayDB
Private waScrToolTip As New XArrayDB
Private waPopUpSub As New XArrayDB
Private wcCombo As Control

Private wsOldCusNo As String
Private wsOldShipCd As String
Private wsOldRmkCd As String
Private wsOldPayCd As String
Private wbReadOnly As Boolean
Private wgsTitle As String
Private wsOldRefDocNo As String

Private Const LINENO = 0
Private Const SONO = 1
Private Const ITMCODE = 2
Private Const BARCODE = 3
Private Const WHSCODE = 4
Private Const LOTNO = 5
Private Const ITMNAME = 6
Private Const Qty = 7
Private Const Price = 8
Private Const ITMID = 9
Private Const SOID = 10

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
Private wiRevNo As Integer
Private wlCusID As Long
Private wlSaleID As Long
Private wlRefDocID As Long
Private wlLineNo As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "soaSPHd"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()
    Dim MyControl As Control
    
    waResult.ReDim 0, -1, LINENO, SOID
    Set tblDetail.Array = waResult
    tblDetail.ReBind
    tblDetail.Bookmark = 0
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

    Call SetButtonStatus("Default")
    Call SetFieldStatus("Default")
    Call SetFieldStatus("AfrActEdit")
    
    Call SetDateMask(medDocDate)
    Call SetDateMask(medETADate)
    
    medDocDate = gsSystemDate
    medETADate = gsSystemDate
    
    wsOldRefDocNo = ""
    wsOldCusNo = ""
    wsOldShipCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""
    
    wlKey = 0
    wlCusID = 0
    wlSaleID = 0
    wlRefDocID = 0
    wlLineNo = 1
    
    wiRevNo = Format(0, "##0")
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    
    FocusMe cboDocNo
    tabDetailInfo.Tab = 0
End Sub



Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
End Sub

Private Sub cboDocNo_GotFocus()
    FocusMe cboDocNo
End Sub

Private Sub cboDocNo_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSQL = "SELECT SPHDDOCNO, CUSCODE, SPHDDOCDATE "
    wsSQL = wsSQL & " FROM soaSPHD, MstCustomer "
    wsSQL = wsSQL & " WHERE SPHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND SPHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SPHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " ORDER BY SPHDDOCNO DESC "
    Call Ini_Combo(3, wsSQL, cboDocNo.Left + tabDetailInfo.Left, cboDocNo.Top + cboDocNo.Height + tabDetailInfo.Top, tblCommon, "SPL001", "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
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
        tabDetailInfo.Tab = 0
        cboDocNo.SetFocus
        Exit Function
    End If
    
    If Chk_TrnHdDocNo(wsTrnCd, cboDocNo, wsStatus) = True Then
        
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
    
    Chk_cboDocNo = True
End Function

Private Sub Ini_Scr_AfrKey()
    If LoadRecord() = False Then
        wiAction = AddRec
        txtRevNo.Text = Format(0, "##0")
        txtRevNo.Enabled = False
        medDocDate.Text = Dsp_Date(Now)
        medETADate.Text = Dsp_Date(Now)
    
        Call SetButtonStatus("AfrKeyAdd")
        Call SetFieldStatus("AfrKey")
        cboRefDocNo.SetFocus
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        txtRevNo.Enabled = True
        wsOldCusNo = cboCusCode.Text
        wsOldShipCd = cboShipCode.Text
        wsOldRmkCd = cboRmkCode.Text
        
        Call SetButtonStatus("AfrKeyEdit")
        Call SetFieldStatus("AfrKey")
        cboCusCode.SetFocus
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    

    
    tabDetailInfo.Tab = 0
    
End Sub



Private Sub cboRefDocNo_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboRefDocNo
    
    wsSQL = "SELECT SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
    wsSQL = wsSQL & " WHERE SOHDSTATUS = '1' "
    wsSQL = wsSQL & " AND SOHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
                
    Call Ini_Combo(2, wsSQL, cboRefDocNo.Left + tabDetailInfo.Left, cboRefDocNo.Top + cboRefDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSONO", Me.Width, Me.Height)
           
            
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
   
End Sub

Private Sub cboRefDocNo_GotFocus()
    
    Set wcCombo = cboRefDocNo
    FocusMe cboRefDocNo
    
End Sub
Private Sub cboRefDocNo_LostFocus()
    FocusMe cboRefDocNo, True
End Sub

Private Sub cboRefDocNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboRefDocNo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_cboRefDocNo() = False Then Exit Sub
        
        If wiAction = AddRec And wsOldRefDocNo <> cboRefDocNo.Text Then Call Get_RefDoc
            tabDetailInfo.Tab = 0
            cboCusCode.SetFocus
            
    End If
    
End Sub
Private Function Chk_cboRefDocNo() As Boolean
    
Dim wsStatus As String
    
    Chk_cboRefDocNo = False
    
    If Trim(cboRefDocNo.Text) = "" Then
        Chk_cboRefDocNo = True
        wlRefDocID = 0
        Exit Function
    End If
    
        
   If Chk_TrnHdDocNo("SO", cboRefDocNo, wsStatus) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
    
        If wsStatus = "3" Then
            gsMsg = "文件已無效!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Function
        End If
        
    End If
    
    
    Chk_cboRefDocNo = True

End Function

Private Sub Form_Activate()
    If OpenDoc = True Then
        OpenDoc = False
        Set wcCombo = cboDocNo
        Call cboDocNo_DropDown
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
            
        
        Case vbKeyF7
            If tbrProcess.Buttons(tcRefresh).Enabled = True Then Call cmdRefresh
       
        
        Case vbKeyF3
            If wiAction = DefaultPage Then Call cmdDel
        
        'Case vbKeyF9
        '    If tbrProcess.Buttons(tcFind).Enabled = True Then Call cmdFind
            
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

Private Function LoadRecord() As Boolean
    Dim rsPL As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
        wsSQL = "SELECT SPHDDOCID, SPHDDOCNO, SPHDREFDOCID, SPHDCUSID, CUSID, CUSCODE, CUSNAME, CUSTEL, CUSFAX, "
        wsSQL = wsSQL & "SPHDDOCDATE, SPHDREVNO, SPHDETADATE, SPDTDOCLINE, "
        wsSQL = wsSQL & "SPHDREFNO, SPHDSHIPPER, SPHDSHIPFROM, SPHDSHIPTO, SPHDSHIPVIA, SPHDSHIPNAME, "
        wsSQL = wsSQL & "SPHDSHIPCODE, SPHDSHIPADR1, SPHDSHIPADR2, SPHDSHIPADR3, SPHDSHIPADR4, "
        wsSQL = wsSQL & "SPHDRMKCODE, SPHDRMK1, SPHDRMK2, SPHDRMK3, SPHDRMK4, SPHDRMK5, "
        wsSQL = wsSQL & "SPHDRMK6, SPHDRMK7, SPHDRMK8, SPHDRMK9, SPHDRMK10, "
        wsSQL = wsSQL & "SPDTITEMID, ITMCODE, SPDTWHSCODE, SPDTLOTNO, ITMBARCODE, SPDTITEMDESC  ITNAME, SPDTQTY, "
        wsSQL = wsSQL & "SOHDDOCNO , SPDTSOID, SPDTUPRICE "
        wsSQL = wsSQL & "FROM  soaSpHd, soaSPDT, MstCustomer, MstItem, soaSoHd "
        wsSQL = wsSQL & "WHERE SPHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND SPHDDOCID = SPDTDOCID "
        wsSQL = wsSQL & "AND SPDTSOID = SOHDDOCID "
        wsSQL = wsSQL & "AND SPHDCUSID = CUSID "
        wsSQL = wsSQL & "AND SPDTITEMID = ITMID "
        wsSQL = wsSQL & "ORDER BY SPDTDOCLINE "
    
    rsPL.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsPL.RecordCount <= 0 Then
        rsPL.Close
        Set rsPL = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsPL, "SPHDDOCID")
    wlRefDocID = ReadRs(rsPL, "SPHDREFDOCID")
    cboRefDocNo.Text = Get_TableInfo("soaSOHD", "SOHDDOCID =" & wlRefDocID, "SOHDDOCNO")
    
    txtRevNo.Text = Format(ReadRs(rsPL, "SPHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsPL, "SPHDREVNO"))
    medDocDate.Text = ReadRs(rsPL, "SPHDDOCDATE")
    medETADate.Text = Dsp_MedDate(ReadRs(rsPL, "SPHDETADATE"))
    wlCusID = ReadRs(rsPL, "CUSID")
    cboCusCode.Text = ReadRs(rsPL, "CUSCODE")
    lblDspCusName.Caption = ReadRs(rsPL, "CUSNAME")
    lblDspCusTel.Caption = ReadRs(rsPL, "CUSTEL")
    lblDspCusFax.Caption = ReadRs(rsPL, "CUSFAX")
    
    wlSaleID = To_Value(ReadRs(rsPL, "SPHDSALEID"))
    
    cboShipCode = ReadRs(rsPL, "SPHDSHIPCODE")
    cboRmkCode = ReadRs(rsPL, "SPHDRMKCODE")
    
    txtCusPo = ReadRs(rsPL, "SPHDREFNO")
    
    txtShipFrom = ReadRs(rsPL, "SPHDSHIPFROM")
    txtShipTo = ReadRs(rsPL, "SPHDSHIPTO")
    txtShipVia = ReadRs(rsPL, "SPHDSHIPVIA")
    txtShipName = ReadRs(rsPL, "SPHDSHIPNAME")
    txtShipPer = ReadRs(rsPL, "SPHDSHIPPER")
    txtShipAdr1 = ReadRs(rsPL, "SPHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsPL, "SPHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsPL, "SPHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsPL, "SPHDSHIPADR4")
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsPL, "SPHDRMK" & i)
    Next i
    
    'lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    
    rsPL.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, SOID
         Do While Not rsPL.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsPL, "SPDTDOCLINE")
             waResult(.UpperBound(1), SONO) = ReadRs(rsPL, "SOHDDOCNO")
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsPL, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsPL, "ITMBARCODE")
             waResult(.UpperBound(1), ITMNAME) = ReadRs(rsPL, "ITNAME")
             waResult(.UpperBound(1), WHSCODE) = ReadRs(rsPL, "SPDTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsPL, "SPDTLOTNO")
             waResult(.UpperBound(1), Qty) = Format(ReadRs(rsPL, "SPDTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), Price) = Format(ReadRs(rsPL, "SPDTUPRICE"), gsUprFmt)
             waResult(.UpperBound(1), ITMID) = ReadRs(rsPL, "SPDTITEMID")
             waResult(.UpperBound(1), SOID) = ReadRs(rsPL, "SPDTSOID")
             rsPL.MoveNext
         Loop
         wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsPL.Close
    
    Set rsPL = Nothing
    
    Call Calc_Total
    
    LoadRecord = True
    
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblRefDocNo.Caption = Get_Caption(waScrItm, "REFNO")
    lblRevNo.Caption = Get_Caption(waScrItm, "REVNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblETADate.Caption = Get_Caption(waScrItm, "ETADATE")
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    lblCusTel.Caption = Get_Caption(waScrItm, "CUSTEL")
    lblCusFax.Caption = Get_Caption(waScrItm, "CUSFAX")
    
    lblTotalQty.Caption = Get_Caption(waScrItm, "TOTALQTY")
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(SONO).Caption = Get_Caption(waScrItm, "SONO")
        .Columns(ITMCODE).Caption = Get_Caption(waScrItm, "ITMCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(WHSCODE).Caption = Get_Caption(waScrItm, "WHSCODE")
        .Columns(LOTNO).Caption = Get_Caption(waScrItm, "LOTNO")
        .Columns(ITMNAME).Caption = Get_Caption(waScrItm, "ITMNAME")
        .Columns(Qty).Caption = Get_Caption(waScrItm, "QTY")
        .Columns(Price).Caption = Get_Caption(waScrItm, "PRICE")
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO03")
    
    lblShipFrom.Caption = Get_Caption(waScrItm, "SHIPFROM")
    lblShipTo.Caption = Get_Caption(waScrItm, "SHIPTO")
    lblShipVia.Caption = Get_Caption(waScrItm, "SHIPVIA")
    lblShipCode.Caption = Get_Caption(waScrItm, "SHIPCODE")
    lblShipPer.Caption = Get_Caption(waScrItm, "SHIPPER")
    lblShipAdr.Caption = Get_Caption(waScrItm, "SHIPADR")
    lblCusPo.Caption = Get_Caption(waScrItm, "CUSPO")
    lblShipName.Caption = Get_Caption(waScrItm, "SHIPNAME")
    lblRmkCode.Caption = Get_Caption(waScrItm, "RMKCODE")
    lblRmk.Caption = Get_Caption(waScrItm, "RMK")
    
  '  btnSOLST.Caption = Get_Caption(waScrItm, "SOLIST")
    
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
    
    wsActNam(1) = Get_Caption(waScrItm, "SPADD")
    wsActNam(2) = Get_Caption(waScrItm, "SPEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SPDELETE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    wgsTitle = Get_Caption(waScrItm, "TITLE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
'    If Button = 2 Then
'        PopupMenu mnuMaster
'    End If

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 9000
        Me.Width = 12000
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    Call UnLockAll(wsConnTime, wsFormID)
    Set waResult = Nothing
    Set waScrToolTip = Nothing
    Set waScrItm = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmSPL001 = Nothing

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
            tabDetailInfo.Tab = 0
            medETADate.SetFocus
        End If
    End If
End Sub

Private Function Chk_medDocDate() As Boolean
    Chk_medDocDate = False
    
    If Trim(medDocDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medDocDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medDocDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medDocDate.SetFocus
        Exit Function
    End If
    
    Chk_medDocDate = True
End Function

Private Function Chk_medETADate() As Boolean
    Chk_medETADate = False
    
    If Trim(medETADate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medETADate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medETADate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medETADate.SetFocus
        Exit Function
    End If
    
    Chk_medETADate = True
End Function

Private Sub medETADate_GotFocus()
    FocusMe medETADate
End Sub

Private Sub medETADate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medETADate Then
            tabDetailInfo.Tab = 0
            txtCusPo.SetFocus
        End If
    End If
End Sub

Private Sub medETADate_LostFocus()
    FocusMe medETADate, True
End Sub





Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
        If cboCusCode.Enabled Then
            cboCusCode.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
        
        If tblDetail.Enabled Then
            tblDetail.Col = SONO
            tblDetail.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 2 Then
    
        If cboRmkCode.Enabled Then
            cboRmkCode.SetFocus
        End If
    
    End If
End Sub

Private Sub tblCommon_DblClick()
    
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        Select Case wcCombo.Col
          Case ITMCODE
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
              Case ITMCODE
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

Private Function Chk_KeyExist() As Boolean
    
    Dim rsSOASDHD As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT SPHDSTATUS FROM soaSPHD WHERE SPHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rsSOASDHD.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsSOASDHD.RecordCount > 0 Then
        Chk_KeyExist = True
    Else
        Chk_KeyExist = False
    End If
    
    rsSOASDHD.Close
    Set rsSOASDHD = Nothing
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
        If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Or wbReadOnly Then
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
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SPL001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlCusID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, txtRevNo.Text)
    Call SetSPPara(adcmdSave, 8, medETADate.Text)
    
    Call SetSPPara(adcmdSave, 9, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 10, cboShipCode.Text)
    Call SetSPPara(adcmdSave, 11, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 12, txtCusPo.Text)
    
    Call SetSPPara(adcmdSave, 13, txtShipFrom.Text)
    Call SetSPPara(adcmdSave, 14, txtShipTo.Text)
    Call SetSPPara(adcmdSave, 15, txtShipVia.Text)
    Call SetSPPara(adcmdSave, 16, txtShipPer.Text)
    Call SetSPPara(adcmdSave, 17, txtShipName.Text)
    Call SetSPPara(adcmdSave, 18, txtShipAdr1.Text)
    Call SetSPPara(adcmdSave, 19, txtShipAdr2.Text)
    Call SetSPPara(adcmdSave, 20, txtShipAdr3.Text)
    Call SetSPPara(adcmdSave, 21, txtShipAdr4.Text)
    
    For i = 1 To 10
        Call SetSPPara(adcmdSave, 22 + i - 1, txtRmk(i).Text)
    Next
    Call SetSPPara(adcmdSave, 32, wlRefDocID)
    Call SetSPPara(adcmdSave, 33, wsFormID)
    
    Call SetSPPara(adcmdSave, 34, gsUserID)
    Call SetSPPara(adcmdSave, 35, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 36)
    wsDocNo = GetSPPara(adcmdSave, 37)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_SPL001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SONO)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SONO))
                Call SetSPPara(adcmdSave, 4, To_Value(waResult(wiCtr, SOID)))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, ITMID))
                Call SetSPPara(adcmdSave, 6, wiCtr + 1)
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, ITMNAME))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, Qty))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, Price))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, WHSCODE))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, LOTNO))
                Call SetSPPara(adcmdSave, 12, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 13, gsUserID)
                Call SetSPPara(adcmdSave, 14, wsGenDte)
                adcmdSave.Execute
                
               
                
            End If
        Next
    End If
    cnCon.CommitTrans
    
    
  
    
    If wiAction = AddRec Then
    If Trim(wsDocNo) <> "" Then
        gsMsg = "文件號 : " & wsDocNo & " 已製作!"
        MsgBox gsMsg, vbOKOnly, gsTitle
    Else
        gsMsg = "文件儲存件敗!"
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

Private Function InputValidation() As Boolean
    
    Dim wsExcRate As String
    Dim wsExcDesc As String

    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not Chk_cboRefDocNo Then Exit Function
    If Not Chk_medETADate Then Exit Function
    If Not chk_cboCusCode() Then Exit Function
    
    If Not Chk_cboShipCode Then Exit Function
    If Not Chk_cboRmkCode Then Exit Function
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, SONO)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tabDetailInfo.Tab = 1
                    tblDetail.Col = SONO
                    tblDetail.SetFocus
                    Exit Function
                End If
                
                If Chk_NoDup2(wlCtr, waResult(wlCtr, SONO), waResult(wlCtr, ITMCODE), waResult(wlCtr, WHSCODE), waResult(wlCtr, LOTNO)) = False Then
                    tblDetail.Row = wlCtr - 1
                    tblDetail.Col = SONO
                    tblDetail.SetFocus
                    tabDetailInfo.Tab = 1
                    Exit Function
                End If
                
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "配貨單沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tabDetailInfo.Tab = 1
            tblDetail.Col = SONO
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

Private Sub cmdNew()
    Dim newForm As New frmSPL001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()

    Dim newForm As New frmSPL001
    
    newForm.OpenDoc = True
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    newForm.Show

End Sub

Private Sub Ini_Form()
    Me.KeyPreview = True
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "SPL001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "SP"
    
End Sub

Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("Default")
    tabDetailInfo.Tab = 0
    cboDocNo.SetFocus
    
End Sub

Private Sub cmdFind()
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
      '  If .Bookmark <> .DestinationRow Then
            If Chk_GrdRow(To_Value(.Bookmark)) = False Then
                Cancel = True
                Exit Sub
            End If
      '  End If
    End With
    
    Exit Sub
    
tblDetail_BeforeRowColChange_Err:
    
    MsgBox "Check tblDeiail BeforeRowColChange!"
    Cancel = True

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

Private Sub txtRevNo_GotFocus()
    FocusMe txtRevNo
End Sub

Private Sub txtRevNo_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtRevNo.Text, False, False)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtRevNo Then
            tabDetailInfo.Tab = 0
            medDocDate.SetFocus
        End If
    End If
End Sub

Private Function chk_txtRevNo() As Boolean
    
    chk_txtRevNo = False
    
    If Trim(txtRevNo) = "" Then
        gsMsg = "修改號超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtRevNo.SetFocus
        Exit Function
    End If
    
    If To_Value(txtRevNo) > wiRevNo + 1 Or _
        To_Value(txtRevNo) < wiRevNo Then
        gsMsg = "修改號錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtRevNo.SetFocus
        Exit Function
    End If
    
    chk_txtRevNo = True

End Function

Private Sub cboCusCode_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusCode
    
    If gsLangID = "1" Then
        wsSQL = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSQL = wsSQL & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
        wsSQL = wsSQL & "AND CUSSTATUS = '1' "
        wsSQL = wsSQL & "ORDER BY CUSCODE "
    Else
        wsSQL = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
        wsSQL = wsSQL & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
        wsSQL = wsSQL & "AND CUSSTATUS = '1' "
        wsSQL = wsSQL & "ORDER BY CUSCODE "
    End If
    Call Ini_Combo(2, wsSQL, cboCusCode.Left + tabDetailInfo.Left, cboCusCode.Top + cboCusCode.Height + tabDetailInfo.Top, tblCommon, "SDN001", "TBLCUSNO", Me.Width, Me.Height)
    
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
            tabDetailInfo.Tab = 0
            medDocDate.SetFocus
            
    End If
    
End Sub

Private Function chk_cboCusCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    Dim wsTel As String
    Dim wsFax As String
    
    chk_cboCusCode = False
    
    If Trim(cboCusCode) = "" Then
        gsMsg = "必需輸入客戶編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
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
        tabDetailInfo.Tab = 0
        cboCusCode.SetFocus
        Exit Function
    End If
    
    chk_cboCusCode = True
End Function

Private Sub Get_DefVal()
    
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcDesc As String
    Dim wsExcRate As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM  mstCUSTOMER "
    wsSQL = wsSQL & "WHERE CUSID = " & wlCusID
    rsDefVal.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        txtShipName = ReadRs(rsDefVal, "CUSSHIPTO")
        txtShipPer = ReadRs(rsDefVal, "CUSSHIPCONTACTPERSON")
        txtShipAdr1 = ReadRs(rsDefVal, "CUSSHIPADD1")
        txtShipAdr2 = ReadRs(rsDefVal, "CUSSHIPADD2")
        txtShipAdr3 = ReadRs(rsDefVal, "CUSSHIPADD3")
        txtShipAdr4 = ReadRs(rsDefVal, "CUSSHIPADD4")
        
    Else
        txtShipName = ""
        txtShipPer = ""
        txtShipAdr1 = ""
        txtShipAdr2 = ""
        txtShipAdr3 = ""
        txtShipAdr4 = ""
        
    End If
    
    rsDefVal.Close
    Set rsDefVal = Nothing
    
    'get Due Date Payment Term
    'medDueDate = Dsp_Date(Get_DueDte(cboPayCode, medDocDate))
End Sub


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
        
        For wiCtr = LINENO To SOID
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case LINENO
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Locked = True
                Case SONO
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 15
                Case ITMCODE
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 30
                Case BARCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case WHSCODE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Visible = False
                Case LOTNO
                    .Columns(wiCtr).Width = 1000
                    '.Columns(wiCtr).Button = False
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Visible = False
                Case ITMNAME
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                                        
                Case Qty
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case Price
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsUprFmt
                Case ITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case SOID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
       ' .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
   
    With tblDetail
        .Update
    End With

End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim wsITMID As String
Dim wsITMCODE As String
Dim wsBarCode As String
Dim wsITMNAME As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double
Dim wsLotNo As String
Dim wsWhsCode As String
Dim wdQty As Double
Dim wsSoId As String
    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
                Case SONO
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdSoNo(.Columns(ColIndex).Text, wsSoId) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(SOID).Text = wsSoId
                
            Case ITMCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdITMCODE(.Columns(SONO).Text, .Columns(ColIndex).Text, wsITMID, wsITMCODE, wsBarCode, wsITMNAME, wsPub, wdPrice, wdDisPer, wsWhsCode, wsLotNo, wdQty) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(LINENO).Text = wlLineNo
                .Columns(ITMID).Text = wsITMID
                .Columns(BARCODE).Text = wsBarCode
                .Columns(ITMNAME).Text = wsITMNAME
                .Columns(WHSCODE).Text = wsWhsCode
                .Columns(LOTNO).Text = wsLotNo
                .Columns(Price).Text = Format(wdPrice, gsUprFmt)
                .Columns(Qty).Text = Format(wdQty, gsQtyFmt)
                wlLineNo = wlLineNo + 1
                
                '.Columns(DisPer).Text = Format(wdDisPer, "0")
                If Trim(.Columns(ColIndex).Text) <> wsITMCODE Then
                    .Columns(ColIndex).Text = wsITMCODE
                End If
                
             Case WHSCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdWhsCode(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
             Case LOTNO
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdLotNo(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
            

            Case Qty, Price
            
                If ColIndex = Qty Then
                    If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                    End If
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
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case SONO
                
                wsSQL = "SELECT SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
                wsSQL = wsSQL & " WHERE SOHDSTATUS = '1' "
                wsSQL = wsSQL & " AND SOHDDOCNO LIKE '%" & Set_Quote(.Columns(SONO).Text) & "%' "
                wsSQL = wsSQL & " AND SOHDCUSID = " & wlCusID
                wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSONO", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case ITMCODE
                
                If gsLangID = 1 Then
                    wsSQL = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM FROM mstITEM, soaSohd, soaSodt "
                    wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                    wsSQL = wsSQL & " AND SOHDDOCNO = '" & Set_Quote(.Columns(SONO).Text) & "' "
                    wsSQL = wsSQL & " AND SOHDDOCID = SODTDOCID "
                    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
                    wsSQL = wsSQL & " ORDER BY ITMCODE "
                Else
                    wsSQL = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM FROM mstITEM, soaSohd, soaSodt "
                    wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                    wsSQL = wsSQL & " AND SOHDDOCNO = '" & Set_Quote(.Columns(SONO).Text) & "' "
                    wsSQL = wsSQL & " AND SOHDDOCID = SODTDOCID "
                    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
                    wsSQL = wsSQL & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(4, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLITMCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case WHSCODE
                
                wsSQL = "SELECT WHSCODE, WHSDESC FROM mstWareHouse "
                wsSQL = wsSQL & " WHERE WHSSTATUS <> '2' AND WHSCODE LIKE '%" & Set_Quote(.Columns(WHSCODE).Text) & "%' "
                wsSQL = wsSQL & " ORDER BY WHSCODE "
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
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
            If .Bookmark = waResult.UpperBound(1) Then Exit Sub
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
                Case Price
                    KeyCode = vbKeyDown
                    .Col = SONO
                Case ITMCODE
                    KeyCode = vbDefault
                       .Col = Qty
                Case LINENO, SONO, Qty
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case BARCODE
                    KeyCode = vbDefault
                       .Col = ITMNAME
                Case ITMNAME
                    KeyCode = vbDefault
                       .Col = Qty
               
            End Select
        Case vbKeyLeft
               KeyCode = vbDefault
            Select Case .Col
                Case Qty
                    .Col = ITMNAME
                Case ITMNAME
                    .Col = BARCODE
                Case Price, BARCODE, ITMCODE
                    .Col = .Col - 1
                
            End Select
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case ITMCODE
                       .Col = Qty
                Case LINENO, Qty, SONO
                   .Col = .Col + 1
                Case BARCODE
                       .Col = ITMNAME
                Case ITMNAME
                       .Col = Qty
               
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
        
        Case Price
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
    End Select
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = SONO
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case SONO
                    Call Chk_grdSoNo(.Columns(SONO).Text, "")
                Case ITMCODE
                    Call Chk_grdITMCODE(.Columns(SONO).Text, .Columns(ITMCODE).Text, "", "", "", "", "", 0, 0, "", "", 0)
                Case WHSCODE
                    Call Chk_grdWhsCode(.Columns(WHSCODE).Text)
                 Case LOTNO
                    Call Chk_grdLotNo(.Columns(LOTNO).Text)
                Case Qty
                    Call Chk_grdQty(.Columns(Qty).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub

Private Function Chk_grdITMCODE(inSoNo As String, inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String, outPub As String, outPrice As Double, outDisPer As Double, outWhsCode As String, outLotNo As String, outQty As Double) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入物料號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
        Exit Function
    End If
    

    
        wsSQL = "SELECT SODTDOCID, SODTITEMID, ITMCODE, SODTITEMDESC ITNAME, ITMBARCODE, SODTUPRICE, SOHDCURR, SODTWHSCODE, SODTLOTNO, SODTQTY - SODTBALQTY SCHQTY, SODTDISPER "
        wsSQL = wsSQL & " FROM mstITEM, soaSoHd, soaSoDt "
        wsSQL = wsSQL & " WHERE SOHDDOCID = SODTDOCID "
        wsSQL = wsSQL & " AND SODTITEMID = ITMID "
        wsSQL = wsSQL & " AND SOHDDOCNO = '" & Set_Quote(inSoNo) & "' "
        wsSQL = wsSQL & " AND (ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "') "
   
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       outAccID = ReadRs(rsDes, "SODTITEMID")
       outAccNo = ReadRs(rsDes, "ITMCODE")
       OutName = ReadRs(rsDes, "ITNAME")
       OutBarCode = ReadRs(rsDes, "ITMBARCODE")
       outPrice = To_Value(ReadRs(rsDes, "SODTUPRICE"))
       wsCurr = ReadRs(rsDes, "SOHDCURR")
       outWhsCode = ReadRs(rsDes, "SODTWHSCODE")
       outLotNo = ReadRs(rsDes, "SODTLOTNO")
       outQty = ReadRs(rsDes, "SCHQTY")
       
       
       'If cboCurr <> wsCurr Then
       'If getExcRate(wsCurr, medDocDate, wsExcr, "") = True Then
       'outPrice = NBRnd(outPrice * To_Value(wsExcr) / txtExcr, giExrDp)
       'End If
       'End If
       
        outDisPer = To_Value(ReadRs(rsDes, "SODTDISPER"))
       
       Chk_grdITMCODE = True
    Else
        outAccID = ""
        OutName = ""
        OutBarCode = ""
        outPub = ""
        outPrice = 0
        outDisPer = 0
        outLotNo = ""
        outWhsCode = ""
        outQty = 0
        gsMsg = "沒有此物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdOrdItm(inSoNo As String, inItmNo As String, inWhsCode As String, InLotNo As String) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    
    wsSQL = "SELECT SODTITEMID "
    wsSQL = wsSQL & " FROM mstITEM, soaSoHd, soaSoDt "
    wsSQL = wsSQL & " WHERE SOHDDOCID = SODTDOCID "
    wsSQL = wsSQL & " AND SODTITEMID = ITMID "
    wsSQL = wsSQL & " AND SOHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSQL = wsSQL & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSQL = wsSQL & " AND SODTWHSCODE = '" & Set_Quote(inWhsCode) & "' "
    wsSQL = wsSQL & " AND SODTLOTNO = '" & Set_Quote(InLotNo) & "' "
    wsSQL = wsSQL & " AND SOHDSTATUS NOT IN ('2' , '3')"
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       Chk_grdOrdItm = True
    Else
        gsMsg = "沒有此物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdOrdItm = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdSoNo(inSoNo As String, ByRef outSoID As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    Chk_grdSoNo = False
    
    outSoID = "0"
    
    wsSQL = "SELECT SOHDDOCID, SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
    wsSQL = wsSQL & " WHERE SOHDSTATUS = '1' "
    wsSQL = wsSQL & " AND SOHDDOCNO = '" & Set_Quote(inSoNo) & "' "
              
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        gsMsg = "沒有此訂單!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
    Set rsRcd = Nothing
        Exit Function
    End If
       
    outSoID = To_Value(ReadRs(rsRcd, "SOHDDOCID"))
       
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_grdSoNo = True

End Function

Private Function Chk_grdWhsCode(inNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdWhsCode = False
    
    If Trim(inNo) = "" Then
        gsMsg = "必需輸入貨倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    wsSQL = "SELECT *  FROM mstWareHouse"
    wsSQL = wsSQL & " WHERE WHSCODE = '" & Set_Quote(inNo) & "' "
    wsSQL = wsSQL & " AND WHSSTATUS = '1' "
       
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
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

Private Function Chk_grdLotNo(inNo As String) As Boolean
    
   
  '  Chk_grdLotNo = False
  '
  '  If Trim(inNo) = "" Then
  '      gsMsg = "必需輸入版次!"
  '      MsgBox gsMsg, vbOKOnly, gsTitle
  '      Exit Function
  '  End If
    
    Chk_grdLotNo = True
    
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
                If Trim(.Columns(SONO)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, SONO)) = "" And _
                   Trim(waResult(inRow, ITMCODE)) = "" And _
                   Trim(waResult(inRow, ITMNAME)) = "" And _
                   Trim(waResult(inRow, Qty)) = "" And _
                   Trim(waResult(inRow, Price)) = "" And _
                   Trim(waResult(inRow, ITMID)) = "" Then
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
        
        If Chk_grdSoNo(waResult(LastRow, SONO), "") = False Then
            .Col = SONO
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdITMCODE(waResult(LastRow, SONO), waResult(LastRow, ITMCODE), "", "", "", "", "", 0, 0, "", "", 0) = False Then
            .Col = ITMCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdWhsCode(waResult(LastRow, WHSCODE)) = False Then
                .Col = WHSCODE
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_grdLotNo(waResult(LastRow, LOTNO)) = False Then
                .Col = LOTNO
                .Row = LastRow
                Exit Function
        End If
        
        
        If Chk_grdQty(waResult(LastRow, Qty)) = False Then
                .Col = Qty
                .Row = LastRow
                Exit Function
        End If
        
        'If Chk_grdDisPer(waResult(LastRow, DisPer)) = False Then
        '        .Col = DisPer
        '        Exit Function
        'End If
        
        'If Chk_Amount(waResult(LastRow, Amt)) = False Then
        '    .Col = Amt
        '    Exit Function
        'End If
        
        If Chk_grdOrdItm(waResult(LastRow, SONO), waResult(LastRow, ITMCODE), waResult(LastRow, WHSCODE), waResult(LastRow, LOTNO)) = False Then
            .Col = ITMCODE
            .Row = LastRow
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
    Dim wiTotalQty As Double
    
    Dim wiRowCtr As Integer
    
    Calc_Total = False
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotalQty = wiTotalQty + To_Value(waResult(wiRowCtr, Qty))
    Next
    
    lblDspTotalQty.Caption = Format(CStr(wiTotalQty), gsQtyFmt)
    
    Calc_Total = True

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
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Or wbReadOnly Then
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
    
    If DelValidation(wlKey) = False Then
       wiAction = CorRec
       MousePointer = vbDefault
       Exit Function
    End If
    
    wiAction = DelRec
    
      cnCon.BeginTrans
    Set adcmdDelete.ActiveConnection = cnCon
        
    adcmdDelete.CommandText = "USP_SPL001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, wlCusID)
    Call SetSPPara(adcmdDelete, 6, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 7, txtRevNo.Text)
    Call SetSPPara(adcmdDelete, 8, medETADate.Text)
    Call SetSPPara(adcmdDelete, 9, "")
    Call SetSPPara(adcmdDelete, 10, "")
    
    Call SetSPPara(adcmdDelete, 11, cboRmkCode.Text)
    
    Call SetSPPara(adcmdDelete, 12, txtCusPo.Text)
    
    Call SetSPPara(adcmdDelete, 13, txtShipFrom.Text)
    Call SetSPPara(adcmdDelete, 14, txtShipTo.Text)
    Call SetSPPara(adcmdDelete, 15, txtShipVia.Text)
    Call SetSPPara(adcmdDelete, 16, txtShipPer.Text)
    Call SetSPPara(adcmdDelete, 17, txtShipName.Text)
    Call SetSPPara(adcmdDelete, 18, txtShipAdr1.Text)
    Call SetSPPara(adcmdDelete, 19, txtShipAdr2.Text)
    Call SetSPPara(adcmdDelete, 20, txtShipAdr3.Text)
    Call SetSPPara(adcmdDelete, 21, txtShipAdr4.Text)
    
    For i = 1 To 10
        Call SetSPPara(adcmdDelete, 22 + i - 1, txtRmk(i).Text)
    Next
    Call SetSPPara(adcmdDelete, 32, wlRefDocID)
    Call SetSPPara(adcmdDelete, 33, wsFormID)
    
    Call SetSPPara(adcmdDelete, 34, gsUserID)
    Call SetSPPara(adcmdDelete, 35, wsGenDte)
    
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 36)
    wsDocNo = GetSPPara(adcmdDelete, 37)
    
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

Private Function SaveData() As Boolean

    Dim wiRet As Long
    
    SaveData = False
    
     If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) And _
        tbrProcess.Buttons(tcSave).Enabled = True Then
        
        gsMsg = "你是否確定不儲存現時之變更而離開?"
        If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbOK Then
        Exit Function
        Else
            If wiAction = DelRec Then
                If cmdDel = True Then
                    Exit Function
                End If
            Else
                If cmdSave = True Then
                    Exit Function
                End If
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
        
            Me.cboDocNo.Enabled = False
            Me.cboRefDocNo.Enabled = False
            
            Me.cboCusCode.Enabled = False
            Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            
            Me.medETADate.Enabled = False
            Me.cboShipCode.Enabled = False
            Me.cboRmkCode.Enabled = False
            Me.txtShipFrom.Enabled = False
            Me.txtShipTo.Enabled = False
            Me.txtShipPer.Enabled = False
            Me.txtShipVia.Enabled = False
            Me.txtShipName.Enabled = False
            Me.txtShipAdr1.Enabled = False
            Me.txtShipAdr2.Enabled = False
            Me.txtShipAdr3.Enabled = False
            Me.txtShipAdr4.Enabled = False
            
            Me.txtCusPo.Enabled = False
            
            Me.picRmk.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
            Me.cboRefDocNo.Enabled = True
            
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
            Me.cboRefDocNo.Enabled = False
            
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            If wiAction = AddRec Then
                Me.cboRefDocNo.Enabled = True
            Else
                Me.cboRefDocNo.Enabled = False
            End If
            
            Me.cboCusCode.Enabled = True
            Me.txtRevNo.Enabled = True
            Me.medDocDate.Enabled = True
            Me.medETADate.Enabled = True
            
            Me.cboShipCode.Enabled = True
            Me.cboRmkCode.Enabled = True
            Me.txtShipFrom.Enabled = True
            Me.txtShipTo.Enabled = True
            Me.txtShipPer.Enabled = True
            Me.txtShipVia.Enabled = True
            Me.txtShipName.Enabled = True
            Me.txtShipAdr1.Enabled = True
            Me.txtShipAdr2.Enabled = True
            Me.txtShipAdr3.Enabled = True
            Me.txtShipAdr4.Enabled = True
            
            Me.txtCusPo.Enabled = True
            
            Me.picRmk.Enabled = True
            
            
           
            Me.tblDetail.Enabled = True
           
    End Select
End Sub

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
    
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "SPHDDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub txtRevNo_LostFocus()
    FocusMe txtRevNo, True
End Sub

Private Sub txtShipFrom_GotFocus()
    FocusMe txtShipFrom
End Sub

Private Sub txtShipFrom_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipFrom, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtShipTo.SetFocus
    End If
End Sub

Private Sub txtShipFrom_LostFocus()
    FocusMe txtShipFrom, True
End Sub

Private Sub txtShipTo_GotFocus()
    FocusMe txtShipTo
End Sub

Private Sub txtShipTo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipTo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtShipVia.SetFocus
       
    End If
End Sub

Private Sub txtShipTo_LostFocus()
    FocusMe txtShipTo, True
End Sub

Private Sub txtShipVia_GotFocus()
    FocusMe txtShipVia
End Sub

Private Sub txtShipVia_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipVia, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        cboShipCode.SetFocus
    End If
End Sub

Private Sub txtShipVia_LostFocus()
        FocusMe txtShipVia, True
End Sub

Private Sub txtCusPo_GotFocus()
        FocusMe txtCusPo
End Sub

Private Sub txtCusPo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusPo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtShipFrom.SetFocus
    End If
End Sub

Private Sub txtCusPo_LostFocus()
    FocusMe txtCusPo, True
End Sub

Private Sub cboShipCode_GotFocus()
    FocusMe cboShipCode
End Sub

Private Sub cboShipCode_LostFocus()
    FocusMe cboShipCode, True
End Sub

Private Sub cboShipCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboShipCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboShipCode = False Then
                Exit Sub
        End If
        
        If wsOldShipCd <> cboShipCode.Text Then
            Get_ShipMark
            wsOldShipCd = cboShipCode.Text
        End If
        
        tabDetailInfo.Tab = 0
        txtShipPer.SetFocus
    End If
End Sub

Private Sub cboShipCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboShipCode
    
    wsSQL = "SELECT ShipCode, ShipName, ShipPer FROM mstShip WHERE ShipCode LIKE '%" & IIf(cboShipCode.SelLength > 0, "", Set_Quote(cboShipCode.Text)) & "%' "
    wsSQL = wsSQL & "AND ShipSTATUS = '1' "
    wsSQL = wsSQL & "AND ShipCardID = " & wlCusID & " "
    wsSQL = wsSQL & "ORDER BY ShipCode "
    Call Ini_Combo(3, wsSQL, cboShipCode.Left + tabDetailInfo.Left, cboShipCode.Top + cboShipCode.Height + tabDetailInfo.Top, tblCommon, "SDN001", "TBLSHIPCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboShipCode() As Boolean

    Chk_cboShipCode = False
     
    If Trim(cboShipCode.Text) = "" Then
        Chk_cboShipCode = True
        Exit Function
    End If
    
    
    If Chk_Ship(cboShipCode) = False Then
        gsMsg = "沒有此貨運編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboShipCode.SetFocus
       Exit Function
    End If
    
    
    Chk_cboShipCode = True
    
End Function

Private Sub txtShipName_GotFocus()
        FocusMe txtShipName
End Sub

Private Sub txtShipName_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtShipName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtShipAdr1.SetFocus
        
    End If
    
End Sub

Private Sub txtShipName_LostFocus()
        FocusMe txtShipName, True
End Sub

Private Sub txtShipPer_GotFocus()
        FocusMe txtShipPer
End Sub

Private Sub txtShipPer_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtShipPer, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtShipName.SetFocus
       
    End If
    
End Sub

Private Sub txtShipPer_LostFocus()
    FocusMe txtShipPer, True
End Sub

Private Sub txtShipAdr1_GotFocus()
    FocusMe txtShipAdr1
End Sub

Private Sub txtShipAdr1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipAdr1, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtShipAdr2.SetFocus
       
    End If
End Sub

Private Sub txtShipAdr1_LostFocus()
    FocusMe txtShipAdr1, True
End Sub

Private Sub txtShipAdr2_GotFocus()
    FocusMe txtShipAdr2
End Sub

Private Sub txtShipAdr2_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtShipAdr2, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtShipAdr3.SetFocus
       
    End If
End Sub

Private Sub txtShipAdr2_LostFocus()
    FocusMe txtShipAdr2, True
End Sub

Private Sub txtShipAdr3_GotFocus()
    FocusMe txtShipAdr3
End Sub

Private Sub txtShipAdr3_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipAdr3, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtShipAdr4.SetFocus
    End If
End Sub

Private Sub txtShipAdr3_LostFocus()
    FocusMe txtShipAdr3, True
End Sub

Private Sub txtShipAdr4_GotFocus()
    FocusMe txtShipAdr4
End Sub

Private Sub txtShipAdr4_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtShipAdr4, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        

            tabDetailInfo.Tab = 1
            tblDetail.Col = SONO
            tblDetail.SetFocus

    End If
End Sub

Private Sub txtShipAdr4_LostFocus()
    FocusMe txtShipAdr4, True
End Sub

Private Sub cboRmkCode_GotFocus()
    FocusMe cboRmkCode
End Sub

Private Sub cboRmkCode_LostFocus()
    FocusMe cboRmkCode, True
End Sub

Private Sub cboRmkCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboRmkCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboRmkCode = False Then
                Exit Sub
        End If
        
        If wsOldRmkCd <> cboRmkCode.Text Then
            Get_Remark
            wsOldRmkCd = cboRmkCode.Text
        End If
        
        tabDetailInfo.Tab = 2
        txtRmk(1).SetFocus
    End If
End Sub

Private Sub cboRmkCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboRmkCode
    
    wsSQL = "SELECT RmkCode FROM mstRemark WHERE RmkCode LIKE '%" & IIf(cboRmkCode.SelLength > 0, "", Set_Quote(cboRmkCode.Text)) & "%' "
    wsSQL = wsSQL & "AND RmkSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY RmkCode "
    Call Ini_Combo(1, wsSQL, cboRmkCode.Left + tabDetailInfo.Left, cboRmkCode.Top + cboRmkCode.Height + tabDetailInfo.Top, tblCommon, "SDN001", "TBLRMKCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboRmkCode() As Boolean

    Chk_cboRmkCode = False
     
    If Trim(cboRmkCode.Text) = "" Then
        Chk_cboRmkCode = True
        Exit Function
    End If
    
    
    If Chk_Remark(cboRmkCode) = False Then
        gsMsg = "沒有此備註!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboRmkCode.SetFocus
       Exit Function
    End If
    
    
    Chk_cboRmkCode = True
    
End Function

Private Sub txtRmk_GotFocus(Index As Integer)
    FocusMe txtRmk(Index)
End Sub

Private Sub txtRmk_KeyPress(Index As Integer, KeyAscii As Integer)
    Call chk_InpLen(txtRmk(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Index = 10 Then
            tabDetailInfo.Tab = 0
            cboCusCode.SetFocus
        Else
            tabDetailInfo.Tab = 2
            txtRmk(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtRmk_LostFocus(Index As Integer)
    FocusMe txtRmk(Index), True
End Sub

Private Sub Get_ShipMark()
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM  mstShip "
    wsSQL = wsSQL & "WHERE ShipCode = '" & Set_Quote(cboShipCode) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        
        txtShipName = ReadRs(rsRcd, "SHIPNAME")
        txtShipPer = ReadRs(rsRcd, "SHIPPER")
        txtShipAdr1 = ReadRs(rsRcd, "SHIPADR1")
        txtShipAdr2 = ReadRs(rsRcd, "SHIPADR2")
        txtShipAdr3 = ReadRs(rsRcd, "SHIPADR3")
        txtShipAdr4 = ReadRs(rsRcd, "SHIPADR4")
        
        Else
        txtShipName = ""
        txtShipPer = ""
        txtShipAdr1 = ""
        txtShipAdr2 = ""
        txtShipAdr3 = ""
        txtShipAdr4 = ""
        
    End If
    rsRcd.Close
    Set rsRcd = Nothing
End Sub

Private Sub Get_Remark()
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim i As Integer
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM  mstReMark "
    wsSQL = wsSQL & "WHERE RmkCode = '" & Set_Quote(cboRmkCode) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        
        For i = 1 To 10
        txtRmk(i) = ReadRs(rsRcd, "RmkDESC" & i)
        Next i
        
        Else
        
        For i = 1 To 10
        txtRmk(i) = ""
        Next i
        
        
    End If
    rsRcd.Close
    Set rsRcd = Nothing
End Sub

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Dim wsCurRecLn2 As String
    Dim wsCurRecLn3 As String
    
    Chk_NoDup = False
    
    wsCurRec = tblDetail.Columns(SONO)
    wsCurRecLn = tblDetail.Columns(ITMCODE)
    wsCurRecLn2 = tblDetail.Columns(WHSCODE)
    wsCurRecLn3 = tblDetail.Columns(LOTNO)
    
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, SONO) And _
                  wsCurRecLn = waResult(wlCtr, ITMCODE) And _
                  wsCurRecLn2 = waResult(wlCtr, WHSCODE) And _
                  wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
                  gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup = True

End Function

Private Function Chk_NoDup2(ByRef inRow As Long, ByVal wsCurRec As String, ByVal wsCurRecLn As String, ByVal wsCurRecLn2 As String, ByVal wsCurRecLn3 As String) As Boolean
    
    Dim wlCtr As Long
     
    Chk_NoDup2 = False
    
    
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, SONO) And _
                  wsCurRecLn = waResult(wlCtr, ITMCODE) And _
                  wsCurRecLn2 = waResult(wlCtr, WHSCODE) And _
                  wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
                  gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
                  MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
                  inRow = To_Value(waResult(wlCtr, LINENO))
                  Exit Function
               End If
            End If
        Next
    
    Chk_NoDup2 = True

End Function

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
    
    '' form delcare
    'Private waPopUpSub As New XArrayDB
    
    '' form unload
    'Set waPopUpSub = Nothing
    
    ''   addin ini_caption
 
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
            
            If .Bookmark = waResult.UpperBound(1) Then Exit Sub
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
    wsSQL = "EXEC usp_RPTSPL002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'SP', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & String(10, "z") & "', "
    wsSQL = wsSQL & "'" & "0000/00/00" & "', "
    wsSQL = wsSQL & "'" & "9999/99/99" & "', "
    wsSQL = wsSQL & "'" & "%" & "', "
    wsSQL = wsSQL & "'N', "
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTSPL002"
    Else
    wsRptName = "RPTSPL002"
    End If
    
    NewfrmPrint.ReportID = "SPL002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SPL002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdRefresh()

    
  If waResult.UpperBound(1) >= 0 Then
        
    
   tblDetail.ReBind
   tblDetail.FirstRow = 0
    
   Call Calc_Total
   
   End If
    
       
    
End Sub
Private Function Get_RefDoc() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    Dim wdBalQty As Double
    
    Get_RefDoc = False
    
        wsSQL = "SELECT SOHDDOCID, SOHDDOCNO, SOHDCUSID, CUSID, CUSCODE, CUSNAME, CUSTEL, CUSFAX, "
        wsSQL = wsSQL & "SOHDDOCDATE, SOHDREVNO, SOHDCURR, SOHDEXCR, "
        wsSQL = wsSQL & "SOHDDUEDATE, SOHDPRCCODE, SOHDSALEID, SOHDMLCODE, SOHDNatureCODE, "
        wsSQL = wsSQL & "SOHDCUSPO, SOHDLCNO, SOHDPORTNO, SOHDSHIPPER, SOHDSHIPFROM, SOHDSHIPTO, SOHDSHIPVIA, SOHDSHIPNAME, "
        wsSQL = wsSQL & "SOHDSHIPCODE, SOHDSHIPADR1,  SOHDSHIPADR2,  SOHDSHIPADR3,  SOHDSHIPADR4, "
        wsSQL = wsSQL & "SOHDRMKCODE, SOHDRMK1,  SOHDRMK2,  SOHDRMK3,  SOHDRMK4, SOHDRMK5, "
        wsSQL = wsSQL & "SOHDRMK6,  SOHDRMK7,  SOHDRMK8,  SOHDRMK9, SOHDRMK10, "
        wsSQL = wsSQL & "SOHDGRSAMT , SOHDGRSAMTL, SOHDDISAMT, SOHDDISAMTL, SOHDNETAMT, SOHDNETAMTL, "
        wsSQL = wsSQL & "SODTITEMID, ITMCODE, SODTWHSCODE, SODTLOTNO, ITMBARCODE, SODTITEMDESC ITNAME, ITMPUBLISHER, SODTTOTQTY - SODTSCHQTY BALQTY, SODTUPRICE, SODTDISPER, SODTAMT, SODTAMTL, SODTDIS, SODTDISL, SODTNET, SODTNETL, "
        wsSQL = wsSQL & "SODTID "
        wsSQL = wsSQL & "FROM  soaSOHD, soaSODT, mstCUSTOMER, mstITEM "
        wsSQL = wsSQL & "WHERE SOHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
        wsSQL = wsSQL & "AND SOHDDOCID = SODTDOCID "
        wsSQL = wsSQL & "AND SOHDCUSID = CUSID "
        wsSQL = wsSQL & "AND SODTITEMID = ITMID "
        wsSQL = wsSQL & "ORDER BY SODTDOCLINE "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    wsOldRefDocNo = cboRefDocNo.Text
    wlRefDocID = ReadRs(rsRcd, "SOHDDOCID")
    wlCusID = ReadRs(rsRcd, "CUSID")
    cboCusCode.Text = ReadRs(rsRcd, "CUSCODE")
    lblDspCusName.Caption = ReadRs(rsRcd, "CUSNAME")
    lblDspCusTel.Caption = ReadRs(rsRcd, "CUSTEL")
    lblDspCusFax.Caption = ReadRs(rsRcd, "CUSFAX")
    
    
    wlSaleID = To_Value(ReadRs(rsRcd, "SOHDSALEID"))
    
    cboShipCode = ReadRs(rsRcd, "SOHDSHIPCODE")
    cboRmkCode = ReadRs(rsRcd, "SOHDRMKCODE")
    
    txtCusPo = ReadRs(rsRcd, "SOHDCUSPO")
    
    
    txtShipFrom = ReadRs(rsRcd, "SOHDSHIPFROM")
    txtShipTo = ReadRs(rsRcd, "SOHDSHIPTO")
    txtShipVia = ReadRs(rsRcd, "SOHDSHIPVIA")
    txtShipName = ReadRs(rsRcd, "SOHDSHIPNAME")
    txtShipPer = ReadRs(rsRcd, "SOHDSHIPPER")
    txtShipAdr1 = ReadRs(rsRcd, "SOHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsRcd, "SOHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsRcd, "SOHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsRcd, "SOHDSHIPADR4")
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsRcd, "SOHDRMK" & i)
    Next i
    
    
    
    
    wlLineNo = 1
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, SOID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             
             '  wdBalQty = Get_SoBalQty(wsTrnCd, 0, ReadRs(rsRcd, "SOHDDOCID"), ReadRs(rsRcd, "SODTITEMID"), ReadRs(rsRcd, "SODTWHSCODE"), ReadRs(rsRcd, "SODTLOTNO"))
 
             .AppendRows
             waResult(.UpperBound(1), LINENO) = wlLineNo
             waResult(.UpperBound(1), SONO) = ReadRs(rsRcd, "SOHDDOCNO")
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "ITMBARCODE")
             waResult(.UpperBound(1), ITMNAME) = ReadRs(rsRcd, "ITNAME")
             waResult(.UpperBound(1), WHSCODE) = ReadRs(rsRcd, "SODTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsRcd, "SODTLOTNO")
             waResult(.UpperBound(1), Qty) = Format(To_Value(ReadRs(rsRcd, "BALQTY")), gsQtyFmt)
             waResult(.UpperBound(1), Price) = Format(ReadRs(rsRcd, "SODTUPRICE"), gsUprFmt)
             waResult(.UpperBound(1), ITMID) = ReadRs(rsRcd, "SODTITEMID")
             waResult(.UpperBound(1), SOID) = ReadRs(rsRcd, "SOHDDOCID")
             wlLineNo = wlLineNo + 1
             rsRcd.MoveNext
         Loop
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    Call Calc_Total
    
    Get_RefDoc = True
    
End Function

Private Function DelValidation(ByVal InDocID As Long) As Boolean
    
    
    DelValidation = False
    
    On Error GoTo DelValidation_Err
    
    
    
 '   If Not chk_txtRevNo Then Exit Function
    If Chk_SpQty(InDocID) = True Then
        
        gsMsg = "配貨單已經發貨!不能刪除"
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    
    End If
    
    DelValidation = True
    
    Exit Function
    
DelValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

