VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmGRV001 
   Caption         =   "訂貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmGRV001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11400
      OleObjectBlob   =   "frmGRV001.frx":030A
      TabIndex        =   40
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
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":66AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGRV001.frx":69CD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   41
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
         NumButtons      =   16
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
            Key             =   "Revise"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save (F10)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Cancel (F11)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Find (F9)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   480
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Header Information"
      TabPicture(0)   =   "frmGRV001.frx":6CE9
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboRefDocNo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboDocNo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboVdrCode"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboPayCode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboPrcCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboMLCode"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboSaleCode"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FraDate"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraInfo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraCode"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraKey"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmGRV001.frx":6D05
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblDspNetAmtOrg"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDspDisAmtOrg"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblDspGrsAmtOrg"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblDspTotalQty"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblNetAmtOrg"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblDisAmtOrg"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblGrsAmtOrg"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblTotalQty"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "tblDetail"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Item Information"
      TabPicture(2)   =   "frmGRV001.frx":6D21
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboShipCode"
      Tab(2).Control(1)=   "fraShip"
      Tab(2).Control(2)=   "cboRmkCode"
      Tab(2).Control(3)=   "fraRmk"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame2 
         Height          =   1680
         Left            =   -74880
         TabIndex        =   93
         Top             =   5280
         Width           =   3975
         Begin VB.CommandButton btnGetDisAmt 
            Caption         =   "Command1"
            Height          =   375
            Left            =   1680
            Picture         =   "frmGRV001.frx":6D3D
            TabIndex        =   14
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtDisAmt 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   13
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtSpecDis 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblDisAmt 
            Caption         =   "EXCR"
            Height          =   495
            Left            =   120
            TabIndex        =   95
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label lblSpecDis 
            Caption         =   "SPECDIS"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   300
            Width           =   1545
         End
      End
      Begin VB.ComboBox cboRefDocNo 
         Height          =   300
         Left            =   -73200
         TabIndex        =   2
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   22
         Top             =   480
         Width           =   2010
      End
      Begin VB.Frame fraShip 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   69
         Top             =   120
         Width           =   11535
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1455
            Left            =   1680
            ScaleHeight     =   1395
            ScaleWidth      =   9555
            TabIndex        =   70
            Top             =   1440
            Width           =   9615
            Begin VB.TextBox txtShipAdr4 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   28
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1080
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr3 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   27
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   720
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr2 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   26
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr1 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   25
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   5865
            End
         End
         Begin VB.TextBox txtShipName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   24
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1080
            Width           =   4305
         End
         Begin VB.TextBox txtShipPer 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   23
            Text            =   "01234567890123457890"
            Top             =   720
            Width           =   4305
         End
         Begin VB.Label lblShipCode 
            Caption         =   "SHIPCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblShipName 
            Caption         =   "SHIPNAME"
            Height          =   240
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   1380
         End
         Begin VB.Label lblShipPer 
            Caption         =   "SHIPPER"
            Height          =   240
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblShipAdr 
            Caption         =   "SHIPADR"
            Height          =   240
            Left            =   120
            TabIndex        =   71
            Top             =   1440
            Width           =   1500
         End
      End
      Begin VB.ComboBox cboRmkCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   29
         Top             =   3600
         Width           =   1890
      End
      Begin VB.ComboBox cboDocNo 
         Height          =   300
         Left            =   -73200
         TabIndex        =   0
         Top             =   420
         Width           =   1935
      End
      Begin VB.ComboBox cboVdrCode 
         Height          =   300
         Left            =   -69720
         TabIndex        =   3
         Top             =   780
         Width           =   2055
      End
      Begin VB.ComboBox cboPayCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   7
         Top             =   2820
         Width           =   2370
      End
      Begin VB.ComboBox cboPrcCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   8
         Top             =   3550
         Width           =   2370
      End
      Begin VB.ComboBox cboMLCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   9
         Top             =   3180
         Width           =   2370
      End
      Begin VB.ComboBox cboSaleCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   6
         Top             =   2460
         Width           =   2370
      End
      Begin VB.Frame FraDate 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   50
         Top             =   4080
         Width           =   3975
         Begin MSMask.MaskEdBox medDueDate 
            Height          =   285
            Left            =   1680
            TabIndex        =   10
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medETADate 
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblETADate 
            Caption         =   "ETADATE"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   660
            Width           =   1425
         End
         Begin VB.Label lblDueDate 
            Caption         =   "DUEDATE"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   2895
         Left            =   -70800
         TabIndex        =   43
         Top             =   4080
         Width           =   7575
         Begin VB.TextBox txtLcNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   19
            Text            =   "0123456789012345789"
            Top             =   1680
            Width           =   5265
         End
         Begin VB.TextBox txtPortNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   20
            Text            =   "0123456789012345789"
            Top             =   2040
            Width           =   5265
         End
         Begin VB.TextBox txtCusPo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   18
            Text            =   "0123456789012345789"
            Top             =   1320
            Width           =   5265
         End
         Begin VB.TextBox txtShipTo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   16
            Text            =   "0123456789012345789"
            Top             =   600
            Width           =   5265
         End
         Begin VB.TextBox txtShipVia 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   17
            Text            =   "0123456789012345789"
            Top             =   960
            Width           =   5265
         End
         Begin VB.TextBox txtShipFrom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   15
            Text            =   "0123456789012345789"
            Top             =   240
            Width           =   5265
         End
         Begin VB.Label lblLcNo 
            Caption         =   "LCNO"
            Height          =   240
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   2100
         End
         Begin VB.Label lblPortNo 
            Caption         =   "PORTNO"
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label lblCusPo 
            Caption         =   "CUSPO"
            Height          =   240
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Label lblShipTo 
            Caption         =   "SHIPTO"
            Height          =   240
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   2100
         End
         Begin VB.Label lblShipVia 
            Caption         =   "SHIPVIA"
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   2100
         End
         Begin VB.Label lblShipFrom 
            Caption         =   "SHIPFROM"
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2100
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   6855
         Left            =   120
         OleObjectBlob   =   "frmGRV001.frx":717F
         TabIndex        =   21
         Top             =   840
         Width           =   11535
      End
      Begin VB.Frame fraCode 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   60
         Top             =   1980
         Width           =   11655
         Begin VB.Label lblMlCode 
            Caption         =   "MLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label lblDspMLDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   67
            Top             =   1200
            Width           =   7335
         End
         Begin VB.Label lblPrcCode 
            Caption         =   "PRCCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   66
            Top             =   1680
            Width           =   1545
         End
         Begin VB.Label lblDspPrcDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   65
            Top             =   1560
            Width           =   7335
         End
         Begin VB.Label lblPayCode 
            Caption         =   "PAYCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   64
            Top             =   900
            Width           =   1545
         End
         Begin VB.Label lblDspPayDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   63
            Top             =   840
            Width           =   7335
         End
         Begin VB.Label lblSaleCode 
            Caption         =   "SALECODE"
            Height          =   240
            Left            =   120
            TabIndex        =   62
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label lblDspSaleDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   61
            Top             =   480
            Width           =   7335
         End
      End
      Begin VB.Frame fraRmk 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   75
         Top             =   3360
         Width           =   11535
         Begin VB.PictureBox picRmk 
            BackColor       =   &H80000009&
            Height          =   3495
            Left            =   1680
            ScaleHeight     =   3435
            ScaleWidth      =   9555
            TabIndex        =   76
            Top             =   600
            Width           =   9615
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   31
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   30
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   32
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   690
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   35
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1740
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   33
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1035
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   34
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1395
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   36
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2085
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   8
               Left            =   0
               TabIndex        =   37
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2430
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   9
               Left            =   0
               TabIndex        =   38
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2775
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   10
               Left            =   0
               TabIndex        =   39
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   3120
               Width           =   7545
            End
         End
         Begin VB.Label lblRmkCode 
            Caption         =   "RMKCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label lblRmk 
            Caption         =   "RMK"
            Height          =   240
            Left            =   120
            TabIndex        =   77
            Top             =   600
            Width           =   1500
         End
      End
      Begin VB.Frame fraKey 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   79
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtExcr 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   9360
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1380
            Width           =   1335
         End
         Begin VB.ComboBox cboCurr 
            Height          =   300
            Left            =   9360
            TabIndex        =   4
            Top             =   1020
            Width           =   1335
         End
         Begin MSMask.MaskEdBox medDocDate 
            Height          =   285
            Left            =   9360
            TabIndex        =   1
            Top             =   660
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblRefDocNo 
            Caption         =   "REFDOCNO"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblVdrCode 
            Caption         =   "VDRCODE"
            Height          =   255
            Left            =   3840
            TabIndex        =   90
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
            TabIndex        =   89
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblDocDate 
            Caption         =   "DOCDATE"
            Height          =   255
            Left            =   7365
            TabIndex        =   88
            Top             =   720
            Width           =   1680
         End
         Begin VB.Label lblDspVdrName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   87
            Top             =   1020
            Width           =   5535
         End
         Begin VB.Label LblCurr 
            Caption         =   "CURR"
            Height          =   255
            Left            =   7365
            TabIndex        =   86
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lblExcr 
            Caption         =   "EXCR"
            Height          =   255
            Left            =   7365
            TabIndex        =   85
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label lblDspVdrTel 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   84
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lblVdrName 
            Caption         =   "VDRNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblDspVdrFax 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   5160
            TabIndex        =   82
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lblVdrFax 
            Caption         =   "VDRFAX"
            Height          =   255
            Left            =   3840
            TabIndex        =   81
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblVdrTel 
            Caption         =   "VDRTEL"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1440
            Width           =   1575
         End
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
         TabIndex        =   59
         Top             =   60
         Width           =   1755
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
         TabIndex        =   58
         Top             =   60
         Width           =   2475
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
         TabIndex        =   57
         Top             =   60
         Width           =   2355
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
         TabIndex        =   56
         Top             =   60
         Width           =   3195
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
         TabIndex        =   55
         Top             =   420
         Width           =   1890
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
         TabIndex        =   54
         Top             =   420
         Width           =   2490
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
         TabIndex        =   53
         Top             =   420
         Width           =   2370
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
         TabIndex        =   52
         Top             =   420
         Width           =   3210
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
Attribute VB_Name = "frmGRV001"
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

Private wsOldVdrNo As String
Private wsOldCurCd As String
Private wsOldShipCd As String
Private wsOldRmkCd As String
Private wsOldPayCd As String
Private wbReadOnly As Boolean
Private wgsTitle As String
Private wsOldRefDocNo As String
Private wbUpdCstOnly As Boolean

Private Const LINENO = 0
Private Const ITMCODE = 1
Private Const ITMTYPE = 2
Private Const WHSCODE = 3
Private Const LOTNO = 4
Private Const ITMNAME = 5
Private Const PUBLISHER = 6
Private Const QTY = 7
Private Const PRICE = 8
Private Const DisPer = 9
Private Const Dis = 10
Private Const Amt = 11
Private Const NET = 12
Private Const Netl = 13
Private Const Disl = 14
Private Const Amtl = 15
Private Const ITMID = 16
Private Const POID = 17

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
Private Const tcRevise = "Revise"


Private wiOpenDoc As Integer
Private wiAction As Integer
Private wiRevNo As Integer
Private wlVdrID As Long
Private wlSaleID As Long
Private wlRefDocID As Long
Private wlLineNo As Long

Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "popGRHD"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, LINENO, POID
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
    Call SetDateMask(medDueDate)
    Call SetDateMask(medETADate)
      
    wsOldRefDocNo = ""
    wsOldVdrNo = ""
    wsOldCurCd = ""
    wsOldShipCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""
    
    wlKey = 0
    wlVdrID = 0
    wlSaleID = 0
    wlRefDocID = 0
    wlLineNo = 1
    wbReadOnly = False
    
    txtSpecDis.Text = Format("0", gsAmtFmt)
    txtDisAmt.Text = Format("0", gsAmtFmt)
    
    wiRevNo = Format(0, "##0")
    tblCommon.Visible = False
    
    Me.Caption = wsFormCaption
    wbUpdCstOnly = False
    Call Ini_UnLockGrid

    
    FocusMe cboDocNo
    tabDetailInfo.Tab = 0
End Sub


Private Sub btnGetDisAmt_Click()
    If To_Value(txtSpecDis.Text) = 0 Then
    lblDspDisAmtOrg = Format(To_Value(txtDisAmt.Text), gsAmtFmt)
    lblDspNetAmtOrg = Format(To_Value(lblDspGrsAmtOrg.Caption) - To_Value(txtDisAmt.Text), gsAmtFmt)
    Else
    txtDisAmt.Text = Format(To_Value(lblDspGrsAmtOrg.Caption) * To_Value(txtSpecDis.Text), gsAmtFmt)
    lblDspDisAmtOrg = Format(To_Value(lblDspGrsAmtOrg.Caption) * To_Value(txtSpecDis.Text), gsAmtFmt)
    lblDspNetAmtOrg = Format(To_Value(lblDspGrsAmtOrg.Caption) * (1 - To_Value(txtSpecDis.Text)), gsAmtFmt)
    End If
    
End Sub

Private Sub cboCurr_GotFocus()
    FocusMe cboCurr
End Sub

Private Sub cboCurr_LostFocus()
FocusMe cboCurr, True
End Sub

Private Sub cboVdrCode_LostFocus()
    FocusMe cboVdrCode, True
End Sub

Private Sub cboCurr_KeyPress(KeyAscii As Integer)
    Dim wsExcRate As String
    Dim wsExcDesc As String
    
    Call chk_InpLen(cboCurr, 3, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCurr = False Then
                Exit Sub
        End If
        
        If getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) = False Then
            gsMsg = "沒有此貨幣!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            txtExcr.Text = Format(0, gsExrFmt)
            cboCurr.SetFocus
            Exit Sub
        End If
        
        If wsOldCurCd <> cboCurr.Text Then
            txtExcr.Text = Format(wsExcRate, gsExrFmt)
            wsOldCurCd = cboCurr.Text
        End If
        
        If UCase(cboCurr) = UCase(wsBaseCurCd) Then
            txtExcr.Text = Format("1", gsExrFmt)
            txtExcr.Enabled = False
        Else
            txtExcr.Enabled = True
        End If
        
        If txtExcr.Enabled Then
            txtExcr.SetFocus
        Else
           If Chk_KeyFld Then
            tabDetailInfo.Tab = 0
            cboSaleCode.SetFocus
           End If
        End If
    End If
    
End Sub

Private Sub cboCurr_DropDown()
    
    Dim wsSQL As String
    Dim wsCtlDte As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCurr
    
    wsCtlDte = IIf(Trim(medDocDate.Text) = "" Or Trim(medDocDate.Text) = "/  /", gsSystemDate, medDocDate.Text)
    wsSQL = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboCurr.SelLength > 0, "", Set_Quote(cboCurr.Text)) & "%' "
    wsSQL = wsSQL & " AND EXCMN = '" & To_Value(Format(wsCtlDte, "MM")) & "' "
    wsSQL = wsSQL & " AND EXCYR = '" & Set_Quote(Format(wsCtlDte, "YYYY")) & "' "
    wsSQL = wsSQL & " AND EXCSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY EXCCURR "
    Call Ini_Combo(2, wsSQL, cboCurr.Left + tabDetailInfo.Left, cboCurr.Top + cboCurr.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCURCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboCurr() As Boolean
    
    Chk_cboCurr = False
     
    If Trim(cboCurr.Text) = "" Then
        gsMsg = "必需輸入貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCurr.SetFocus
        Exit Function
    End If
    
    
    If Chk_Curr(cboCurr, medDocDate.Text) = False Then
        gsMsg = "沒有此貨幣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCurr.SetFocus
       Exit Function
    End If
    
    
    Chk_cboCurr = True
    
End Function

Private Sub cboDocNo_GotFocus()
    
    FocusMe cboDocNo

End Sub

Private Sub cboDocNo_DropDown()
    
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSQL = "SELECT GRHDDOCNO, VDRCODE, GRHDDOCDATE "
    wsSQL = wsSQL & " FROM popGRHD, MstVendor "
    wsSQL = wsSQL & " WHERE GRHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND GRHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND GRHDSTATUS IN ('1','4')"
    wsSQL = wsSQL & " ORDER BY GRHDDOCNO DESC "
    Call Ini_Combo(3, wsSQL, cboDocNo.Left + tabDetailInfo.Left, cboDocNo.Top + cboDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboDocNo_LostFocus()
    FocusMe cboDocNo, True
End Sub

Private Sub cboDocNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLenC(cboDocNo, 15, KeyAscii, True, True)
    
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
        
        'If wsStatus = "4" Then
        '    gsMsg = "文件已入數, 祇可以更新基本資料!"
        '    MsgBox gsMsg, vbOKOnly, gsTitle
        '    wbReadOnly = True
        'End If
        
        
        If Get_TableInfo("POPGRHD", "GRHDDOCNO = '" & Set_Quote(cboDocNo) & "'", "GRHDUPDFLG") = "Y" Then
            
            gsMsg = "文件已會計入數!現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
            
        End If
            
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數!現在以更正成本"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Call Ini_LockGrid
            wbUpdCstOnly = True
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
        wiRevNo = Format(0, "##0")
        medDocDate.Text = Dsp_Date(Now)
    
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

        wsOldVdrNo = cboVdrCode.Text
        wsOldCurCd = cboCurr.Text
        wsOldShipCd = cboShipCode.Text
        wsOldRmkCd = cboRmkCode.Text
        wsOldPayCd = cboPayCode.Text
        
        
        Call SetButtonStatus("AfrKeyEdit")
        Call SetFieldStatus("AfrKey")
        cboVdrCode.SetFocus
    End If
    
    Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    If UCase(cboCurr) = UCase(wsBaseCurCd) Then
            txtExcr.Text = Format("1", gsExrFmt)
            txtExcr.Enabled = False
        Else
            txtExcr.Enabled = True
    End If

    
    tabDetailInfo.Tab = 0
    
        
End Sub

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
        
    
    Call Ini_Form
    Call Ini_Grid
    Call Ini_Caption
    Call Ini_Scr
  
  
    MousePointer = vbDefault

End Sub

Private Function LoadRecord() As Boolean
    Dim rsInvoice As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
        wsSQL = "SELECT GRHDDOCID, GRHDDOCNO, GRHDREFDOCID, GRHDVDRID, VDRID, VDRCODE, VDRNAME, VDRTEL, VDRFAX, "
        wsSQL = wsSQL & "GRHDDOCDATE, GRHDREVNO, GRHDCURR, GRHDEXCR, GRDTDOCLINE, "
        wsSQL = wsSQL & "GRHDDUEDATE, GRHDETADATE, GRHDDEPNO, GRHDDEPAMT, GRHDDEPML, GRHDPAYCODE, GRHDCRML, GRHDSALEID, GRHDMLCODE, "
        wsSQL = wsSQL & "GRHDCUSPO, GRHDLCNO, GRHDPORTNO, GRHDSHIPPER, GRHDSHIPFROM, GRHDSHIPTO, GRHDSHIPVIA, GRHDSHIPNAME, "
        wsSQL = wsSQL & "GRHDSHIPCODE, GRHDSHIPADR1,  GRHDSHIPADR2,  GRHDSHIPADR3,  GRHDSHIPADR4, "
        wsSQL = wsSQL & "GRHDRMKCODE, GRHDRMK1,  GRHDRMK2,  GRHDRMK3,  GRHDRMK4, GRHDRMK5, "
        wsSQL = wsSQL & "GRHDRMK6,  GRHDRMK7,  GRHDRMK8,  GRHDRMK9, GRHDRMK10, "
        wsSQL = wsSQL & "GRHDGRSAMT , GRHDGRSAMTL, GRHDDISAMT, GRHDDISAMTL, GRHDNETAMT, GRHDNETAMTL, "
        wsSQL = wsSQL & "GRDTITEMID, ITMCODE, GRDTWHSCODE, GRDTLOTNO, ITMITMTYPECODE, GRDTITEMDESC ITNAME,  GRDTQTY, GRDTUPRICE, GRDTDISPER, GRDTAMT, GRDTAMTL, GRDTDIS, GRDTDISL, GRDTNET, GRDTNETL, "
        wsSQL = wsSQL & "GRDTPOID,GRHDSPECDIS "
        wsSQL = wsSQL & "FROM  popGRHD, popGRDT, MstVendor, mstITEM "
        wsSQL = wsSQL & "WHERE GRHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND GRHDDOCID = GRDTDOCID "
        wsSQL = wsSQL & "AND GRHDVDRID = VDRID "
        wsSQL = wsSQL & "AND GRDTITEMID = ITMID "
        wsSQL = wsSQL & "ORDER BY GRDTDOCLINE "
    
    rsInvoice.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount <= 0 Then
        rsInvoice.Close
        Set rsInvoice = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsInvoice, "GRHDDOCID")
    wiRevNo = To_Value(ReadRs(rsInvoice, "GRHDREVNO"))
    medDocDate.Text = ReadRs(rsInvoice, "GRHDDOCDATE")
    wlRefDocID = ReadRs(rsInvoice, "GRHDREFDOCID")
    cboRefDocNo.Text = Get_TableInfo("popPOHD", "POHDDOCID =" & wlRefDocID, "POHDDOCNO")
    wlVdrID = ReadRs(rsInvoice, "VDRID")
    
    cboVdrCode.Text = ReadRs(rsInvoice, "VDRCODE")
    lblDspVdrName.Caption = ReadRs(rsInvoice, "VDRNAME")
    lblDspVdrTel.Caption = ReadRs(rsInvoice, "VDRTEL")
    lblDspVdrFax.Caption = ReadRs(rsInvoice, "VDRFAX")
    cboCurr.Text = ReadRs(rsInvoice, "GRHDCURR")
    txtExcr.Text = Format(ReadRs(rsInvoice, "GRHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "GRHDDUEDATE"))
    medETADate.Text = Dsp_MedDate(ReadRs(rsInvoice, "GRHDETADATE"))
    
    wlSaleID = To_Value(ReadRs(rsInvoice, "GRHDSALEID"))
    
    cboPayCode = ReadRs(rsInvoice, "GRHDPAYCODE")
    cboPrcCode = ReadRs(rsInvoice, "GRHDCRML")
    cboMLCode = ReadRs(rsInvoice, "GRHDMLCODE")
    cboShipCode = ReadRs(rsInvoice, "GRHDSHIPCODE")
    cboRmkCode = ReadRs(rsInvoice, "GRHDRMKCODE")
    
    txtCusPo = ReadRs(rsInvoice, "GRHDCUSPO")
    txtLcNo = ReadRs(rsInvoice, "GRHDLCNO")
    txtPortNo = ReadRs(rsInvoice, "GRHDPORTNO")
    
    txtSpecDis = Format(ReadRs(rsInvoice, "GRHDSPECDIS"), gsExrFmt)
    txtDisAmt.Text = Format(To_Value(ReadRs(rsInvoice, "GRHDDISAMT")), gsAmtFmt)
    
    
    txtShipFrom = ReadRs(rsInvoice, "GRHDSHIPFROM")
    txtShipTo = ReadRs(rsInvoice, "GRHDSHIPTO")
    txtShipVia = ReadRs(rsInvoice, "GRHDSHIPVIA")
    txtShipName = ReadRs(rsInvoice, "GRHDSHIPNAME")
    txtShipPer = ReadRs(rsInvoice, "GRHDSHIPPER")
    txtShipAdr1 = ReadRs(rsInvoice, "GRHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsInvoice, "GRHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsInvoice, "GRHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsInvoice, "GRHDSHIPADR4")
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsInvoice, "GRHDRMK" & i)
    Next i
    
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboPrcCode.Text) & "'", "MLDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    rsInvoice.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, POID
         Do While Not rsInvoice.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsInvoice, "GRDTDOCLINE")
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsInvoice, "ITMCODE")
             waResult(.UpperBound(1), ITMTYPE) = ReadRs(rsInvoice, "ITMITMTYPECODE")
             waResult(.UpperBound(1), ITMNAME) = ReadRs(rsInvoice, "ITNAME")
             waResult(.UpperBound(1), WHSCODE) = ReadRs(rsInvoice, "GRDTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsInvoice, "GRDTLOTNO")
             waResult(.UpperBound(1), PUBLISHER) = ""
             'waResult(.UpperBound(1), Qty) = Format(ReadRs(rsInvoice, "GRDTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), QTY) = Format(ReadRs(rsInvoice, "GRDTQTY"), gsAmtFmt)
             waResult(.UpperBound(1), PRICE) = Format(ReadRs(rsInvoice, "GRDTUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsInvoice, "GRDTDISPER"), gsAmtFmt)
             waResult(.UpperBound(1), Amt) = Format(ReadRs(rsInvoice, "GRDTAMT"), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsInvoice, "GRDTAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(ReadRs(rsInvoice, "GRDTDIS"), gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(ReadRs(rsInvoice, "GRDTDISL"), gsAmtFmt)
             waResult(.UpperBound(1), NET) = Format(ReadRs(rsInvoice, "GRDTNET"), gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(ReadRs(rsInvoice, "GRDTNETL"), gsAmtFmt)
             waResult(.UpperBound(1), ITMID) = ReadRs(rsInvoice, "GRDTITEMID")
             waResult(.UpperBound(1), POID) = ReadRs(rsInvoice, "GRDTPOID")
             
             rsInvoice.MoveNext
         Loop
         wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsInvoice.Close
    
    Set rsInvoice = Nothing
    
    Call Calc_Total
    
    LoadRecord = True
    
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP_M", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblRefDocNo.Caption = Get_Caption(waScrItm, "REFNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblVdrCode.Caption = Get_Caption(waScrItm, "VDRCODE")
    lblVdrName.Caption = Get_Caption(waScrItm, "VDRNAME")
    lblVdrTel.Caption = Get_Caption(waScrItm, "VDRTEL")
    lblVdrFax.Caption = Get_Caption(waScrItm, "VDRFAX")
    LblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblPayCode.Caption = Get_Caption(waScrItm, "PAYCODE")
    lblPrcCode.Caption = Get_Caption(waScrItm, "PRCCODE")
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblDueDate.Caption = Get_Caption(waScrItm, "DUEDATE")
    lblETADate.Caption = Get_Caption(waScrItm, "ETADATE")
    
    lblGrsAmtOrg.Caption = Get_Caption(waScrItm, "GRSAMTORG")
    lblNetAmtOrg.Caption = Get_Caption(waScrItm, "NETAMTORG")
    lblDisAmtOrg.Caption = Get_Caption(waScrItm, "DISAMTORG")
    lblTotalQty.Caption = Get_Caption(waScrItm, "TOTALQTY")
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(ITMCODE).Caption = Get_Caption(waScrItm, "ITMCODE")
        .Columns(ITMTYPE).Caption = Get_Caption(waScrItm, "ITMTYPE")
        .Columns(WHSCODE).Caption = Get_Caption(waScrItm, "WHSCODE")
        .Columns(LOTNO).Caption = Get_Caption(waScrItm, "LOTNO")
        .Columns(ITMNAME).Caption = Get_Caption(waScrItm, "ITMNAME")
        .Columns(PUBLISHER).Caption = Get_Caption(waScrItm, "PUBLISHER")
        .Columns(QTY).Caption = Get_Caption(waScrItm, "QTY")
        .Columns(PRICE).Caption = Get_Caption(waScrItm, "PRICE")
        .Columns(DisPer).Caption = Get_Caption(waScrItm, "DISPER")
        .Columns(Dis).Caption = Get_Caption(waScrItm, "DIS")
        .Columns(NET).Caption = Get_Caption(waScrItm, "NET")
        .Columns(Amt).Caption = Get_Caption(waScrItm, "AMT")
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
    lblLcNo.Caption = Get_Caption(waScrItm, "LCNO")
    lblPortNo.Caption = Get_Caption(waScrItm, "PORTNO")
    lblShipName.Caption = Get_Caption(waScrItm, "SHIPNAME")
    lblRmkCode.Caption = Get_Caption(waScrItm, "RMKCODE")
    lblRmk.Caption = Get_Caption(waScrItm, "RMK")
    
    lblSpecDis.Caption = Get_Caption(waScrItm, "SPECDIS")
    lblDisAmt.Caption = Get_Caption(waScrItm, "DISAMTORG")
    btnGetDisAmt.Caption = Get_Caption(waScrItm, "GETDISAMT")
    
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
    tbrProcess.Buttons(tcRevise).ToolTipText = Get_Caption(waScrToolTip, tcRevise)
    
    wsActNam(1) = Get_Caption(waScrItm, "PVADD")
    wsActNam(2) = Get_Caption(waScrItm, "PVEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PVDELETE")
    wgsTitle = Get_Caption(waScrItm, "TITLE")
    
    
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
    Set frmGRV001 = Nothing

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
        cboCurr.SetFocus
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

Private Sub medDueDate_GotFocus()
    
  FocusMe medDueDate
    
End Sub

Private Sub medDueDate_LostFocus()

    FocusMe medDueDate, True
    
End Sub

Private Sub medDueDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medDueDate Then
            tabDetailInfo.Tab = 0
            medETADate.SetFocus
        End If
    End If
End Sub

Private Function Chk_medDueDate() As Boolean
    Chk_medDueDate = False
    
    If Trim(medDueDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medDueDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medDueDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medDueDate.SetFocus
        Exit Function
    End If
    
    
    Chk_medDueDate = True

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
            Me.txtSpecDis.SetFocus
        End If
    End If
End Sub

Private Sub medETADate_LostFocus()
    FocusMe medETADate, True
End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
        If cboVdrCode.Enabled Then
            cboVdrCode.SetFocus
        End If
               
    ElseIf tabDetailInfo.Tab = 1 Then
        
        If tblDetail.Enabled Then
            tblDetail.Col = ITMCODE
            tblDetail.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 2 Then
    
        If cboShipCode.Enabled Then
            cboShipCode.SetFocus
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
    
    Dim rspopGRHD As New ADODB.Recordset
    Dim wsSQL As String

    wsSQL = "SELECT GRHDSTATUS FROM popGRHD WHERE GRHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rspopGRHD.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rspopGRHD.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rspopGRHD.Close
    Set rspopGRHD = Nothing
End Function

Private Function Chk_KeyFld() As Boolean
    Chk_KeyFld = False
    
    If Chk_cboRefDocNo = False Then
        Exit Function
    End If
    
    If chk_cboVdrCode = False Then
        Exit Function
    End If
    
    If Chk_medDocDate = False Then
        Exit Function
    End If
    
    If Chk_cboCurr = False Then
        Exit Function
    End If
    
    If txtExcr.Enabled = True Then
        If chk_txtExcr = False Then
            Exit Function
        End If
    End If
    
    tblDetail.Enabled = True
    
    Chk_KeyFld = True
End Function
Private Function cmdSave() As Boolean

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
    Dim wiCtr As Integer
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
       If wiAction = RevRec Then wiAction = CorRec
       Exit Function
    End If
    
    '' Last Check when Add
    
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
    'If lblDspNetAmtOrg.Caption > Get_CreditLimit(wlVdrID, Trim(medDocDate.Text)) Then
    '    gsMsg = "已超過信貸額!"
    '    MsgBox gsMsg, vbOKOnly, gsTitle
    '    MousePointer = vbDefault
    '    Exit Function
    'End If
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    If wbReadOnly = True Then
    wiAction = CorRO
    End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_GRV001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlVdrID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, wiRevNo)
    Call SetSPPara(adcmdSave, 8, cboCurr.Text)
    Call SetSPPara(adcmdSave, 9, txtExcr.Text)
    Call SetSPPara(adcmdSave, 10, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medDueDate.Text))
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medETADate.Text))
    
    Call SetSPPara(adcmdSave, 13, wlSaleID)
    
    Call SetSPPara(adcmdSave, 14, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 15, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 16, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 17, cboShipCode.Text)
    Call SetSPPara(adcmdSave, 18, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 19, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 20, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 21, txtPortNo.Text)
    
    Call SetSPPara(adcmdSave, 22, "")
    Call SetSPPara(adcmdSave, 23, txtSpecDis.Text)
    Call SetSPPara(adcmdSave, 24, "")
    
    Call SetSPPara(adcmdSave, 25, txtShipFrom.Text)
    Call SetSPPara(adcmdSave, 26, txtShipTo.Text)
    Call SetSPPara(adcmdSave, 27, txtShipVia.Text)
    Call SetSPPara(adcmdSave, 28, txtShipPer.Text)
    Call SetSPPara(adcmdSave, 29, txtShipName.Text)
    Call SetSPPara(adcmdSave, 30, txtShipAdr1.Text)
    Call SetSPPara(adcmdSave, 31, txtShipAdr2.Text)
    Call SetSPPara(adcmdSave, 32, txtShipAdr3.Text)
    Call SetSPPara(adcmdSave, 33, txtShipAdr4.Text)
    
    For i = 1 To 10
        Call SetSPPara(adcmdSave, 34 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdSave, 44, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 45, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 46, lblDspNetAmtOrg)
    Call SetSPPara(adcmdSave, 47, wlRefDocID)
    
    Call SetSPPara(adcmdSave, 48, wsFormID)
    
    Call SetSPPara(adcmdSave, 49, gsUserID)
    Call SetSPPara(adcmdSave, 50, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 51)
    wsDocNo = GetSPPara(adcmdSave, 52)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If wbReadOnly = False Then
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_GRV001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, cboRefDocNo.Text)
                Call SetSPPara(adcmdSave, 4, To_Value(waResult(wiCtr, POID)))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, ITMID))
                Call SetSPPara(adcmdSave, 6, wiCtr + 1)
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, ITMNAME))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, QTY))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, PRICE))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, DisPer))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, WHSCODE))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, LOTNO))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, Amt))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, Amtl))
                Call SetSPPara(adcmdSave, 15, waResult(wiCtr, Dis))
                Call SetSPPara(adcmdSave, 16, waResult(wiCtr, Disl))
                Call SetSPPara(adcmdSave, 17, waResult(wiCtr, NET))
                Call SetSPPara(adcmdSave, 18, waResult(wiCtr, Netl))
                Call SetSPPara(adcmdSave, 19, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 20, gsUserID)
                Call SetSPPara(adcmdSave, 21, wsGenDte)
                adcmdSave.Execute
                
                
                
            End If
        Next
    End If
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
    
    If wiAction = CorRO Then
        gsMsg = "基本資料已儲存!"
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
    
    
    
    If Not Chk_medDocDate Then Exit Function
    If Not Chk_cboRefDocNo Then Exit Function
    
    If Not chk_cboVdrCode() Then Exit Function
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    If Not Chk_cboPayCode Then Exit Function
    If Not Chk_cboPrcCode Then Exit Function
    If Not Chk_cboMLCode Then Exit Function
    
    If Not Chk_medDueDate Then Exit Function
    
    If Not Chk_txtSpecDis Then Exit Function
    If Not chk_txtDisAmt Then Exit Function
     
    If Not Chk_cboShipCode Then Exit Function
    If Not Chk_cboRmkCode Then Exit Function
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, ITMCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tabDetailInfo.Tab = 1
                    tblDetail.Col = ITMCODE
                    tblDetail.SetFocus
                    Exit Function
                End If
            
                If Chk_NoDup2(wlCtr, waResult(wlCtr, ITMCODE), waResult(wlCtr, WHSCODE), waResult(wlCtr, LOTNO)) = False Then
                    tblDetail.Row = wlCtr - 1
                    tblDetail.Col = ITMCODE
                    tblDetail.SetFocus
                    tabDetailInfo.Tab = 1
                    Exit Function
                End If
                
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tabDetailInfo.Tab = 1
            tblDetail.Col = ITMCODE
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    Call Calc_Total
    
    InputValidation = True
    
    Exit Function
    
InputValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Private Sub cmdNew()

    Dim newForm As New frmGRV001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmGRV001
    
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
    wsFormID = "GRV001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "GR"
End Sub

Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("Default")
    tabDetailInfo.Tab = 0
    cboDocNo.SetFocus
    
End Sub

Private Sub cmdFind()
    
    Call OpenPromptForm
    
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
           If MsgBox("你是否確定儲存現時之變更而離開?", vbYesNo, gsTitle) = vbNo Then
                Call cmdCancel
           End If
           Else
                Call cmdCancel
           End If
        Case tcRefresh
            Call cmdRefresh
        Case tcPrint
            Call cmdPrint
        Case tcRevise
            Call cmdRevise
        Case tcExit
            Unload Me
    End Select
    
End Sub

Private Sub txtExcr_GotFocus()

    FocusMe txtExcr
    
End Sub

Private Sub txtExcr_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtExcr.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtExcr Then
            If Chk_KeyFld Then
            tabDetailInfo.Tab = 0
            cboSaleCode.SetFocus
            End If
        End If
    End If

End Sub

Private Function chk_txtExcr() As Boolean
    
    chk_txtExcr = False
    
    If Trim(txtExcr.Text) = "" Or Trim(To_Value(txtExcr.Text)) = 0 Then
        gsMsg = "必需輸入對換率!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtExcr.SetFocus
        Exit Function
    End If
    
    If To_Value(txtExcr.Text) > 9999.999999 Then
        gsMsg = "對換率超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtExcr.SetFocus
        Exit Function
    End If
    txtExcr.Text = Format(txtExcr.Text, gsExrFmt)
    
    chk_txtExcr = True
    
End Function

Private Sub txtExcr_LostFocus()
FocusMe txtExcr, True
End Sub

Private Sub cboRefDocNo_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboRefDocNo
    
    wsSQL = "SELECT POHDDOCNO, POHDDOCDATE , VDRCODE, VDRNAME FROM popPOHd, mstVendor "
    wsSQL = wsSQL & " WHERE POHDSTATUS = '1' "
    wsSQL = wsSQL & " AND POHDVDRID = VDRID "
    wsSQL = wsSQL & " AND POHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDPGMNO <> 'PN001' "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO "
                
    Call Ini_Combo(4, wsSQL, cboRefDocNo.Left + tabDetailInfo.Left, cboRefDocNo.Top + cboRefDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLPONO", Me.Width, Me.Height)
            
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
            If Trim(cboVdrCode.Text) = "" Then
                tabDetailInfo.Tab = 0
                cboVdrCode.SetFocus
                Exit Sub
            End If
            If Chk_KeyFld Then
                tabDetailInfo.Tab = 0
                cboSaleCode.SetFocus
            End If
    End If
    
End Sub

Private Function Chk_cboRefDocNo() As Boolean
    
Dim wsStatus As String
Dim wsPgmNo As String
    
    Chk_cboRefDocNo = False
    
    If Trim(cboRefDocNo.Text) = "" Then
        Chk_cboRefDocNo = True
        wlRefDocID = 0
        Exit Function
    End If
        
    If Chk_PoHdDocNo(cboRefDocNo, wsStatus, wsPgmNo) = True Then
        
    '    If wsStatus = "4" Then
    '        gsMsg = "文件已入數!"
    '        MsgBox gsMsg, vbOKOnly, gsTitle
    '        Exit Function
    '    End If
        
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
        
        
        If wsPgmNo = "PN001" Then
            gsMsg = "文件類別不同!不能開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboRefDocNo.SetFocus
            wlRefDocID = 0
            Exit Function
        End If
        
        
    End If
    
    Chk_cboRefDocNo = True

End Function

Private Sub cboVdrCode_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboVdrCode
    
    wsSQL = "SELECT VDRCODE, VDRNAME FROM MstVendor "
    wsSQL = wsSQL & "WHERE VDRCODE LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
    wsSQL = wsSQL & "AND VDRSTATUS = '1' "
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & "ORDER BY VDRCODE "
    Call Ini_Combo(2, wsSQL, cboVdrCode.Left + tabDetailInfo.Left, cboVdrCode.Top + cboVdrCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLVDRNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
   
End Sub

Private Sub cboVdrCode_GotFocus()
    
    Set wcCombo = cboVdrCode
    'TREtoolsbar1.ButtonEnabled(tcCusSrh) = True
    FocusMe cboVdrCode
    
End Sub

Private Sub cboVdrCode_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboVdrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_cboVdrCode() = False Then Exit Sub
        
        If wiAction = AddRec Or wsOldVdrNo <> cboVdrCode.Text Then Call Get_DefVal
            If Chk_KeyFld Then
                tabDetailInfo.Tab = 0
                cboCurr.SetFocus
            End If
    End If
    
End Sub

Private Function chk_cboVdrCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    Dim wsTel As String
    Dim wsFax As String
    
    
    chk_cboVdrCode = False
    
    If Trim(cboVdrCode) = "" Then
        gsMsg = "必需輸入供應商編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboVdrCode.SetFocus
        Exit Function
    End If
    
    If Chk_VdrCode(cboVdrCode, wlID, wsName, wsTel, wsFax) Then
        wlVdrID = wlID
        lblDspVdrName.Caption = wsName
        lblDspVdrTel.Caption = wsTel
        lblDspVdrFax.Caption = wsFax
    Else
        wlVdrID = 0
        lblDspVdrName.Caption = ""
        lblDspVdrTel.Caption = ""
        lblDspVdrFax.Caption = ""
        gsMsg = "供應商不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboVdrCode.SetFocus
        Exit Function
    End If
    
    chk_cboVdrCode = True

End Function

Private Sub Get_DefVal()
    
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsExcDesc As String
    Dim wsExcRate As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM MstVendor "
    wsSQL = wsSQL & "WHERE VDRID = " & wlVdrID
    rsDefVal.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        cboCurr.Text = ReadRs(rsDefVal, "VDRCURR")
        cboPayCode.Text = ReadRs(rsDefVal, "VDRPAYCODE")
        cboMLCode.Text = ReadRs(rsDefVal, "VDRMLCODE")
        wlSaleID = ReadRs(rsDefVal, "VDRSALEID")
        txtShipName = ReadRs(rsDefVal, "VDRSHIPTO")
        txtShipPer = ReadRs(rsDefVal, "VDRSHIPCONTACTPERSON")
        txtShipAdr1 = ReadRs(rsDefVal, "VDRSHIPADD1")
        txtShipAdr2 = ReadRs(rsDefVal, "VDRSHIPADD2")
        txtShipAdr3 = ReadRs(rsDefVal, "VDRSHIPADD3")
        txtShipAdr4 = ReadRs(rsDefVal, "VDRSHIPADD4")
        
    Else
        cboCurr.Text = ""
        cboPayCode.Text = ""
        cboMLCode.Text = ""
        wlSaleID = 0
        txtShipName = ""
        txtShipPer = ""
        txtShipAdr1 = ""
        txtShipAdr2 = ""
        txtShipAdr3 = ""
        txtShipAdr4 = ""
        
    End If
    rsDefVal.Close
    Set rsDefVal = Nothing
    
    ' get currency code description
    If getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) = True Then
        txtExcr.Text = Format(wsExcRate, gsExrFmt)
    Else
        txtExcr.Text = Format("0", gsExrFmt)
    End If

    If UCase(cboCurr) = UCase(wsBaseCurCd) Then
        txtExcr.Text = Format("1", gsExrFmt)
        txtExcr.Enabled = False
    Else
        txtExcr.Enabled = True
    End If
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    
    'get Due Date Payment Term
    medDueDate = Dsp_Date(Get_DueDte(cboPayCode, medDocDate))

End Sub

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
       ' .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = LINENO To POID
            .Columns(wiCtr).AllowSizing = True
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
                Case ITMCODE
                    .Columns(wiCtr).Width = 2500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 30
                Case ITMTYPE
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
                    '.Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 20
                  '  .Columns(wiCtr).Visible = False
                Case ITMNAME
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case PUBLISHER
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).Visible = False
                Case QTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case PRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case DisPer
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case NET
                    .Columns(wiCtr).Width = 1000
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
                Case ITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case POID
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
Dim wsITMTYPE As String
Dim wsITMNAME As String
Dim wsPub As String
Dim wdPrice As Double
Dim wdDisPer As Double
Dim wsLotNo As String
Dim wsWhsCode As String
Dim wdQty As Double
Dim wsPoId As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
                Case ITMCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdITMCODE(.Columns(ColIndex).Text, wsITMID, wsITMCODE, wsITMTYPE, wsITMNAME, wsPub, wdPrice, wdDisPer, wsWhsCode, wsLotNo, wdQty) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(LINENO).Text = wlLineNo
                .Columns(ITMID).Text = wsITMID
                .Columns(ITMNAME).Text = wsITMNAME
                .Columns(ITMTYPE).Text = wsITMTYPE
                .Columns(PUBLISHER).Text = wsPub
                .Columns(WHSCODE).Text = wsWhsCode
                .Columns(LOTNO).Text = wsLotNo
                .Columns(PRICE).Text = Format(wdPrice, gsAmtFmt)
                .Columns(QTY).Text = Format(wdQty, gsQtyFmt)
                .Columns(DisPer).Text = Format(wdDisPer, gsAmtFmt)
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(ColIndex).Text) <> wsITMCODE Then
                    .Columns(ColIndex).Text = wsITMCODE
                End If
                If Trim(.Columns(PRICE).Text) <> "" Then
                .Columns(Amt).Text = Format(To_Value(.Columns(PRICE).Text) * To_Value(.Columns(QTY).Text), gsAmtFmt)
                End If
                If Trim(txtExcr.Text) <> "" Then
                .Columns(Amtl).Text = Format(To_Value(.Columns(PRICE).Text) * To_Value(.Columns(QTY).Text) * To_Value(txtExcr.Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Dis).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Disl).Text = Format(To_Value(.Columns(Amtl).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(NET).Text = Format(To_Value(.Columns(Amt).Text) * (1 - (To_Value(.Columns(DisPer).Text) / 100)), gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) * (1 - (To_Value(.Columns(DisPer).Text) / 100)), gsAmtFmt)
                End If
        
             Case WHSCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
            '    If Chk_grdWhsCode(.Columns(ColIndex).Text) = False Then
            '            GoTo Tbl_BeforeColUpdate_Err
            '    End If
             Case LOTNO
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdLotNo(.Columns(WHSCODE).Text, .Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
            

            Case QTY, PRICE, DisPer
            
                If ColIndex = QTY Then
                        If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                        End If
                ElseIf ColIndex = PRICE Then
                        If Chk_grdUPrice(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                        End If
                
                ElseIf ColIndex = DisPer Then
                        If Chk_grdDisPer(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                        End If
                End If
                    
                If Trim(.Columns(PRICE).Text) <> "" Then
                .Columns(Amt).Text = Format(To_Value(.Columns(PRICE).Text) * To_Value(.Columns(QTY).Text), gsAmtFmt)
                End If
                If Trim(txtExcr.Text) <> "" Then
                .Columns(Amtl).Text = Format(To_Value(.Columns(PRICE).Text) * To_Value(.Columns(QTY).Text) * To_Value(txtExcr.Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Dis).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Disl).Text = Format(To_Value(.Columns(Amtl).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(NET).Text = Format(To_Value(.Columns(Amt).Text) * (1 - (To_Value(.Columns(DisPer).Text) / 100)), gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) * (1 - (To_Value(.Columns(DisPer).Text) / 100)), gsAmtFmt)
                End If
                
                Case Dis
                                
                If Trim(txtExcr.Text) <> "" Then
                .Columns(Disl).Text = Format(To_Value(.Columns(Dis).Text) * To_Value(txtExcr.Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(NET).Text = Format(To_Value(.Columns(Amt).Text) * (1 - (To_Value(.Columns(DisPer).Text) / 100)), gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) * (1 - (To_Value(.Columns(DisPer).Text) / 100)), gsAmtFmt)
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
                
            Case ITMCODE
                
                If Trim(cboRefDocNo.Text) = "" Then
                
                
                wsSQL = "SELECT ITMCODE, ITMITMTYPECODE, ITMENGNAME, ITMCHINAME "
                wsSQL = wsSQL & " FROM mstITEM, mstVdrItem "
                wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' "
                wsSQL = wsSQL & " AND VDRITEMSTATUS <> '2' "
                wsSQL = wsSQL & " AND ITMINACTIVE = 'N' "
                wsSQL = wsSQL & " AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                wsSQL = wsSQL & " AND ITMID = VDRITEMITMID "
                wsSQL = wsSQL & " AND VDRITEMCURR = '" & Set_Quote(cboCurr.Text) & "' "
                wsSQL = wsSQL & " AND VDRITEMVDRID = " & To_Value(wlVdrID) & " "
                
                If waResult.UpperBound(1) > -1 Then
                          wsSQL = wsSQL & " AND ITMCODE NOT IN ( "
                          For wiCtr = 0 To waResult.UpperBound(1)
                                wsSQL = wsSQL & " '" & Set_Quote(waResult(wiCtr, ITMCODE)) & IIf(wiCtr = waResult.UpperBound(1), "' )", "' ,")
                          Next
                End If
                wsSQL = wsSQL & " ORDER BY ITMCODE "
                
            '    wsSQL = "SELECT ITMCODE, ITMITMTYPECODE, ITMENGNAME, ITMCHINAME "
            '    wsSQL = wsSQL & "FROM mstITEM "
            '    wsSQL = wsSQL & "WHERE ITMSTATUS <> '2' "
            '    wsSQL = wsSQL & "AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
            '    wsSQL = wsSQL & " ORDER BY ITMCODE "
                
                Else
                
                wsSQL = "SELECT ITMCODE, ITMITMTYPECODE, ITMENGNAME, ITMCHINAME "
                wsSQL = wsSQL & "FROM mstITEM, popPODT "
                wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                wsSQL = wsSQL & " AND PODTDOCID = " & wlRefDocID & " "
                wsSQL = wsSQL & " AND PODTITEMID = ITMID "
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
            If wbUpdCstOnly = True Then Exit Sub
            If IsEmptyRow Then Exit Sub
            waResult.InsertRows IIf(IsNull(.Bookmark), 0, .Bookmark)
            .ReBind
            .SetFocus
            
        Case vbKeyF8        ' DELETE LINE
            KeyCode = vbDefault
            If IsNull(.Bookmark) Then Exit Sub
            If wbUpdCstOnly = True Then Exit Sub
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
                Case LINENO
                    KeyCode = vbDefault
                    .Col = ITMCODE
                Case ITMCODE
                    KeyCode = vbDefault
                    .Col = QTY
                Case QTY
                    KeyCode = vbDefault
                       .Col = PRICE
                Case PRICE
                    KeyCode = vbDefault
                    .Col = DisPer
                Case DisPer
                     KeyCode = vbKeyDown
                    .Col = ITMCODE
                Case NET
                    KeyCode = vbKeyDown
                    .Col = ITMCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            Select Case .Col
                Case ITMTYPE
                    .Col = ITMCODE
                Case ITMNAME
                    .Col = ITMTYPE
                Case QTY
                    .Col = ITMNAME
                Case PRICE
                    .Col = QTY
                Case DisPer
                    .Col = PRICE
                Case NET
                    .Col = DisPer
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case LINENO
                    .Col = ITMCODE
                Case ITMCODE
                    .Col = ITMTYPE
                Case ITMTYPE
                    .Col = ITMNAME
                Case ITMNAME
                    .Col = QTY
                Case QTY
                    .Col = PRICE
                Case PRICE
                    .Col = DisPer
                Case DisPer
                    .Col = NET
            End Select
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    Select Case tblDetail.Col
        
    '    Case Qty
    '        Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        Case QTY, PRICE, DisPer, Dis
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
            
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = ITMCODE
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case ITMCODE
                    Call Chk_grdITMCODE(.Columns(ITMCODE).Text, "", "", "", "", "", 0, 0, "", "", 0)
                Case WHSCODE
                 '   Call Chk_grdWhsCode(.Columns(WHSCODE).Text)
                 Case LOTNO
                    Call Chk_grdLotNo(.Columns(WHSCODE).Text, .Columns(LOTNO).Text)
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

Private Function Chk_grdITMCODE(inAccNo As String, outAccID As String, outAccNo As String, OutItmType As String, OutName As String, outPub As String, outPrice As Double, outDisPer As Double, outWhsCode As String, outLotNo As String, outQty As Double) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String

    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
        Exit Function
    End If
    
    If wlRefDocID = 0 Then
    
   'wsSQL = "SELECT ITMID ITEMID, ITMCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME, ITMITMTYPECODE, ITMBOTTOMPRICE PRICE, ITMUNITPRICE UPRICE, ITMCURR CURR, 0 DISPER, ITMPVDRID VDRID "
   'wsSQL = wsSQL & " FROM mstITEM "
   'wsSQL = wsSQL & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' "
   'wsSQL = wsSQL & " AND ITMSTATUS <> '2' "
    
   wsSQL = "SELECT ITMID ITEMID, ITMCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITNAME, ITMITMTYPECODE, VDRITEMCOST PRICE, 1 BALQTY, 0 DISPER, VDRITEMCURR CURR "
   wsSQL = wsSQL & " FROM mstITEM, mstVdrItem "
   wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' "
   wsSQL = wsSQL & " AND VDRITEMSTATUS <> '2' "
   wsSQL = wsSQL & " AND ITMINACTIVE = 'N' "
   wsSQL = wsSQL & " AND ITMCODE = '" & Set_Quote(inAccNo) & "' "
   wsSQL = wsSQL & " AND ITMID = VDRITEMITMID "
   wsSQL = wsSQL & " AND VDRITEMCURR = '" & Set_Quote(cboCurr.Text) & "' "
   wsSQL = wsSQL & " AND VDRITEMVDRID = " & To_Value(wlVdrID) & " "
    
    
    Else
    
    wsSQL = "SELECT PODTITEMID ITEMID, ITMCODE, PODTITEMDESC ITNAME, ITMITMTYPECODE, PODTUPRICE PRICE, ITMUNITPRICE UPRICE, POHDCURR CURR, PODTDISPER DISPER, POHDVDRID VDRID, PODTQTY - PODTSCHQTY BALQTY "
    wsSQL = wsSQL & " FROM mstITEM, popPOHD, popPODT "
    wsSQL = wsSQL & " WHERE POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND POHDDOCID = " & wlRefDocID & " "
    wsSQL = wsSQL & " AND ITMCODE = '" & Set_Quote(inAccNo) & "' "
    
    End If
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITEMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITNAME")
        OutItmType = ReadRs(rsDes, "ITMITMTYPECODE")
        outPub = ""
        outPrice = To_Value(ReadRs(rsDes, "PRICE"))

        wsCurr = ReadRs(rsDes, "CURR")
        
        outWhsCode = ""
        outLotNo = ""
        outQty = To_Value(ReadRs(rsDes, "BALQTY"))
       
        If cboCurr <> wsCurr Then
            If getExcRate(wsCurr, medDocDate, wsExcr, "") = True Then
                outPrice = NBRnd(outPrice * To_Value(wsExcr) / txtExcr, giUprDp)
            End If
        End If
       
        outDisPer = To_Value(ReadRs(rsDes, "DISPER"))
       
        Chk_grdITMCODE = True
    Else
        outAccID = ""
        OutName = ""
        OutItmType = ""
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
Private Function Chk_grdOrdItm(inPoNo As Long, inItmNo As String, inWhsCode As String, InLotNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset

    
    If To_Value(inPoNo) = 0 Then
        Chk_grdOrdItm = True
        Exit Function
    End If
    
    wsSQL = "SELECT PODTITEMID "
    wsSQL = wsSQL & " FROM mstITEM, popPOHD, popPODT "
    wsSQL = wsSQL & " WHERE POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND PODTITEMID = ITMID "
    wsSQL = wsSQL & " AND POHDDOCID = " & To_Value(inPoNo) & " "
    wsSQL = wsSQL & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSQL = wsSQL & " AND POHDSTATUS NOT IN ('2' , '3')"
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       Chk_grdOrdItm = True
    Else
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdOrdItm = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdPoNo(inPoNo As String, ByRef outPoID As String) As Boolean
    
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    
    Chk_grdPoNo = False
    
    outPoID = "0"
    
    wsSQL = "SELECT POHDDOCID, POHDDOCNO, POHDDOCDATE FROM popPOHD "
    wsSQL = wsSQL & " WHERE POHDSTATUS = '1' "
    wsSQL = wsSQL & " AND POHDDOCNO = '" & inPoNo & "' "

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        gsMsg = "沒有此訂單!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
    Set rsRcd = Nothing
        Exit Function
    End If
       
    outPoID = To_Value(ReadRs(rsRcd, "POHDDOCID"))
       
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_grdPoNo = True

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


Private Function Chk_grdLotNo(inWhs As String, inNo As String) As Boolean
    
  
    Chk_grdLotNo = False
    
    If Chk_LotEnabled(inWhs) = False Then
        Chk_grdLotNo = True
        Exit Function
    End If
    
    If wbUpdCstOnly = True Then
        Chk_grdLotNo = True
        Exit Function
    End If
    
    
    If Chk_LotB(inWhs, inNo) = False Then
         gsMsg = "不能輸入 " & inNo & " 貨架!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    If Trim(inNo) = "" Then
        gsMsg = "必需輸入貨架!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
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


Private Function Chk_grdUPrice(inCode As String) As Boolean
    
    Chk_grdUPrice = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入單價!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdUPrice = False
        Exit Function
    End If

'    If To_Value(inCode) = 0 Then
'        gsMsg = "單價必需大於零!"
'        MsgBox gsMsg, vbOKOnly, gsTitle
'        Chk_grdUPrice = False
'        Exit Function
'    End If
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
                If Trim(.Columns(ITMCODE)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, ITMCODE)) = "" And _
                   Trim(waResult(inRow, ITMNAME)) = "" And _
                   Trim(waResult(inRow, PUBLISHER)) = "" And _
                   Trim(waResult(inRow, QTY)) = "" And _
                   Trim(waResult(inRow, PRICE)) = "" And _
                   Trim(waResult(inRow, DisPer)) = "" And _
                   Trim(waResult(inRow, Amt)) = "" And _
                   Trim(waResult(inRow, Amtl)) = "" And _
                   Trim(waResult(inRow, Dis)) = "" And _
                   Trim(waResult(inRow, Disl)) = "" And _
                   Trim(waResult(inRow, NET)) = "" And _
                   Trim(waResult(inRow, Netl)) = "" And _
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
        
        
        If Chk_grdITMCODE(waResult(LastRow, ITMCODE), "", "", "", "", "", 0, 0, "", "", 0) = False Then
            .Col = ITMCODE
            .Row = LastRow
            Exit Function
        End If
        
       ' If Chk_grdWhsCode(waResult(LastRow, WHSCODE)) = False Then
       '         .Col = WHSCODE
       '         .Row = LastRow
       '         Exit Function
       ' End If
        
        If Chk_grdLotNo(waResult(LastRow, WHSCODE), waResult(LastRow, LOTNO)) = False Then
                .Col = LOTNO
                .Row = LastRow
                Exit Function
        End If
        
        
        If Chk_grdQty(waResult(LastRow, QTY)) = False Then
                .Col = QTY
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_grdUPrice(waResult(LastRow, PRICE)) = False Then
                .Col = PRICE
                .Row = LastRow
                Exit Function
        End If
        
        
        If Chk_grdDisPer(waResult(LastRow, DisPer)) = False Then
                .Col = DisPer
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_Amount(waResult(LastRow, Amt)) = False Then
            .Col = Amt
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdOrdItm(wlRefDocID, waResult(LastRow, ITMCODE), waResult(LastRow, WHSCODE), waResult(LastRow, LOTNO)) = False Then
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
        wiTotalGrs = wiTotalGrs + To_Value(waResult(wiRowCtr, Amt)) - To_Value(waResult(wiRowCtr, Dis))
       ' wiTotalDis = wiTotalDis + To_Value(waResult(wiRowCtr, Dis))
        wiTotalNet = wiTotalNet + To_Value(waResult(wiRowCtr, NET))
        wiTotalQty = wiTotalQty + To_Value(waResult(wiRowCtr, QTY))
    Next
    
    lblDspGrsAmtOrg.Caption = Format(CStr(wiTotalGrs), gsAmtFmt)
    'lblDspDisAmtOrg.Caption = Format(CStr(wiTotalDis), gsAmtFmt)
    lblDspNetAmtOrg.Caption = Format(CStr(wiTotalNet), gsAmtFmt)
    lblDspTotalQty.Caption = Format(CStr(wiTotalQty), gsQtyFmt)
    
    btnGetDisAmt_Click
    
    Calc_Total = True

End Function




Private Function cmdDel() As Boolean

    Dim wsGenDte As String
    Dim adcmdDelete As New ADODB.Command
    Dim i As Integer
    
    cmdDel = False
    
    MousePointer = vbHourglass
    
    On Error GoTo cmdDelete_Err
    
    MousePointer = vbHourglass
    wsGenDte = gsSystemDate
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Or wbReadOnly Or wbUpdCstOnly Then
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
        
    adcmdDelete.CommandText = "USP_GRV001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, wlVdrID)
    Call SetSPPara(adcmdDelete, 6, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 7, wiRevNo)
    Call SetSPPara(adcmdDelete, 8, cboCurr.Text)
    Call SetSPPara(adcmdDelete, 9, txtExcr.Text)
    Call SetSPPara(adcmdDelete, 10, "")
    
    Call SetSPPara(adcmdDelete, 11, Set_MedDate(medDueDate.Text))
    Call SetSPPara(adcmdDelete, 12, Set_MedDate(medETADate.Text))
    
    Call SetSPPara(adcmdDelete, 13, wlSaleID)
    
    Call SetSPPara(adcmdDelete, 14, cboPayCode.Text)
    Call SetSPPara(adcmdDelete, 15, cboPrcCode.Text)
    Call SetSPPara(adcmdDelete, 16, cboMLCode.Text)
    Call SetSPPara(adcmdDelete, 17, cboShipCode.Text)
    Call SetSPPara(adcmdDelete, 18, cboRmkCode.Text)
    
    Call SetSPPara(adcmdDelete, 19, txtCusPo.Text)
    Call SetSPPara(adcmdDelete, 20, txtLcNo.Text)
    Call SetSPPara(adcmdDelete, 21, txtPortNo.Text)
    
    Call SetSPPara(adcmdDelete, 22, "")
    Call SetSPPara(adcmdDelete, 23, txtSpecDis.Text)
    Call SetSPPara(adcmdDelete, 24, "")
    
    
    Call SetSPPara(adcmdDelete, 25, txtShipFrom.Text)
    Call SetSPPara(adcmdDelete, 26, txtShipTo.Text)
    Call SetSPPara(adcmdDelete, 27, txtShipVia.Text)
    Call SetSPPara(adcmdDelete, 28, txtShipPer.Text)
    Call SetSPPara(adcmdDelete, 29, txtShipName.Text)
    Call SetSPPara(adcmdDelete, 30, txtShipAdr1.Text)
    Call SetSPPara(adcmdDelete, 31, txtShipAdr2.Text)
    Call SetSPPara(adcmdDelete, 32, txtShipAdr3.Text)
    Call SetSPPara(adcmdDelete, 33, txtShipAdr4.Text)
    
    For i = 1 To 10
    Call SetSPPara(adcmdDelete, 34 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdDelete, 44, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdDelete, 45, lblDspDisAmtOrg)
    Call SetSPPara(adcmdDelete, 46, lblDspNetAmtOrg)
    Call SetSPPara(adcmdDelete, 47, wlRefDocID)
    
    Call SetSPPara(adcmdDelete, 48, wsFormID)
    
    Call SetSPPara(adcmdDelete, 49, gsUserID)
    Call SetSPPara(adcmdDelete, 50, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 51)
    wsDocNo = GetSPPara(adcmdDelete, 52)
    
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
        
        gsMsg = "你是否確定要儲存現時之作業?"
        If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
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
            Me.cboVdrCode.Enabled = False

            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            
            Me.medDueDate.Enabled = False
            Me.medETADate.Enabled = False
            
            Me.cboSaleCode.Enabled = False
            Me.cboPayCode.Enabled = False
            Me.cboPrcCode.Enabled = False
            Me.cboMLCode.Enabled = False
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
            Me.txtLcNo.Enabled = False
            Me.txtPortNo.Enabled = False
            
            Me.picRmk.Enabled = False
            Me.txtSpecDis.Enabled = False
            Me.txtDisAmt.Enabled = False
            Me.btnGetDisAmt.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
            Me.cboRefDocNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
            
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            If wiAction = AddRec Then
                Me.cboRefDocNo.Enabled = True
            Else
                Me.cboRefDocNo.Enabled = False
            End If
            
            
            Me.cboVdrCode.Enabled = True
  
            Me.medDocDate.Enabled = True
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.medDueDate.Enabled = True
            Me.medETADate.Enabled = True
            
            Me.cboSaleCode.Enabled = True
            Me.cboPayCode.Enabled = True
            Me.cboPrcCode.Enabled = True
            Me.cboMLCode.Enabled = True
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
            Me.txtLcNo.Enabled = True
            Me.txtPortNo.Enabled = True
            
            Me.txtSpecDis.Enabled = True
            Me.txtDisAmt.Enabled = True
            Me.btnGetDisAmt.Enabled = True
            
            Me.picRmk.Enabled = True
            
            
            If wiAction <> AddRec Then
                Me.tblDetail.Enabled = True
            End If
            
       
            
    End Select
End Sub

Private Sub GetNewKey()
    Dim Newfrm As New frmKeyInput
    
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    With Newfrm
    
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "IVHDDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub


Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim wsSQL As String
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "Doc No."
    vFilterAry(1, 2) = "IVHDDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "IVHDDocDate"
    
    vFilterAry(3, 1) = "Vendor #"
    vFilterAry(3, 2) = "VdrCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "IVHDDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "IVHDDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Vendor #"
    vAry(3, 2) = "VdrCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Vendor Name"
    vAry(4, 2) = "VdrName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT popGRHD.GRHDDocNo, popGRHD.GRHDDocDate, MstVendor.VdrCode,  MstVendor.VdrName "
        wsSQL = wsSQL + "FROM MstVendor, popGRHD "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE popGRHD.GRHDStatus = '1' And popGRHD.GRHDVdrID = MstVendor.VdrID "
        .sBindOrderSQL = "ORDER BY popGRHD.GRHDDocNo"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboDocNo Then
        cboDocNo = Trim(frmShareSearch.Tag)
        cboDocNo.SetFocus
        SendKeys "{Enter}"
    End If
    Unload frmShareSearch
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
        
        tabDetailInfo.Tab = 0
        cboPayCode.SetFocus
       
    End If
    
End Sub

Private Sub cboSaleCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSaleCode
    
    wsSQL = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' "
    wsSQL = wsSQL & " AND SaleType = 'S' "
    wsSQL = wsSQL & "AND SaleStatus = '1' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSQL, cboSaleCode.Left + tabDetailInfo.Left, cboSaleCode.Top + cboSaleCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSALECOD", Me.Width, Me.Height)
    
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
        tabDetailInfo.Tab = 0
        cboSaleCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_Salesman(cboSaleCode, wlSaleID, wsDesc) = False Then
        gsMsg = "沒有此營業員!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboSaleCode.SetFocus
        lblDspSaleDesc = ""
       Exit Function
    End If
    
    lblDspSaleDesc = wsDesc
    
    Chk_cboSaleCode = True
    
End Function


Private Sub cboPayCode_GotFocus()
    FocusMe cboPayCode
End Sub

Private Sub cboPayCode_LostFocus()
FocusMe cboPayCode, True
End Sub


Private Sub cboPayCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboPayCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPayCode = False Then
                Exit Sub
        End If
        
        If wsOldPayCd <> cboPayCode.Text Then
            medDueDate = Dsp_Date(Get_DueDte(cboPayCode, medDocDate))
            wsOldPayCd = cboPayCode.Text
        End If
        
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
       
    End If
    
End Sub

Private Sub cboPayCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboPayCode
    
    wsSQL = "SELECT PAYCODE, PAYDESC FROM mstPayTerm WHERE PAYCODE LIKE '%" & IIf(cboPayCode.SelLength > 0, "", Set_Quote(cboPayCode.Text)) & "%' "
    wsSQL = wsSQL & "AND PAYSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY PAYCODE "
    Call Ini_Combo(2, wsSQL, cboPayCode.Left + tabDetailInfo.Left, cboPayCode.Top + cboPayCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLPAYCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboPayCode() As Boolean
Dim wsDesc As String

    Chk_cboPayCode = False
     
    If Trim(cboPayCode.Text) = "" Then
        gsMsg = "必需輸入付款條款!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPayCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_PayTerm(cboPayCode, wsDesc) = False Then
        gsMsg = "沒有此付款條款!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPayCode.SetFocus
        lblDspPayDesc = ""
       Exit Function
    End If
    
    lblDspPayDesc = wsDesc
    
    Chk_cboPayCode = True
    
End Function


Private Sub cboPrcCode_GotFocus()
    FocusMe cboPrcCode
End Sub

Private Sub cboPrcCode_LostFocus()
    FocusMe cboPrcCode, True
End Sub


Private Sub cboPrcCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboPrcCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboPrcCode = False Then
                Exit Sub
        End If

        
     '  txtPortNo = Get_TableInfo("MstPriceTerm", "PrcCode = '" & Set_Quote(cboPrcCode.Text) & "'", "PricePort")
        
        tabDetailInfo.Tab = 0
        medDueDate.SetFocus
       
    End If
    
End Sub

Private Sub cboPrcCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboPrcCode
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass WHERE MLCode LIKE '%" & IIf(cboPrcCode.SelLength > 0, "", Set_Quote(cboPrcCode.Text)) & "%' "
    wsSQL = wsSQL & "AND MLTYPE = 'P' "
    wsSQL = wsSQL & "AND MLSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboPrcCode.Left + tabDetailInfo.Left, cboPrcCode.Top + cboPrcCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboPrcCode() As Boolean
Dim wsDesc As String

    Chk_cboPrcCode = False
     
    If Trim(cboPrcCode.Text) = "" Then
        gsMsg = "必需輸入會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    
    If Chk_MClass(cboPrcCode, "P", wsDesc) = False Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
        lblDspPrcDesc = ""
       Exit Function
    End If
    
    lblDspPrcDesc = wsDesc
    Chk_cboPrcCode = True
    
End Function

Private Sub cboMLCode_GotFocus()
    FocusMe cboMLCode
End Sub

Private Sub cboMLCode_LostFocus()
    FocusMe cboMLCode, True
End Sub


Private Sub cboMLCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboMLCode = False Then
                Exit Sub
        End If
        
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
       
    End If
    
End Sub

Private Sub cboMLCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboMLCode
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass WHERE MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSQL = wsSQL & "AND MLTYPE = 'R' "
    wsSQL = wsSQL & "AND MLSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboMLCode.Left + tabDetailInfo.Left, cboMLCode.Top + cboMLCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboMLCode() As Boolean
    Dim wsDesc As String

    Chk_cboMLCode = False
     
    If Trim(cboMLCode.Text) = "" Then
        gsMsg = "必需輸入會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboMLCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_MClass(cboMLCode, "R", wsDesc) = False Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboMLCode.SetFocus
        lblDspMLDesc = ""
       Exit Function
    End If
    
    lblDspMLDesc = wsDesc
    
    Chk_cboMLCode = True
    
End Function



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
        txtCusPo.SetFocus
       
    End If
    
End Sub

Private Sub txtShipVia_LostFocus()
        FocusMe txtShipVia, True
End Sub

Private Sub txtCusPo_GotFocus()
        FocusMe txtCusPo
End Sub

Private Sub txtCusPo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtCusPo, 15, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtLcNo.SetFocus
       
    End If
    
End Sub

Private Sub txtCusPo_LostFocus()
        FocusMe txtCusPo, True
End Sub

Private Sub txtLcNo_GotFocus()
        FocusMe txtLcNo
End Sub

Private Sub txtLcNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtLcNo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtPortNo.SetFocus
       
    End If
    
End Sub

Private Sub txtLcNo_LostFocus()
        FocusMe txtLcNo, True
End Sub

Private Sub txtPortNo_GotFocus()
        FocusMe txtPortNo
End Sub

Private Sub txtPortNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtPortNo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        If Chk_KeyFld = True Then
            tabDetailInfo.Tab = 1
            tblDetail.Col = ITMCODE
            tblDetail.SetFocus
        End If
       
    End If
    
End Sub

Private Sub txtPortNo_LostFocus()
        FocusMe txtPortNo, True
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
        
        
        tabDetailInfo.Tab = 2
        txtShipPer.SetFocus
    End If
    
End Sub

Private Sub cboShipCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboShipCode
    
    wsSQL = "SELECT ShipCode, ShipName, ShipPer FROM mstShip WHERE ShipCode LIKE '%" & IIf(cboShipCode.SelLength > 0, "", Set_Quote(cboShipCode.Text)) & "%' "
    wsSQL = wsSQL & "AND ShipSTATUS = '1' "
    wsSQL = wsSQL & "AND ShipCardID = " & wlVdrID & " "
    wsSQL = wsSQL & "ORDER BY ShipCode "
    Call Ini_Combo(3, wsSQL, cboShipCode.Left + tabDetailInfo.Left, cboShipCode.Top + cboShipCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSHIPCOD", Me.Width, Me.Height)
    
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
        tabDetailInfo.Tab = 2
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
        
        
        tabDetailInfo.Tab = 2
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
        
        
        tabDetailInfo.Tab = 2
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
        
        
        tabDetailInfo.Tab = 2
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
        
        
        tabDetailInfo.Tab = 2
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
        
        
        tabDetailInfo.Tab = 2
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
        
        
        If Chk_KeyFld = True Then
        tabDetailInfo.Tab = 2
        cboRmkCode.SetFocus
        End If
        
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
    Call Ini_Combo(1, wsSQL, cboRmkCode.Left + tabDetailInfo.Left, cboRmkCode.Top + cboRmkCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLRMKCOD", Me.Width, Me.Height)
    
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
    
    Call chk_InpLen(txtRmk(Index), 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        If Index = 10 Then
            tabDetailInfo.Tab = 0
            cboVdrCode.SetFocus
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
    Dim wsCurRecLn As String
    Dim wsCurRecLn2 As String
    Dim wsCurRecLn3 As String
    
    Chk_NoDup = False
    
    wsCurRecLn = tblDetail.Columns(ITMCODE)
    wsCurRecLn2 = tblDetail.Columns(WHSCODE)
    wsCurRecLn3 = tblDetail.Columns(LOTNO)
    
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRecLn = waResult(wlCtr, ITMCODE) And _
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

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
    

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
           If wbUpdCstOnly = True Then Exit Sub
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
            If wbUpdCstOnly = True Then Exit Sub
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
                .Buttons(tcRevise).Enabled = False
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
                .Buttons(tcRevise).Enabled = False
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
                .Buttons(tcRevise).Enabled = True
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
                .Buttons(tcRevise).Enabled = False
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
    wsSQL = "EXEC usp_RPTGRV002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'GR', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & String(10, "z") & "', "
    wsSQL = wsSQL & "'" & "000000" & "', "
    wsSQL = wsSQL & "'" & "999999" & "', "
    wsSQL = wsSQL & "'" & "%" & "', "
    wsSQL = wsSQL & "'N', "
    wsSQL = wsSQL & gsLangID
    
    
    If gsLangID = "2" Then
    wsRptName = "C" + "RPTGRV002"
    Else
    wsRptName = "RPTGRV002"
    End If
    
    NewfrmPrint.ReportID = "GRV002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "GRV002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdRefresh()
Dim wiCtr As Integer
Dim wsITMID As String
Dim wdDisPer As Double
Dim wdNewDisPer As Double

    
  If waResult.UpperBound(1) >= 0 Then
        
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                wsITMID = waResult(wiCtr, ITMID)
                wdDisPer = waResult(wiCtr, DisPer)
                'wdNewDisPer = Get_SaleDiscount(cboNatureCode.Text, wlVdrID, wsITMID)
                'If wdDisPer <> wdNewDisPer Then
                '    waResult(wiCtr, DisPer) = Format(wdNewDisPer, gsAmtFmt)
                '    waResult(wiCtr, Dis) = Format(To_Value(waResult(wiCtr, Amt)) * To_Value(waResult(wiCtr, DisPer)) / 100, gsAmtFmt)
                '    waResult(wiCtr, Disl) = Format(To_Value(waResult(wiCtr, Amtl)) * To_Value(waResult(wiCtr, DisPer)) / 100, gsAmtFmt)
                '    waResult(wiCtr, Net) = Format(To_Value(waResult(wiCtr, Amt)) - To_Value(waResult(wiCtr, Dis)), gsAmtFmt)
                '    waResult(wiCtr, Netl) = Format(To_Value(waResult(wiCtr, Amtl)) - To_Value(waResult(wiCtr, Disl)), gsAmtFmt)
                'End If
            End If
        Next
   
   
   
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
    
    wsSQL = "SELECT POHDDOCID, POHDDOCNO, POHDVDRID, VDRID, VDRCODE, VDRNAME, VDRTEL, VDRFAX, "
    wsSQL = wsSQL & "POHDDOCDATE, POHDREVNO, POHDCURR, POHDEXCR, POHDETADATE, "
    wsSQL = wsSQL & "POHDDUEDATE, POHDPRCCODE, POHDSALEID, POHDMLCODE, "
    wsSQL = wsSQL & "POHDCUSPO, POHDLCNO, POHDREFNO, POHDSHIPPER, POHDSHIPFROM, POHDSHIPTO, POHDSHIPVIA, POHDSHIPNAME, "
    wsSQL = wsSQL & "POHDSHIPCODE, POHDSHIPADR1,  POHDSHIPADR2,  POHDSHIPADR3,  POHDSHIPADR4, "
    wsSQL = wsSQL & "POHDRMKCODE, POHDRMK1,  POHDRMK2,  POHDRMK3,  POHDRMK4, POHDRMK5, "
    wsSQL = wsSQL & "POHDRMK6,  POHDRMK7,  POHDRMK8,  POHDRMK9, POHDRMK10, "
    wsSQL = wsSQL & "POHDGRSAMT , POHDGRSAMTL, POHDDISAMT, POHDDISAMTL, POHDNETAMT, POHDNETAMTL, "
    wsSQL = wsSQL & "PODTITEMID, ITMCODE, PODTWHSCODE, PODTLOTNO, ITMITMTYPECODE, PODTITEMDESC ITNAME, ITMPUBLISHER,  PODTQTY - PODTSCHQTY BALQTY, PODTUPRICE, PODTDISPER, PODTAMT, PODTAMTL, PODTDIS, PODTDISL, PODTNET, PODTNETL, "
    wsSQL = wsSQL & "PODTID, PODTDOCLINE, POHDSPECDIS "
    wsSQL = wsSQL & "FROM  popPOHD, popPODT, MstVendor, mstITEM "
    wsSQL = wsSQL & "WHERE POHDDOCNO = '" & Set_Quote(cboRefDocNo) & "' "
    wsSQL = wsSQL & "AND POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & "AND POHDVDRID = VDRID "
    wsSQL = wsSQL & "AND PODTITEMID = ITMID "
    wsSQL = wsSQL & "ORDER BY PODTDOCLINE "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    wsOldRefDocNo = cboRefDocNo.Text
    wlRefDocID = ReadRs(rsRcd, "POHDDOCID")
    wlVdrID = ReadRs(rsRcd, "VDRID")
    cboVdrCode.Text = ReadRs(rsRcd, "VDRCODE")
     wsOldVdrNo = cboVdrCode
    lblDspVdrName.Caption = ReadRs(rsRcd, "VDRNAME")
    lblDspVdrTel.Caption = ReadRs(rsRcd, "VDRTEL")
    lblDspVdrFax.Caption = ReadRs(rsRcd, "VDRFAX")
    cboCurr.Text = ReadRs(rsRcd, "POHDCURR")
    txtExcr.Text = Format(ReadRs(rsRcd, "POHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsRcd, "POHDDUEDATE"))
    medETADate.Text = Dsp_MedDate(ReadRs(rsRcd, "POHDETADATE"))
    
    wlSaleID = To_Value(ReadRs(rsRcd, "POHDSALEID"))
    
    cboPayCode = ReadRs(rsRcd, "POHDPAYCODE")
    cboPrcCode = ReadRs(rsRcd, "POHDPRCCODE")
    cboMLCode = ReadRs(rsRcd, "POHDMLCODE")
    cboShipCode = ReadRs(rsRcd, "POHDSHIPCODE")
    cboRmkCode = ReadRs(rsRcd, "POHDRMKCODE")
    
    txtCusPo = ReadRs(rsRcd, "POHDCUSPO")
    txtLcNo = ReadRs(rsRcd, "POHDLCNO")
    txtPortNo = ReadRs(rsRcd, "POHDREFNO")
    
    txtSpecDis = Format(ReadRs(rsRcd, "POHDSPECDIS"), gsExrFmt)
    
    
     
    txtShipFrom = ReadRs(rsRcd, "POHDSHIPFROM")
    txtShipTo = ReadRs(rsRcd, "POHDSHIPTO")
    txtShipVia = ReadRs(rsRcd, "POHDSHIPVIA")
    txtShipName = ReadRs(rsRcd, "POHDSHIPNAME")
    txtShipPer = ReadRs(rsRcd, "POHDSHIPPER")
    txtShipAdr1 = ReadRs(rsRcd, "POHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsRcd, "POHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsRcd, "POHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsRcd, "POHDSHIPADR4")
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsRcd, "POHDRMK" & i)
    Next i
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    wsOldVdrNo = cboVdrCode
    wsOldCurCd = cboCurr
    wsOldShipCd = cboShipCode
    wsOldRmkCd = cboRmkCode
    wsOldPayCd = cboPayCode
    
    
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, POID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             
           '   wdBalQty = Get_PoBalQty(wsTrnCd, 0, ReadRs(rsRcd, "POHDDOCID"), ReadRs(rsRcd, "PODTITEMID"), ReadRs(rsRcd, "PODTWHSCODE"), ReadRs(rsRcd, "PODTLOTNO"))

             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsRcd, "PODTDOCLINE")
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), ITMTYPE) = ReadRs(rsRcd, "ITMITMTYPECODE")
             waResult(.UpperBound(1), ITMNAME) = ReadRs(rsRcd, "ITNAME")
             waResult(.UpperBound(1), WHSCODE) = ReadRs(rsRcd, "PODTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsRcd, "PODTLOTNO")
             waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsRcd, "ITMPUBLISHER")
            ' waResult(.UpperBound(1), Qty) = Format(ReadRs(rsRcd, "BALQTY"), gsQtyFmt)
             waResult(.UpperBound(1), QTY) = Format(ReadRs(rsRcd, "BALQTY"), gsAmtFmt)
             waResult(.UpperBound(1), PRICE) = Format(ReadRs(rsRcd, "PODTUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsRcd, "PODTDISPER"), gsAmtFmt)
             waResult(.UpperBound(1), Amt) = Format(To_Value(ReadRs(rsRcd, "PODTUPRICE")) * To_Value(ReadRs(rsRcd, "BALQTY")), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(To_Value(ReadRs(rsRcd, "PODTUPRICE")) * To_Value(ReadRs(rsRcd, "BALQTY")) * To_Value(txtExcr.Text), gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(waResult(.UpperBound(1), Amt) * To_Value(ReadRs(rsRcd, "PODTDISPER")) / 100, gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(waResult(.UpperBound(1), Amtl) * To_Value(ReadRs(rsRcd, "PODTDISPER")) / 100, gsAmtFmt)
             waResult(.UpperBound(1), NET) = Format(waResult(.UpperBound(1), Amt) * (1 - To_Value(ReadRs(rsRcd, "PODTDISPER")) / 100), gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(waResult(.UpperBound(1), Amtl) * (1 - To_Value(ReadRs(rsRcd, "PODTDISPER")) / 100), gsAmtFmt)
             waResult(.UpperBound(1), ITMID) = ReadRs(rsRcd, "PODTITEMID")
             waResult(.UpperBound(1), POID) = ReadRs(rsRcd, "POHDDOCID")
             
             rsRcd.MoveNext
         Loop
         
         wlLineNo = waResult(.UpperBound(1), LINENO) + 1
         
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsRcd.Close
    
    Set rsRcd = Nothing
    
    Call Calc_Total
    
    Get_RefDoc = True
    
End Function

Private Function Chk_NoDup2(ByRef inRow As Long, ByVal wsCurRecLn As String, ByVal wsCurRecLn2 As String, ByVal wsCurRecLn3 As String) As Boolean
    Dim wlCtr As Long
     
    Chk_NoDup2 = False
    
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           If wsCurRecLn = waResult(wlCtr, ITMCODE) And _
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


Private Sub cmdRevise()

     
    On Error GoTo cmdRevise_Err
    
    
    If wbReadOnly Or wbUpdCstOnly Then
        gsMsg = "記錄已被鎖定, 不能改正此檔案!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Sub
    End If
    
    gsMsg = "你是否確認要改正此檔案?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       Exit Sub
    End If
    
    wiAction = RevRec
    
    If cmdSave = True Then
       cboDocNo.Text = wsDocNo
       Call Ini_Scr_AfrKey
    End If
    
    Exit Sub
    
cmdRevise_Err:
    MsgBox Err.Description
    
End Sub


Private Sub txtSpecDis_GotFocus()
    FocusMe txtSpecDis
End Sub

Private Sub txtSpecDis_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtSpecDis.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtSpecDis Then
            tabDetailInfo.Tab = 0
            txtDisAmt.SetFocus
            
            btnGetDisAmt_Click
            
        End If
    End If
End Sub

Private Sub txtSpecDis_LostFocus()
    txtSpecDis = Format(To_Value(txtSpecDis), gsAmtFmt)
    FocusMe txtSpecDis, True
End Sub

Private Function Chk_txtSpecDis() As Boolean
    
    Chk_txtSpecDis = False
    
    If Trim(txtSpecDis.Text) = "" Then
        gsMsg = "必需輸入特別折扣!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtSpecDis.SetFocus
        Exit Function
    End If
    
    If To_Value(txtSpecDis.Text) > 1 Then
        gsMsg = "特別折扣超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtSpecDis.SetFocus
        Exit Function
    End If
    
    txtSpecDis.Text = Format(txtSpecDis.Text, gsExrFmt)
    
    Chk_txtSpecDis = True
    
End Function


Private Sub txtDisAmt_GotFocus()

    FocusMe txtDisAmt
    
End Sub

Private Sub txtDisAmt_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtDisAmt.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
      '  If chk_txtDisAmt Then
            tabDetailInfo.Tab = 0
            txtShipName.SetFocus
            
            btnGetDisAmt_Click
            
       ' End If
    End If

End Sub

Private Function chk_txtDisAmt() As Boolean
    
    chk_txtDisAmt = False
    
    
    If To_Value(txtDisAmt.Text) < 0 Then
        gsMsg = "錯誤!一定大於零"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtDisAmt.SetFocus
        Exit Function
    End If
    txtDisAmt.Text = Format(txtDisAmt.Text, gsAmtFmt)
    
    chk_txtDisAmt = True
    
End Function

Private Sub txtDisAmt_LostFocus()
txtDisAmt.Text = Format(txtDisAmt.Text, gsAmtFmt)
FocusMe txtDisAmt, True
End Sub

Private Sub Ini_LockGrid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = False
        .AllowAddNew = False
        .AllowDelete = False
        
        For wiCtr = LINENO To POID
            .Columns(wiCtr).Locked = True
            
            Select Case wiCtr
                Case LINENO
                    .Columns(wiCtr).Locked = True
                Case ITMCODE
                    .Columns(wiCtr).Button = False
                Case ITMTYPE
                    .Columns(wiCtr).Button = False
                Case ITMNAME
                    .Columns(wiCtr).Locked = True
                Case NET
                    .Columns(wiCtr).Locked = True
                Case PRICE
                    .Columns(wiCtr).Locked = False
             End Select
        Next
       ' .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
Private Sub Ini_UnLockGrid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .AllowAddNew = True
        .AllowDelete = True
        
        For wiCtr = LINENO To POID
            .Columns(wiCtr).Locked = False
            
            Select Case wiCtr
                Case LINENO
                    .Columns(wiCtr).Locked = True
                Case ITMCODE
                    .Columns(wiCtr).Button = True
                Case ITMTYPE
                    .Columns(wiCtr).Button = True
                Case ITMNAME
                    .Columns(wiCtr).Locked = True
                Case NET
                    .Columns(wiCtr).Locked = True
             End Select
             
        Next
       ' .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub
