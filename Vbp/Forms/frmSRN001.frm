VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSRN001 
   Caption         =   "銷售退貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmSRN001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11400
      OleObjectBlob   =   "frmSRN001.frx":030A
      TabIndex        =   36
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
            Picture         =   "frmSRN001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSRN001.frx":66AD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   37
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
      TabIndex        =   38
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
      TabPicture(0)   =   "frmSRN001.frx":69C9
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraKey"
      Tab(0).Control(1)=   "fraCode"
      Tab(0).Control(2)=   "fraInfo"
      Tab(0).Control(3)=   "FraDate"
      Tab(0).Control(4)=   "cboSaleCode"
      Tab(0).Control(5)=   "cboPayCode"
      Tab(0).Control(6)=   "cboCusCode"
      Tab(0).Control(7)=   "cboDocNo"
      Tab(0).Control(8)=   "cboPrcCode"
      Tab(0).Control(9)=   "cboCurr"
      Tab(0).Control(10)=   "cboMLCode"
      Tab(0).Control(11)=   "cboRefDocNo"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmSRN001.frx":69E5
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
      Tab(1).Control(9)=   "btnSOLST"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Item Information"
      TabPicture(2)   =   "frmSRN001.frx":6A01
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraRmk"
      Tab(2).Control(1)=   "cboRmkCode"
      Tab(2).Control(2)=   "fraShip"
      Tab(2).Control(3)=   "cboShipCode"
      Tab(2).ControlCount=   4
      Begin VB.ComboBox cboRefDocNo 
         Height          =   300
         Left            =   -73200
         TabIndex        =   3
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox cboMLCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   89
         Top             =   3540
         Width           =   2370
      End
      Begin VB.ComboBox cboCurr 
         Height          =   300
         Left            =   -65520
         TabIndex        =   86
         Top             =   1130
         Width           =   1335
      End
      Begin VB.ComboBox cboPrcCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   8
         Top             =   3180
         Width           =   2370
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   18
         Top             =   480
         Width           =   2010
      End
      Begin VB.Frame fraShip 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   63
         Top             =   120
         Width           =   11535
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1455
            Left            =   1680
            ScaleHeight     =   1395
            ScaleWidth      =   9555
            TabIndex        =   64
            Top             =   1440
            Width           =   9615
            Begin VB.TextBox txtShipAdr4 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   24
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1080
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr3 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   23
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   720
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr2 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   22
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr1 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   21
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   5865
            End
         End
         Begin VB.TextBox txtShipName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   20
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1080
            Width           =   4305
         End
         Begin VB.TextBox txtShipPer 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   19
            Text            =   "01234567890123457890"
            Top             =   720
            Width           =   4305
         End
         Begin VB.Label lblShipCode 
            Caption         =   "SHIPCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblShipName 
            Caption         =   "SHIPNAME"
            Height          =   240
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   1380
         End
         Begin VB.Label lblShipPer 
            Caption         =   "SHIPPER"
            Height          =   240
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblShipAdr 
            Caption         =   "SHIPADR"
            Height          =   240
            Left            =   120
            TabIndex        =   65
            Top             =   1440
            Width           =   1500
         End
      End
      Begin VB.ComboBox cboRmkCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   25
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
      Begin VB.ComboBox cboCusCode 
         Height          =   300
         Left            =   -69720
         TabIndex        =   4
         Top             =   780
         Width           =   1935
      End
      Begin VB.CommandButton btnSOLST 
         Caption         =   "ITEMPRICE"
         Height          =   675
         Left            =   10200
         TabIndex        =   16
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cboPayCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   7
         Top             =   2820
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
         Height          =   615
         Left            =   -74880
         TabIndex        =   46
         Top             =   4080
         Width           =   3975
         Begin MSMask.MaskEdBox medDueDate 
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDueDate 
            Caption         =   "DUEDATE"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1545
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   39
         Top             =   4680
         Width           =   11655
         Begin VB.TextBox txtLcNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   14
            Text            =   "0123456789012345789"
            Top             =   1680
            Width           =   5265
         End
         Begin VB.TextBox txtPortNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   15
            Text            =   "0123456789012345789"
            Top             =   2040
            Width           =   5265
         End
         Begin VB.TextBox txtCusPo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   13
            Text            =   "0123456789012345789"
            Top             =   1320
            Width           =   5265
         End
         Begin VB.TextBox txtShipTo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   11
            Text            =   "0123456789012345789"
            Top             =   600
            Width           =   5265
         End
         Begin VB.TextBox txtShipVia 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   12
            Text            =   "0123456789012345789"
            Top             =   960
            Width           =   5265
         End
         Begin VB.TextBox txtShipFrom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   10
            Text            =   "0123456789012345789"
            Top             =   240
            Width           =   5265
         End
         Begin VB.Label lblLcNo 
            Caption         =   "LCNO"
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   1680
            Width           =   2100
         End
         Begin VB.Label lblPortNo 
            Caption         =   "PORTNO"
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   2040
            Width           =   2100
         End
         Begin VB.Label lblCusPo 
            Caption         =   "CUSPO"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Width           =   2100
         End
         Begin VB.Label lblShipTo 
            Caption         =   "SHIPTO"
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   2100
         End
         Begin VB.Label lblShipVia 
            Caption         =   "SHIPVIA"
            Height          =   240
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   2100
         End
         Begin VB.Label lblShipFrom 
            Caption         =   "SHIPFROM"
            Height          =   240
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   2100
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   6855
         Left            =   120
         OleObjectBlob   =   "frmSRN001.frx":6A1D
         TabIndex        =   17
         Top             =   840
         Width           =   11535
      End
      Begin VB.Frame fraCode 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   56
         Top             =   1980
         Width           =   11655
         Begin VB.Label lblDspMLDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   88
            Top             =   1560
            Width           =   7335
         End
         Begin VB.Label lblMlCode 
            Caption         =   "MLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   87
            Top             =   1620
            Width           =   1545
         End
         Begin VB.Label lblPrcCode 
            Caption         =   "PRCCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   62
            Top             =   1260
            Width           =   1545
         End
         Begin VB.Label lblDspPrcDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   61
            Top             =   1200
            Width           =   7335
         End
         Begin VB.Label lblPayCode 
            Caption         =   "PAYCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   900
            Width           =   1545
         End
         Begin VB.Label lblDspPayDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   59
            Top             =   840
            Width           =   7335
         End
         Begin VB.Label lblSaleCode 
            Caption         =   "SALECODE"
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label lblDspSaleDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   57
            Top             =   480
            Width           =   7335
         End
      End
      Begin VB.Frame fraRmk 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   69
         Top             =   3360
         Width           =   11535
         Begin VB.PictureBox picRmk 
            BackColor       =   &H80000009&
            Height          =   3495
            Left            =   1680
            ScaleHeight     =   3435
            ScaleWidth      =   9555
            TabIndex        =   70
            Top             =   600
            Width           =   9615
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   27
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   26
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   28
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   690
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   31
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1740
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   29
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1035
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   30
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1395
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   32
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2085
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   8
               Left            =   0
               TabIndex        =   33
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2430
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   9
               Left            =   0
               TabIndex        =   34
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2775
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   10
               Left            =   0
               TabIndex        =   35
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   3120
               Width           =   7545
            End
         End
         Begin VB.Label lblRmkCode 
            Caption         =   "RMKCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label lblRmk 
            Caption         =   "RMK"
            Height          =   240
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   1500
         End
      End
      Begin VB.Frame fraKey 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   73
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtRevNo 
            Height          =   324
            Left            =   5160
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "12345678901234567890"
            Top             =   300
            Width           =   408
         End
         Begin VB.TextBox txtExcr 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   9360
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1380
            Width           =   1335
         End
         Begin MSMask.MaskEdBox medDocDate 
            Height          =   285
            Left            =   9360
            TabIndex        =   2
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
            Caption         =   "CUSCODE"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label lblCusCode 
            Caption         =   "CUSCODE"
            Height          =   255
            Left            =   3840
            TabIndex        =   85
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
            TabIndex        =   84
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblRevNo 
            Caption         =   "REVNO"
            Height          =   255
            Left            =   3840
            TabIndex        =   83
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblDocDate 
            Caption         =   "DOCDATE"
            Height          =   255
            Left            =   7365
            TabIndex        =   82
            Top             =   720
            Width           =   1680
         End
         Begin VB.Label lblDspCusName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   81
            Top             =   1020
            Width           =   5535
         End
         Begin VB.Label LblCurr 
            Caption         =   "CURR"
            Height          =   255
            Left            =   7365
            TabIndex        =   80
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lblExcr 
            Caption         =   "EXCR"
            Height          =   255
            Left            =   7365
            TabIndex        =   79
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label lblDspCusTel 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   78
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lblCusName 
            Caption         =   "CUSNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblDspCusFax 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   5160
            TabIndex        =   76
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lblCusFax 
            Caption         =   "CUSFAX"
            Height          =   255
            Left            =   3840
            TabIndex        =   75
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCusTel 
            Caption         =   "CUSTEL"
            Height          =   255
            Left            =   120
            TabIndex        =   74
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
Attribute VB_Name = "frmSRN001"
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
Private wsOldCurCd As String
Private wsOldShipCd As String
Private wsOldRmkCd As String
Private wsOldPayCd As String
Private wbReadOnly As Boolean
Private wgsTitle As String
Private wsOldRefDocNo As String

Private Const LINENO = 0
Private Const SONO = 1
Private Const BOOKCODE = 2
Private Const BARCODE = 3
Private Const WhsCode = 4
Private Const LOTNO = 5
Private Const BOOKNAME = 6
Private Const PUBLISHER = 7
Private Const Qty = 8
Private Const Price = 9
Private Const DisPer = 10
Private Const Dis = 11
Private Const Amt = 12
Private Const Net = 13
Private Const Netl = 14
Private Const Disl = 15
Private Const Amtl = 16
Private Const BOOKID = 17
Private Const SOID = 18

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
Private Const wsKeyType = "soaSRHd"
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
    Call SetDateMask(medDueDate)
    
    medDocDate = gsSystemDate
    medDueDate = gsSystemDate
    
    wsOldCusNo = ""
    wsOldCurCd = ""
    wsOldShipCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""
    wsOldRefDocNo = ""
    
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

Private Sub btnSOLST_Click()
    If wiAction <> AddRec And wiAction <> CorRec Then Exit Sub
    
    frmSOLST.InDocID = wlKey
    frmSOLST.inTrnCd = wsTrnCd
    frmSOLST.InCurr = cboCurr.Text
    frmSOLST.inExcr = txtExcr.Text
    frmSOLST.InCusID = wlCusID
    frmSOLST.InvDoc = waResult
    frmSOLST.InLineNo = wlLineNo
    frmSOLST.Show vbModal
    waResult.ReDim 0, frmSOLST.InvDoc.UpperBound(1), LINENO, SOID
    Set waResult = frmSOLST.InvDoc
    wlLineNo = frmSOLST.InLineNo
    Unload frmSOLST
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    Call Calc_Total
End Sub

Private Sub cboCurr_GotFocus()
    FocusMe cboCurr
End Sub

Private Sub cboCurr_LostFocus()
    FocusMe cboCurr, True
End Sub

Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
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
    
    Dim wsSql As String
    Dim wsCtlDte As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCurr
    
    wsCtlDte = IIf(Trim(medDocDate.Text) = "" Or Trim(medDocDate.Text) = "/  /", gsSystemDate, medDocDate.Text)
    wsSql = "SELECT EXCCURR, EXCDESC FROM mstEXCHANGERATE WHERE EXCCURR LIKE '%" & IIf(cboCurr.SelLength > 0, "", Set_Quote(cboCurr.Text)) & "%' "
    wsSql = wsSql & " AND EXCMN = '" & To_Value(Format(wsCtlDte, "MM")) & "' "
    wsSql = wsSql & " AND EXCYR = '" & Set_Quote(Format(wsCtlDte, "YYYY")) & "' "
    wsSql = wsSql & " AND EXCSTATUS = '1' "
    wsSql = wsSql & "ORDER BY EXCCURR "
    Call Ini_Combo(2, wsSql, cboCurr.Left + tabDetailInfo.Left, cboCurr.Top + cboCurr.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLCURCOD", Me.Width, Me.Height)
    
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
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboDocNo
  
    wsSql = "SELECT SRHDDOCNO, CUSCODE, SRHDDOCDATE "
    wsSql = wsSql & " FROM soaSRHD, MstCustomer "
    wsSql = wsSql & " WHERE SRHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSql = wsSql & " AND SRHDCUSID  = CUSID "
    wsSql = wsSql & " AND SRHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY SRHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNo.Left + tabDetailInfo.Left, cboDocNo.Top + cboDocNo.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLDOCNO", Me.Width, Me.Height)
    
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
        wsOldCurCd = cboCurr.Text
        wsOldShipCd = cboShipCode.Text
        wsOldRmkCd = cboRmkCode.Text
        wsOldPayCd = cboPayCode.Text
        
        
        Call SetButtonStatus("AfrKeyEdit")
        Call SetFieldStatus("AfrKey")
        cboCusCode.SetFocus
        
        
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

Private Sub cboMLCode_GotFocus()
    FocusMe cboMLCode
End Sub

Private Sub cboMLCode_LostFocus()
    FocusMe cboMLCode, True
End Sub

Private Sub cboRefDocNo_DropDown()
   
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboRefDocNo
    
    wsSql = "SELECT SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
    wsSql = wsSql & " WHERE SOHDSTATUS = '1' "
    wsSql = wsSql & " AND SOHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSql = wsSql & " ORDER BY SOHDDOCNO "
                
    Call Ini_Combo(2, wsSql, cboRefDocNo.Left + tabDetailInfo.Left, cboRefDocNo.Top + cboRefDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSONO", Me.Width, Me.Height)
           
            
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
            
            If Chk_KeyFld Then
                tabDetailInfo.Tab = 0
                cboSaleCode.SetFocus
            End If
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
    Dim rsSR As New ADODB.Recordset
    Dim wsSql As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
        wsSql = "SELECT SRHDDOCID, SRHDDOCNO, SRHDREFDOCID, SRHDCUSID, CUSID, CUSCODE, CUSNAME, CUSTEL, CUSFAX, "
        wsSql = wsSql & "SRHDDOCDATE, SRHDREVNO, SRHDCURR, SRHDEXCR, SRDTDOCLINE, "
        wsSql = wsSql & "SRHDDUEDATE, SRHDPAYCODE, SRHDPRCCODE, SRHDSALEID, "
        wsSql = wsSql & "SRHDCUSPO, SRHDLCNO, SRHDPORTNO, SRHDSHIPPER, SRHDSHIPFROM, SRHDSHIPTO, SRHDSHIPVIA, SRHDSHIPNAME, "
        wsSql = wsSql & "SRHDSHIPCODE, SRHDSHIPADR1, SRHDSHIPADR2, SRHDSHIPADR3, SRHDSHIPADR4, "
        wsSql = wsSql & "SRHDRMKCODE, SRHDRMK1,  SRHDRMK2,  SRHDRMK3,  SRHDRMK4, SRHDRMK5, "
        wsSql = wsSql & "SRHDRMK6, SRHDRMK7, SRHDRMK8, SRHDRMK9, SRHDRMK10, "
        wsSql = wsSql & "SRHDGRSAMT , SRHDGRSAMTL, SRHDDISAMT, SRHDDISAMTL, SRHDNETAMT, SRHDNETAMTL, "
        wsSql = wsSql & "SRDTITEMID, ITMCODE, SRDTWHSCODE, SRDTLOTNO, ITMBARCODE, SRDTITEMDESC ITNAME, ITMPUBLISHER,  SRDTQTY, SRDTUPRICE, SRDTDISPER, SRDTAMT, SRDTAMTL, SRDTDIS, SRDTDISL, SRDTNET, SRDTNETL, "
        wsSql = wsSql & "SOHDDOCNO , SRDTSOID, SRHDMLCODE "
        wsSql = wsSql & "FROM  soaSrHd, soaSrDT, MstCustomer, MstItem, soaSoHd "
        wsSql = wsSql & "WHERE SRHDDOCNO = '" & cboDocNo & "' "
        wsSql = wsSql & "AND SRHDDOCID = SRDTDOCID "
        wsSql = wsSql & "AND SRDTSOID = SOHDDOCID "
        wsSql = wsSql & "AND SRHDCUSID = CUSID "
        wsSql = wsSql & "AND SRDTITEMID = ITMID "
        wsSql = wsSql & "ORDER BY SRDTDOCLINE "
   
    rsSR.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsSR.RecordCount <= 0 Then
        rsSR.Close
        Set rsSR = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsSR, "SRHDDOCID")
    wlRefDocID = ReadRs(rsSR, "SRHDREFDOCID")
    cboRefDocNo.Text = Get_TableInfo("soaSOHD", "SOHDDOCID =" & wlRefDocID, "SOHDDOCNO")
    txtRevNo.Text = Format(ReadRs(rsSR, "SRHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsSR, "SRHDREVNO"))
    medDocDate.Text = ReadRs(rsSR, "SRHDDOCDATE")
    wlCusID = ReadRs(rsSR, "CUSID")
    cboCusCode.Text = ReadRs(rsSR, "CUSCODE")
    lblDspCusName.Caption = ReadRs(rsSR, "CUSNAME")
    lblDspCusTel.Caption = ReadRs(rsSR, "CUSTEL")
    lblDspCusFax.Caption = ReadRs(rsSR, "CUSFAX")
    cboCurr.Text = ReadRs(rsSR, "SRHDCURR")
    txtExcr.Text = Format(ReadRs(rsSR, "SRHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsSR, "SRHDDUEDATE"))
    
    wlSaleID = To_Value(ReadRs(rsSR, "SRHDSALEID"))
    
    cboPayCode = ReadRs(rsSR, "SRHDPAYCODE")
    cboPrcCode = ReadRs(rsSR, "SRHDPRCCODE")
    cboShipCode = ReadRs(rsSR, "SRHDSHIPCODE")
    cboRmkCode = ReadRs(rsSR, "SRHDRMKCODE")
    
    txtCusPo = ReadRs(rsSR, "SRHDCUSPO")
    txtLcNo = ReadRs(rsSR, "SRHDLCNO")
    txtPortNo = ReadRs(rsSR, "SRHDPORTNO")
    
    cboMLCode = ReadRs(rsSR, "SRHDMLCODE")
    
    txtShipFrom = ReadRs(rsSR, "SRHDSHIPFROM")
    txtShipTo = ReadRs(rsSR, "SRHDSHIPTO")
    txtShipVia = ReadRs(rsSR, "SRHDSHIPVIA")
    txtShipName = ReadRs(rsSR, "SRHDSHIPNAME")
    txtShipPer = ReadRs(rsSR, "SRHDSHIPPER")
    txtShipAdr1 = ReadRs(rsSR, "SRHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsSR, "SRHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsSR, "SRHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsSR, "SRHDSHIPADR4")
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsSR, "SRHDRMK" & i)
    Next i
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    lblDspMLDesc = Get_TableInfo("MstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
        
    rsSR.MoveFirst
    With waResult
        .ReDim 0, -1, LINENO, SOID
        Do While Not rsSR.EOF
            wiCtr = wiCtr + 1
            .AppendRows
            waResult(.UpperBound(1), LINENO) = ReadRs(rsSR, "SRDTDOCLINE")
            waResult(.UpperBound(1), SONO) = ReadRs(rsSR, "SOHDDOCNO")
            waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsSR, "ITMCODE")
            waResult(.UpperBound(1), BARCODE) = ReadRs(rsSR, "ITMBARCODE")
            waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsSR, "ITNAME")
            waResult(.UpperBound(1), WhsCode) = ReadRs(rsSR, "SRDTWHSCODE")
            waResult(.UpperBound(1), LOTNO) = ReadRs(rsSR, "SRDTLOTNO")
            waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsSR, "ITMPUBLISHER")
            waResult(.UpperBound(1), Qty) = Format(ReadRs(rsSR, "SRDTQTY"), gsQtyFmt)
            waResult(.UpperBound(1), Price) = Format(ReadRs(rsSR, "SRDTUPRICE"), gsAmtFmt)
            waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsSR, "SRDTDISPER"), "0.0")
            waResult(.UpperBound(1), Amt) = Format(ReadRs(rsSR, "SRDTAMT"), gsAmtFmt)
            waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsSR, "SRDTAMTL"), gsAmtFmt)
            waResult(.UpperBound(1), Dis) = Format(ReadRs(rsSR, "SRDTDIS"), gsAmtFmt)
            waResult(.UpperBound(1), Disl) = Format(ReadRs(rsSR, "SRDTDISL"), gsAmtFmt)
            waResult(.UpperBound(1), Net) = Format(ReadRs(rsSR, "SRDTNET"), gsAmtFmt)
            waResult(.UpperBound(1), Netl) = Format(ReadRs(rsSR, "SRDTNETL"), gsAmtFmt)
            waResult(.UpperBound(1), BOOKID) = ReadRs(rsSR, "SRDTITEMID")
            waResult(.UpperBound(1), SOID) = ReadRs(rsSR, "SRDTSOID")
            
            rsSR.MoveNext
        Loop
        wlLineNo = waResult(.UpperBound(1), LINENO) + 1
    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsSR.Close
    
    Set rsSR = Nothing
    
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
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    lblCusTel.Caption = Get_Caption(waScrItm, "CUSTEL")
    lblCusFax.Caption = Get_Caption(waScrItm, "CUSFAX")
    LblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblPayCode.Caption = Get_Caption(waScrItm, "PAYCODE")
    lblPrcCode.Caption = Get_Caption(waScrItm, "PRCCODE")
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    
    lblDueDate.Caption = Get_Caption(waScrItm, "DUEDATE")
    
    lblGrsAmtOrg.Caption = Get_Caption(waScrItm, "GRSAMTORG")
    lblNetAmtOrg.Caption = Get_Caption(waScrItm, "NETAMTORG")
    lblDisAmtOrg.Caption = Get_Caption(waScrItm, "DISAMTORG")
    lblTotalQty.Caption = Get_Caption(waScrItm, "TOTALQTY")
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(SONO).Caption = Get_Caption(waScrItm, "SONO")
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(WhsCode).Caption = Get_Caption(waScrItm, "WHSCODE")
        .Columns(LOTNO).Caption = Get_Caption(waScrItm, "LOTNO")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(PUBLISHER).Caption = Get_Caption(waScrItm, "PUBLISHER")
        .Columns(Qty).Caption = Get_Caption(waScrItm, "QTY")
        .Columns(Price).Caption = Get_Caption(waScrItm, "PRICE")
        .Columns(DisPer).Caption = Get_Caption(waScrItm, "DISPER")
        .Columns(Dis).Caption = Get_Caption(waScrItm, "DIS")
        .Columns(Net).Caption = Get_Caption(waScrItm, "NET")
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
    
    btnSOLST.Caption = Get_Caption(waScrItm, "SOLIST")
    
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
    
    wsActNam(1) = Get_Caption(waScrItm, "SRADD")
    wsActNam(2) = Get_Caption(waScrItm, "SREDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SRDELETE")
    wgsTitle = Get_Caption(waScrItm, "TITLE")
    
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
    Set frmSRN001 = Nothing

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
            txtShipFrom.SetFocus
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
    
        If cboShipCode.Enabled Then
            cboShipCode.SetFocus
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

Private Function Chk_KeyExist() As Boolean
    
    Dim rssoaSrHd As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT SRHDSTATUS FROM soaSrHd WHERE SRHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rssoaSrHd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    If rssoaSrHd.RecordCount > 0 Then
        Chk_KeyExist = True
    Else
        Chk_KeyExist = False
    End If
    
    rssoaSrHd.Close
    Set rssoaSrHd = Nothing
End Function

Private Function Chk_KeyFld() As Boolean
    Chk_KeyFld = False
    
    If chk_cboCusCode = False Then
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
    
    If lblDspNetAmtOrg.Caption > Get_CreditLimit(wlCusID, Trim(medDocDate.Text)) Then
       gsMsg = "已超過信貸額!"
       MsgBox gsMsg, vbOKOnly, gsTitle
       MousePointer = vbDefault
       Exit Function
    End If
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SRN001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlCusID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, txtRevNo.Text)
    Call SetSPPara(adcmdSave, 8, cboCurr.Text)
    Call SetSPPara(adcmdSave, 9, txtExcr.Text)
    Call SetSPPara(adcmdSave, 10, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medDueDate.Text))
    
    Call SetSPPara(adcmdSave, 12, wlSaleID)
    
    Call SetSPPara(adcmdSave, 13, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 14, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 15, cboMLCode.Text)
    
    Call SetSPPara(adcmdSave, 16, cboShipCode.Text)
    Call SetSPPara(adcmdSave, 17, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 18, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 19, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 20, txtPortNo.Text)
    
    Call SetSPPara(adcmdSave, 21, txtShipFrom.Text)
    Call SetSPPara(adcmdSave, 22, txtShipTo.Text)
    Call SetSPPara(adcmdSave, 23, txtShipVia.Text)
    Call SetSPPara(adcmdSave, 24, txtShipPer.Text)
    Call SetSPPara(adcmdSave, 25, txtShipName.Text)
    Call SetSPPara(adcmdSave, 26, txtShipAdr1.Text)
    Call SetSPPara(adcmdSave, 27, txtShipAdr2.Text)
    Call SetSPPara(adcmdSave, 28, txtShipAdr3.Text)
    Call SetSPPara(adcmdSave, 29, txtShipAdr4.Text)
    
    For i = 1 To 10
        Call SetSPPara(adcmdSave, 30 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdSave, 40, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 41, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 42, lblDspNetAmtOrg)
    Call SetSPPara(adcmdSave, 43, wlRefDocID)
    Call SetSPPara(adcmdSave, 44, wsFormID)
    
    Call SetSPPara(adcmdSave, 45, gsUserID)
    Call SetSPPara(adcmdSave, 46, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 47)
    wsDocNo = GetSPPara(adcmdSave, 48)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_SRN001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, SONO)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, SONO))
                Call SetSPPara(adcmdSave, 4, To_Value(waResult(wiCtr, SOID)))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, BOOKID))
                Call SetSPPara(adcmdSave, 6, wiCtr + 1)
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, BOOKNAME))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, Qty))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, Price))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, DisPer))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, WhsCode))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, LOTNO))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, Amt))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, Amtl))
                Call SetSPPara(adcmdSave, 15, waResult(wiCtr, Dis))
                Call SetSPPara(adcmdSave, 16, waResult(wiCtr, Disl))
                Call SetSPPara(adcmdSave, 17, waResult(wiCtr, Net))
                Call SetSPPara(adcmdSave, 18, waResult(wiCtr, Netl))
                Call SetSPPara(adcmdSave, 19, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 20, gsUserID)
                Call SetSPPara(adcmdSave, 21, wsGenDte)
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
    
    If Not Chk_cboRefDocNo() Then Exit Function
    If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboCusCode() Then Exit Function
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    If Not Chk_cboPayCode Then Exit Function
    If Not Chk_cboPrcCode Then Exit Function
    
    If Not Chk_medDueDate Then Exit Function
    
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
                
                If Chk_NoDup2(wlCtr, waResult(wlCtr, SONO), waResult(wlCtr, BOOKCODE), waResult(wlCtr, WhsCode), waResult(wlCtr, LOTNO)) = False Then
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
        gsMsg = "銷售單沒有詳細資料!"
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
    Dim newForm As New frmSRN001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show
End Sub

Private Sub cmdOpen()

    Dim newForm As New frmSRN001
    
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
    wsFormID = "SRN001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "SR"
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
        gsMsg = "對換率超出範圍!"
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
    Call Ini_Combo(2, wsSql, cboCusCode.Left + tabDetailInfo.Left, cboCusCode.Top + cboCusCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLCUSNO", Me.Width, Me.Height)
    
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
                tabDetailInfo.Tab = 0
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
        cboCurr.Text = ReadRs(rsDefVal, "CUSCURR")
        cboPayCode.Text = ReadRs(rsDefVal, "CUSPAYCODE")
        wlSaleID = ReadRs(rsDefVal, "CUSSALEID")
        txtShipName = ReadRs(rsDefVal, "CUSSHIPTO")
        txtShipPer = ReadRs(rsDefVal, "CUSSHIPCONTACTPERSON")
        txtShipAdr1 = ReadRs(rsDefVal, "CUSSHIPADD1")
        txtShipAdr2 = ReadRs(rsDefVal, "CUSSHIPADD2")
        txtShipAdr3 = ReadRs(rsDefVal, "CUSSHIPADD3")
        txtShipAdr4 = ReadRs(rsDefVal, "CUSSHIPADD4")
        
          Else
        cboCurr.Text = ""
        cboPayCode.Text = ""
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
                Case BOOKCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 13
                Case BARCODE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case WhsCode
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Visible = False
                Case LOTNO
                    .Columns(wiCtr).Width = 1000
                    '.Columns(wiCtr).Button = False
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Visible = False
                Case BOOKNAME
                    .Columns(wiCtr).Width = 2500
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case PUBLISHER
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).Visible = False
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
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case DisPer
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Net
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Dis
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Visible = False
                Case Amt
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
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
Dim wsBookID As String
Dim wsBookCode As String
Dim wsBarCode As String
Dim wsBookName As String
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
                
                If Chk_grdSoNo(.Columns(ColIndex).Text, "") = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(SOID).Text = wsSoId
                
                
            Case BOOKCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdBookCode(.Columns(SONO).Text, .Columns(ColIndex).Text, wsBookID, wsBookCode, wsBarCode, wsBookName, wsPub, wdPrice, wdDisPer, wsWhsCode, wsLotNo, wdQty) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                .Columns(LINENO).Text = wlLineNo
                .Columns(BOOKID).Text = wsBookID
                .Columns(BARCODE).Text = wsBarCode
                .Columns(BOOKNAME).Text = wsBookName
                .Columns(PUBLISHER).Text = wsPub
                .Columns(WhsCode).Text = wsWhsCode
                .Columns(LOTNO).Text = wsLotNo
                .Columns(Price).Text = Format(wdPrice, gsAmtFmt)
                .Columns(Qty).Text = Format(wdQty, gsQtyFmt)
                .Columns(DisPer).Text = Format(wdDisPer, "0")
                wlLineNo = wlLineNo + 1
                
                If Trim(.Columns(ColIndex).Text) <> wsBookCode Then
                    .Columns(ColIndex).Text = wsBookCode
                End If
                If Trim(.Columns(Price).Text) <> "" Then
                .Columns(Amt).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text), gsAmtFmt)
                End If
                If Trim(txtExcr.Text) <> "" Then
                .Columns(Amtl).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text) * To_Value(txtExcr.Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Dis).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Disl).Text = Format(To_Value(.Columns(Amtl).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(Dis).Text) <> "" Then
                .Columns(Net).Text = Format(To_Value(.Columns(Amt).Text) - To_Value(.Columns(Dis).Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(Disl).Text) <> "" Then
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) - To_Value(.Columns(Disl).Text), gsAmtFmt)
                End If
        
             Case WhsCode
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
                End If
                If Trim(txtExcr.Text) <> "" Then
                .Columns(Amtl).Text = Format(To_Value(.Columns(Price).Text) * To_Value(.Columns(Qty).Text) * To_Value(txtExcr.Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Dis).Text = Format(To_Value(.Columns(Amt).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(DisPer).Text) <> "" Then
                .Columns(Disl).Text = Format(To_Value(.Columns(Amtl).Text) * To_Value(.Columns(DisPer).Text) / 100, gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(Dis).Text) <> "" Then
                .Columns(Net).Text = Format(To_Value(.Columns(Amt).Text) - To_Value(.Columns(Dis).Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(Disl).Text) <> "" Then
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) - To_Value(.Columns(Disl).Text), gsAmtFmt)
                End If
                
                Case Dis
                                
                If Trim(txtExcr.Text) <> "" Then
                .Columns(Disl).Text = Format(To_Value(.Columns(Dis).Text) * To_Value(txtExcr.Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amt).Text) <> "" And Trim(.Columns(Dis).Text) <> "" Then
                .Columns(Net).Text = Format(To_Value(.Columns(Amt).Text) - To_Value(.Columns(Dis).Text), gsAmtFmt)
                End If
                If Trim(.Columns(Amtl).Text) <> "" And Trim(.Columns(Disl).Text) <> "" Then
                .Columns(Netl).Text = Format(To_Value(.Columns(Amtl).Text) - To_Value(.Columns(Disl).Text), gsAmtFmt)
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
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
            Case SONO
                
                wsSql = "SELECT SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
                wsSql = wsSql & " WHERE SOHDSTATUS = '1' "
                wsSql = wsSql & " AND SOHDDOCNO LIKE '%" & Set_Quote(.Columns(SONO).Text) & "%' "
                wsSql = wsSql & " AND SOHDCUSID = " & wlCusID
                wsSql = wsSql & " ORDER BY SOHDDOCNO "
                
                Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSONO", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case BOOKCODE
                
                If gsLangID = 1 Then
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM FROM mstITEM, soaSohd, soaSodt "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND SOHDDOCNO = '" & Set_Quote(.Columns(SONO).Text) & "' "
                    wsSql = wsSql & " AND SOHDDOCID = SODTDOCID "
                    wsSql = wsSql & " AND SODTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                    wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM FROM mstITEM, soaSohd, soaSodt "
                    wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                    wsSql = wsSql & " AND SOHDDOCNO = '" & Set_Quote(.Columns(SONO).Text) & "' "
                    wsSql = wsSql & " AND SOHDDOCID = SODTDOCID "
                    wsSql = wsSql & " AND SODTITEMID = ITMID "
                    wsSql = wsSql & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(4, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLBOOKCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
                
            Case WhsCode
                
                wsSql = "SELECT WHSCODE, WHSDESC FROM mstWareHouse "
                wsSql = wsSql & " WHERE WHSSTATUS <> '2' AND WHSCODE LIKE '%" & Set_Quote(.Columns(WhsCode).Text) & "%' "
                wsSql = wsSql & " ORDER BY WHSCODE "
                
                Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLWHSCODE", Me.Width, Me.Height)
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
                Case Net
                    KeyCode = vbKeyDown
                    .Col = SONO
                Case BOOKCODE
                    KeyCode = vbDefault
                       .Col = Qty
                Case LINENO, SONO, Qty, Price
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
                Case Qty
                    .Col = BOOKNAME
                Case BOOKNAME
                    .Col = BARCODE
                Case Net
                    .Col = DisPer
                Case DisPer, Price, BARCODE, BOOKCODE
                    .Col = .Col - 1
                
            End Select
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case BOOKCODE
                       .Col = Qty
                Case LINENO, Qty, Price, SONO
                   .Col = .Col + 1
                Case BARCODE
                       .Col = BOOKNAME
                Case BOOKNAME
                       .Col = Qty
                Case DisPer
                    KeyCode = vbDefault
                       .Col = Net
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
        
        Case Price, DisPer, Dis
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
                Case BOOKCODE
                    Call Chk_grdBookCode(.Columns(SONO).Text, .Columns(BOOKCODE).Text, "", "", "", "", "", 0, 0, "", "", 0)
                Case WhsCode
                    Call Chk_grdWhsCode(.Columns(WhsCode).Text)
                 Case LOTNO
                    Call Chk_grdLotNo(.Columns(LOTNO).Text)
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

Private Function Chk_grdBookCode(inSoNo As String, inAccNo As String, outAccID As String, outAccNo As String, OutBarCode As String, OutName As String, outPub As String, outPrice As Double, outDisPer As Double, outWhsCode As String, outLotNo As String, outQty As Double) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    Dim wsCurr As String
    Dim wsExcr As String
    Dim wdPrice As Double
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入書號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
        Exit Function
    End If
    

    
        wsSql = "SELECT SODTDOCID, SODTITEMID, ITMCODE, SODTITEMDESC ITNAME, ITMBARCODE, ITMPUBLISHER, SODTUPRICE, SOHDCURR, SODTWHSCODE, SODTLOTNO, SODTQTY, SODTBALQTY, SODTDISPER "
        wsSql = wsSql & " FROM mstITEM, soaSoHd, soaSoDt "
        wsSql = wsSql & " WHERE SOHDDOCID = SODTDOCID "
        wsSql = wsSql & " AND SODTITEMID = ITMID "
        wsSql = wsSql & " AND SOHDDOCNO = '" & Set_Quote(inSoNo) & "' "
        wsSql = wsSql & " AND (ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "') "
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       outAccID = ReadRs(rsDes, "SODTITEMID")
       outAccNo = ReadRs(rsDes, "ITMCODE")
       OutName = ReadRs(rsDes, "ITNAME")
       OutBarCode = ReadRs(rsDes, "ITMBARCODE")
       outPub = ReadRs(rsDes, "ITMPUBLISHER")
       outPrice = To_Value(ReadRs(rsDes, "SODTUPRICE"))
       wsCurr = ReadRs(rsDes, "SOHDCURR")
       outWhsCode = ReadRs(rsDes, "SODTWHSCODE")
       outLotNo = ReadRs(rsDes, "SODTLOTNO")
       outQty = Get_SoBalQty(wsTrnCd, wlKey, ReadRs(rsDes, "SODTDOCID"), outAccID, outWhsCode, outLotNo)
       
       If cboCurr <> wsCurr Then
       If getExcRate(wsCurr, medDocDate, wsExcr, "") = True Then
       outPrice = NBRnd(outPrice * To_Value(wsExcr) / txtExcr, giExrDp)
       End If
       End If
       
        outDisPer = To_Value(ReadRs(rsDes, "SODTDISPER"))
       
       Chk_grdBookCode = True
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
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdOrdItm(inSoNo As String, inItmNo As String, inWhsCode As String, InLotNo As String) As Boolean
    Dim wsSql As String
    Dim rsDes As New ADODB.Recordset
    
    wsSql = "SELECT SODTITEMID "
    wsSql = wsSql & " FROM mstITEM, soaSoHd, soaSoDt "
    wsSql = wsSql & " WHERE SOHDDOCID = SODTDOCID "
    wsSql = wsSql & " AND SODTITEMID = ITMID "
    wsSql = wsSql & " AND SOHDDOCNO = '" & Set_Quote(inSoNo) & "' "
    wsSql = wsSql & " AND ITMCODE = '" & Set_Quote(inItmNo) & "' "
    wsSql = wsSql & " AND SODTWHSCODE = '" & Set_Quote(inWhsCode) & "' "
    wsSql = wsSql & " AND SODTLOTNO = '" & Set_Quote(InLotNo) & "' "
    wsSql = wsSql & " AND SOHDSTATUS NOT IN ('2' , '3')"
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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

Private Function Chk_grdSoNo(inSoNo As String, ByRef outSoID As String) As Boolean
    
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
    Chk_grdSoNo = False
    
    outSoID = "0"
    
    wsSql = "SELECT SOHDDOCID, SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
    wsSql = wsSql & " WHERE SOHDSTATUS = '1' "
    wsSql = wsSql & " AND SOHDDOCNO = '" & inSoNo & "' "
              
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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

Private Function Chk_grdLotNo(inNo As String) As Boolean
    
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
  
    Chk_grdLotNo = False
    
    If Trim(inNo) = "" Then
        gsMsg = "必需輸入版次!"
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
                If Trim(.Columns(SONO)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, SONO)) = "" And _
                   Trim(waResult(inRow, BOOKCODE)) = "" And _
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
        
        If Chk_grdSoNo(waResult(LastRow, SONO), "") = False Then
            .Col = SONO
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdBookCode(waResult(LastRow, SONO), waResult(LastRow, BOOKCODE), "", "", "", "", "", 0, 0, "", "", 0) = False Then
            .Col = BOOKCODE
            .Row = LastRow
            Exit Function
        End If
        
        If Chk_grdWhsCode(waResult(LastRow, WhsCode)) = False Then
                .Col = WhsCode
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
        
        If Chk_grdOrdItm(waResult(LastRow, SONO), waResult(LastRow, BOOKCODE), waResult(LastRow, WhsCode), waResult(LastRow, LOTNO)) = False Then
            .Col = BOOKCODE
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
    
    wiAction = DelRec
    
      cnCon.BeginTrans
    Set adcmdDelete.ActiveConnection = cnCon
        
    adcmdDelete.CommandText = "USP_SRN001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
    
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, wlCusID)
    Call SetSPPara(adcmdDelete, 6, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 7, txtRevNo.Text)
    Call SetSPPara(adcmdDelete, 8, cboCurr.Text)
    Call SetSPPara(adcmdDelete, 9, txtExcr.Text)
    Call SetSPPara(adcmdDelete, 10, "")
    
    Call SetSPPara(adcmdDelete, 11, Set_MedDate(medDueDate.Text))
    
    Call SetSPPara(adcmdDelete, 12, wlSaleID)
    
    Call SetSPPara(adcmdDelete, 13, cboPayCode.Text)
    Call SetSPPara(adcmdDelete, 14, cboPrcCode.Text)
    Call SetSPPara(adcmdDelete, 15, cboMLCode.Text)
    
    Call SetSPPara(adcmdDelete, 16, cboShipCode.Text)
    Call SetSPPara(adcmdDelete, 17, cboRmkCode.Text)
    
    Call SetSPPara(adcmdDelete, 18, txtCusPo.Text)
    Call SetSPPara(adcmdDelete, 19, txtLcNo.Text)
    Call SetSPPara(adcmdDelete, 20, txtPortNo.Text)
    
    Call SetSPPara(adcmdDelete, 21, txtShipFrom.Text)
    Call SetSPPara(adcmdDelete, 22, txtShipTo.Text)
    Call SetSPPara(adcmdDelete, 23, txtShipVia.Text)
    Call SetSPPara(adcmdDelete, 24, txtShipPer.Text)
    Call SetSPPara(adcmdDelete, 25, txtShipName.Text)
    Call SetSPPara(adcmdDelete, 26, txtShipAdr1.Text)
    Call SetSPPara(adcmdDelete, 27, txtShipAdr2.Text)
    Call SetSPPara(adcmdDelete, 28, txtShipAdr3.Text)
    Call SetSPPara(adcmdDelete, 29, txtShipAdr4.Text)
    
    For i = 1 To 10
        Call SetSPPara(adcmdDelete, 30 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdDelete, 40, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdDelete, 41, lblDspDisAmtOrg)
    Call SetSPPara(adcmdDelete, 42, lblDspNetAmtOrg)
    Call SetSPPara(adcmdDelete, 43, wlRefDocID)
    Call SetSPPara(adcmdDelete, 44, wsFormID)
    
    Call SetSPPara(adcmdDelete, 45, gsUserID)
    Call SetSPPara(adcmdDelete, 46, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 47)
    wsDocNo = GetSPPara(adcmdDelete, 48)
      
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
Public Sub SetFieldStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
        Case "Default"
        
            Me.cboDocNo.Enabled = False
            Me.cboRefDocNo.Enabled = False
            
            Me.cboCusCode.Enabled = False
            Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            
            Me.medDueDate.Enabled = False
            Me.cboSaleCode.Enabled = False
            Me.cboPayCode.Enabled = False
            Me.cboPrcCode.Enabled = False
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
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.medDueDate.Enabled = True
            Me.cboSaleCode.Enabled = True
            Me.cboPayCode.Enabled = True
            Me.cboPrcCode.Enabled = True
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
        .TableKey = "SRHDDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cbosaleCode_GotFocus()
    FocusMe cboSaleCode
End Sub

Private Sub cboSaleCode_LostFocus()
    FocusMe cboSaleCode, True
End Sub

Private Sub cbosaleCode_KeyPress(KeyAscii As Integer)
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

Private Sub cbosaleCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboSaleCode
    
    wsSql = "SELECT SALECODE, SALENAME FROM mstSalesman WHERE SaleCode LIKE '%" & IIf(cboSaleCode.SelLength > 0, "", Set_Quote(cboSaleCode.Text)) & "%' "
    wsSql = wsSql & "AND SaleStatus = '1' "
    wsSql = wsSql & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSql, cboSaleCode.Left + tabDetailInfo.Left, cboSaleCode.Top + cboSaleCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLSALECOD", Me.Width, Me.Height)
    
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
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboPayCode
    
    wsSql = "SELECT PAYCODE, PAYDESC FROM mstPayTerm WHERE PAYCODE LIKE '%" & IIf(cboPayCode.SelLength > 0, "", Set_Quote(cboPayCode.Text)) & "%' "
    wsSql = wsSql & "AND PAYSTATUS = '1' "
    wsSql = wsSql & "ORDER BY PAYCODE "
    Call Ini_Combo(2, wsSql, cboPayCode.Left + tabDetailInfo.Left, cboPayCode.Top + cboPayCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLPAYCOD", Me.Width, Me.Height)
    
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
        
        txtPortNo = Get_TableInfo("MstPriceTerm", "PrcCode = '" & Set_Quote(cboPrcCode.Text) & "'", "PricePort")
        
        tabDetailInfo.Tab = 0
        cboMLCode.SetFocus
       
    End If
End Sub

Private Sub cboPrcCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboPrcCode
    
    wsSql = "SELECT PrcCode, PRCDESC FROM mstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCode.SelLength > 0, "", Set_Quote(cboPrcCode.Text)) & "%' "
    wsSql = wsSql & "AND PRCSTATUS = '1' "
    wsSql = wsSql & "ORDER BY PrcCode "
    Call Ini_Combo(2, wsSql, cboPrcCode.Left + tabDetailInfo.Left, cboPrcCode.Top + cboPrcCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLPRCCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboPrcCode() As Boolean
    Dim wsDesc As String

    Chk_cboPrcCode = False
     
    If Trim(cboPrcCode.Text) = "" Then
        gsMsg = "必需輸入銷售條款!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
        Exit Function
    End If
    
    If Chk_PriceTerm(cboPrcCode, wsDesc) = False Then
        gsMsg = "沒有此銷售條款!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
        lblDspPrcDesc = ""
       Exit Function
    End If
    
    lblDspPrcDesc = wsDesc
    Chk_cboPrcCode = True
    
End Function

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
    
    Call chk_InpLen(txtCusPo, 20, KeyAscii)
    
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
            tblDetail.Col = SONO
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
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboShipCode
    
    wsSql = "SELECT ShipCode, ShipName, ShipPer FROM mstShip WHERE ShipCode LIKE '%" & IIf(cboShipCode.SelLength > 0, "", Set_Quote(cboShipCode.Text)) & "%' "
    wsSql = wsSql & "AND ShipSTATUS = '1' "
    wsSql = wsSql & "AND ShipCardID = " & wlCusID & " "
    wsSql = wsSql & "ORDER BY ShipCode "
    Call Ini_Combo(3, wsSql, cboShipCode.Left + tabDetailInfo.Left, cboShipCode.Top + cboShipCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLSHIPCOD", Me.Width, Me.Height)
    
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
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboRmkCode
    
    wsSql = "SELECT RmkCode FROM mstRemark WHERE RmkCode LIKE '%" & IIf(cboRmkCode.SelLength > 0, "", Set_Quote(cboRmkCode.Text)) & "%' "
    wsSql = wsSql & "AND RmkSTATUS = '1' "
    wsSql = wsSql & "ORDER BY RmkCode "
    Call Ini_Combo(1, wsSql, cboRmkCode.Left + tabDetailInfo.Left, cboRmkCode.Top + cboRmkCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLRMKCOD", Me.Width, Me.Height)
    
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
    Dim wsSql As String
    
    wsSql = "SELECT * "
    wsSql = wsSql & "FROM  mstShip "
    wsSql = wsSql & "WHERE ShipCode = '" & Set_Quote(cboShipCode) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    Dim wsSql As String
    Dim i As Integer
    
    wsSql = "SELECT * "
    wsSql = wsSql & "FROM  mstReMark "
    wsSql = wsSql & "WHERE RmkCode = '" & Set_Quote(cboRmkCode) & "'"
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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
    wsCurRecLn = tblDetail.Columns(BOOKCODE)
    wsCurRecLn2 = tblDetail.Columns(WhsCode)
    wsCurRecLn3 = tblDetail.Columns(LOTNO)
    
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, SONO) And _
                  wsCurRecLn = waResult(wlCtr, BOOKCODE) And _
                  wsCurRecLn2 = waResult(wlCtr, WhsCode) And _
                  wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
                  gsMsg = "重覆書本於第 " & waResult(wlCtr, LINENO) & " 行!"
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
                  wsCurRecLn = waResult(wlCtr, BOOKCODE) And _
                  wsCurRecLn2 = waResult(wlCtr, WhsCode) And _
                  wsCurRecLn3 = waResult(wlCtr, LOTNO) Then
                  gsMsg = "重覆書本於第 " & waResult(wlCtr, LINENO) & " 行!"
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
    '    Call Ini_PgmMenu(mnuPopUpSub, "POPUP", waPopUpSub)
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

Private Sub cboMLCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboMLCode
    
    wsSql = "SELECT MLCode, MLDESC FROM mstMerchClass WHERE MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSql = wsSql & "AND MLSTATUS = '1' "
    wsSql = wsSql & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSql, cboMLCode.Left + tabDetailInfo.Left, cboMLCode.Top + cboMLCode.Height + tabDetailInfo.Top, tblCommon, "SRN001", "TBLMLCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

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
        medDueDate.SetFocus
    End If
    
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
    
    
    If Chk_MerchClass(cboMLCode, wsDesc) = False Then
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

Public Sub SetButtonStatus(ByVal SSTATUS As String)
    Select Case SSTATUS
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
    wsSql = "EXEC usp_RPTSRN002 '" & Set_Quote(gsUserID) & "', "
    wsSql = wsSql & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSql = wsSql & "'" & wgsTitle & "', "
    wsSql = wsSql & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSql = wsSql & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSql = wsSql & "'" & "" & "', "
    wsSql = wsSql & "'" & String(10, "z") & "', "
    wsSql = wsSql & "'" & "0000/00/00" & "', "
    wsSql = wsSql & "'" & "9999/99/99" & "', "
    wsSql = wsSql & "'" & "%" & "', "
    wsSql = wsSql & gsLangID
    
    
    If gsLangID = "2" Then wsRptName = "C" + "RPTSRN002"
    
    NewfrmPrint.ReportID = "SRN002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SRN002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSql
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
    Dim wsSql As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    Dim wdBalQty As Double
    
    Get_RefDoc = False
    
        wsSql = "SELECT SOHDDOCID, SOHDDOCNO, SOHDCUSID, CUSID, CUSCODE, CUSNAME, CUSTEL, CUSFAX, "
        wsSql = wsSql & "SOHDDOCDATE, SOHDREVNO, SOHDCURR, SOHDEXCR, "
        wsSql = wsSql & "SOHDDUEDATE, SOHDPAYCODE, SOHDPRCCODE, SOHDSALEID, SOHDMLCODE, SOHDNatureCODE, "
        wsSql = wsSql & "SOHDCUSPO, SOHDLCNO, SOHDPORTNO, SOHDSHIPPER, SOHDSHIPFROM, SOHDSHIPTO, SOHDSHIPVIA, SOHDSHIPNAME, "
        wsSql = wsSql & "SOHDSHIPCODE, SOHDSHIPADR1,  SOHDSHIPADR2,  SOHDSHIPADR3,  SOHDSHIPADR4, "
        wsSql = wsSql & "SOHDRMKCODE, SOHDRMK1,  SOHDRMK2,  SOHDRMK3,  SOHDRMK4, SOHDRMK5, "
        wsSql = wsSql & "SOHDRMK6,  SOHDRMK7,  SOHDRMK8,  SOHDRMK9, SOHDRMK10, "
        wsSql = wsSql & "SOHDGRSAMT , SOHDGRSAMTL, SOHDDISAMT, SOHDDISAMTL, SOHDNETAMT, SOHDNETAMTL, "
        wsSql = wsSql & "SODTITEMID, ITMCODE, SODTWHSCODE, SODTLOTNO, ITMBARCODE, SODTITEMDESC ITNAME, ITMPUBLISHER,  SODTBALQTY, SODTUPRICE, SODTDISPER, SODTAMT, SODTAMTL, SODTDIS, SODTDISL, SODTNET, SODTNETL, "
        wsSql = wsSql & "SODTID "
        wsSql = wsSql & "FROM  soaSOHD, soaSODT, mstCUSTOMER, mstITEM "
        wsSql = wsSql & "WHERE SOHDDOCNO = '" & cboRefDocNo & "' "
        wsSql = wsSql & "AND SOHDDOCID = SODTDOCID "
        wsSql = wsSql & "AND SOHDCUSID = CUSID "
        wsSql = wsSql & "AND SODTITEMID = ITMID "
        wsSql = wsSql & "ORDER BY SODTDOCLINE "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

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
    cboCurr.Text = ReadRs(rsRcd, "SOHDCURR")
    txtExcr.Text = Format(ReadRs(rsRcd, "SOHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsRcd, "SOHDDUEDATE"))
    
    wlSaleID = To_Value(ReadRs(rsRcd, "SOHDSALEID"))
    
    cboPayCode = ReadRs(rsRcd, "SOHDPAYCODE")
    cboPrcCode = ReadRs(rsRcd, "SOHDPRCCODE")
    cboMLCode = ReadRs(rsRcd, "SOHDMLCODE")
    cboShipCode = ReadRs(rsRcd, "SOHDSHIPCODE")
    cboRmkCode = ReadRs(rsRcd, "SOHDRMKCODE")
    
    txtCusPo = ReadRs(rsRcd, "SOHDCUSPO")
    txtLcNo = ReadRs(rsRcd, "SOHDLCNO")
    txtPortNo = ReadRs(rsRcd, "SOHDPORTNO")
    
    
    
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
    
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    wsOldCusNo = cboCusCode
    wsOldCurCd = cboCurr
    wsOldShipCd = cboShipCode
    wsOldRmkCd = cboRmkCode
    wsOldPayCd = cboPayCode
    
    wlLineNo = 1
    rsRcd.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, SOID
         Do While Not rsRcd.EOF
             wiCtr = wiCtr + 1
             
              wdBalQty = Get_SoBalQty(wsTrnCd, 0, ReadRs(rsRcd, "SOHDDOCID"), ReadRs(rsRcd, "SODTITEMID"), ReadRs(rsRcd, "SODTWHSCODE"), ReadRs(rsRcd, "SODTLOTNO"))
   
             .AppendRows
             waResult(.UpperBound(1), LINENO) = wlLineNo
             waResult(.UpperBound(1), SONO) = ReadRs(rsRcd, "SOHDDOCNO")
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsRcd, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsRcd, "ITMBARCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsRcd, "ITNAME")
             waResult(.UpperBound(1), WhsCode) = ReadRs(rsRcd, "SODTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsRcd, "SODTLOTNO")
             waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsRcd, "ITMPUBLISHER")
             waResult(.UpperBound(1), Qty) = Format(wdBalQty, gsQtyFmt)
             waResult(.UpperBound(1), Price) = Format(ReadRs(rsRcd, "SODTUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsRcd, "SODTDISPER"), "0.0")
             waResult(.UpperBound(1), Amt) = Format(ReadRs(rsRcd, "SODTAMT"), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsRcd, "SODTAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(ReadRs(rsRcd, "SODTDIS"), gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(ReadRs(rsRcd, "SODTDISL"), gsAmtFmt)
             waResult(.UpperBound(1), Net) = Format(ReadRs(rsRcd, "SODTNET"), gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(ReadRs(rsRcd, "SODTNETL"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsRcd, "SODTITEMID")
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


