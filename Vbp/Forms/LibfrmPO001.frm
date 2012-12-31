VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPO001 
   Caption         =   "訂貨單"
   ClientHeight    =   6615
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   9795
   Icon            =   "frmPO001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   9840
      OleObjectBlob   =   "frmPO001.frx":030A
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCurr 
      Height          =   300
      Left            =   8280
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtExcr 
      Alignment       =   1  '靠右對齊
      Height          =   288
      Left            =   8280
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox cboVdrCode 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtRevNo 
      Height          =   324
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "12345678901234567890"
      Top             =   480
      Width           =   408
   End
   Begin VB.ComboBox cboDocNo 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin MSMask.MaskEdBox medDocDate 
      Height          =   288
      Left            =   8280
      TabIndex        =   3
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   5640
      Top             =   480
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
            Picture         =   "frmPO001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":6385
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   4455
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2040
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Header Information"
      TabPicture(0)   =   "frmPO001.frx":66AD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboVdrTyp"
      Tab(0).Control(1)=   "cboMLCode"
      Tab(0).Control(2)=   "cboPrcCode"
      Tab(0).Control(3)=   "cboPayCode"
      Tab(0).Control(4)=   "medDueDate"
      Tab(0).Control(5)=   "medOnDate"
      Tab(0).Control(6)=   "medEtaDate"
      Tab(0).Control(7)=   "lblNetAmtLoc"
      Tab(0).Control(8)=   "lblDisAmtLoc"
      Tab(0).Control(9)=   "lblGrsAmtLoc"
      Tab(0).Control(10)=   "lblDspNetAmtLoc"
      Tab(0).Control(11)=   "lblDspDisAmtLoc"
      Tab(0).Control(12)=   "lblDspGrsAmtLoc"
      Tab(0).Control(13)=   "lblDspNetAmtOrg"
      Tab(0).Control(14)=   "lblNetAmtOrg"
      Tab(0).Control(15)=   "lblDisAmtOrg"
      Tab(0).Control(16)=   "lblDspDisAmtOrg"
      Tab(0).Control(17)=   "lblDspGrsAmtOrg"
      Tab(0).Control(18)=   "lblGrsAmtOrg"
      Tab(0).Control(19)=   "lblEtaDate"
      Tab(0).Control(20)=   "lblOnDate"
      Tab(0).Control(21)=   "lblDueDate"
      Tab(0).Control(22)=   "lblDspPayDesc"
      Tab(0).Control(23)=   "lblPayCode"
      Tab(0).Control(24)=   "lblDspPrcDesc"
      Tab(0).Control(25)=   "lblPrcCode"
      Tab(0).Control(26)=   "lblDspMLDesc"
      Tab(0).Control(27)=   "lblDspVdrTypDesc"
      Tab(0).Control(28)=   "lblMlCode"
      Tab(0).Control(29)=   "lblVdrTyp"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmPO001.frx":66C9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblShipCode"
      Tab(1).Control(1)=   "lblShipFrom"
      Tab(1).Control(2)=   "lblShipVia"
      Tab(1).Control(3)=   "lblShipTo"
      Tab(1).Control(4)=   "lblShipName"
      Tab(1).Control(5)=   "lblShipPer"
      Tab(1).Control(6)=   "lblShipAdr"
      Tab(1).Control(7)=   "lblCusPo"
      Tab(1).Control(8)=   "lblPortNo"
      Tab(1).Control(9)=   "lblLcNo"
      Tab(1).Control(10)=   "Picture1"
      Tab(1).Control(11)=   "cboShipCode"
      Tab(1).Control(12)=   "txtShipFrom"
      Tab(1).Control(13)=   "txtShipVia"
      Tab(1).Control(14)=   "txtShipTo"
      Tab(1).Control(15)=   "txtShipName"
      Tab(1).Control(16)=   "txtShipPer"
      Tab(1).Control(17)=   "txtCusPo"
      Tab(1).Control(18)=   "txtPortNo"
      Tab(1).Control(19)=   "txtLcNo"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Item Information"
      TabPicture(2)   =   "frmPO001.frx":66E5
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "tblDetail"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Remark"
      TabPicture(3)   =   "frmPO001.frx":6701
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblRmkCode"
      Tab(3).Control(1)=   "lblRmk"
      Tab(3).Control(2)=   "cboRmkCode"
      Tab(3).Control(3)=   "picRmk"
      Tab(3).ControlCount=   4
      Begin VB.PictureBox picRmk 
         BackColor       =   &H80000009&
         Height          =   3495
         Left            =   -73320
         ScaleHeight     =   3435
         ScaleWidth      =   7635
         TabIndex        =   67
         Top             =   480
         Width           =   7695
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   10
            Left            =   0
            TabIndex        =   77
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   3120
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   9
            Left            =   0
            TabIndex        =   76
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   2775
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   8
            Left            =   0
            TabIndex        =   75
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   2430
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   7
            Left            =   0
            TabIndex        =   74
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   2085
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   5
            Left            =   0
            TabIndex        =   73
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1395
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   4
            Left            =   0
            TabIndex        =   72
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1035
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   6
            Left            =   0
            TabIndex        =   71
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1740
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   3
            Left            =   0
            TabIndex        =   70
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   690
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   1
            Left            =   0
            TabIndex        =   69
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   0
            Width           =   7545
         End
         Begin VB.TextBox txtRmk 
            BorderStyle     =   0  '沒有框線
            Height          =   300
            Index           =   2
            Left            =   0
            TabIndex        =   68
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   345
            Width           =   7545
         End
      End
      Begin VB.ComboBox cboRmkCode 
         Height          =   300
         Left            =   -73320
         TabIndex        =   64
         Top             =   120
         Width           =   1890
      End
      Begin VB.TextBox txtLcNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -68880
         TabIndex        =   62
         Text            =   "0123456789012345789"
         Top             =   600
         Width           =   3465
      End
      Begin VB.TextBox txtPortNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -68880
         TabIndex        =   60
         Text            =   "0123456789012345789"
         Top             =   960
         Width           =   3465
      End
      Begin VB.TextBox txtCusPo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -68880
         TabIndex        =   58
         Text            =   "0123456789012345789"
         Top             =   240
         Width           =   3465
      End
      Begin VB.TextBox txtShipPer 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73320
         TabIndex        =   50
         Text            =   "01234567890123457890"
         Top             =   1680
         Width           =   2745
      End
      Begin VB.TextBox txtShipName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -68880
         TabIndex        =   48
         Text            =   "012345678901234578901234567890123457890123456789"
         Top             =   1320
         Width           =   3465
      End
      Begin VB.TextBox txtShipTo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73320
         TabIndex        =   46
         Text            =   "0123456789012345789"
         Top             =   600
         Width           =   2745
      End
      Begin VB.TextBox txtShipVia 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73320
         TabIndex        =   44
         Text            =   "0123456789012345789"
         Top             =   960
         Width           =   2745
      End
      Begin VB.TextBox txtShipFrom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73320
         TabIndex        =   36
         Text            =   "0123456789012345789"
         Top             =   240
         Width           =   2745
      End
      Begin VB.ComboBox cboVdrTyp 
         Height          =   300
         Left            =   -73200
         TabIndex        =   25
         Top             =   1320
         Width           =   2370
      End
      Begin VB.ComboBox cboMLCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   24
         Top             =   1680
         Width           =   2370
      End
      Begin VB.ComboBox cboPrcCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   23
         Top             =   960
         Width           =   2370
      End
      Begin VB.ComboBox cboPayCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   22
         Top             =   600
         Width           =   2370
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Left            =   -73320
         TabIndex        =   21
         Top             =   1320
         Width           =   2730
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   3855
         Left            =   120
         OleObjectBlob   =   "frmPO001.frx":671D
         TabIndex        =   35
         Top             =   120
         Width           =   9495
      End
      Begin MSMask.MaskEdBox medDueDate 
         Height          =   285
         Left            =   -73200
         TabIndex        =   38
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOnDate 
         Height          =   285
         Left            =   -70200
         TabIndex        =   40
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEtaDate 
         Height          =   285
         Left            =   -66960
         TabIndex        =   42
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   1455
         Left            =   -73320
         ScaleHeight     =   1395
         ScaleWidth      =   7875
         TabIndex        =   53
         Top             =   2040
         Width           =   7935
         Begin VB.TextBox txtShipAdr4 
            BorderStyle     =   0  '沒有框線
            Enabled         =   0   'False
            Height          =   300
            Left            =   0
            TabIndex        =   57
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1080
            Width           =   7785
         End
         Begin VB.TextBox txtShipAdr3 
            BorderStyle     =   0  '沒有框線
            Enabled         =   0   'False
            Height          =   300
            Left            =   0
            TabIndex        =   56
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   720
            Width           =   7785
         End
         Begin VB.TextBox txtShipAdr2 
            BorderStyle     =   0  '沒有框線
            Enabled         =   0   'False
            Height          =   300
            Left            =   0
            TabIndex        =   55
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   360
            Width           =   7785
         End
         Begin VB.TextBox txtShipAdr1 
            BorderStyle     =   0  '沒有框線
            Enabled         =   0   'False
            Height          =   300
            Left            =   0
            TabIndex        =   54
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   0
            Width           =   7785
         End
      End
      Begin VB.Label lblNetAmtLoc 
         Caption         =   "NETAMTLOC"
         Height          =   255
         Left            =   -68760
         TabIndex        =   89
         Top             =   3600
         Width           =   1755
      End
      Begin VB.Label lblDisAmtLoc 
         Caption         =   "DISAMTLOC"
         Height          =   255
         Left            =   -68760
         TabIndex        =   88
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label lblGrsAmtLoc 
         Caption         =   "GRSAMTLOC"
         Height          =   255
         Left            =   -68760
         TabIndex        =   87
         Top             =   2880
         Width           =   1755
      End
      Begin VB.Label lblDspNetAmtLoc 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -66960
         TabIndex        =   86
         Top             =   3600
         Width           =   1290
      End
      Begin VB.Label lblDspDisAmtLoc 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         Height          =   300
         Left            =   -66960
         TabIndex        =   85
         Top             =   3240
         Width           =   1290
      End
      Begin VB.Label lblDspGrsAmtLoc 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         Height          =   300
         Left            =   -66960
         TabIndex        =   84
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label lblDspNetAmtOrg 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -70200
         TabIndex        =   83
         Top             =   3600
         Width           =   1290
      End
      Begin VB.Label lblNetAmtOrg 
         Caption         =   "NETAMTORG"
         Height          =   255
         Left            =   -72120
         TabIndex        =   82
         Top             =   3600
         Width           =   1515
      End
      Begin VB.Label lblDisAmtOrg 
         Caption         =   "DISAMTORG"
         Height          =   255
         Left            =   -72120
         TabIndex        =   81
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label lblDspDisAmtOrg 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         Height          =   300
         Left            =   -70200
         TabIndex        =   80
         Top             =   3240
         Width           =   1290
      End
      Begin VB.Label lblDspGrsAmtOrg 
         Alignment       =   1  '靠右對齊
         BorderStyle     =   1  '單線固定
         Caption         =   "9.999.999.999.99"
         Height          =   300
         Left            =   -70200
         TabIndex        =   79
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label lblGrsAmtOrg 
         Caption         =   "GRSAMTORG"
         Height          =   255
         Left            =   -72120
         TabIndex        =   78
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Label lblRmk 
         Caption         =   "RMK"
         Height          =   240
         Left            =   -74880
         TabIndex        =   66
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblRmkCode 
         Caption         =   "RMKCODE"
         Height          =   240
         Left            =   -74880
         TabIndex        =   65
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblLcNo 
         Caption         =   "LCNO"
         Height          =   240
         Left            =   -70440
         TabIndex        =   63
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label lblPortNo 
         Caption         =   "PORTNO"
         Height          =   240
         Left            =   -70440
         TabIndex        =   61
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lblCusPo 
         Caption         =   "CUSPO"
         Height          =   240
         Left            =   -70440
         TabIndex        =   59
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblShipAdr 
         Caption         =   "SHIPADR"
         Height          =   240
         Left            =   -74880
         TabIndex        =   52
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label lblShipPer 
         Caption         =   "SHIPPER"
         Height          =   240
         Left            =   -74880
         TabIndex        =   51
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label lblShipName 
         Caption         =   "SHIPNAME"
         Height          =   240
         Left            =   -70440
         TabIndex        =   49
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label lblShipTo 
         Caption         =   "SHIPTO"
         Height          =   240
         Left            =   -74880
         TabIndex        =   47
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label lblShipVia 
         Caption         =   "SHIPVIA"
         Height          =   240
         Left            =   -74880
         TabIndex        =   45
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lblEtaDate 
         Caption         =   "ETADATE"
         Height          =   255
         Left            =   -68640
         TabIndex        =   43
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lblOnDate 
         Caption         =   "ONDATE"
         Height          =   255
         Left            =   -71760
         TabIndex        =   41
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lblDueDate 
         Caption         =   "DUEDATE"
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   2460
         Width           =   1545
      End
      Begin VB.Label lblShipFrom 
         Caption         =   "SHIPFROM"
         Height          =   240
         Left            =   -74880
         TabIndex        =   37
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label lblDspPayDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -70800
         TabIndex        =   34
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lblPayCode 
         Caption         =   "PAYCODE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   33
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label lblDspPrcDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -70800
         TabIndex        =   32
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label lblPrcCode 
         Caption         =   "PRCCODE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   31
         Top             =   1020
         Width           =   1545
      End
      Begin VB.Label lblDspMLDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -70800
         TabIndex        =   30
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label lblDspVdrTypDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -70800
         TabIndex        =   29
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label lblShipCode 
         Caption         =   "SHIPCODE"
         Height          =   240
         Left            =   -74880
         TabIndex        =   28
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label lblMlCode 
         Caption         =   "MLCODE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   27
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label lblVdrTyp 
         Caption         =   "CUSTYP"
         Height          =   240
         Left            =   -74760
         TabIndex        =   26
         Top             =   1380
         Width           =   1545
      End
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   12
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
   Begin VB.Label lblVdrTel 
      Caption         =   "VDRTEL"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblVdrFax 
      Caption         =   "VDRFAX"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblDspVdrFax 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   4800
      TabIndex        =   17
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblVdrName 
      Caption         =   "VDRNAME"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label lblDspVdrTel 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblExcr 
      Caption         =   "EXCR"
      Height          =   255
      Left            =   7000
      TabIndex        =   14
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label LblCurr 
      Caption         =   "CURR"
      Height          =   255
      Left            =   7000
      TabIndex        =   13
      Top             =   1260
      Width           =   1200
   End
   Begin VB.Label lblDspVdrName 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label lblVdrCode 
      Caption         =   "VDRCODE"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label lblDocDate 
      Caption         =   "DOCDATE"
      Height          =   255
      Left            =   7000
      TabIndex        =   8
      Top             =   900
      Width           =   1200
   End
   Begin VB.Label lblRevNo 
      Caption         =   "REVNO"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   540
      Width           =   1215
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
      TabIndex        =   6
      Top             =   540
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
Attribute VB_Name = "frmPO001"
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




Private Const BOOKCODE = 0
Private Const BARCODE = 1
Private Const WhsCode = 2
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
Private Const SOID = 16


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
Private wlVdrID As Long
Private wlSoID As Long
Private wlVdrTyp As Long

Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "POPPOHD"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String





Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, BOOKCODE, SOID
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

    Call SetButtonStatus("AfrActEdit")
    Call SetFieldStatus("Default")
    Call SetFieldStatus("AfrActEdit")
    
    Call SetDateMask(medDocDate)
    Call SetDateMask(medDueDate)
    Call SetDateMask(medOnDate)
    Call SetDateMask(medETADate)
      
    
    wsOldVdrNo = ""
    wsOldCurCd = ""
    wsOldShipCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""

    
    wlKey = 0
    wlVdrID = 0
    wlSoID = 0
    
    
    wiRevNo = Format(0, "##0")
    tblCommon.Visible = False

    
    Me.Caption = wsFormCaption
    
    FocusMe cboDocNo
    tabDetailInfo.Tab = 0
    

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
        
        If getExcPRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) = False Then
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
            cboPayCode.SetFocus
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
    Call Ini_Combo(2, wsSql, cboCurr.Left, cboCurr.Top + cboCurr.Height, tblCommon, "PO001", "TBLCURCOD", Me.Width, Me.Height)
    
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
  
    wsSql = "SELECT POHDDOCNO, VdrCode, POHDDOCDATE "
    wsSql = wsSql & " FROM POPPOHD, mstVendor "
    wsSql = wsSql & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSql = wsSql & " AND POHDVDRID  = VDRID "
    wsSql = wsSql & " AND POHDSTATUS  <> '2' "
    wsSql = wsSql & " ORDER BY POHDDOCNO "
    Call Ini_Combo(3, wsSql, cboDocNo.Left, cboDocNo.Top + cboDocNo.Height, tblCommon, "PO001", "TBLDOCNO", Me.Width, Me.Height)
    
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
        cboDocNo.SetFocus
        Exit Function
    End If
    
        
   If Chk_PoHdDocNo(cboDocNo, wsStatus) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            Exit Function
        End If
        
        If wsStatus = "2" Then
            gsMsg = "文件已被刪除!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            cboDocNo.SetFocus
            Exit Function
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
        Me.Caption = wsFormCaption & " - Add"
        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        txtRevNo.Enabled = True
        wsOldVdrNo = cboVdrCode.Text
        wsOldCurCd = cboCurr.Text
        wsOldShipCd = cboShipCode.Text
        wsOldRmkCd = cboRmkCode.Text
        wsOldPayCd = cboPayCode.Text
        
    
        If UCase(cboCurr) = UCase(wsBaseCurCd) Then
            txtExcr.Text = Format("1", gsExrFmt)
            txtExcr.Enabled = False
        Else
            txtExcr.Enabled = True
        End If
        Call SetButtonStatus("AfrKeyEdit")
    End If
    
     Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    
    Call SetFieldStatus("AfrKey")
    
    cboVdrCode.SetFocus
        
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
            
        
        'Case vbKeyF5
        '    If wiAction = DefaultPage Then Call cmdEdit
       
        
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
    Dim wsSql As String
    Dim wsExcRate As String
    Dim wsExcDesc As String
    Dim wiCtr As Long
    
    LoadRecord = False
    
    If gsLangID = "1" Then
        wsSql = "SELECT POHDDOCID, POHDDOCNO, POHDVDRID, VDRID, VdrCode, VdrName, VdrTel, VdrFax, "
        wsSql = wsSql & "POHDDOCDATE, POHDREVNO, POHDCURR, POHDEXCR, "
        wsSql = wsSql & "POHDDUEDATE, POHDONDATE, POHDETADATE, POHDPAYCODE, POHDPRCCODE, POHDVDRTYP, POHDMLCODE, "
        wsSql = wsSql & "POHDCUSPO, POHDLCNO, POHDPORTNO, POHDSHIPPER, POHDSHIPFROM, POHDSHIPTO, POHDSHIPVIA, POHDSHIPNAME, "
        wsSql = wsSql & "POHDSHIPCODE, POHDSHIPADR1,  POHDSHIPADR2,  POHDSHIPADR3,  POHDSHIPADR4, "
        wsSql = wsSql & "POHDRMKCODE, POHDRMK1,  POHDRMK2,  POHDRMK3,  POHDRMK4, POHDRMK5, "
        wsSql = wsSql & "POHDRMK6,  POHDRMK7,  POHDRMK8,  POHDRMK9, POHDRMK10, "
        wsSql = wsSql & "POHDGRSAMT , POHDGRSAMTL, POHDDISAMT, POHDDISAMTL, POHDNETAMT, POHDNETAMTL, "
        wsSql = wsSql & "PODTSOID, PODTITEMID, ITMCODE, PODTWHSCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMPUBLISHER, PODTWANTED, PODTQTY, PODTUPRICE, PODTDISPER, PODTAMT, PODTAMTL, PODTDIS, PODTDISL, PODTNET, PODTNETL "
        wsSql = wsSql & "FROM  POPPOHD, POPPODT, mstVendor, mstITEM "
        wsSql = wsSql & "WHERE POHDDOCNO = '" & cboDocNo & "' "
        wsSql = wsSql & "AND POHDDOCID = PODTDOCID "
        wsSql = wsSql & "AND POHDVDRID = VDRID "
        wsSql = wsSql & "AND PODTITEMID = ITMID "
        wsSql = wsSql & "ORDER BY PODTDOCLINE "
        
    Else
        wsSql = "SELECT POHDDOCID, POHDDOCNO, POHDVDRID, VDRID, VdrCode, VdrName, VdrTel, VdrFax, "
        wsSql = wsSql & "POHDDOCDATE, POHDREVNO, POHDCURR, POHDEXCR, "
        wsSql = wsSql & "POHDDUEDATE, POHDONDATE, POHDETADATE, POHDPAYCODE, POHDPRCCODE, POHDVDRTYP, POHDMLCODE, "
        wsSql = wsSql & "POHDCUSPO, POHDLCNO, POHDPORTNO, POHDSHIPPER, POHDSHIPFROM, POHDSHIPTO, POHDSHIPVIA, POHDSHIPNAME, "
        wsSql = wsSql & "POHDSHIPCODE, POHDSHIPADR1,  POHDSHIPADR2,  POHDSHIPADR3,  POHDSHIPADR4, "
        wsSql = wsSql & "POHDRMKCODE, POHDRMK1,  POHDRMK2,  POHDRMK3,  POHDRMK4, POHDRMK5, "
        wsSql = wsSql & "POHDRMK6,  POHDRMK7,  POHDRMK8,  POHDRMK9, POHDRMK10, "
        wsSql = wsSql & "POHDGRSAMT , POHDGRSAMTL, POHDDISAMT, POHDDISAMTL, POHDNETAMT, POHDNETAMTL, "
        wsSql = wsSql & "PODTSOID, PODTITEMID, ITMCODE, PODTWHSCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMPUBLISHER, PODTWANTED, PODTQTY, PODTUPRICE, PODTDISPER, PODTAMT, PODTAMTL, PODTDIS, PODTDISL, PODTNET, PODTNETL "
        wsSql = wsSql & "FROM  POPPOHD, POPPODT, mstVendor, mstITEM "
        wsSql = wsSql & "WHERE POHDDOCNO = '" & cboDocNo & "' "
        wsSql = wsSql & "AND POHDDOCID = PODTDOCID "
        wsSql = wsSql & "AND POHDVDRID = VDRID "
        wsSql = wsSql & "AND PODTITEMID = ITMID "
        wsSql = wsSql & "ORDER BY PODTDOCLINE "
    End If
    
    rsInvoice.Open wsSql, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount <= 0 Then
        rsInvoice.Close
        Set rsInvoice = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsInvoice, "POHDDOCID")
    txtRevNo.Text = Format(ReadRs(rsInvoice, "POHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsInvoice, "POHDREVNO"))
    medDocDate.Text = ReadRs(rsInvoice, "POHDDOCDATE")
    wlVdrID = ReadRs(rsInvoice, "VDRID")
    cboVdrCode.Text = ReadRs(rsInvoice, "VdrCode")
    lblDspVdrName.Caption = ReadRs(rsInvoice, "VdrName")
    lblDspVdrTel.Caption = ReadRs(rsInvoice, "VdrTel")
    lblDspVdrFax.Caption = ReadRs(rsInvoice, "VdrFax")
    cboCurr.Text = ReadRs(rsInvoice, "POHDCURR")
    txtExcr.Text = Format(ReadRs(rsInvoice, "POHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "POHDDUEDATE"))
    medOnDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "POHDONDATE"))
    medETADate.Text = Dsp_MedDate(ReadRs(rsInvoice, "POHDETADATE"))
    
    wlVdrTyp = To_Value(ReadRs(rsInvoice, "POHDVDRTYP"))
    
    cboPayCode = ReadRs(rsInvoice, "POHDPAYCODE")
    cboPrcCode = ReadRs(rsInvoice, "POHDPRCCODE")
    cboMLCode = ReadRs(rsInvoice, "POHDMLCODE")
    cboShipCode = ReadRs(rsInvoice, "POHDSHIPCODE")
    cboRmkCode = ReadRs(rsInvoice, "POHDRMKCODE")
    
    txtCusPo = ReadRs(rsInvoice, "POHDCUSPO")
    txtLcNo = ReadRs(rsInvoice, "POHDLCNO")
    txtPortNo = ReadRs(rsInvoice, "POHDPORTNO")
    
    txtShipFrom = ReadRs(rsInvoice, "POHDSHIPFROM")
    txtShipTo = ReadRs(rsInvoice, "POHDSHIPTO")
    txtShipVia = ReadRs(rsInvoice, "POHDSHIPVIA")
    txtShipName = ReadRs(rsInvoice, "POHDSHIPNAME")
    txtShipPer = ReadRs(rsInvoice, "POHDSHIPPER")
    txtShipAdr1 = ReadRs(rsInvoice, "POHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsInvoice, "POHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsInvoice, "POHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsInvoice, "POHDSHIPADR4")
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsInvoice, "POHDRMK" & i)
    Next i
    
    
    cboVdrTyp.Text = Get_TableInfo("mstType", "TypID =" & wlVdrTyp, "TYPCODE")
    lblDspVdrTypDesc = Get_TableInfo("mstType", "TypID =" & wlVdrTyp, "TYPDESC")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    
    
    rsInvoice.MoveFirst
    With waResult
         .ReDim 0, -1, BOOKCODE, SOID
         Do While Not rsInvoice.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), BOOKCODE) = ReadRs(rsInvoice, "ITMCODE")
             waResult(.UpperBound(1), BARCODE) = ReadRs(rsInvoice, "ITMBARCODE")
             waResult(.UpperBound(1), BOOKNAME) = ReadRs(rsInvoice, "ITNAME")
             waResult(.UpperBound(1), WhsCode) = ReadRs(rsInvoice, "PODTWHSCODE")
             waResult(.UpperBound(1), PUBLISHER) = ReadRs(rsInvoice, "ITMPUBLISHER")
             waResult(.UpperBound(1), WANTED) = Dsp_MedDate(ReadRs(rsInvoice, "PODTWANTED"))
             waResult(.UpperBound(1), Qty) = Format(ReadRs(rsInvoice, "PODTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), Price) = Format(ReadRs(rsInvoice, "PODTUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsInvoice, "PODTDISPER"), "0.0")
             waResult(.UpperBound(1), Amt) = Format(ReadRs(rsInvoice, "PODTAMT"), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsInvoice, "PODTAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(ReadRs(rsInvoice, "PODTDIS"), gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(ReadRs(rsInvoice, "PODTDISL"), gsAmtFmt)
             waResult(.UpperBound(1), Net) = Format(ReadRs(rsInvoice, "PODTNET"), gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(ReadRs(rsInvoice, "PODTNETL"), gsAmtFmt)
             waResult(.UpperBound(1), BOOKID) = ReadRs(rsInvoice, "PODTITEMID")
             waResult(.UpperBound(1), SOID) = ReadRs(rsInvoice, "PODTSOID")
             rsInvoice.MoveNext
         Loop
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
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblRevNo.Caption = Get_Caption(waScrItm, "REVNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblVdrCode.Caption = Get_Caption(waScrItm, "VdrCode")
    lblVdrName.Caption = Get_Caption(waScrItm, "VdrName")
    lblVdrTel.Caption = Get_Caption(waScrItm, "VdrTel")
    lblVdrFax.Caption = Get_Caption(waScrItm, "VdrFax")
    LblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblPayCode.Caption = Get_Caption(waScrItm, "PAYCODE")
    lblPrcCode.Caption = Get_Caption(waScrItm, "PRCCODE")
    lblVdrTyp.Caption = Get_Caption(waScrItm, "VDRTYP")
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblDueDate.Caption = Get_Caption(waScrItm, "DUEDATE")
    lblOnDate.Caption = Get_Caption(waScrItm, "ONDATE")
    lblETADate.Caption = Get_Caption(waScrItm, "ETADATE")
    
    lblGrsAmtOrg.Caption = Get_Caption(waScrItm, "GRSAMTORG")
    lblNetAmtOrg.Caption = Get_Caption(waScrItm, "NETAMTORG")
    lblDisAmtOrg.Caption = Get_Caption(waScrItm, "DISAMTORG")
    
    lblGrsAmtLoc.Caption = Get_Caption(waScrItm, "GRSAMTLOC")
    lblNetAmtLoc.Caption = Get_Caption(waScrItm, "NETAMTLOC")
    lblDisAmtLoc.Caption = Get_Caption(waScrItm, "DISAMTLOC")
    
    With tblDetail
        .Columns(BOOKCODE).Caption = Get_Caption(waScrItm, "BOOKCODE")
        .Columns(BARCODE).Caption = Get_Caption(waScrItm, "BARCODE")
        .Columns(WhsCode).Caption = Get_Caption(waScrItm, "WHSCODE")
        .Columns(BOOKNAME).Caption = Get_Caption(waScrItm, "BOOKNAME")
        .Columns(WANTED).Caption = Get_Caption(waScrItm, "WANTED")
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
    tabDetailInfo.TabCaption(3) = Get_Caption(waScrItm, "TABDETAILINFO04")
    
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
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "POADD")
    wsActNam(2) = Get_Caption(waScrItm, "POEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PODELETE")
    
    Call Ini_PgmMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
  '  If Button = 2 Then
  '      PopupMenu mnuMaster
  '  End If

End Sub



Private Sub Form_Resize()
    If Me.WindowState = 0 Then
        Me.Height = 7020
        Me.Width = 9915
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If SaveData = True Then
        Cancel = True
        Exit Sub
    End If
    Call UnLockAll(wsConnTime, wsFormID)
    Set waResult = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmPO001 = Nothing

End Sub





Private Sub medDocDate_GotFocus()
    
  FocusMe medDocDate
    
End Sub


Private Sub medDocDate_LostFocus()

    FocusMe medDocDate, True
    
End Sub


Private Sub medDocDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medDocDate Then cboCurr.SetFocus
    End If
End Sub

Private Function Chk_medDocDate() As Boolean

    
    Chk_medDocDate = False
    
    If Trim(medDocDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        medDocDate.SetFocus
        Exit Function
    End If
    
    If Chk_Date(medDocDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
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
            medOnDate.SetFocus
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

Private Sub medOnDate_GotFocus()
    
  FocusMe medOnDate
    
End Sub


Private Sub medOnDate_LostFocus()

    FocusMe medOnDate, True
    
End Sub


Private Sub medOnDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medOnDate Then
        tabDetailInfo.Tab = 0
        medETADate.SetFocus
        End If
    End If
End Sub

Private Function Chk_medOnDate() As Boolean

    
    Chk_medOnDate = False
    
    If Trim(medOnDate.Text) = "/  /" Then
        Chk_medOnDate = True
        Exit Function
    End If
    
    If Chk_Date(medOnDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medOnDate.SetFocus
        Exit Function
    End If
    
    
    Chk_medOnDate = True

End Function


Private Sub medEtaDate_GotFocus()
    
  FocusMe medETADate
    
End Sub


Private Sub medEtaDate_LostFocus()

    FocusMe medETADate, True
    
End Sub


Private Sub medEtaDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medEtaDate Then
        tabDetailInfo.Tab = 1
        txtShipFrom.SetFocus
        End If
    End If
End Sub

Private Function Chk_medEtaDate() As Boolean

    
    Chk_medEtaDate = False
    
    If Trim(medETADate.Text) = "/  /" Then
        Chk_medEtaDate = True
        Exit Function
    End If
    
    If Chk_Date(medETADate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medETADate.SetFocus
        Exit Function
    End If
    
    
    Chk_medEtaDate = True

End Function



Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
        If cboPayCode.Enabled Then
            cboPayCode.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
    
        If txtShipFrom.Enabled Then
            txtShipFrom.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 2 Then
        
        If tblDetail.Enabled Then
            tblDetail.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 3 Then
    
        If cboRmkCode.Enabled Then
            cboRmkCode.SetFocus
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
    
    Dim rsPOPPOHD As New ADODB.Recordset
    Dim wsSql As String

    
    wsSql = "SELECT POHDSTATUS FROM POPPOHD WHERE POHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rsPOPPOHD.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    If rsPOPPOHD.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rsPOPPOHD.Close
    Set rsPOPPOHD = Nothing
    

End Function

Private Function Chk_KeyFld() As Boolean
    
        
    Chk_KeyFld = False
    
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
    
   
     
    
 '   If lblDspNetAmtLoc.Caption > Get_CreditLimit(wlVDRID, wlKey, Trim(medDocDate.Text)) Then
 '      gsMsg = "已超過信貸額!"
 '      MsgBox gsMsg, vbOKOnly, gsTitle
 '      MousePointer = vbDefault
 '      Exit Function
 '   End If
    
    
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_PO001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlVdrID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, txtRevNo.Text)
    Call SetSPPara(adcmdSave, 8, cboCurr.Text)
    Call SetSPPara(adcmdSave, 9, txtExcr.Text)
    Call SetSPPara(adcmdSave, 10, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medDueDate.Text))
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medOnDate.Text))
    Call SetSPPara(adcmdSave, 13, Set_MedDate(medETADate.Text))
    
    Call SetSPPara(adcmdSave, 14, wlVdrTyp)
    
    Call SetSPPara(adcmdSave, 15, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 16, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 17, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 18, cboShipCode.Text)
    Call SetSPPara(adcmdSave, 19, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 20, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 21, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 22, txtPortNo.Text)
    Call SetSPPara(adcmdSave, 23, "")
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
    Call SetSPPara(adcmdSave, 45, lblDspGrsAmtLoc)
    Call SetSPPara(adcmdSave, 46, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 47, lblDspDisAmtLoc)
    Call SetSPPara(adcmdSave, 48, lblDspNetAmtOrg)
    Call SetSPPara(adcmdSave, 49, lblDspNetAmtLoc)
    
    Call SetSPPara(adcmdSave, 50, wsFormID)
    
    Call SetSPPara(adcmdSave, 51, gsUserID)
    Call SetSPPara(adcmdSave, 52, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 53)
    wsDocNo = GetSPPara(adcmdSave, 54)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_PO001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, BOOKCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, waResult(wiCtr, SOID))
                Call SetSPPara(adcmdSave, 3, wlKey)
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, BOOKID))
                Call SetSPPara(adcmdSave, 5, wiCtr + 1)
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, BOOKNAME))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, Qty))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, Price))
                Call SetSPPara(adcmdSave, 9, waResult(wiCtr, DisPer))
                Call SetSPPara(adcmdSave, 10, Set_MedDate(waResult(wiCtr, WANTED)))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, WhsCode))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, Amt))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, Amtl))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, Dis))
                Call SetSPPara(adcmdSave, 15, waResult(wiCtr, Disl))
                Call SetSPPara(adcmdSave, 16, waResult(wiCtr, Net))
                Call SetSPPara(adcmdSave, 17, waResult(wiCtr, Netl))
                Call SetSPPara(adcmdSave, 18, IIf(wlRowCtr = wiCtr, "Y", "N"))
                adcmdSave.Execute
            End If
        Next
    End If
    cnCon.CommitTrans
    
    If wiAction = AddRec Then
    If Trim(wsDocNo) <> "" Then
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

Private Function InputValidation() As Boolean
    
    Dim wsExcRate As String
    Dim wsExcDesc As String

    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    
    
    If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboVdrCode() Then Exit Function
    If Not getExcPRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboPayCode Then Exit Function
    If Not Chk_cboPrcCode Then Exit Function
    If Not Chk_cboVDRTYP Then Exit Function
    If Not Chk_cboMLCode Then Exit Function
    
    If Not Chk_medDueDate Then Exit Function
    If Not Chk_medOnDate Then Exit Function
    If Not Chk_medEtaDate Then Exit Function
    
    If Not Chk_cboShipCode Then Exit Function
    If Not Chk_cboRmkCode Then Exit Function
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, BOOKCODE)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tblDetail.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    
    If wiEmptyGrid = True Then
        gsMsg = "採購單沒有詳細資料!"
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

    Dim newForm As New frmPO001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmPO001
    
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
    wsFormID = "PO001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "PO"
    
    


End Sub



Private Sub cmdCancel()
    
    Call Ini_Scr
    Call UnLockAll(wsConnTime, wsFormID)
    Call SetButtonStatus("AfrActEdit")
    Call SetButtonStatus("AfrActEdit")
  
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
           If MsgBox("你是否確定要放棄現時之作業?", vbYesNo, gsTitle) = vbYes Then
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
            cboPayCode.SetFocus
            End If
        End If
    End If

End Sub

Private Function chk_txtExcr() As Boolean
    
    chk_txtExcr = False
    
    If Trim(txtExcr.Text) = "" Or Trim(To_Value(txtExcr.Text)) = 0 Then
        gsMsg = "必需輸入對換率!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtExcr.SetFocus
        Exit Function
    End If
    
    If To_Value(txtExcr.Text) > 9999.999999 Then
        gsMsg = "對換率超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtExcr.SetFocus
        Exit Function
    End If
    txtExcr.Text = Format(txtExcr.Text, gsExrFmt)
    
    chk_txtExcr = True
    
End Function

Private Sub txtExcr_LostFocus()
FocusMe txtExcr, True
End Sub

Private Sub txtRevNo_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtRevNo.Text, False, False)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtRevNo Then
            medDocDate.SetFocus
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
        gsMsg = "修改號不正確!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        txtRevNo.SetFocus
        Exit Function
    End If
    
    chk_txtRevNo = True

End Function

Private Sub cboVdrCode_DropDown()
   
    Dim wsSql As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboVdrCode
    
    If gsLangID = "1" Then
        wsSql = "SELECT VdrCode, VdrName FROM mstVendor "
        wsSql = wsSql & "WHERE VdrCode LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
        wsSql = wsSql & "AND VDRSTATUS = '1' "
        wsSql = wsSql & "ORDER BY VdrCode "
    Else
        wsSql = "SELECT VdrCode, VdrName FROM mstVendor "
        wsSql = wsSql & "WHERE VdrCode LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
        wsSql = wsSql & "AND VDRSTATUS = '1' "
        wsSql = wsSql & "ORDER BY VdrCode "
    End If
    Call Ini_Combo(2, wsSql, cboVdrCode.Left, cboVdrCode.Top + cboVdrCode.Height, tblCommon, "PO001", "TBLVDRNO", Me.Width, Me.Height)
    
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
           
            cboCurr.SetFocus
            
    End If
    
End Sub

Private Function chk_cboVdrCode() As Boolean
    Dim wlID As Long
    Dim wsName As String
    Dim wsTel As String
    Dim wsFax As String
    
    
    chk_cboVdrCode = False
    
    
    If Trim(cboVdrCode) = "" Then
        gsMsg = "請輸入供應商編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
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
        cboVdrCode.SetFocus
        Exit Function
    End If
    
    chk_cboVdrCode = True

End Function

Private Sub Get_DefVal()
    
    Dim rsDefVal As New ADODB.Recordset
    Dim wsSql As String
    Dim wsExcDesc As String
    Dim wsExcRate As String
    Dim wsCode As String
    Dim wsName As String
    
    wsSql = "SELECT * "
    wsSql = wsSql & "FROM  mstVendor "
    wsSql = wsSql & "WHERE VDRID = " & wlVdrID
    rsDefVal.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        cboCurr.Text = ReadRs(rsDefVal, "CUSCURR")
        cboPayCode.Text = ReadRs(rsDefVal, "CUSPAYCODE")
        wlVdrTyp = To_Value(ReadRs(rsDefVal, "VDRTYPID"))
        txtShipName = ReadRs(rsDefVal, "CUSSHIPTO")
        txtShipPer = ReadRs(rsDefVal, "CUSSHIPCONTACTPERSON")
        txtShipAdr1 = ReadRs(rsDefVal, "CUSSHIPADD1")
        txtShipAdr2 = ReadRs(rsDefVal, "CUSSHIPADD2")
        txtShipAdr3 = ReadRs(rsDefVal, "CUSSHIPADD3")
        txtShipAdr4 = ReadRs(rsDefVal, "CUSSHIPADD4")
        
          Else
        cboCurr.Text = ""
        cboPayCode.Text = ""
        wlVdrTyp = 0
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
    If getExcPRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) = True Then
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
    
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    cboVdrTyp.Text = Get_TableInfo("mstType", "TypID =" & wlVdrTyp, "TYPCODE")
    lblDspVdrTypDesc = Get_TableInfo("mstType", "TypID =" & wlVdrTyp, "TYPDESC")
    
    
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
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = BOOKCODE To SOID
            .Columns(wiCtr).AllowSizing = False
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case BOOKCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 13
                Case BARCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Locked = True
                Case WhsCode
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                Case BOOKNAME
                    .Columns(wiCtr).Width = 2500
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = False
                Case WANTED
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = False
                    .Columns(wiCtr).EditMask = "####/##/##"
                Case PUBLISHER
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Locked = True
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
                    .Columns(wiCtr).Locked = False
                Case Net
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Dis
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                   ' .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case Amt
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
               
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
        .Styles("EvenRow").BackColor = &H8000000F
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
                .Columns(DisPer).Text = Format(wdDisPer, "0")
                .Columns(WANTED).Text = medETADate
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
            Case WANTED
                If Chk_grdWantedDate(.Columns(ColIndex).Text) = False Then
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
            Case BOOKCODE
                
                If gsLangID = 1 Then
                wsSql = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME ITNAME, ITMGRPENGNAME ITGRPNAM FROM mstITEM "
                wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
                wsSql = wsSql & " ORDER BY ITMCODE "
                Else
                wsSql = "SELECT ITMCODE, ITMBARCODE, ITMCHINAME ITNAME, ITMGRPCHINAME ITGRPNAM FROM mstITEM "
                wsSql = wsSql & " WHERE ITMSTATUS <> '2' AND ITMCODE LIKE '%" & Set_Quote(.Columns(BOOKCODE).Text) & "%' "
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
                Case BOOKCODE
                    KeyCode = vbDefault
                       .Col = WhsCode
                   ' KeyCode = vbKeyDown
                   ' .Col = BOOKCODE
                Case BOOKNAME, BARCODE, WhsCode, PUBLISHER, WANTED, Qty, DisPer, Amt, Dis
                    KeyCode = vbDefault
                    .Col = .Col + 1
                Case Price, Net, Amtl
                    KeyCode = vbKeyDown
                    .Col = BOOKCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> BOOKCODE Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> Net Then
                  .Col = .Col + 1
            End If
            
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
           .Col = BOOKCODE
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case BOOKCODE
                    Call Chk_grdBookCode(.Columns(BOOKCODE).Text, "", "", "", "", "", 0, 0)
                Case WhsCode
                    Call Chk_grdWhsCode(.Columns(WhsCode).Text)
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
        gsMsg = "沒有輸入書號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
        Exit Function
    End If
    

    
    If gsLangID = "1" Then
    wsSql = "SELECT ITMID, ITMCODE, ITMENGNAME ITNAME, ITMBARCODE, ITMPUBLISHER, ITMDEFAULTPRICE, ITMCURR  FROM mstITEM"
    wsSql = wsSql & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "' "
    Else
    wsSql = "SELECT ITMID, ITMCODE, ITMCHINAME ITNAME, ITMBARCODE, ITMPUBLISHER, ITMDEFAULTPRICE , ITMCURR  FROM mstITEM"
    wsSql = wsSql & " WHERE ITMCODE = '" & Set_Quote(inAccNo) & "' OR ITMBARCODE = '" & Set_Quote(inAccNo) & "' "
    
    End If
    
    rsDes.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
       outAccID = ReadRs(rsDes, "ITMID")
       outAccNo = ReadRs(rsDes, "ITMCODE")
       OutName = ReadRs(rsDes, "ITNAME")
       OutBarCode = ReadRs(rsDes, "ITMBARCODE")
       outPub = ReadRs(rsDes, "ITMPUBLISHER")
       outPrice = To_Value(ReadRs(rsDes, "ITMDEFAULTPRICE"))
       wsCurr = ReadRs(rsDes, "ITMCURR")
       
       wdPrice = getVdrItemPrice(wlVdrID, outAccID, cboCurr.Text)
       
       If wdPrice = 0 Then
       If cboCurr <> wsCurr Then
       If getExcPRate(wsCurr, medDocDate, wsExcr, "") = True Then
       outPrice = NBRnd(outPrice * To_Value(wsExcr) / txtExcr, giExrDp)
       End If
       End If
       Else
        outPrice = wdPrice
       End If
       
        outDisPer = 0
       
       Chk_grdBookCode = True
    Else
        outAccID = ""
        OutName = ""
        OutBarCode = ""
        outPub = ""
        outPrice = 0
        outDisPer = 0
        gsMsg = "沒有此書!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdBookCode = False
    End If
    rsDes.Close
    Set rsDes = Nothing

End Function

Private Function Chk_grdWhsCode(inNo As String) As Boolean
    
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
  
    
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
        gsMsg = "折扣必需大於零及小於一百!"
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
        gsMsg = "日期不正確!"
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
                   Trim(waResult(inRow, BOOKID)) = "" And _
                   Trim(waResult(inRow, SOID)) = "" Then
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
        
        If Chk_grdWhsCode(waResult(LastRow, WhsCode)) = False Then
                .Col = WhsCode
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
    
    Dim wiRowCtr As Integer
    
    Calc_Total = False
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotalGrs = wiTotalGrs + To_Value(waResult(wiRowCtr, Amt))
        wiTotalDis = wiTotalDis + To_Value(waResult(wiRowCtr, Dis))
        wiTotalNet = wiTotalNet + To_Value(waResult(wiRowCtr, Net))
    Next
    
    lblDspGrsAmtOrg.Caption = Format(CStr(wiTotalGrs), gsAmtFmt)
    lblDspGrsAmtLoc.Caption = Format(CStr(wiTotalGrs * To_Value(txtExcr)), gsAmtFmt)
    lblDspDisAmtOrg.Caption = Format(CStr(wiTotalDis), gsAmtFmt)
    lblDspDisAmtLoc.Caption = Format(CStr(wiTotalDis * To_Value(txtExcr)), gsAmtFmt)
    lblDspNetAmtOrg.Caption = Format(CStr(wiTotalNet), gsAmtFmt)
    lblDspNetAmtLoc.Caption = Format(CStr(wiTotalNet * To_Value(txtExcr)), gsAmtFmt)
    
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
    
    If ReadOnlyMode(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID) Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
    End If
    
    gsMsg = "你是不確定要刪除此檔案?"
    If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then
       wiAction = CorRec
       MousePointer = vbDefault
       Exit Function
    End If
    
    wiAction = DelRec
    
      cnCon.BeginTrans
    Set adcmdDelete.ActiveConnection = cnCon
        
    adcmdDelete.CommandText = "USP_PO001A"
    adcmdDelete.CommandType = adCmdStoredProc
    adcmdDelete.Parameters.Refresh
      
    Call SetSPPara(adcmdDelete, 1, wiAction)
    Call SetSPPara(adcmdDelete, 2, wsTrnCd)
    Call SetSPPara(adcmdDelete, 3, wlKey)
    Call SetSPPara(adcmdDelete, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdDelete, 5, wlVdrID)
    Call SetSPPara(adcmdDelete, 6, medDocDate.Text)
    Call SetSPPara(adcmdDelete, 7, txtRevNo.Text)
    Call SetSPPara(adcmdDelete, 8, cboCurr.Text)
    Call SetSPPara(adcmdDelete, 9, txtExcr.Text)
    Call SetSPPara(adcmdDelete, 10, "")
    
    Call SetSPPara(adcmdDelete, 11, Set_MedDate(medDueDate.Text))
    Call SetSPPara(adcmdDelete, 12, Set_MedDate(medOnDate.Text))
    Call SetSPPara(adcmdDelete, 13, Set_MedDate(medETADate.Text))
    
    Call SetSPPara(adcmdDelete, 14, wlVdrTyp)
    
    Call SetSPPara(adcmdDelete, 15, cboPayCode.Text)
    Call SetSPPara(adcmdDelete, 16, cboPrcCode.Text)
    Call SetSPPara(adcmdDelete, 17, cboMLCode.Text)
    Call SetSPPara(adcmdDelete, 18, cboShipCode.Text)
    Call SetSPPara(adcmdDelete, 19, cboRmkCode.Text)
    
    Call SetSPPara(adcmdDelete, 20, txtCusPo.Text)
    Call SetSPPara(adcmdDelete, 21, txtLcNo.Text)
    Call SetSPPara(adcmdDelete, 22, txtPortNo.Text)
    Call SetSPPara(adcmdDelete, 23, "")
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
    Call SetSPPara(adcmdDelete, 45, lblDspGrsAmtLoc)
    Call SetSPPara(adcmdDelete, 46, lblDspDisAmtOrg)
    Call SetSPPara(adcmdDelete, 47, lblDspDisAmtLoc)
    Call SetSPPara(adcmdDelete, 48, lblDspNetAmtOrg)
    Call SetSPPara(adcmdDelete, 49, lblDspNetAmtLoc)
    
    Call SetSPPara(adcmdDelete, 50, wsFormID)
    
    Call SetSPPara(adcmdDelete, 51, gsUserID)
    Call SetSPPara(adcmdDelete, 52, wsGenDte)
    adcmdDelete.Execute
    wlKey = GetSPPara(adcmdDelete, 53)
    wsDocNo = GetSPPara(adcmdDelete, 54)
    
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
            Me.cboVdrCode.Enabled = False
            Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            
            Me.medDueDate.Enabled = False
            Me.medOnDate.Enabled = False
            Me.medETADate.Enabled = False
            Me.cboPayCode.Enabled = False
            Me.cboPrcCode.Enabled = False
            Me.cboMLCode.Enabled = False
            Me.cboVdrTyp.Enabled = False
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
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            Me.cboVdrCode.Enabled = True
            Me.txtRevNo.Enabled = True
            Me.medDocDate.Enabled = True
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.medDueDate.Enabled = True
            Me.medOnDate.Enabled = True
            Me.medETADate.Enabled = True
            Me.cboPayCode.Enabled = True
            Me.cboPrcCode.Enabled = True
            Me.cboMLCode.Enabled = True
            Me.cboVdrTyp.Enabled = True
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
        .TableKey = "PoHdDocNo"
        .KeyLen = 15
        Set .ctlKey = cboDocNo
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub


Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim wsSql As String
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "Doc No."
    vFilterAry(1, 2) = "POHDDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "POHDDocDate"
    
    vFilterAry(3, 1) = "Vendor #"
    vFilterAry(3, 2) = "VdrCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "POHDDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "POHDDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Vendor#"
    vAry(3, 2) = "VdrCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Vendor Name"
    vAry(4, 2) = "VdrName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSql = "SELECT POPPOHD.POHDDocNo, POPPOHD.POHDDocDate, mstVendor.VdrCode,  mstVendor.VdrName "
        wsSql = wsSql + "FROM mstVendor, POPPOHD "
        .sBindSQL = wsSql
        .sBindWhereSQL = "WHERE POPPOHD.POHDStatus = '1' And POPPOHD.POHDVDRID = mstVendor.VDRID "
        .sBindOrderSQL = "ORDER BY POPPOHD.POHDDocNo"
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
    
End Sub







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
    Call Ini_Combo(2, wsSql, cboPayCode.Left + tabDetailInfo.Left, cboPayCode.Top + cboPayCode.Height + tabDetailInfo.Top, tblCommon, "PO001", "TBLPAYCOD", Me.Width, Me.Height)
    
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
        cboVdrTyp.SetFocus
       
    End If
    
End Sub

Private Sub cboPrcCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboPrcCode
    
    wsSql = "SELECT PrcCode, PRCDESC FROM mstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCode.SelLength > 0, "", Set_Quote(cboPrcCode.Text)) & "%' "
    wsSql = wsSql & "AND PRCSTATUS = '1' "
    wsSql = wsSql & "ORDER BY PrcCode "
    Call Ini_Combo(2, wsSql, cboPrcCode.Left + tabDetailInfo.Left, cboPrcCode.Top + cboPrcCode.Height + tabDetailInfo.Top, tblCommon, "PO001", "TBLPRCCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboPrcCode() As Boolean
Dim wsDesc As String

    Chk_cboPrcCode = False
     
    If Trim(cboPrcCode.Text) = "" Then
        gsMsg = "沒有此銷售條款!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_PriceTerm(cboPrcCode, wsDesc) = False Then
        gsMsg = "必需輸入銷售條款!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboPrcCode.SetFocus
        lblDspPrcDesc = ""
       Exit Function
    End If
    
    lblDspPrcDesc = wsDesc
    Chk_cboPrcCode = True
    
End Function



Private Sub cboVDRTYP_GotFocus()
    FocusMe cboVdrTyp
End Sub

Private Sub cboVDRTYP_LostFocus()
    FocusMe cboVdrTyp, True
End Sub


Private Sub cboVDRTYP_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboVdrTyp, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboVDRTYP = False Then
                Exit Sub
        End If
        
        tabDetailInfo.Tab = 0
        cboMLCode.SetFocus
       
    End If
    
End Sub

Private Sub cboVDRTYP_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrTyp
    
    wsSql = "SELECT TYPCODE, TYPDESC FROM mstType WHERE TypCode LIKE '%" & IIf(cboVdrTyp.SelLength > 0, "", Set_Quote(cboVdrTyp.Text)) & "%' "
    wsSql = wsSql & "AND TYPCLASS = '2' "
    wsSql = wsSql & "AND TYPSTATUS = '1' "
    wsSql = wsSql & "ORDER BY TypCode "
    Call Ini_Combo(2, wsSql, cboVdrTyp.Left + tabDetailInfo.Left, cboVdrTyp.Top + cboVdrTyp.Height + tabDetailInfo.Top, tblCommon, "PO001", "TBLVDRTYP", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboVDRTYP() As Boolean
Dim wsDesc As String

    Chk_cboVDRTYP = False
     
    If Trim(cboVdrTyp.Text) = "" Then
        gsMsg = "必需輸入供應商類別!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboVdrTyp.SetFocus
        Exit Function
    End If
    
    
    If Chk_Type(cboVdrTyp, "2", wlVdrTyp, wsDesc) = False Then
        gsMsg = "沒有此供應商類別!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboVdrTyp.SetFocus
        lblDspVdrTypDesc = ""
       Exit Function
    End If
    
    lblDspVdrTypDesc = wsDesc
    
    Chk_cboVDRTYP = True
    
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
        medDueDate.SetFocus
       
    End If
    
End Sub

Private Sub cboMLCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboMLCode
    
    wsSql = "SELECT MLCode, MLDESC FROM mstMerchClass WHERE MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSql = wsSql & "AND MLSTATUS = '1' "
    wsSql = wsSql & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSql, cboMLCode.Left + tabDetailInfo.Left, cboMLCode.Top + cboMLCode.Height + tabDetailInfo.Top, tblCommon, "PO001", "TBLMLCOD", Me.Width, Me.Height)
    
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









Private Sub txtShipFrom_GotFocus()
        FocusMe txtShipFrom
End Sub

Private Sub txtShipFrom_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtShipFrom, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
        cboShipCode.SetFocus
       
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
        
        
        tabDetailInfo.Tab = 1
        txtShipName.SetFocus
       
    End If
    
End Sub

Private Sub cboShipCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboShipCode
    
    wsSql = "SELECT ShipCode, ShipName, ShipPer FROM mstShip WHERE ShipCode LIKE '%" & IIf(cboShipCode.SelLength > 0, "", Set_Quote(cboShipCode.Text)) & "%' "
    wsSql = wsSql & "AND ShipSTATUS = '1' "
    wsSql = wsSql & "AND ShipCardID = " & wlVdrID & " "
    wsSql = wsSql & "ORDER BY ShipCode "
    Call Ini_Combo(3, wsSql, cboShipCode.Left + tabDetailInfo.Left, cboShipCode.Top + cboShipCode.Height + tabDetailInfo.Top, tblCommon, "PO001", "TBLSHIPCOD", Me.Width, Me.Height)
    
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
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
        txtShipPer.SetFocus
       
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
        
        
        tabDetailInfo.Tab = 1
        txtShipAdr1.SetFocus
       
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
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
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
        
        
        tabDetailInfo.Tab = 1
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
        tblDetail.SetFocus
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
        
        tabDetailInfo.Tab = 3
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
    Call Ini_Combo(1, wsSql, cboRmkCode.Left + tabDetailInfo.Left, cboRmkCode.Top + cboRmkCode.Height + tabDetailInfo.Top, tblCommon, "PO001", "TBLRMKCOD", Me.Width, Me.Height)
    
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
        tabDetailInfo.Tab = 3
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
        cboPayCode.SetFocus
        Else
        tabDetailInfo.Tab = 3
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
    Chk_NoDup = False
    
    wsCurRec = tblDetail.Columns(BOOKCODE)
    wsCurRecLn = tblDetail.Columns(WhsCode)
   
        For wlCtr = 0 To waResult.UpperBound(1)
            If inRow <> wlCtr Then
               If wsCurRec = waResult(wlCtr, BOOKCODE) And _
                  wsCurRecLn = waResult(wlCtr, WhsCode) Then
                  gsMsg = "重覆書本!"
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
