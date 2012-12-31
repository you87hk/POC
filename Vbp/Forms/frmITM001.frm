VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmITM001 
   BackColor       =   &H8000000A&
   Caption         =   "ITEM"
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10875
   Icon            =   "frmITM001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   10875
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10800
      OleObjectBlob   =   "frmITM001.frx":08CA
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboItemClassCode 
      Height          =   300
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   1425
   End
   Begin VB.ComboBox cboItmCode 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   2610
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   960
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
            Picture         =   "frmITM001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":6945
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":6C6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmITM001.frx":6F91
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraHeaderInfo 
      Caption         =   "HEADERINFO"
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   10695
      Begin VB.TextBox txtItmEngName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   9015
      End
      Begin VB.TextBox txtItmChiName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   9015
      End
      Begin VB.TextBox txtItmCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   0
         Tag             =   "K"
         Top             =   240
         Width           =   2610
      End
      Begin VB.Label lblDspItemClassCode 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6885
         TabIndex        =   41
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label lblItemClassCode 
         Caption         =   "ITEMCLASSCODE"
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
         Left            =   4320
         TabIndex        =   40
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblItmEngName 
         Caption         =   "ITMENGNAME"
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   1005
         Width           =   1380
      End
      Begin VB.Label lblItmChiName 
         Caption         =   "ITMCHINAME"
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   645
         Width           =   1305
      End
      Begin VB.Label lblItmCode 
         Caption         =   "國際書號 :"
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
         TabIndex        =   27
         Top             =   300
         Width           =   1380
      End
   End
   Begin MSComDlg.CommonDialog cdlgDir 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "iglProcess"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "開新視窗 (F8)"
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
            Key             =   "Key"
            Object.ToolTipText     =   "Change Key (F8)"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "退出 (F12)"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   4575
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8070
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "索書資料"
      TabPicture(0)   =   "frmITM001.frx":72AD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPrice"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboItmAccTypeCode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboItmCurr"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboItmUOMCode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboItmTypeCode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboItmPVdrCode"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "附加資料"
      TabPicture(1)   =   "frmITM001.frx":72C9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContent"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "BOM"
      TabPicture(2)   =   "frmITM001.frx":72E5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "chkOwnEdition"
      Tab(2).Control(2)=   "tblDetail"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame3 
         Height          =   450
         Left            =   -74880
         TabIndex        =   62
         Top             =   120
         Width           =   6135
         Begin VB.Label lblKeyDesc 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   360
            TabIndex        =   66
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblComboPrompt 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   1920
            TabIndex        =   65
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblInsertLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   3360
            TabIndex        =   64
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblDeleteLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   4800
            TabIndex        =   63
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkOwnEdition 
         Alignment       =   1  '靠右對齊
         Caption         =   "OWNEDITION"
         Height          =   180
         Left            =   -66360
         TabIndex        =   60
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboItmPVdrCode 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmITM001.frx":7301
         Left            =   6840
         List            =   "frmITM001.frx":7303
         TabIndex        =   12
         Top             =   945
         Width           =   1785
      End
      Begin VB.ComboBox cboItmTypeCode 
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   3705
      End
      Begin VB.Frame fraContent 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   32
         Top             =   240
         Width           =   9855
         Begin VB.TextBox txtItmMaxQty 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   23
            Top             =   600
            Width           =   945
         End
         Begin VB.TextBox txtItmPORepuQty 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   24
            Top             =   960
            Width           =   945
         End
         Begin VB.CheckBox chkItmReorderInd 
            Alignment       =   1  '靠右對齊
            Caption         =   "再版指標 :"
            Height          =   180
            Left            =   240
            TabIndex        =   21
            Top             =   1010
            Width           =   1455
         End
         Begin VB.CheckBox chkItmInvItemFlg 
            Alignment       =   1  '靠右對齊
            Caption         =   "非存貨 :"
            Height          =   180
            Left            =   240
            TabIndex        =   20
            Top             =   650
            Width           =   1455
         End
         Begin VB.CheckBox chkItmInActive 
            Alignment       =   1  '靠右對齊
            Caption         =   "暫停發貨 :"
            Height          =   180
            Left            =   240
            TabIndex        =   19
            Top             =   290
            Width           =   1455
         End
         Begin VB.TextBox txtItmReorderQty 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   22
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblDspStkOnHand 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   8400
            TabIndex        =   76
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label lblStkOnHand 
            Caption         =   "ICTRNQTY"
            Height          =   240
            Left            =   6000
            TabIndex        =   75
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label lblDspStkIndent 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   8400
            TabIndex        =   74
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label lblStkIndent 
            Caption         =   "ICTRNQTY"
            Height          =   240
            Left            =   6000
            TabIndex        =   73
            Top             =   1320
            Width           =   2115
         End
         Begin VB.Label lblDspStkOnOrder 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   8400
            TabIndex        =   72
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label lblStkOnOrder 
            Caption         =   "ICTRNQTY"
            Height          =   240
            Left            =   6000
            TabIndex        =   71
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label lblDspStkAllocated 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   8400
            TabIndex        =   70
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lblStkAllocated 
            Caption         =   "ICTRNQTY"
            Height          =   240
            Left            =   6000
            TabIndex        =   69
            Top             =   600
            Width           =   2115
         End
         Begin VB.Label lblDspStkAvailable 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   8400
            TabIndex        =   68
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label lblStkAvailable 
            Caption         =   "ICTRNQTY"
            Height          =   240
            Left            =   6000
            TabIndex        =   67
            Top             =   1680
            Width           =   2115
         End
         Begin VB.Label lblItmMaxQty 
            Caption         =   "再版指標 :"
            Height          =   240
            Left            =   2160
            TabIndex        =   39
            Top             =   637
            Width           =   1140
         End
         Begin VB.Label lblItmPORepuQty 
            Caption         =   "再版數量 :"
            Height          =   240
            Left            =   2160
            TabIndex        =   34
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label lblItmReorderQty 
            Caption         =   "再版指標 :"
            Height          =   240
            Left            =   2160
            TabIndex        =   33
            Top             =   315
            Width           =   1140
         End
      End
      Begin VB.ComboBox cboItmUOMCode 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   3705
      End
      Begin VB.ComboBox cboItmCurr 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmITM001.frx":7305
         Left            =   6840
         List            =   "frmITM001.frx":7307
         TabIndex        =   11
         Top             =   480
         Width           =   1785
      End
      Begin VB.ComboBox cboItmAccTypeCode 
         Height          =   300
         Left            =   1440
         TabIndex        =   10
         Top             =   3000
         Width           =   3705
      End
      Begin VB.Frame fraPrice 
         Height          =   3855
         Left            =   5400
         TabIndex        =   52
         Top             =   120
         Width           =   5055
         Begin VB.TextBox txtItmUnitPrice 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1440
            TabIndex        =   13
            Top             =   1320
            Width           =   1785
         End
         Begin VB.TextBox txtItmDefaultPrice 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1440
            TabIndex        =   17
            Top             =   3225
            Width           =   1785
         End
         Begin VB.TextBox txtItmDiscount 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1440
            TabIndex        =   14
            Top             =   1800
            Width           =   1785
         End
         Begin VB.TextBox txtItmMarkUp 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1440
            TabIndex        =   16
            Top             =   2745
            Width           =   1785
         End
         Begin VB.TextBox txtItmBottomPrice 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1440
            TabIndex        =   15
            Top             =   2265
            Width           =   1785
         End
         Begin VB.CommandButton btnItemPrice 
            Caption         =   "ITEMPRICE"
            Enabled         =   0   'False
            Height          =   555
            Left            =   3360
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblItmDefaultPrice 
            Caption         =   "ITMDEFAULTPRICE"
            Height          =   240
            Left            =   360
            TabIndex        =   59
            Top             =   3300
            Width           =   1020
         End
         Begin VB.Label lblItmDiscount 
            Caption         =   "ITMDISCOUNT"
            Height          =   240
            Left            =   360
            TabIndex        =   58
            Top             =   1860
            Width           =   1020
         End
         Begin VB.Label lblItmPVdrCode 
            Caption         =   "ITMPVDRID"
            Height          =   240
            Left            =   360
            TabIndex        =   57
            Top             =   885
            Width           =   1020
         End
         Begin VB.Label lblItmMarkUp 
            Caption         =   "ITMMARKUP"
            Height          =   240
            Left            =   360
            TabIndex        =   56
            Top             =   2805
            Width           =   1020
         End
         Begin VB.Label lblUnitPrice 
            Caption         =   "UNITPRICE"
            Height          =   240
            Left            =   360
            TabIndex        =   55
            Top             =   1365
            Width           =   1020
         End
         Begin VB.Label lblItmCurrCode 
            Caption         =   "貨幣 :"
            Height          =   240
            Left            =   360
            TabIndex        =   54
            Top             =   420
            Width           =   1020
         End
         Begin VB.Label lblItmBottomPrice 
            Caption         =   "BOTTOMPRICE"
            Height          =   240
            Left            =   360
            TabIndex        =   53
            Top             =   2345
            Width           =   1020
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   3855
         Left            =   240
         TabIndex        =   42
         Top             =   120
         Width           =   5055
         Begin VB.TextBox txtItmBarCode 
            Height          =   300
            Left            =   1200
            TabIndex        =   7
            Top             =   1785
            Width           =   3690
         End
         Begin VB.TextBox txtItmBinNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   9
            Top             =   2505
            Width           =   3690
         End
         Begin VB.TextBox txtItmSeriesNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   8
            Top             =   2145
            Width           =   3705
         End
         Begin VB.Label lblDspItmUomDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   4785
         End
         Begin VB.Label lblItmBarCode 
            Caption         =   "條碼編號 :"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1830
            Width           =   1095
         End
         Begin VB.Label lblItmBinNo 
            Caption         =   "ITMBINNO"
            Height          =   240
            Left            =   120
            TabIndex        =   49
            Top             =   2565
            Width           =   900
         End
         Begin VB.Label lblItmUomCode 
            Caption         =   "PACKTYPE"
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label lblDspItmAccTypeDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   120
            TabIndex        =   47
            Top             =   3240
            Width           =   4785
         End
         Begin VB.Label lblItmSeriesNo 
            Caption         =   "ITMSERIESNO"
            Height          =   240
            Left            =   120
            TabIndex        =   46
            Top             =   2220
            Width           =   1020
         End
         Begin VB.Label lblDspItmTypeDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   4785
         End
         Begin VB.Label lblItmTypeCode 
            Caption         =   "ITMTYPE"
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   420
            Width           =   1020
         End
         Begin VB.Label lblItmAccTypeCode 
            Caption         =   "ITMACCTYPE"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   2880
            Width           =   1260
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   3375
         Left            =   -74880
         OleObjectBlob   =   "frmITM001.frx":7309
         TabIndex        =   61
         Top             =   720
         Width           =   10455
      End
   End
   Begin VB.Label lblItmLastUpd 
      Caption         =   "最後修改人 :"
      Height          =   240
      Left            =   120
      TabIndex        =   38
      Top             =   6720
      Width           =   1380
   End
   Begin VB.Label lblItmLastUpdDate 
      Caption         =   "最後修改日期 :"
      Height          =   240
      Left            =   6120
      TabIndex        =   37
      Top             =   6720
      Width           =   1380
   End
   Begin VB.Label lblDspItmLastUpd 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   1680
      TabIndex        =   36
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label lblDspItmLastUpdDate 
      BorderStyle     =   1  '單線固定
      Height          =   300
      Left            =   7680
      TabIndex        =   35
      Top             =   6720
      Width           =   2895
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
Attribute VB_Name = "frmITM001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private wsFormCaption As String
Private waResult As New XArrayDB
Private waPopUpSub As New XArrayDB
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
Private Const tcKey = "Key"
Private Const tcCopy = "Copy"

Private wsActNam(4) As String
Private wiAction As Integer
Private wlKey As Long
Private wbAfrKey As Boolean

Private Const ITMCODE = 0
Private Const ITMDESC = 1
Private Const QTY = 2
Private Const ITMID = 3

Dim wcCombo As Control

Private Const wsKeyType = "MstItem"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private wdOldPrice As Double
Private wbErr As Boolean
Private wlVdrID As Long

Private Sub btnItemPrice_Click()

    
    Me.MousePointer = vbHourglass
    frmIP001.ITMCODE = cboITMCODE
    frmIP001.Show
    Me.MousePointer = vbNormal
End Sub

Private Sub btnPriceChange_Click()
    'frmB0011.InBookName = Me.txtItmChiName
    'frmB0011.inISBN = Me.cboItmCode
    'frmB0011.InItemID = wlKey
    'frmB0011.Show vbModal
End Sub

Private Sub cboItmPVdrCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmPVdrCode
    
    wsSQL = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrStatus = '1'"
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " AND VdrCode LIKE '%" & IIf(cboItmPVdrCode.SelLength > 0, "", Set_Quote(cboItmPVdrCode.Text)) & "%' "
   
    wsSQL = wsSQL & "ORDER BY VdrCode "
    Call Ini_Combo(2, wsSQL, cboItmPVdrCode.Left + Me.tabDetailInfo.Left, cboItmPVdrCode.Top + cboItmPVdrCode.Height + Me.tabDetailInfo.Top, tblCommon, wsFormID, "TBLV", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmPVdrCode_GotFocus()
    FocusMe cboItmPVdrCode
End Sub

Private Sub cboItmPVdrCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmPVdrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboVdrCode() = True Then
            txtItmUnitPrice.SetFocus
        End If
    End If
End Sub

Private Sub cboItmPVdrCode_LostFocus()
    FocusMe cboItmPVdrCode, True
End Sub

Private Sub cboItmUOMCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmUOMCode
    
    wsSQL = "SELECT UOMCode, UOMDesc FROM MstUOM WHERE UOMStatus = '1'"
    wsSQL = wsSQL & "ORDER BY UOMCode "
    Call Ini_Combo(2, wsSQL, cboItmUOMCode.Left + tabDetailInfo.Left, cboItmUOMCode.Top + cboItmUOMCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLUOM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmUOMCode_GotFocus()
    FocusMe cboItmUOMCode
End Sub

Private Sub cboItmUOMCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmUOMCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmUOMCode() = False Then
            Exit Sub
        End If
           
        tabDetailInfo.Tab = 0
        txtItmBarCode.SetFocus
    End If
End Sub

Private Sub cboItmUOMCode_LostFocus()
    FocusMe cboItmUOMCode, True
End Sub

Private Sub chkItmInActive_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        chkItmInvItemFlg.SetFocus
    End If
End Sub

Private Sub chkItmInvItemFlg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        chkItmReorderInd.SetFocus
    End If
End Sub



Private Sub chkItmReorderInd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        txtItmReorderQty.SetFocus
        
    End If
End Sub

Private Sub chkOwnEdition_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If tblDetail.Enabled = True Then
            tabDetailInfo.Tab = 2
            tblDetail.SetFocus
        Else
            txtItmChiName.SetFocus
        End If
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        Case vbKeyPageDown
            KeyCode = 0
            If tabDetailInfo.Tab < tabDetailInfo.Tabs - 1 Then
                If tabDetailInfo.TabVisible(tabDetailInfo.Tab + 1) = True Then
                    tabDetailInfo.Tab = tabDetailInfo.Tab + 1
                End If
                Exit Sub
            End If
        Case vbKeyPageUp
            KeyCode = 0
            If tabDetailInfo.Tab > 0 Then
                If tabDetailInfo.TabVisible(tabDetailInfo.Tab - 1) = True Then
                tabDetailInfo.Tab = tabDetailInfo.Tab - 1
                End If
                Exit Sub
            End If
        
        Case vbKeyF7
            If tbrProcess.Buttons(tcKey).Enabled = True Then
                Call cmdChangeKey(CorRec)
            End If
            
        
        Case vbKeyF2
            If wiAction = DefaultPage Then Call cmdNew
            
        
        Case vbKeyF5
            If wiAction = DefaultPage Then Call cmdEdit
       
        
        Case vbKeyF3
            If wiAction = DefaultPage Then Call cmdDel
        
        Case vbKeyF9
        
        If tbrProcess.Buttons(tcFind).Enabled = True Then
            Call cmdFind
        End If
            
        Case vbKeyF10
        If tbrProcess.Buttons(tcSave).Enabled = True Then
            Call cmdSave
        End If
            
        Case vbKeyF11
        
            If wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec Then Call cmdCancel
        
        Case vbKeyF12
        
            Unload Me
            
    End Select
End Sub

Private Sub Form_Load()
    Dim iCounter As Integer
    Dim iTabs As Integer
    Dim vToolTip As Variant
    
    MousePointer = vbHourglass
  
    wsFormCaption = Me.Caption
    
    IniForm
    Ini_Caption
    Ini_Grid
    Ini_Scr
    
    MousePointer = vbDefault
  
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 7500
        Me.Width = 11000
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
                .Buttons(tcKey).Enabled = False
                .Buttons(tcCopy).Enabled = False
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
                .Buttons(tcFind).Enabled = True
                .Buttons(tcExit).Enabled = True
                .Buttons(tcKey).Enabled = False
                .Buttons(tcCopy).Enabled = False
                
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
                .Buttons(tcKey).Enabled = False
                .Buttons(tcCopy).Enabled = False
                
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
                .Buttons(tcKey).Enabled = False
                .Buttons(tcCopy).Enabled = False
                
            End With
        
        Case "AfrKeyEdit"
            With tbrProcess
                .Buttons(tcOpen).Enabled = True
                .Buttons(tcAdd).Enabled = False
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
                .Buttons(tcKey).Enabled = True
                .Buttons(tcCopy).Enabled = True
                
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
                .Buttons(tcKey).Enabled = False
                .Buttons(tcCopy).Enabled = False
                
            End With
    End Select
End Sub

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
Dim i As Integer
    Select Case sStatus
        Case "Default"
            Me.cboITMCODE.Enabled = False
            Me.cboItemClassCode.Enabled = False
            Me.txtItmCode.Enabled = False
            Me.txtItmChiName.Enabled = False
            Me.txtItmEngName.Enabled = False
            
            'Tab 0 fields
            Me.cboItmTypeCode.Enabled = False
            Me.cboItmUOMCode.Enabled = False
            Me.txtItmBarCode.Enabled = False
            Me.txtItmSeriesNo.Enabled = False
            Me.txtItmBinNo.Enabled = False
            Me.cboItmAccTypeCode.Enabled = False
            
            Me.cboItmCurr.Enabled = False
            Me.cboItmPVdrCode.Enabled = False
            
            Me.txtItmUnitPrice.Enabled = False
            Me.txtItmDiscount.Enabled = False
            Me.txtItmBottomPrice.Enabled = False
            Me.txtItmMarkUp.Enabled = False
            Me.txtItmDefaultPrice.Enabled = False
            
            Me.btnItemPrice.Enabled = False
            
            fraInfo.Visible = True
            fraPrice.Visible = True
            cboItmAccTypeCode.Visible = True
            cboItmCurr.Visible = True
            cboItmPVdrCode.Visible = True
            cboItmTypeCode.Visible = True
            cboItmUOMCode.Visible = True
            
            'Tab 1 fields
            Me.chkItmInActive.Enabled = False
            Me.chkItmInvItemFlg.Enabled = False
            Me.chkItmReorderInd.Enabled = False
            
            Me.txtItmReorderQty.Enabled = False
            Me.txtItmMaxQty.Enabled = False
            Me.txtItmPORepuQty.Enabled = False
            
            fraContent.Visible = True
            
            'Tab 2 fields
            Me.chkOwnEdition.Enabled = False
            Me.tblDetail.Enabled = False
            
            tabDetailInfo.TabVisible(2) = True
            
        Case "AfrActAdd"
            Me.cboITMCODE.Enabled = False
            Me.cboITMCODE.Visible = False
            
            Me.txtItmCode.Enabled = True
            Me.txtItmCode.Visible = True
            
            Me.cboItemClassCode.Enabled = True
            
            
       Case "AfrActEdit"
            Me.cboITMCODE.Enabled = True
            Me.cboITMCODE.Visible = True
            
            Me.txtItmCode.Enabled = False
            Me.txtItmCode.Visible = False
            
            'Me.cboItemClassCode.Enabled = True
            Me.cboItemClassCode.Enabled = False
            
        Case "AfrKey"
            Me.txtItmCode.Enabled = False
            Me.cboITMCODE.Enabled = False
            Me.cboItemClassCode.Enabled = False
        '    Me.cboItemClassCode.Enabled = True
            
            Me.txtItmChiName.Enabled = True
            Me.txtItmEngName.Enabled = True
            
            Select Case UCase(cboItemClassCode)
                Case "P"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    
                    Me.txtItmUnitPrice.Enabled = True
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                    
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = True
                    Me.chkItmInvItemFlg.Enabled = True
                    Me.chkItmReorderInd.Enabled = True
                    
                    Me.txtItmReorderQty.Enabled = True
                    Me.txtItmMaxQty.Enabled = True
                    Me.txtItmPORepuQty.Enabled = True
                    
                    fraContent.Visible = True
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                    
                    tabDetailInfo.TabVisible(2) = False
                        
                Case "N"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    
                    Me.txtItmUnitPrice.Enabled = True
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                    
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = False
                    Me.chkItmInvItemFlg.Enabled = False
                    Me.chkItmReorderInd.Enabled = False
                    
                    Me.txtItmReorderQty.Enabled = False
                    Me.txtItmMaxQty.Enabled = False
                    Me.txtItmPORepuQty.Enabled = False
                    
                    fraContent.Visible = False
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                    
                    tabDetailInfo.TabVisible(2) = False
                    
                Case "D"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = False
                    Me.cboItmUOMCode.Enabled = False
                    Me.txtItmBarCode.Enabled = False
                    Me.txtItmSeriesNo.Enabled = False
                    Me.txtItmBinNo.Enabled = False
                    Me.cboItmAccTypeCode.Enabled = False
                    
                    Me.cboItmCurr.Enabled = False
                    Me.cboItmPVdrCode.Enabled = False
                    
                    Me.txtItmUnitPrice.Enabled = False
                    Me.txtItmDiscount.Enabled = False
                    Me.txtItmBottomPrice.Enabled = False
                    Me.txtItmMarkUp.Enabled = False
                    Me.txtItmDefaultPrice.Enabled = False
                    
                    Me.btnItemPrice.Enabled = False
                    
                    fraInfo.Visible = False
                    fraPrice.Visible = False
                    cboItmAccTypeCode.Visible = False
                    cboItmCurr.Visible = False
                    cboItmPVdrCode.Visible = False
                    cboItmTypeCode.Visible = False
                    cboItmUOMCode.Visible = False
                    
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = False
                    Me.chkItmInvItemFlg.Enabled = False
                    Me.chkItmReorderInd.Enabled = False
                    
                    Me.txtItmReorderQty.Enabled = False
                    Me.txtItmMaxQty.Enabled = False
                    Me.txtItmPORepuQty.Enabled = False
                    
                    fraContent.Visible = False
                    
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                    
                     tabDetailInfo.TabVisible(2) = False
                    
                Case "S"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    
                    Me.txtItmUnitPrice.Enabled = True
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                  
                    
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = False
                    Me.chkItmInvItemFlg.Enabled = False
                    Me.chkItmReorderInd.Enabled = False
                    
                    Me.txtItmReorderQty.Enabled = False
                    Me.txtItmMaxQty.Enabled = False
                    Me.txtItmPORepuQty.Enabled = False
                    
                    fraContent.Visible = False
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                    
                    tabDetailInfo.TabVisible(2) = False
                    
                Case "L"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    
                    Me.txtItmUnitPrice.Enabled = True
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = False
                    Me.chkItmInvItemFlg.Enabled = False
                    Me.chkItmReorderInd.Enabled = False
                    
                    Me.txtItmReorderQty.Enabled = False
                    Me.txtItmMaxQty.Enabled = False
                    Me.txtItmPORepuQty.Enabled = False
                    
                    fraContent.Visible = False
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                
                    tabDetailInfo.TabVisible(2) = False
                
                Case "A"
                    'Tab 0 fields
                   'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    
                    Me.txtItmUnitPrice.Enabled = True
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                    
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = True
                    Me.chkItmInvItemFlg.Enabled = True
                    Me.chkItmReorderInd.Enabled = True
                    
                    Me.txtItmReorderQty.Enabled = True
                    Me.txtItmMaxQty.Enabled = True
                    Me.txtItmPORepuQty.Enabled = True
                    
                    fraContent.Visible = True
                    
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = True
                    Me.tblDetail.Enabled = True
                    
                    tabDetailInfo.TabVisible(2) = True
                    
                Case "T"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    
                    Me.txtItmUnitPrice.Enabled = True
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                   
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = False
                    Me.chkItmInvItemFlg.Enabled = False
                    Me.chkItmReorderInd.Enabled = False
                    
                    Me.txtItmReorderQty.Enabled = False
                    Me.txtItmMaxQty.Enabled = False
                    Me.txtItmPORepuQty.Enabled = False
                    
                    fraContent.Visible = False
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                    
                    tabDetailInfo.TabVisible(2) = False
                
                
                Case "C"
                    'Tab 0 fields
                    Me.cboItmTypeCode.Enabled = True
                    Me.cboItmUOMCode.Enabled = True
                    Me.txtItmBarCode.Enabled = True
                    Me.txtItmSeriesNo.Enabled = True
                    Me.txtItmBinNo.Enabled = True
                    Me.cboItmAccTypeCode.Enabled = True
                    
                    Me.cboItmCurr.Enabled = True
                    Me.cboItmPVdrCode.Enabled = True
                    Me.txtItmUnitPrice.Enabled = True
                    
                    Me.txtItmDiscount.Enabled = True
                    Me.txtItmBottomPrice.Enabled = True
                    Me.txtItmMarkUp.Enabled = True
                    Me.txtItmDefaultPrice.Enabled = True
                    
                    If wiAction = CorRec Then
                        Me.btnItemPrice.Enabled = True
                    End If
                    
                    fraInfo.Visible = True
                    fraPrice.Visible = True
                    cboItmAccTypeCode.Visible = True
                    cboItmCurr.Visible = True
                    cboItmPVdrCode.Visible = True
                    cboItmTypeCode.Visible = True
                    cboItmUOMCode.Visible = True
                    
                    'Tab 1 fields
                    Me.chkItmInActive.Enabled = False
                    Me.chkItmInvItemFlg.Enabled = False
                    Me.chkItmReorderInd.Enabled = False
                    
                    Me.txtItmReorderQty.Enabled = False
                    Me.txtItmMaxQty.Enabled = False
                    Me.txtItmPORepuQty.Enabled = False
                    
                    fraContent.Visible = False
                    
                    'Tab 2 fields
                    Me.chkOwnEdition.Enabled = False
                    Me.tblDetail.Enabled = False
                    
                    tabDetailInfo.TabVisible(2) = False
                    
                    
            End Select
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
On Error GoTo InputValidation_Err
        
    InputValidation = False
    
    
    If Chk_txtItmChiName = False Then
        Exit Function
    End If
    
    If Chk_cboItmTypeCode = False Then
        Exit Function
    End If
    
    
    If chk_cboVdrCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmUOMCode = False Then
        Exit Function
    End If
    
    
    If Chk_cboItmCurr = False Then
        Exit Function
    End If
    
    If Chk_cboItmAccTypeCode() = False Then
        Exit Function
    End If
    
    If Chk_txtItmDiscount = False Then
        Exit Function
    End If
    
    InputValidation = True
    
Exit Function

InputValidation_Err:

    MsgBox "Error on InputValidation_Err " & Err.Description
    InputValidation = False

    
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim rsRcd As New ADODB.Recordset
    Dim wiCtr As Integer
        
    wsSQL = "SELECT * "
    wsSQL = wsSQL + "From MstItem "
    wsSQL = wsSQL + "WHERE (((MstItem.ItmCode)='" + Set_Quote(cboITMCODE) + "') AND ((MstItem.ItmStatus)='1'));"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "ItmID")
        
        Me.cboItemClassCode = ReadRs(rsRcd, "ItmClass")
        Me.lblDspItemClassCode = Get_TableInfo("MstItemClass", "ItemClassCode='" & cboItemClassCode & "'", IIf(gsLangID = "1", "ItemClassEDesc", "ItemClassCDesc"))
        Me.txtItmChiName = ReadRs(rsRcd, "ItmChiName")
        Me.txtItmEngName = ReadRs(rsRcd, "ItmEngName")
        Me.lblDspItmLastUpd = ReadRs(rsRcd, "ItmLastUpd")
        Me.lblDspItmLastUpdDate = ReadRs(rsRcd, "ItmLastUpdDate")
        
        'Tab 0
        Me.cboItmTypeCode = ReadRs(rsRcd, "ItmItmTypeCode")
        lblDspItmTypeDesc = LoadDescByCode("MstItemType", "ItmTypeCode", "ItmTypeChiDesc", cboItmTypeCode, True)
        Me.cboItmUOMCode = ReadRs(rsRcd, "ItmUOMCode")
        lblDspItmUomDesc = LoadDescByCode("MstUOM", "UomCode", "UomDesc", cboItmUOMCode, True)
        Me.txtItmBarCode = ReadRs(rsRcd, "ItmBarCode")
        Me.txtItmSeriesNo = ReadRs(rsRcd, "ItmSeriesNo")
        Me.txtItmBinNo = ReadRs(rsRcd, "ItmBinNo")
        Me.cboItmAccTypeCode = ReadRs(rsRcd, "ItmAccTypeCode")
        lblDspItmAccTypeDesc = LoadDescByCode("MstAccountType", "AccTypeCode", "AccTypeDesc", cboItmAccTypeCode, True)
        
        Me.cboItmCurr = ReadRs(rsRcd, "ItmCurr")
        wlVdrID = ReadRs(rsRcd, "ItmPVdrID")
        cboItmPVdrCode = Get_TableInfo("MstVendor", "VdrID=" & wlVdrID, "VdrCode")
        
        txtItmUnitPrice = Format(To_Value(ReadRs(rsRcd, "ItmUnitPrice")), gsUprFmt)
        txtItmDiscount = Format(To_Value(ReadRs(rsRcd, "ItmDiscount")), gsUprFmt)
        txtItmBottomPrice = Format(To_Value(ReadRs(rsRcd, "ItmBottomPrice")), gsUprFmt)
        txtItmMarkUp = Format(To_Value(ReadRs(rsRcd, "ItmMarkUp")), gsUprFmt)
        txtItmDefaultPrice = Format(To_Value(ReadRs(rsRcd, "ItmDefaultPrice")), gsUprFmt)
        
        'Tab 1 fields
        Me.lblDspStkOnHand.Caption = Get_IcTrnQty("INOUT", wlKey, "", "")
        Me.lblDspStkAllocated.Caption = Get_IcTrnQty("STKALL", wlKey, "", "")
        Me.lblDspStkOnOrder.Caption = Get_IcTrnQty("STKORD", wlKey, "", "")
        Me.lblDspStkIndent.Caption = Get_IcTrnQty("STKIND", wlKey, "", "")
        Me.lblDspStkAvailable.Caption = Get_IcTrnQty("%", wlKey, "", "")
       ' Me.lblDspUnitPrice.Caption = Get_AvgCost(wlKey, "")
        
        Call Set_CheckValue(chkItmInActive, ReadRs(rsRcd, "ItmInActive"))
        Call Set_CheckValue(chkItmInvItemFlg, ReadRs(rsRcd, "ItmInvItemFlg"))
        Call Set_CheckValue(chkItmReorderInd, ReadRs(rsRcd, "ItmReorderInd"))
        
        Me.txtItmReorderQty = Format(To_Value(ReadRs(rsRcd, "ItmReorderQty")), gsQtyFmt)
        Me.txtItmMaxQty = Format(To_Value(ReadRs(rsRcd, "ItmMaxQty")), gsQtyFmt)
        Me.txtItmPORepuQty = Format(To_Value(ReadRs(rsRcd, "ItmPORepuQty")), gsQtyFmt)
        
        'Tab 2 fields
        Call Set_CheckValue(chkOwnEdition, ReadRs(rsRcd, "ItmOwnEdition"))
        rsRcd.Close
        
        wsSQL = "SELECT MstBOM.*, MstItem.ItmCode, MstItem.ItmChiName, MstItem.ItmEngName "
        wsSQL = wsSQL + "From MstBOM, MstItem "
        wsSQL = wsSQL + "WHERE (((MstBOM.BOMItmID)=" + CStr(wlKey) + ") AND ((MstItem.ItmID)=MstBOM.BOMDTItmID));"

        rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        If rsRcd.RecordCount <> 0 Then
            rsRcd.MoveFirst
            With waResult
                .ReDim 0, -1, ITMCODE, ITMID
                Do While Not rsRcd.EOF
                    wiCtr = wiCtr + 1
                    .AppendRows
                    waResult(.UpperBound(1), ITMCODE) = ReadRs(rsRcd, "ITMCODE")
                    waResult(.UpperBound(1), ITMDESC) = IIf(gsLangID = "1", ReadRs(rsRcd, "ItmEngName"), ReadRs(rsRcd, "ItmChiName"))
                    waResult(.UpperBound(1), QTY) = ReadRs(rsRcd, "BOMQTY")
                    waResult(.UpperBound(1), ITMID) = ReadRs(rsRcd, "BOMITMID")
                    rsRcd.MoveNext
                Loop
            End With
            
            tblDetail.ReBind
            tblDetail.FirstRow = 0
        End If
        
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
    Set waResult = Nothing
    Set waPopUpSub = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
   ' Set waPgmItm = Nothing
    Set frmITM001 = Nothing

End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        If cboItmCurr.Enabled = True Then
            cboItmCurr.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 1 Then
        If txtItmReorderQty.Enabled = True Then
            txtItmReorderQty.SetFocus
        End If
    End If
End Sub

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
                If MsgBox("你是否確定儲存現時之變更而離開?", vbYesNo, gsTitle) = vbNo Then
                    Call cmdCancel
                End If
            Else
                Call cmdCancel
            End If
        
        Case tcFind
            
            Call OpenPromptForm
            
        Case tcKey
            
            Call cmdChangeKey(CorRec)
            
        Case tcCopy
            
            Call cmdChangeKey(AddRec)
            
        Case tcExit
        
            Unload Me
            
    End Select
End Sub

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub IniForm()
    Me.KeyPreview = True
  '  Me.Left = 0
  '  Me.Top = 0
  '  Me.Width = Screen.Width
  '  Me.Height = Screen.Height
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "ITM001"
    wsTrnCd = ""
End Sub


Private Sub Ini_Caption()
Dim i As Integer
On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP_M", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    fraHeaderInfo.Caption = Get_Caption(waScrItm, "fraHeaderInfo")
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO0")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO1")
   
    'not added to insert script
    lblItmCode.Caption = Get_Caption(waScrItm, "ITMCODE")
    lblItemClassCode.Caption = Get_Caption(waScrItm, "ITEMCLASSCODE")
    lblItmBarCode.Caption = Get_Caption(waScrItm, "ITMBARCODE")
    lblItmChiName.Caption = Get_Caption(waScrItm, "ITMCHINAME")
    lblItmEngName.Caption = Get_Caption(waScrItm, "ITMENGNAME")
    lblItmTypeCode.Caption = Get_Caption(waScrItm, "ITMTYPECODE")
    lblItmLastUpd.Caption = Get_Caption(waScrItm, "ITMLASTUPD")
    lblItmLastUpdDate.Caption = Get_Caption(waScrItm, "ITMLASTUPDDATE")
    lblItmPVdrCode.Caption = Get_Caption(waScrItm, "VDRCODE")
    lblItmDiscount.Caption = Get_Caption(waScrItm, "ITMDISCOUNT")
    lblItmDefaultPrice.Caption = Get_Caption(waScrItm, "ITMDEFAULTPRICE")
        
    lblItmUomCode.Caption = Get_Caption(waScrItm, "ITMUOMCODE")
    lblItmSeriesNo.Caption = Get_Caption(waScrItm, "ITMSERIESNO")
    chkItmInActive.Caption = Get_Caption(waScrItm, "ITMINACTIVE")
    chkItmInvItemFlg.Caption = Get_Caption(waScrItm, "ITMINVITEMFLG")
    chkOwnEdition.Caption = Get_Caption(waScrItm, "OWNEDITION")
    
    lblItmCurrCode.Caption = Get_Caption(waScrItm, "ITMCURRCODE")
    lblItmBottomPrice.Caption = Get_Caption(waScrItm, "ITMBOTTOMPRICE")
    lblItmMarkUp.Caption = Get_Caption(waScrItm, "ITMMARKUP")
    
    lblItmReorderQty.Caption = Get_Caption(waScrItm, "ITMREORDERQTY")
    chkItmReorderInd.Caption = Get_Caption(waScrItm, "ITMREORDERIND")
    lblItmMaxQty.Caption = Get_Caption(waScrItm, "ITMMAXQTY")
    lblItmPORepuQty.Caption = Get_Caption(waScrItm, "ITMPOREPUQTY")
    
    lblItmAccTypeCode.Caption = Get_Caption(waScrItm, "ITMACCTYPECODE")
    lblUnitPrice.Caption = Get_Caption(waScrItm, "UNITPRICE")
    lblItmBinNo.Caption = Get_Caption(waScrItm, "BINNO")
    
    With tblDetail
        .Columns(ITMCODE).Caption = Get_Caption(waScrItm, "GDITMCODE")
        .Columns(ITMDESC).Caption = Get_Caption(waScrItm, "GDITMDESC")
        .Columns(QTY).Caption = Get_Caption(waScrItm, "GDQTY")
    End With
    
    btnItemPrice.Caption = Get_Caption(waScrItm, "ITMPRICE")
    
        
    lblStkOnHand.Caption = Get_Caption(waScrItm, "STKONHAND")
    lblStkIndent.Caption = Get_Caption(waScrItm, "STKINDENT")
    lblStkOnOrder.Caption = Get_Caption(waScrItm, "STKONORDER")
    lblStkAllocated.Caption = Get_Caption(waScrItm, "STKALLOCATED")
    lblStkAvailable.Caption = Get_Caption(waScrItm, "STKAVAILABLE")
    
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    tbrProcess.Buttons(tcKey).ToolTipText = Get_Caption(waScrToolTip, tcKey) & "(F7)"
    tbrProcess.Buttons(tcCopy).ToolTipText = Get_Caption(waScrToolTip, tcCopy)
    
    
    wsActNam(1) = Get_Caption(waScrItm, "ITMADD")
    wsActNam(2) = Get_Caption(waScrItm, "ITMEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "ITMDELETE")
    
    lblKeyDesc = Get_Caption(waScrToolTip, "KEYDESC")
    lblComboPrompt = Get_Caption(waScrToolTip, "COMBOPROMPT")
    lblInsertLine = Get_Caption(waScrToolTip, "INSERTLINE")
    lblDeleteLine = Get_Caption(waScrToolTip, "DELETELINE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP", waPopUpSub)
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub
Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, ITMCODE, ITMID
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

    
    wiAction = DefaultPage
    wlKey = 0
    wlVdrID = 0
    
    Me.txtItmUnitPrice = Format(0, gsUprFmt)
    Me.txtItmBottomPrice = Format(0, gsUprFmt)
    Me.txtItmMarkUp = Format(0, gsUprFmt)
    Me.txtItmReorderQty = 0
    Me.txtItmMaxQty = 0
    Me.txtItmPORepuQty = 0
    Me.txtItmDiscount = Format(0, gsUprFmt)
    Me.txtItmDefaultPrice = Format(0, gsUprFmt)
    
    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    
    wbAfrKey = False
    Me.tabDetailInfo.Tab = 0
    Me.Caption = wsFormCaption
    tblCommon.Visible = False
 
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
              
       Me.Caption = wsFormCaption + " - ADD"
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
        txtItmCode.SetFocus
       
    Case CorRec
           
        Me.Caption = wsFormCaption + " - EDIT"
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboITMCODE.SetFocus
    
    Case DelRec
    
        Me.Caption = wsFormCaption + " - DELETE"
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboITMCODE.SetFocus
    
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub

Private Sub Ini_Scr_AfrKey()
    Select Case wiAction
        Case CorRec, DelRec

            If LoadRecord() = False Then
                gsMsg = "存取檔案失敗! 請聯絡系統管理員或無限系統顧問!"
                MsgBox gsMsg, vbOKOnly, gsTitle
                Exit Sub
            Else
                If RowLock(wsConnTime, wsKeyType, cboITMCODE, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                End If
            End If
        wbAfrKey = True
        Call SetFieldStatus("AfrKey")
        Call SetButtonStatus("AfrKeyEdit")
    
        Case AddRec
        wbAfrKey = True
        Call SetFieldStatus("AfrKey")
        Call SetButtonStatus("AfrKeyAdd")
        
    End Select
    
    txtItmChiName.SetFocus
End Sub

Private Function Chk_txtItmCode() As Boolean
    Dim wsStatus As String

    Chk_txtItmCode = False
    
    If Trim(txtItmCode.Text) = "" Then
        Chk_txtItmCode = True
        Exit Function
    End If

    If Chk_ItmCode(txtItmCode.Text, wsStatus) = True Then
    
    If wsStatus = "2" Then
        gsMsg = "物料已存在但已無效!"
    Else
        gsMsg = "物料已存在!"
    End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtItmCode.SetFocus
        Exit Function
        
    End If
    
    Chk_txtItmCode = True
End Function

Private Function Chk_cboItmCode() As Boolean
    Dim wsStatus As String

    Chk_cboItmCode = False
    
        If Trim(cboITMCODE.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboITMCODE.SetFocus
            Exit Function
        End If
    
        If Chk_ItmCode(cboITMCODE.Text, wsStatus) = False Then
            gsMsg = "物料不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboITMCODE.SetFocus
            Exit Function
        Else
        If wsStatus = "2" Then
            gsMsg = "物料已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboITMCODE.SetFocus
            Exit Function
        End If
        End If
    
    Chk_cboItmCode = True
End Function
Private Sub cmdOpen()

    Dim newForm As New frmITM001
    
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
    Dim i As Integer
    Dim wiCtr As Integer
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboITMCODE, wsFormID) Then
            gsMsg = "記錄已被鎖定, 現在以唯讀模式開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            MousePointer = vbDefault
            Exit Function
        End If
    End If
    
    If wiAction = DelRec Then
        If MsgBox("你是否確定要刪除此檔案?", vbYesNo, gsTitle) = vbNo Then
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
    
  '  If wiAction = AddRec Then
  '      If Chk_KeyExist() = True Then
  '          Call GetNewKey
  '      End If
  '  End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_ITM001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
    
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, UCase(txtItmCode.Text), UCase(cboITMCODE.Text)))
    Call SetSPPara(adcmdSave, 4, txtItmBarCode.Text)
    Call SetSPPara(adcmdSave, 5, "")
    Call SetSPPara(adcmdSave, 6, "")
    Call SetSPPara(adcmdSave, 7, txtItmChiName)
    Call SetSPPara(adcmdSave, 8, "")
    Call SetSPPara(adcmdSave, 9, "")
    Call SetSPPara(adcmdSave, 10, "")
    Call SetSPPara(adcmdSave, 11, cboItmTypeCode)
    Call SetSPPara(adcmdSave, 12, "")
    Call SetSPPara(adcmdSave, 13, txtItmSeriesNo)
    Call SetSPPara(adcmdSave, 14, "")
    Call SetSPPara(adcmdSave, 15, "")
    Call SetSPPara(adcmdSave, 16, "")
    Call SetSPPara(adcmdSave, 17, "")
    Call SetSPPara(adcmdSave, 18, "")
    Call SetSPPara(adcmdSave, 19, "")
    Call SetSPPara(adcmdSave, 20, gsUserID)
    Call SetSPPara(adcmdSave, 21, wsGenDte)
    Call SetSPPara(adcmdSave, 22, "")
    Call SetSPPara(adcmdSave, 23, "")
    Call SetSPPara(adcmdSave, 24, "")
    Call SetSPPara(adcmdSave, 25, "")
    Call SetSPPara(adcmdSave, 26, "")
    Call SetSPPara(adcmdSave, 27, "")
    Call SetSPPara(adcmdSave, 28, "")
    Call SetSPPara(adcmdSave, 29, "")
    Call SetSPPara(adcmdSave, 30, txtItmBinNo)
    Call SetSPPara(adcmdSave, 31, "")
    Call SetSPPara(adcmdSave, 32, "")
    Call SetSPPara(adcmdSave, 33, "")
    Call SetSPPara(adcmdSave, 34, "")
    Call SetSPPara(adcmdSave, 35, "")
    Call SetSPPara(adcmdSave, 36, 0)
    Call SetSPPara(adcmdSave, 37, 0)
    Call SetSPPara(adcmdSave, 38, 0)
    Call SetSPPara(adcmdSave, 39, 0)
    Call SetSPPara(adcmdSave, 40, 0)
    Call SetSPPara(adcmdSave, 41, 0)
    Call SetSPPara(adcmdSave, 42, 0)
    Call SetSPPara(adcmdSave, 43, cboItmCurr)
    Call SetSPPara(adcmdSave, 44, txtItmDefaultPrice)
    Call SetSPPara(adcmdSave, 45, txtItmBottomPrice)
    Call SetSPPara(adcmdSave, 46, "")
    Call SetSPPara(adcmdSave, 47, cboItmAccTypeCode)
    Call SetSPPara(adcmdSave, 48, Get_CheckValue(chkItmInActive))
    Call SetSPPara(adcmdSave, 49, Get_CheckValue(chkItmInvItemFlg))
    Call SetSPPara(adcmdSave, 50, "N")
    Call SetSPPara(adcmdSave, 51, "N")
    Call SetSPPara(adcmdSave, 52, txtItmReorderQty)
    
    For i = 0 To 11
        Call SetSPPara(adcmdSave, 53 + i, "")
    Next i
    
    Call SetSPPara(adcmdSave, 65, Get_CheckValue(chkItmReorderInd))
    Call SetSPPara(adcmdSave, 66, txtItmPORepuQty)
    Call SetSPPara(adcmdSave, 67, "")
    Call SetSPPara(adcmdSave, 68, Get_CheckValue(chkOwnEdition))
    Call SetSPPara(adcmdSave, 69, cboItmUOMCode)
    Call SetSPPara(adcmdSave, 70, txtItmEngName)
    Call SetSPPara(adcmdSave, 71, txtItmMarkUp)
    Call SetSPPara(adcmdSave, 72, txtItmMaxQty)
    Call SetSPPara(adcmdSave, 73, txtItmUnitPrice)
    Call SetSPPara(adcmdSave, 74, txtItmDiscount)
    Call SetSPPara(adcmdSave, 75, cboItemClassCode)
    Call SetSPPara(adcmdSave, 76, wlVdrID)
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 77)
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_ITM001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, wiCtr + 1)
                Call SetSPPara(adcmdSave, 4, waResult(wiCtr, ITMID))
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, QTY))
                Call SetSPPara(adcmdSave, 6, gsUserID)
                Call SetSPPara(adcmdSave, 7, wsGenDte)
                adcmdSave.Execute
            End If
        Next
    End If
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - ITM001!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
    Else
        If wiAction = DelRec Then
        gsMsg = "已成功刪除!"
        Else
        gsMsg = "已成功儲存!"
        End If
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
       If MsgBox("你是否確定要儲存現時之作業?", vbYesNo, gsTitle) = vbNo Then
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
    Dim wsSQL As String
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "物料編碼"
    vFilterAry(1, 2) = "ItmCode"
    
If gsLangID = "1" Then
    vFilterAry(2, 1) = "物料英文名稱"
    vFilterAry(2, 2) = "ItmEngName"
Else
    vFilterAry(2, 1) = "物料名稱"
    vFilterAry(2, 2) = "ItmChiName"
End If
    
    vFilterAry(3, 1) = "物料分類"
    vFilterAry(3, 2) = "ItmItmTypeCode"
    
    
    ReDim vAry(3, 3)
    vAry(1, 1) = "物料編碼"
    vAry(1, 2) = "ItmCode"
    vAry(1, 3) = "2500"
    
    
If gsLangID = "1" Then
    vAry(2, 1) = "物料英文名稱"
    vAry(2, 2) = "ItmEngName"
    vAry(2, 3) = "3500"
Else
    vAry(2, 1) = "物料名稱"
    vAry(2, 2) = "ItmChiName"
    vAry(2, 3) = "3500"
End If

    vAry(3, 1) = "物料分類"
    vAry(3, 2) = "ItmItmTypeCode"
    vAry(3, 3) = "1200"
    
    

    
    
    'frmShareSearch.Show vbModal
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT MstItem.ItmCode, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & ", MstItem.ItmItmTypeCode "
        wsSQL = wsSQL + "FROM MstItem "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE MstItem.ItmStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstItem.ItmCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And frmShareSearch.Tag <> cboITMCODE Then
        cboITMCODE = frmShareSearch.Tag
       If cboITMCODE.Enabled = False Then
        LoadRecord
        txtItmBarCode.Text = ""
        txtItmCode.SetFocus
       Else
        cboITMCODE.SetFocus
        SendKeys "{Enter}"
       End If
    End If
    Unload frmShareSearch

    
End Sub

Private Sub txtItmBarCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmBarCode, 13, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtItmSeriesNo.SetFocus
    End If
End Sub

Private Sub txtItmBarCode_LostFocus()
    FocusMe txtItmBarCode, True
End Sub

Private Sub txtItmBinNo_GotFocus()
    FocusMe txtItmBinNo
End Sub

Private Sub txtItmBinNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmBinNo, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        cboItmAccTypeCode.SetFocus
    End If
End Sub

Private Sub txtItmBinNo_LostFocus()
    FocusMe txtItmBinNo, True
End Sub



Private Sub txtItmChiName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmChiName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtItmChiName() = True Then
            txtItmEngName.SetFocus
        End If
    End If
End Sub

Private Sub txtItmChiName_LostFocus()
    FocusMe txtItmChiName, True
End Sub

Private Sub txtItmCode_KeyPress(KeyAscii As Integer)
'    Call chk_InpLenA(txtItmCode, 30, KeyAscii, True)
    Call chk_InpLenC(txtItmCode, 30, KeyAscii, True, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtItmCode() = True Then
            cboItemClassCode.SetFocus
            'Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub txtItmCode_LostFocus()
    FocusMe txtItmCode, True
End Sub

Private Sub txtItmDefaultPrice_GotFocus()
    FocusMe txtItmDefaultPrice
End Sub

Private Sub txtItmDefaultPrice_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmDefaultPrice, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Select Case UCase(cboItemClassCode)
            Case "P"
                tabDetailInfo.Tab = 1
                chkItmInActive.SetFocus
                    
            Case "N"
                txtItmChiName.SetFocus
                
            Case "D"
                txtItmChiName.SetFocus
                
            Case "S"
                txtItmChiName.SetFocus
                
            Case "L"
                txtItmChiName.SetFocus
            
            Case "A"
                tabDetailInfo.Tab = 2
                tblDetail.SetFocus
                
            Case "T"
                txtItmChiName.SetFocus
            
            Case "C"
                txtItmChiName.SetFocus
                
        End Select
    End If
End Sub

Private Sub txtItmDefaultPrice_LostFocus()
    txtItmDefaultPrice = Format(txtItmDefaultPrice, gsUprFmt)
    FocusMe txtItmDefaultPrice
End Sub



Private Sub txtItmDiscount_Change()
Call Calc_Price
End Sub

Private Sub txtItmMarkUp_Change()
Call Calc_Price
End Sub

Private Sub txtItmUnitPrice_GotFocus()
    FocusMe txtItmUnitPrice
End Sub

Private Sub txtItmUnitPrice_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmUnitPrice, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call Calc_Price
        
        tabDetailInfo.Tab = 0
        txtItmDiscount.SetFocus
    End If
End Sub

Private Sub txtItmUnitPrice_LostFocus()
    txtItmUnitPrice = Format(txtItmUnitPrice, gsUprFmt)
    FocusMe txtItmUnitPrice, True
End Sub

Private Sub txtItmDiscount_GotFocus()
    FocusMe txtItmDiscount
End Sub

Private Sub txtItmDiscount_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmDiscount, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        If Chk_txtItmDiscount = False Then Exit Sub
        Call Calc_Price
        
        tabDetailInfo.Tab = 0
        txtItmBottomPrice.SetFocus
    End If
End Sub

Private Sub txtItmDiscount_LostFocus()
    txtItmDiscount = Format(txtItmDiscount, gsUprFmt)
    FocusMe txtItmDiscount, True
End Sub

Private Sub txtItmEngName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmEngName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Select Case UCase(cboItemClassCode)
            Case "P"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
                    
            Case "N"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
                
            Case "D"
                txtItmChiName.SetFocus
                
            Case "S"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
                
            Case "L"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
            
            Case "A"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
                
            Case "T"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
            
            Case "C"
                tabDetailInfo.Tab = 0
                cboItmTypeCode.SetFocus
                
        End Select
    End If
End Sub

Private Sub txtItmEngName_LostFocus()
    FocusMe txtItmEngName, True
End Sub

Private Sub txtItmReorderQty_GotFocus()
    FocusMe txtItmReorderQty
End Sub

Private Sub txtItmReorderQty_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmReorderQty, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        txtItmMaxQty.SetFocus
    End If
End Sub

Private Sub txtItmReorderQty_LostFocus()
    txtItmReorderQty = Format(txtItmReorderQty, gsQtyFmt)
    FocusMe txtItmReorderQty, True
End Sub

Private Sub txtItmMaxQty_GotFocus()
    FocusMe txtItmMaxQty
End Sub

Private Sub txtItmMaxQty_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmMaxQty, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        txtItmPORepuQty.SetFocus
    End If
End Sub

Private Sub txtItmMaxQty_LostFocus()
    txtItmMaxQty = Format(txtItmMaxQty, gsQtyFmt)
    FocusMe txtItmMaxQty, True
End Sub

Private Sub txtItmPORepuQty_GotFocus()
    FocusMe txtItmPORepuQty
End Sub

Private Sub txtItmPORepuQty_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmPORepuQty, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Select Case UCase(cboItemClassCode)
            Case "P"
                tabDetailInfo.Tab = 0
                txtItmChiName.SetFocus
                    
            Case "N"
                txtItmChiName.SetFocus
                
            Case "D"
                txtItmChiName.SetFocus
                
            Case "S"
                txtItmChiName.SetFocus
                
            Case "L"
                txtItmChiName.SetFocus
            
            Case "A"
                tabDetailInfo.Tab = 2
                tblDetail.SetFocus
                
            Case "T"
                txtItmChiName.SetFocus
            
            Case "C"
                txtItmChiName.SetFocus
                
        End Select
    End If
End Sub

Private Sub txtItmPORepuQty_LostFocus()
    txtItmPORepuQty = Format(txtItmPORepuQty, gsQtyFmt)
    FocusMe txtItmPORepuQty, True
End Sub

Private Sub txtItmSeriesNo_GotFocus()
    FocusMe txtItmSeriesNo
End Sub

Private Sub txtItmSeriesNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmSeriesNo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtItmBinNo.SetFocus
    End If
End Sub

Private Sub txtItmSeriesNo_LostFocus()
    FocusMe txtItmSeriesNo, True
End Sub

Private Sub txtItmBarCode_GotFocus()
    FocusMe txtItmBarCode
End Sub

Private Sub txtItmBottomPrice_GotFocus()
    FocusMe txtItmBottomPrice
End Sub

Private Sub txtItmBottomPrice_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmBottomPrice, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtItmMarkUp.SetFocus
    End If
End Sub

Private Sub txtItmBottomPrice_LostFocus()
    txtItmBottomPrice = Format(txtItmBottomPrice, gsUprFmt)
    FocusMe txtItmBottomPrice, True
End Sub


Private Sub txtItmMarkUp_GotFocus()
    FocusMe txtItmMarkUp
End Sub

Private Sub txtItmMarkUp_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmMarkUp, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        Call Calc_Price
        
        tabDetailInfo.Tab = 0
        txtItmDefaultPrice.SetFocus
    End If
End Sub

Private Sub txtItmMarkUp_LostFocus()
    txtItmMarkUp = Format(txtItmMarkUp, gsUprFmt)
    FocusMe txtItmMarkUp, True
End Sub
Private Sub txtItmChiName_GotFocus()
    FocusMe txtItmChiName
End Sub

Private Sub txtItmCode_GotFocus()
    FocusMe txtItmCode
End Sub

Private Sub txtItmEngName_GotFocus()
    FocusMe txtItmEngName
End Sub


Private Function Chk_ItmCurr() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_ItmCurr = False
    
    sSQL = "SELECT ExcCurr FROM MstExchangeRate WHERE ExcCurr='" & Set_Quote(cboItmCurr.Text) + "' And ExcStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    Chk_ItmCurr = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboItmCurr() As Boolean
  
    Chk_cboItmCurr = False
    
    If UCase(cboItemClassCode) = "D" Then
            Chk_cboItmCurr = True
            Exit Function
    End If
        
    If Trim(cboItmCurr.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        Me.tabDetailInfo.Tab = 0
        cboItmCurr.SetFocus
        Exit Function
    End If
    
    If Chk_ItmCurr() = False Then
        gsMsg = "貨幣不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboItmCurr.SetFocus
        Exit Function
    End If

    
    Chk_cboItmCurr = True
End Function


Private Function Chk_cboItmTypeCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    If UCase(cboItemClassCode) = "D" Then
            Chk_cboItmTypeCode = True
            Exit Function
    End If
    

    Chk_cboItmTypeCode = False
    
    If Trim(cboItmTypeCode.Text) = "" Then
         gsMsg = "沒有輸入須要之資料!"
         MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
         tabDetailInfo.Tab = 0
         cboItmTypeCode.SetFocus
         Exit Function
    End If
    
    If Chk_ItmTypeCode(cboItmTypeCode.Text, wsRetName) = False Then
            lblDspItmTypeDesc = ""
            gsMsg = "物料類別編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboItmTypeCode.SetFocus
            Exit Function
    Else
            lblDspItmTypeDesc = wsRetName
    End If
   
    
    Chk_cboItmTypeCode = True
    
End Function

Private Function Chk_ItmTypeCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_ItmTypeCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT ItmTypeChiDesc "
    wsSQL = wsSQL & " FROM MstItemType WHERE MstItemType.ItmTypeCode = '" & Set_Quote(inCode) & "'  And ItmTypeStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "ItmTypeChiDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmTypeCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_ItmUOMCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_ItmUOMCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT UomDesc "
    wsSQL = wsSQL & " FROM MstUOM WHERE MstUOM.UomCode = '" & Set_Quote(inCode) & "' And UomStatus = '1' "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "UomDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmUOMCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function


Private Function Chk_cboItmUOMCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmUOMCode = False
    
    If UCase(cboItemClassCode) = "D" Then
            Chk_cboItmUOMCode = True
            Exit Function
    End If
    
    If Trim(cboItmUOMCode.Text) = "" Then
        Chk_cboItmUOMCode = True
        Exit Function
    End If
    
    If Chk_ItmUOMCode(cboItmUOMCode.Text, wsRetName) = False Then
        gsMsg = "量度單位編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboItmUOMCode.SetFocus
        Exit Function
    Else
        lblDspItmUomDesc = wsRetName
    End If

    
    Chk_cboItmUOMCode = True
End Function


Private Function Chk_txtItmEngName() As Boolean
    Chk_txtItmEngName = False
    
    If Trim(txtItmEngName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtItmEngName.SetFocus
        Exit Function
    End If
    
    Chk_txtItmEngName = True
End Function

Private Function Chk_txtItmChiName() As Boolean
     
    Chk_txtItmChiName = False
    
    If Trim(txtItmChiName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtItmChiName.SetFocus
        Exit Function
    End If
    
    Chk_txtItmChiName = True
End Function

Private Sub tblCommon_DblClick()
    If wcCombo.Name = tblDetail.Name Then
        tblDetail.EditActive = True
        wcCombo.Text = tblCommon.Columns(0).Text
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
            wcCombo.Text = tblCommon.Columns(0).Text
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

Private Sub cboItmCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboITMCODE, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmCode() = False Then
            Exit Sub
        Else
            'Call Chk_KeyFld
            Call Ini_Scr_AfrKey
            'cboItemClassCode.SetFocus
        End If
    End If
End Sub

Private Sub cboItmCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboITMCODE
    
    If gsLangID = "1" Then
        wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmEngName FROM MstItem WHERE ItmStatus = '1'"
        wsSQL = wsSQL & " AND ItmCode LIKE '%" & IIf(cboITMCODE.SelLength > 0, "", Set_Quote(cboITMCODE.Text)) & "%' "
        wsSQL = wsSQL & " AND ItmClass LIKE '%" & IIf(cboItemClassCode.SelLength > 0, "", Set_Quote(cboItemClassCode.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    Else
        wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmChiName FROM MstItem WHERE ItmStatus = '1'"
        wsSQL = wsSQL & " AND ItmCode LIKE '%" & IIf(cboITMCODE.SelLength > 0, "", Set_Quote(cboITMCODE.Text)) & "%' "
        wsSQL = wsSQL & " AND ItmClass LIKE '%" & IIf(cboItemClassCode.SelLength > 0, "", Set_Quote(cboItemClassCode.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItmCode "
    End If
    
    Call Ini_Combo(3, wsSQL, cboITMCODE.Left, cboITMCODE.Top + cboITMCODE.Height, tblCommon, wsFormID, "TBLB", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmCode_GotFocus()

    FocusMe cboITMCODE
End Sub

Private Sub cboItmCode_LostFocus()
    FocusMe cboITMCODE, True
End Sub

Private Sub cboItmTypeCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    
    If wbAfrKey = False Then
    
     Set wcCombo = cboITMCODE
    
    If gsLangID = "1" Then
    
    wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmEngName FROM MstItem WHERE ItmStatus = '1'"
    wsSQL = wsSQL & " AND ItmItmTypeCode LIKE '%" & IIf(cboItmTypeCode.SelLength > 0, "", Set_Quote(cboItmTypeCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY ItmItmTypeCode "
    
    Else
    
    wsSQL = "SELECT ItmCode, ItmItmTypeCode, ItmChiName FROM MstItem WHERE ItmStatus = '1'"
    wsSQL = wsSQL & " AND ItmItmTypeCode LIKE '%" & IIf(cboItmTypeCode.SelLength > 0, "", Set_Quote(cboItmTypeCode.Text)) & "%' "
    wsSQL = wsSQL & "ORDER BY ItmItmTypeCode "
    
    End If
    
    Call Ini_Combo(3, wsSQL, cboItmTypeCode.Left + tabDetailInfo.Left, cboItmTypeCode.Top + tabDetailInfo.Top + cboItmTypeCode.Height, tblCommon, wsFormID, "TBLB", Me.Width, Me.Height)
 
    
    Else
    
    Set wcCombo = cboItmTypeCode
    
    If gsLangID = "1" Then
    wsSQL = "SELECT ItmTypeCode, ItmTypeEngDesc FROM MstItemType WHERE ItmTypeStatus = '1'"
    wsSQL = wsSQL & "ORDER BY ItmTypeCode "
    Else
    wsSQL = "SELECT ItmTypeCode, ItmTypeChiDesc FROM MstItemType WHERE ItmTypeStatus = '1'"
    wsSQL = wsSQL & "ORDER BY ItmTypeCode "
    End If
    
    Call Ini_Combo(2, wsSQL, cboItmTypeCode.Left + Me.tabDetailInfo.Left, cboItmTypeCode.Top + cboItmTypeCode.Height + Me.tabDetailInfo.Top, tblCommon, wsFormID, "TBLIT", Me.Width, Me.Height)
    
    End If
    
    
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        'If wbAfrKey = False Then
        '   cboItmCode.SetFocus
        '   Exit Sub
        'End If
        
        If Chk_cboItmTypeCode() = False Then
            Exit Sub
        End If
            
        tabDetailInfo.Tab = 0
        cboItmUOMCode.SetFocus
    End If
End Sub

Private Sub cboItmTypeCode_GotFocus()

    FocusMe cboItmTypeCode
End Sub

Private Sub cboItmTypeCode_LostFocus()
    FocusMe cboItmTypeCode, True
End Sub


Private Sub cboItmCurr_DropDown()
    
    Dim wsSQL As String


    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmCurr
    
    wsSQL = "SELECT DISTINCT ExcCurr FROM MstExchangeRate WHERE ExcStatus = '1'"
    wsSQL = wsSQL & "ORDER BY ExcCurr "
    Call Ini_Combo(1, wsSQL, cboItmCurr.Left + tabDetailInfo.Left, cboItmCurr.Top + cboItmCurr.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCURR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmCurr() = False Then
            Exit Sub
        End If
            
        tabDetailInfo.Tab = 0
        
        cboItmPVdrCode.SetFocus
    End If
End Sub

Private Sub cboItmCurr_GotFocus()
    FocusMe cboItmCurr
End Sub

Private Sub cboItmCurr_LostFocus()
    FocusMe cboItmCurr, True
End Sub


Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT ItmStatus FROM MstItem WHERE ItmCode = '" & Set_Quote(txtItmCode) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
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
    
    'Create Selection wsSql
    With Newfrm
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "ItmCode"
        .KeyLen = 15
        Set .ctlKey = txtItmCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmAccTypeCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmAccTypeCode
    
    wsSQL = "SELECT AccTypeCode, AccTypeDesc FROM MstAccountType WHERE AccTypeStatus = '1'"
    wsSQL = wsSQL & "ORDER BY AccTypeCode "
    Call Ini_Combo(2, wsSQL, cboItmAccTypeCode.Left + tabDetailInfo.Left, cboItmAccTypeCode.Top + cboItmAccTypeCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLACCTYPE", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmAccTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmAccTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmAccTypeCode() = False Then
            Exit Sub
        End If
            
        tabDetailInfo.Tab = 0
        cboItmCurr.SetFocus
        
    End If
End Sub

Private Sub cboItmAccTypeCode_GotFocus()
    FocusMe cboItmAccTypeCode
End Sub

Private Sub cboItmAccTypeCode_LostFocus()
    FocusMe cboItmAccTypeCode, True
End Sub

Private Function Chk_cboItmAccTypeCode() As Boolean
Dim wsDesc As String
    Chk_cboItmAccTypeCode = False
    
    If UCase(cboItemClassCode) = "D" Then
            Chk_cboItmAccTypeCode = True
            Exit Function
    End If
    
 '   If Trim(cboItmAccTypeCode.Text) = "" Then
 '       Chk_cboItmAccTypeCode = True
 '       Exit Function
 '   End If
    
    If Chk_ItmAccTypeCode(wsDesc) = False Then
        gsMsg = "會計分類不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboItmAccTypeCode.SetFocus
        Exit Function
    End If
    
    lblDspItmAccTypeDesc = wsDesc
    Chk_cboItmAccTypeCode = True
End Function

Private Function Chk_ItmAccTypeCode(ByRef OutDesc As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    Chk_ItmAccTypeCode = False
    
    OutDesc = ""
    sSQL = "SELECT AccTypeCode, AccTypeDesc FROM MstAccountType WHERE MstAccountType.AccTypeCode = '" & Set_Quote(cboItmAccTypeCode.Text) + "' And AccTypeStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
    
    rsRcd.Close
    Set rsRcd = Nothing
    Exit Function
    End If
    
    OutDesc = ReadRs(rsRcd, "AccTypeDesc")
    Chk_ItmAccTypeCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Sub LoadForm(f As Form)
   f.WindowState = 0
   f.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
   f.Show
   f.ZOrder 0
   
End Sub

Private Sub cboItemClassCode_DropDown()
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboItemClassCode

    If gsLangID = "1" Then
        wsSQL = "SELECT ItemClassCode, ItemClassEDesc FROM MstItemClass WHERE ItemClassCode LIKE '%" & IIf(cboItemClassCode.SelLength > 0, "", Set_Quote(cboItemClassCode.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItemClassCode "
    Else
        wsSQL = "SELECT ItemClassCode, ItemClassCDesc FROM MstItemClass WHERE ItemClassCode LIKE '%" & IIf(cboItemClassCode.SelLength > 0, "", Set_Quote(cboItemClassCode.Text)) & "%' "
        wsSQL = wsSQL & "ORDER BY ItemClassCode "
    End If

    Call Ini_Combo(2, wsSQL, cboItemClassCode.Left, cboItemClassCode.Top + cboItemClassCode.Height, tblCommon, wsFormID, "TBLITEMCLASS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItemClassCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItemClassCode, 1, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItemClassCode() = False Then
            Exit Sub
        End If
        'Call Chk_KeyFld
        'txtItmChiName.SetFocus
        If wiAction = AddRec Then
            Call Ini_Scr_AfrKey
        Else
           ' cboItmCode.SetFocus
            txtItmChiName.SetFocus
        End If
    End If
End Sub

Private Sub cboItemClassCode_GotFocus()
    FocusMe cboItemClassCode
End Sub

Private Sub cboItemClassCode_LostFocus()
    FocusMe cboItemClassCode, True
End Sub

Private Function Chk_cboItemClassCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItemClassCode = False
    
    If Trim(cboItemClassCode.Text) = "" Then
        If wiAction = AddRec Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboItemClassCode.SetFocus
            Exit Function
        Else
            lblDspItemClassCode = ""
            Chk_cboItemClassCode = True
            Exit Function
        End If
    End If
    
    If Chk_ItemClassCode(cboItemClassCode.Text, wsRetName) = False Then
        lblDspItemClassCode = ""
        gsMsg = "物料分類編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboItemClassCode.SetFocus
        Exit Function
    Else
        lblDspItemClassCode = wsRetName
    End If
    
    Chk_cboItemClassCode = True
    
End Function

Private Function Chk_ItemClassCode(ByVal inCode As String, ByRef OutName As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_ItemClassCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    If gsLangID = "1" Then
        wsSQL = "SELECT ItemClassEDesc ItemClassDesc "
        wsSQL = wsSQL & " FROM MstItemClass WHERE MstItemClass.ItemClassCode = '" & Set_Quote(inCode) & "' "
    Else
        wsSQL = "SELECT ItemClassCDesc ItemClassDesc "
        wsSQL = wsSQL & " FROM MstItemClass WHERE MstItemClass.ItemClassCode = '" & Set_Quote(inCode) & "' "
    End If
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "ItemClassDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItemClassCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_KeyFld() As Boolean
    Chk_KeyFld = False
    
    If Chk_cboItmCode = False Then
        Exit Function
    End If
    
    If Chk_cboItemClassCode = False Then
        Exit Function
    End If
    
    Call Ini_Scr_AfrKey
    Chk_KeyFld = True
End Function

Private Function chk_cboVdrCode() As Boolean
    chk_cboVdrCode = False
    
    If UCase(cboItemClassCode) = "D" Then
            chk_cboVdrCode = True
            Exit Function
    End If
    
    If Trim(cboItmPVdrCode.Text) = "" Then
        chk_cboVdrCode = True
        Exit Function
    End If
        
    If Chk_VdrCode(cboItmPVdrCode.Text, wlVdrID, "", "", "") = False Then
        gsMsg = "供應商編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboItmPVdrCode.SetFocus
        Exit Function
    End If
    
    chk_cboVdrCode = True
End Function

Private Sub Calc_Price()
    txtItmBottomPrice = Format(To_Value(txtItmUnitPrice) * To_Value(txtItmDiscount), gsUprFmt)
    txtItmDefaultPrice = Format(To_Value(txtItmBottomPrice) * To_Value(txtItmMarkUp), gsUprFmt)

End Sub


Private Sub tblDetail_AfterColUpdate(ByVal ColIndex As Integer)
    Dim sTemp As String
   
    With tblDetail
        sTemp = .Columns(ColIndex)
        .UPDATE
    End With

End Sub

Private Sub tblDetail_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim wsITMCODE As String
    Dim wsItmDesc As String
    Dim wsITMID As String

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            Case ITMCODE
                'If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                '    GoTo Tbl_BeforeColUpdate_Err
                'End If
                
                If Chk_grdITMCODE(.Columns(ColIndex).Text, wsITMCODE, wsItmDesc, wsITMID) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If

                .Columns(ITMCODE).Text = wsITMCODE
                .Columns(ITMDESC).Text = wsItmDesc
                .Columns(QTY).Text = 0
                .Columns(ITMID).Text = wsITMID
                
                If Trim(.Columns(ColIndex).Text) <> wsITMCODE Then
                    .Columns(ColIndex).Text = wsITMCODE
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
    Dim sEndCode As String
    
    On Error GoTo tblDetail_ButtonClick_Err
    
    If wiAction = AddRec Then
        sEndCode = "', '" & txtItmCode & "' )"
    Else
        sEndCode = "', '" & cboITMCODE & "' )"
    End If

    With tblDetail
        Select Case ColIndex
            Case ITMCODE
                If gsLangID = 1 Then
                    wsSQL = "SELECT ITMCODE, ITMENGNAME FROM MSTITEM "
                    wsSQL = wsSQL & " WHERE ITMSTATUS = '1' AND ITMCLASS = 'P' AND ITMINACTIVE = 'N' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                    If waResult.UpperBound(1) > -1 Then
                        wsSQL = wsSQL & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                            wsSQL = wsSQL & " '" & Set_Quote(waResult(wiCtr, ITMCODE)) & IIf(wiCtr = waResult.UpperBound(1), sEndCode, "' ,")
                        Next
                    Else
                        If wiAction = AddRec Then
                            wsSQL = wsSQL & " AND ITMCODE <> '" & txtItmCode & "' "
                        Else
                            wsSQL = wsSQL & " AND ITMCODE <> '" & cboITMCODE & "' "
                        End If
                    End If
                    wsSQL = wsSQL & " ORDER BY ITMCODE "
                Else
                    wsSQL = "SELECT ITMCODE, ITMCHINAME FROM MSTITEM "
                    wsSQL = wsSQL & " WHERE ITMSTATUS = '1' AND ITMCLASS = 'P' AND ITMINACTIVE = 'N' AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                    If waResult.UpperBound(1) > -1 Then
                        wsSQL = wsSQL & " AND ITMCODE NOT IN ( "
                        For wiCtr = 0 To waResult.UpperBound(1)
                              wsSQL = wsSQL & " '" & Set_Quote(waResult(wiCtr, ITMCODE)) & IIf(wiCtr = waResult.UpperBound(1), sEndCode, "' ,")
                        Next
                    Else
                        If wiAction = AddRec Then
                            wsSQL = wsSQL & " AND ITMCODE <> '" & txtItmCode & "' "
                        Else
                            wsSQL = wsSQL & " AND ITMCODE <> '" & cboITMCODE & "' "
                        End If
                    End If
                    wsSQL = wsSQL & " ORDER BY ITMCODE "
                End If
                
                Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLB", Me.Width, Me.Height)
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
            .UPDATE
            If .Row = -1 Then
                .Row = 0
            End If
            .Refresh
            .SetFocus

        Case vbKeyReturn
            Select Case .Col
                Case ITMCODE
                    KeyCode = vbDefault
                    .Col = QTY
                Case QTY
                    KeyCode = vbKeyDown
                    .Col = ITMCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
              Select Case .Col
                Case ITMDESC, QTY
                    .Col = .Col - 1
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case ITMCODE, ITMDESC
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
        
        Case QTY
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
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
           .Col = ITMCODE
        End If
        
        'Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case ITMCODE
                    Call Chk_grdITMCODE(.Columns(ITMCODE).Text, "", "", "")
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

Private Function Chk_grdITMCODE(inAccNo As String, outAccNo As String, outAccDesc As String, OutID As String) As Boolean
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
    
    If gsLangID = "1" Then
        wsSQL = "SELECT ITMID, ITMCODE, ITMENGNAME ITMDESC FROM MstItem"
        wsSQL = wsSQL & " WHERE (ITMCODE = '" & Set_Quote(inAccNo) & "') AND ITMCLASS = 'P' AND ITMINACTIVE = 'N' "
    Else
        wsSQL = "SELECT ITMID, ITMCODE, ITMCHINAME ITMDESC FROM MstItem"
        wsSQL = wsSQL & " WHERE (ITMCODE = '" & Set_Quote(inAccNo) & "') AND ITMCLASS = 'P' AND ITMINACTIVE = 'N' "
    End If
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccNo = ReadRs(rsDes, "ITMCODE")
        outAccDesc = ReadRs(rsDes, "ITMDESC")
        OutID = ReadRs(rsDes, "ITMID")
       
        Chk_grdITMCODE = True
    Else
        outAccDesc = ""
        OutID = 0
        
        gsMsg = "沒有此物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
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
            If Trim(.Columns(ITMCODE)) = "" Then
                Exit Function
            End If
        End With
    Else
        If waResult.UpperBound(1) >= 0 Then
            If Trim(waResult(inRow, ITMCODE)) = "" And _
               Trim(waResult(inRow, ITMDESC)) = "" And _
               Trim(waResult(inRow, QTY)) = "" And _
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
        
        If Chk_grdITMCODE(waResult(LastRow, ITMCODE), "", "", "") = False Then
            .Col = ITMCODE
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

Private Sub Ini_Grid()
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
      '  .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = ITMCODE To ITMID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case ITMCODE
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).DataWidth = 30
                    .Columns(wiCtr).Button = True
                Case ITMDESC
                    .Columns(wiCtr).Width = 5000
                    .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).Locked = True
                Case QTY
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case ITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
      '  .Styles("EvenRow").BackColor = &H8000000F
    End With
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
                .UPDATE
                
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

Private Sub cmdChangeKey(ByVal wiAct As Integer)
    Dim Newfrm As New frmChangeKey
    Dim outResult As Boolean
    Dim wsNew As String
    
    Me.MousePointer = vbHourglass
    
    'Create Selection wsSql
    With Newfrm
        .KeyID = wlKey
        .KeyType = cboItemClassCode.Text
       ' Set .ctlKey = txtItmCode
        .Show vbModal
        outResult = .Result
        wsNew = .NewKey
        
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
    
    If outResult = True Then
        wiAction = wiAct
        If wiAction = AddRec Then
        txtItmCode.Text = wsNew
        Else
        Call UnLockAll(wsConnTime, wsFormID)
        cboITMCODE.Text = wsNew
        If UCase(cboItemClassCode) = "T" Then UCase(cboItemClassCode) = "P"
        If RowLock(wsConnTime, wsKeyType, cboITMCODE, wsFormID, wsUsrId) = False Then
                    gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                    MsgBox gsMsg, vbOKOnly, gsTitle
                    Exit Sub
        End If
        End If
        Call cmdSave
    
    End If
End Sub

Private Function Chk_txtItmDiscount() As Boolean
  
    Chk_txtItmDiscount = False
    
    If UCase(cboItemClassCode) = "D" Then
        Chk_txtItmDiscount = True
        Exit Function
    End If
    
    If To_Value(txtItmDiscount.Text) = 0 Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        Me.tabDetailInfo.Tab = 0
        txtItmDiscount.SetFocus
        Exit Function
    End If
    
    Chk_txtItmDiscount = True
    
End Function

