VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmB001 
   BackColor       =   &H8000000A&
   Caption         =   "書本"
   ClientHeight    =   6195
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   9885
   Icon            =   "B001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   9885
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10680
      OleObjectBlob   =   "B001.frx":08CA
      TabIndex        =   70
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboItmCode 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   2250
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   0
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
            Picture         =   "B001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   3615
      Left            =   240
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6376
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "索書資料"
      TabPicture(0)   =   "B001.frx":6C6D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblItmAuthor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblItmTranslator"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblItmLastUpd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblItmLastUpdDate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblItmCatCode"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDspItmCatDesc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblDspItmLastUpd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblDspItmLastUpdDate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblArt"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDraw"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPhoto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblEditor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblItmTypeCode"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblDspItmTypeDesc"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblItmLevelCode"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblDspItmLevelDesc"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDspItmLangDesc"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblItmLangCode"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblItmVolume"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblItmSeriesNo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblDspIctrnQty"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblIctrnQty"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtItmAuthor"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtItmTranslator"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cboItmCatCode"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtArt"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtDraw"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtPhoto"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtEditor"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cboItmTypeCode"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cboItmLevelCode"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cboItmLangCode"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtItmVolume"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtItmSeriesNo"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cboItmUOMCode"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "附加資料"
      TabPicture(1)   =   "B001.frx":6C89
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContent"
      Tab(1).Control(1)=   "fraCover"
      Tab(1).Control(2)=   "txtItmTextDir"
      Tab(1).Control(3)=   "btnItmTextDir"
      Tab(1).Control(4)=   "btnItmDir"
      Tab(1).Control(5)=   "txtItmDir"
      Tab(1).Control(6)=   "txtItmBinNo"
      Tab(1).Control(7)=   "cboItmPrintSizeCode"
      Tab(1).Control(8)=   "cboItmPackTypeCode"
      Tab(1).Control(9)=   "txtItmWeight"
      Tab(1).Control(10)=   "txtItmPage"
      Tab(1).Control(11)=   "txtItmWidth"
      Tab(1).Control(12)=   "txtItmSize"
      Tab(1).Control(13)=   "lblItmBinNo"
      Tab(1).Control(14)=   "lblWeight"
      Tab(1).Control(15)=   "lblPages"
      Tab(1).Control(16)=   "lblItmWidth"
      Tab(1).Control(17)=   "lblItmHeight"
      Tab(1).Control(18)=   "lblItmPrintSizeCode"
      Tab(1).Control(19)=   "lblDspItmPrintSizeDesc"
      Tab(1).Control(20)=   "lblDspItmPackTypeDesc"
      Tab(1).Control(21)=   "lblItmWeight"
      Tab(1).Control(22)=   "lblItmPage"
      Tab(1).Control(23)=   "lblItmSize"
      Tab(1).Control(24)=   "lblItmPackTypeCode"
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "包裝資料"
      TabPicture(2)   =   "B001.frx":6CA5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboItmPackUOMCode"
      Tab(2).Control(1)=   "txtItmVersion"
      Tab(2).Control(2)=   "txtItmPrint"
      Tab(2).Control(3)=   "txtItmPublisher"
      Tab(2).Control(4)=   "txtItmPackQty"
      Tab(2).Control(5)=   "txtItmExtWidth"
      Tab(2).Control(6)=   "txtItmExtLength"
      Tab(2).Control(7)=   "txtItmExtHeight"
      Tab(2).Control(8)=   "txtItmIntWidth"
      Tab(2).Control(9)=   "txtItmIntLength"
      Tab(2).Control(10)=   "txtItmIntHeight"
      Tab(2).Control(11)=   "medItmPrintDate"
      Tab(2).Control(12)=   "lblPackInfo"
      Tab(2).Control(13)=   "lblDspItmPackUOMDesc"
      Tab(2).Control(14)=   "lblItmPackUOMCode"
      Tab(2).Control(15)=   "lblItmVersion"
      Tab(2).Control(16)=   "lblItmPrint"
      Tab(2).Control(17)=   "lblItmPrintDate"
      Tab(2).Control(18)=   "lblItmPublisher"
      Tab(2).Control(19)=   "lblItmPackQty"
      Tab(2).Control(20)=   "lblItmExternal"
      Tab(2).Control(21)=   "lblItmInternal"
      Tab(2).Control(22)=   "lblWidth"
      Tab(2).Control(23)=   "lblHeight"
      Tab(2).Control(24)=   "lblLength"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "出版資料"
      TabPicture(3)   =   "B001.frx":6CC1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "btnPriceChange"
      Tab(3).Control(1)=   "chkItmOwnEdition"
      Tab(3).Control(2)=   "txtItmReorderQty"
      Tab(3).Control(3)=   "btnItemPrice"
      Tab(3).Control(4)=   "cboItmAccTypeCode"
      Tab(3).Control(5)=   "cboItmCurr"
      Tab(3).Control(6)=   "txtItmDefaultPrice"
      Tab(3).Control(7)=   "txtItmBottomPrice"
      Tab(3).Control(8)=   "cboItmCDisCode"
      Tab(3).Control(9)=   "txtAdj(11)"
      Tab(3).Control(10)=   "txtAdj(10)"
      Tab(3).Control(11)=   "txtAdj(9)"
      Tab(3).Control(12)=   "txtAdj(8)"
      Tab(3).Control(13)=   "txtAdj(7)"
      Tab(3).Control(14)=   "txtAdj(6)"
      Tab(3).Control(15)=   "txtAdj(5)"
      Tab(3).Control(16)=   "txtAdj(4)"
      Tab(3).Control(17)=   "txtAdj(3)"
      Tab(3).Control(18)=   "txtAdj(2)"
      Tab(3).Control(19)=   "txtAdj(1)"
      Tab(3).Control(20)=   "txtAdj(0)"
      Tab(3).Control(21)=   "chkItmInActive"
      Tab(3).Control(22)=   "chkItmInvItemFlg"
      Tab(3).Control(23)=   "chkItmTaxFlg"
      Tab(3).Control(24)=   "chkItmReorderFlg"
      Tab(3).Control(25)=   "chkItmReorderInd"
      Tab(3).Control(26)=   "txtItmPORtnDate"
      Tab(3).Control(27)=   "txtItmPORepuQty"
      Tab(3).Control(28)=   "lblItmReorderQty"
      Tab(3).Control(29)=   "lblDspItmAccTypeDesc"
      Tab(3).Control(30)=   "lblItmAccTypeCode"
      Tab(3).Control(31)=   "lblItmCurrCode"
      Tab(3).Control(32)=   "lblItmDefaultPrice"
      Tab(3).Control(33)=   "lblItmBottomPrice"
      Tab(3).Control(34)=   "lblDspItmCDisDesc"
      Tab(3).Control(35)=   "lblItmCDisCode"
      Tab(3).Control(36)=   "lblAdj(11)"
      Tab(3).Control(37)=   "lblAdj(10)"
      Tab(3).Control(38)=   "lblAdj(9)"
      Tab(3).Control(39)=   "lblAdj(8)"
      Tab(3).Control(40)=   "lblAdj(7)"
      Tab(3).Control(41)=   "lblAdj(6)"
      Tab(3).Control(42)=   "lblAdj(5)"
      Tab(3).Control(43)=   "lblAdj(4)"
      Tab(3).Control(44)=   "lblAdj(3)"
      Tab(3).Control(45)=   "lblAdj(2)"
      Tab(3).Control(46)=   "lblAdj(1)"
      Tab(3).Control(47)=   "lblAdj(0)"
      Tab(3).Control(48)=   "lblItmPORtnDate"
      Tab(3).Control(49)=   "lblItmPORepuQty"
      Tab(3).ControlCount=   50
      Begin VB.CommandButton btnPriceChange 
         Caption         =   "PRICECHANGE"
         Enabled         =   0   'False
         Height          =   555
         Left            =   -67320
         TabIndex        =   48
         Top             =   1020
         Width           =   1455
      End
      Begin VB.ComboBox cboItmUOMCode 
         Height          =   300
         Left            =   3120
         TabIndex        =   154
         Top             =   2460
         Width           =   930
      End
      Begin VB.CheckBox chkItmOwnEdition 
         Alignment       =   1  '靠右對齊
         Caption         =   "暫停發貨 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   49
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Frame fraContent 
         Height          =   2655
         Left            =   -72600
         TabIndex        =   153
         Top             =   120
         Width           =   2535
         Begin RichTextLib.RichTextBox rtContent 
            Height          =   2295
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   4048
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"B001.frx":6CDD
         End
      End
      Begin VB.Frame fraCover 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   151
         Top             =   120
         Width           =   2295
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   2535
            Left            =   2280
            TabIndex        =   152
            Top             =   120
            Width           =   15
         End
         Begin VB.Image imgCover 
            Height          =   2265
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.TextBox txtItmTextDir 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -72600
         TabIndex        =   21
         Top             =   2880
         Width           =   2235
      End
      Begin VB.CommandButton btnItmTextDir 
         Caption         =   "..."
         Height          =   315
         Left            =   -70320
         TabIndex        =   22
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton btnItmDir 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   -72960
         TabIndex        =   20
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox txtItmDir 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -74760
         TabIndex        =   19
         Top             =   2880
         Width           =   1755
      End
      Begin VB.TextBox txtItmReorderQty 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -66840
         TabIndex        =   54
         Top             =   1665
         Width           =   945
      End
      Begin VB.TextBox txtItmBinNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69000
         TabIndex        =   29
         Top             =   2340
         Width           =   930
      End
      Begin VB.CommandButton btnItemPrice 
         Caption         =   "ITEMPRICE"
         Enabled         =   0   'False
         Height          =   555
         Left            =   -68880
         TabIndex        =   47
         Top             =   1020
         Width           =   1455
      End
      Begin VB.ComboBox cboItmAccTypeCode 
         Height          =   300
         Left            =   -68880
         TabIndex        =   46
         Top             =   660
         Width           =   930
      End
      Begin VB.ComboBox cboItmCurr 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "B001.frx":7048
         Left            =   -73680
         List            =   "B001.frx":704A
         TabIndex        =   42
         Top             =   300
         Width           =   2985
      End
      Begin VB.TextBox txtItmDefaultPrice 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73680
         TabIndex        =   43
         Top             =   660
         Width           =   2985
      End
      Begin VB.TextBox txtItmBottomPrice 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73680
         TabIndex        =   44
         Top             =   1020
         Width           =   2985
      End
      Begin VB.ComboBox cboItmPackUOMCode 
         Height          =   300
         Left            =   -73560
         TabIndex        =   34
         Top             =   2340
         Width           =   930
      End
      Begin VB.TextBox txtItmVersion 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -70320
         TabIndex        =   32
         Top             =   900
         Width           =   885
      End
      Begin VB.TextBox txtItmPrint 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -67200
         TabIndex        =   33
         Top             =   900
         Width           =   1185
      End
      Begin VB.TextBox txtItmPublisher 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73560
         TabIndex        =   30
         Top             =   540
         Width           =   7545
      End
      Begin VB.ComboBox cboItmCDisCode 
         Height          =   300
         Left            =   -68880
         TabIndex        =   45
         Top             =   300
         Width           =   930
      End
      Begin VB.ComboBox cboItmPrintSizeCode 
         Height          =   300
         Left            =   -69000
         TabIndex        =   24
         Top             =   900
         Width           =   930
      End
      Begin VB.ComboBox cboItmPackTypeCode 
         Height          =   300
         Left            =   -69000
         TabIndex        =   23
         Top             =   540
         Width           =   930
      End
      Begin VB.TextBox txtItmWeight 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69000
         TabIndex        =   27
         Top             =   1620
         Width           =   1635
      End
      Begin VB.TextBox txtItmPage 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69000
         TabIndex        =   28
         Top             =   1980
         Width           =   1905
      End
      Begin VB.TextBox txtItmWidth 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -67560
         TabIndex        =   26
         Top             =   1260
         Width           =   930
      End
      Begin VB.TextBox txtItmSize 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69000
         TabIndex        =   25
         Top             =   1260
         Width           =   930
      End
      Begin VB.TextBox txtItmSeriesNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   12
         Top             =   2100
         Width           =   2985
      End
      Begin VB.TextBox txtItmVolume 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   18
         Top             =   2100
         Width           =   885
      End
      Begin VB.ComboBox cboItmLangCode 
         Height          =   300
         Left            =   6120
         TabIndex        =   17
         Top             =   1740
         Width           =   930
      End
      Begin VB.ComboBox cboItmLevelCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   11
         Top             =   1740
         Width           =   930
      End
      Begin VB.ComboBox cboItmTypeCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   10
         Top             =   1380
         Width           =   930
      End
      Begin VB.TextBox txtEditor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   9
         Top             =   1020
         Width           =   2985
      End
      Begin VB.TextBox txtPhoto 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   15
         Top             =   1020
         Width           =   2985
      End
      Begin VB.TextBox txtDraw 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   13
         Top             =   300
         Width           =   2985
      End
      Begin VB.TextBox txtArt 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   14
         Top             =   660
         Width           =   2985
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   -66840
         TabIndex        =   66
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   -67560
         TabIndex        =   65
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   -68280
         TabIndex        =   64
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   -69000
         TabIndex        =   63
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   -69720
         TabIndex        =   62
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   -70440
         TabIndex        =   61
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   -71160
         TabIndex        =   60
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   -71880
         TabIndex        =   59
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   -72600
         TabIndex        =   58
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   -73320
         TabIndex        =   57
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   -74040
         TabIndex        =   56
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtAdj 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   -74760
         TabIndex        =   55
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtItmPackQty 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -73560
         TabIndex        =   38
         Top             =   2700
         Width           =   915
      End
      Begin VB.TextBox txtItmExtWidth 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69000
         TabIndex        =   35
         Top             =   2340
         Width           =   995
      End
      Begin VB.TextBox txtItmExtLength 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -67035
         TabIndex        =   37
         Top             =   2340
         Width           =   995
      End
      Begin VB.TextBox txtItmExtHeight 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -68025
         TabIndex        =   36
         Top             =   2340
         Width           =   995
      End
      Begin VB.TextBox txtItmIntWidth 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -69000
         TabIndex        =   39
         Top             =   2700
         Width           =   995
      End
      Begin VB.TextBox txtItmIntLength 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -67035
         TabIndex        =   41
         Top             =   2700
         Width           =   995
      End
      Begin VB.TextBox txtItmIntHeight 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -68025
         TabIndex        =   40
         Top             =   2700
         Width           =   995
      End
      Begin VB.CheckBox chkItmInActive 
         Alignment       =   1  '靠右對齊
         Caption         =   "暫停發貨 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   50
         Top             =   1740
         Width           =   1455
      End
      Begin VB.CheckBox chkItmInvItemFlg 
         Alignment       =   1  '靠右對齊
         Caption         =   "非存貨 :"
         Height          =   180
         Left            =   -73080
         TabIndex        =   51
         Top             =   1740
         Width           =   1455
      End
      Begin VB.CheckBox chkItmTaxFlg 
         Alignment       =   1  '靠右對齊
         Caption         =   "稅收 :"
         Height          =   180
         Left            =   -71160
         TabIndex        =   52
         Top             =   1740
         Width           =   1455
      End
      Begin VB.ComboBox cboItmCatCode 
         Height          =   300
         Left            =   6120
         TabIndex        =   16
         Top             =   1380
         Width           =   930
      End
      Begin VB.CheckBox chkItmReorderFlg 
         Alignment       =   1  '靠右對齊
         Caption         =   "低存量 :"
         Height          =   180
         Left            =   -69240
         TabIndex        =   53
         Top             =   1740
         Width           =   975
      End
      Begin VB.CheckBox chkItmReorderInd 
         Alignment       =   1  '靠右對齊
         Caption         =   "再版指標 :"
         Height          =   180
         Left            =   -74760
         TabIndex        =   67
         Top             =   2910
         Width           =   1215
      End
      Begin VB.TextBox txtItmPORtnDate 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -67080
         TabIndex        =   69
         Top             =   2820
         Width           =   1185
      End
      Begin VB.TextBox txtItmPORepuQty 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -71520
         TabIndex        =   68
         Top             =   2820
         Width           =   1065
      End
      Begin VB.TextBox txtItmTranslator 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   660
         Width           =   2985
      End
      Begin VB.TextBox txtItmAuthor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   7
         Top             =   300
         Width           =   2985
      End
      Begin MSMask.MaskEdBox medItmPrintDate 
         Height          =   285
         Left            =   -73560
         TabIndex        =   31
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblItmReorderQty 
         Caption         =   "再版指標 :"
         Height          =   240
         Left            =   -67920
         TabIndex        =   148
         Top             =   1740
         Width           =   1140
      End
      Begin VB.Label lblItmBinNo 
         Caption         =   "ITMBINNO"
         Height          =   240
         Left            =   -69960
         TabIndex        =   147
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblDspItmAccTypeDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -67920
         TabIndex        =   146
         Top             =   660
         Width           =   2025
      End
      Begin VB.Label lblItmAccTypeCode 
         Caption         =   "會計版別 :"
         Height          =   240
         Left            =   -69840
         TabIndex        =   145
         Top             =   705
         Width           =   900
      End
      Begin VB.Label lblItmCurrCode 
         Caption         =   "貨幣 :"
         Height          =   240
         Left            =   -74880
         TabIndex        =   144
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lblItmDefaultPrice 
         Caption         =   "預設售價 :"
         Height          =   240
         Left            =   -74880
         TabIndex        =   143
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblItmBottomPrice 
         Caption         =   "底價 :"
         Height          =   240
         Left            =   -74880
         TabIndex        =   142
         Top             =   1065
         Width           =   1020
      End
      Begin VB.Label lblPackInfo 
         Caption         =   "印刷次數 :"
         Height          =   240
         Left            =   -72480
         TabIndex        =   141
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblDspItmPackUOMDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71520
         TabIndex        =   140
         Top             =   2340
         Width           =   1065
      End
      Begin VB.Label lblItmPackUOMCode 
         Caption         =   "包裝類別 :"
         Height          =   240
         Left            =   -74760
         TabIndex        =   139
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblItmVersion 
         Caption         =   "版次 :"
         Height          =   240
         Left            =   -71160
         TabIndex        =   138
         Top             =   975
         Width           =   780
      End
      Begin VB.Label lblItmPrint 
         Caption         =   "印刷次數 :"
         Height          =   240
         Left            =   -68160
         TabIndex        =   137
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblItmPrintDate 
         Caption         =   "初版日期 :"
         Height          =   240
         Left            =   -74760
         TabIndex        =   136
         Top             =   975
         Width           =   900
      End
      Begin VB.Label lblItmPublisher 
         Caption         =   "出版社 :"
         Height          =   240
         Left            =   -74760
         TabIndex        =   135
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblDspItmCDisDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -67920
         TabIndex        =   134
         Top             =   300
         Width           =   2025
      End
      Begin VB.Label lblItmCDisCode 
         Caption         =   "折扣類別 :"
         Height          =   240
         Left            =   -69840
         TabIndex        =   133
         Top             =   345
         Width           =   900
      End
      Begin VB.Label lblWeight 
         Caption         =   "重量 :"
         Height          =   240
         Left            =   -66840
         TabIndex        =   132
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lblPages 
         Caption         =   "頁數 :"
         Height          =   240
         Left            =   -66840
         TabIndex        =   131
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lblItmWidth 
         Caption         =   "闊"
         Height          =   240
         Left            =   -66600
         TabIndex        =   130
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblItmHeight 
         Caption         =   "闊"
         Height          =   240
         Left            =   -67920
         TabIndex        =   129
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label lblItmPrintSizeCode 
         Caption         =   "開度 :"
         Height          =   240
         Left            =   -69960
         TabIndex        =   128
         Top             =   975
         Width           =   900
      End
      Begin VB.Label lblDspItmPrintSizeDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -68040
         TabIndex        =   127
         Top             =   900
         Width           =   2025
      End
      Begin VB.Label lblDspItmPackTypeDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -68040
         TabIndex        =   126
         Top             =   540
         Width           =   2025
      End
      Begin VB.Label lblItmWeight 
         Caption         =   "重量 :"
         Height          =   240
         Left            =   -69960
         TabIndex        =   125
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label lblItmPage 
         Caption         =   "頁數 :"
         Height          =   240
         Left            =   -69960
         TabIndex        =   124
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label lblItmSize 
         Caption         =   "尺寸 :"
         Height          =   240
         Left            =   -69960
         TabIndex        =   123
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblItmPackTypeCode 
         Caption         =   "裝幀 :"
         Height          =   240
         Left            =   -69960
         TabIndex        =   122
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblIctrnQty 
         Caption         =   "ICTRNQTY"
         Height          =   240
         Left            =   120
         TabIndex        =   120
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label lblDspIctrnQty 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1560
         TabIndex        =   119
         Top             =   2460
         Width           =   1425
      End
      Begin VB.Label lblItmSeriesNo 
         Caption         =   "套書書號 :"
         Height          =   240
         Left            =   120
         TabIndex        =   118
         Top             =   2145
         Width           =   900
      End
      Begin VB.Label lblItmVolume 
         Caption         =   "冊次 :"
         Height          =   240
         Left            =   4680
         TabIndex        =   117
         Top             =   2145
         Width           =   660
      End
      Begin VB.Label lblItmLangCode 
         Caption         =   "語種 :"
         Height          =   240
         Left            =   4680
         TabIndex        =   116
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Label lblDspItmLangDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   7080
         TabIndex        =   115
         Top             =   1740
         Width           =   2025
      End
      Begin VB.Label lblDspItmLevelDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2520
         TabIndex        =   114
         Top             =   1740
         Width           =   2025
      End
      Begin VB.Label lblItmLevelCode 
         Caption         =   "程度 :"
         Height          =   240
         Left            =   120
         TabIndex        =   113
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lblDspItmTypeDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2520
         TabIndex        =   112
         Top             =   1380
         Width           =   2025
      End
      Begin VB.Label lblItmTypeCode 
         Caption         =   "圖書分類 :"
         Height          =   240
         Left            =   120
         TabIndex        =   111
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label lblEditor 
         Caption         =   "EDITOR"
         Height          =   240
         Left            =   120
         TabIndex        =   110
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblPhoto 
         Caption         =   "PHOTO"
         Height          =   240
         Left            =   4680
         TabIndex        =   109
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label lblDraw 
         Caption         =   "DRAW"
         Height          =   240
         Left            =   4680
         TabIndex        =   108
         Top             =   375
         Width           =   1380
      End
      Begin VB.Label lblArt 
         Caption         =   "ART"
         Height          =   240
         Left            =   4680
         TabIndex        =   107
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   11
         Left            =   -66840
         TabIndex        =   106
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   10
         Left            =   -67560
         TabIndex        =   105
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   9
         Left            =   -68280
         TabIndex        =   104
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   8
         Left            =   -69000
         TabIndex        =   103
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   7
         Left            =   -69720
         TabIndex        =   102
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   6
         Left            =   -70440
         TabIndex        =   101
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   5
         Left            =   -71160
         TabIndex        =   100
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   4
         Left            =   -71880
         TabIndex        =   99
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   3
         Left            =   -72600
         TabIndex        =   98
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   2
         Left            =   -73320
         TabIndex        =   97
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   1
         Left            =   -74040
         TabIndex        =   96
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblAdj 
         Alignment       =   2  '置中對齊
         BackColor       =   &H80000010&
         Caption         =   "闊"
         Height          =   240
         Index           =   0
         Left            =   -74760
         TabIndex        =   95
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblItmPackQty 
         Caption         =   "數量 :"
         Height          =   240
         Left            =   -74760
         TabIndex        =   94
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblItmExternal 
         Caption         =   "外 :"
         Height          =   240
         Left            =   -69360
         TabIndex        =   93
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label lblItmInternal 
         Caption         =   "內 :"
         Height          =   240
         Left            =   -69360
         TabIndex        =   92
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label lblWidth 
         Alignment       =   2  '置中對齊
         Caption         =   "闊"
         Height          =   240
         Left            =   -69000
         TabIndex        =   91
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lblHeight 
         Alignment       =   2  '置中對齊
         Caption         =   "高"
         Height          =   240
         Left            =   -68040
         TabIndex        =   90
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lblLength 
         Alignment       =   2  '置中對齊
         Caption         =   "長"
         Height          =   240
         Left            =   -67080
         TabIndex        =   89
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lblItmPORtnDate 
         Caption         =   "再版日期 (回倉日期) :"
         Height          =   240
         Left            =   -68880
         TabIndex        =   88
         Top             =   2895
         Width           =   1740
      End
      Begin VB.Label lblItmPORepuQty 
         Caption         =   "再版數量 :"
         Height          =   240
         Left            =   -72600
         TabIndex        =   87
         Top             =   2895
         Width           =   1020
      End
      Begin VB.Label lblDspItmLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   6120
         TabIndex        =   86
         Top             =   2820
         Width           =   3135
      End
      Begin VB.Label lblDspItmLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   1560
         TabIndex        =   85
         Top             =   2820
         Width           =   3015
      End
      Begin VB.Label lblDspItmCatDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   7080
         TabIndex        =   84
         Top             =   1380
         Width           =   2025
      End
      Begin VB.Label lblItmCatCode 
         Caption         =   "杜威分類 :"
         Height          =   240
         Left            =   4680
         TabIndex        =   83
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label lblItmLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4800
         TabIndex        =   82
         Top             =   2895
         Width           =   1260
      End
      Begin VB.Label lblItmLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   120
         TabIndex        =   81
         Top             =   2895
         Width           =   1380
      End
      Begin VB.Label lblItmTranslator 
         Caption         =   "譯者 :"
         Height          =   240
         Left            =   120
         TabIndex        =   80
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblItmAuthor 
         Caption         =   "作者姓/名:"
         Height          =   240
         Left            =   120
         TabIndex        =   79
         Top             =   375
         Width           =   1380
      End
   End
   Begin VB.Frame fraCustomerInfo 
      Caption         =   "書本資料"
      Height          =   1695
      Left            =   240
      TabIndex        =   71
      Top             =   600
      Width           =   9375
      Begin VB.TextBox txtItmEngName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   6
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtItmChiName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtItmGrpEngName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   7575
      End
      Begin VB.TextBox txtItmGrpChiName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   600
         Width           =   7575
      End
      Begin VB.TextBox txtItmCode 
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Tag             =   "K"
         Top             =   240
         Width           =   2250
      End
      Begin VB.TextBox txtItmBarCode 
         Height          =   300
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   2250
      End
      Begin VB.Label lblUnitPrice 
         Caption         =   "UNITPRICE"
         Height          =   240
         Left            =   7320
         TabIndex        =   150
         Top             =   300
         Width           =   660
      End
      Begin VB.Label lblDspUnitPrice 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   8040
         TabIndex        =   149
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblItmEngName 
         Caption         =   "書名 (英文) :"
         Height          =   240
         Left            =   4800
         TabIndex        =   77
         Top             =   1365
         Width           =   1020
      End
      Begin VB.Label lblItmBarCode 
         Caption         =   "條碼編號 :"
         Height          =   255
         Left            =   3960
         TabIndex        =   76
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblItmChiName 
         Caption         =   "ITMCHINAME"
         Height          =   240
         Left            =   120
         TabIndex        =   75
         Top             =   1365
         Width           =   1140
      End
      Begin VB.Label lblItmGrpEngName 
         Caption         =   "叢書名稱 (英文):"
         Height          =   240
         Left            =   120
         TabIndex        =   74
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Label lblItmGrpChiName 
         Caption         =   "叢書名稱 (中文):"
         Height          =   240
         Left            =   120
         TabIndex        =   73
         Top             =   660
         Width           =   1500
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
         TabIndex        =   72
         Top             =   300
         Width           =   1020
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
      TabIndex        =   155
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
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
Attribute VB_Name = "frmB001"
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

Dim wcCombo As Control

Private Const wsKeyType = "MstItem"
Private wsUsrId As String
Private wsTrnCd As String
Private wsFormID As String
Private wsConnTime As String

Private wdOldPrice As Double

Private Sub btnItmDir_Click()
    Dim wsFilePath As String
    
    cdlgDir.FileName = ""
    cdlgDir.ShowOpen
    wsFilePath = cdlgDir.FileName
    
    If Trim(wsFilePath) = "" Then Exit Sub
    
    If Chk_Load_Cover(wsFilePath) Then
        txtItmDir = wsFilePath
        txtItmTextDir.SetFocus
    End If
End Sub

Private Sub btnItemPrice_Click()
    Me.MousePointer = vbHourglass
    frmIP001.ItmCode = cboItmCode
    frmIP001.Show
    Me.MousePointer = vbNormal
End Sub

Private Sub btnItmTextDir_Click()
    Dim wsFilePath As String
    
    cdlgDir.FileName = ""
    cdlgDir.ShowOpen
    
    wsFilePath = cdlgDir.FileName
    
    If Trim(wsFilePath) = "" Then Exit Sub
    
    If Chk_Load_Content(wsFilePath) Then
        txtItmTextDir = wsFilePath
        cboItmPackTypeCode.SetFocus
    End If
End Sub

Private Sub btnPriceChange_Click()
    frmB0011.InBookName = Me.txtItmChiName
    frmB0011.inISBN = Me.cboItmCode
    frmB0011.InItemID = wlKey
    frmB0011.Show vbModal
End Sub

Private Sub cboItmUOMCode_DropDown()
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmUOMCode
    
    wsSql = "SELECT UOMCode, UOMDesc FROM MstUOM WHERE UOMStatus = '1'"
    wsSql = wsSql & "ORDER BY UOMCode "
    Call Ini_Combo(2, wsSql, cboItmUOMCode.Left + tabDetailInfo.Left, cboItmUOMCode.Top + cboItmUOMCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLUOM", Me.Width, Me.Height)
    
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
        txtDraw.SetFocus
    End If
End Sub

Private Sub cboItmUOMCode_LostFocus()
    FocusMe cboItmUOMCode, True
End Sub

Private Sub chkItmInActive_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        chkItmInvItemFlg.SetFocus
    End If
End Sub

Private Sub chkItmInvItemFlg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        chkItmTaxFlg.SetFocus
    End If
End Sub

Private Sub chkItmOwnEdition_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        chkItmInActive.SetFocus
    End If
End Sub

Private Sub chkItmReorderFlg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        txtItmReorderQty.SetFocus
    End If
End Sub

Private Sub chkItmReorderInd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        txtItmPORepuQty.SetFocus
    End If
End Sub


Private Sub chkItmTaxFlg_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        chkItmReorderFlg.SetFocus
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
        
        Case vbKeyF8
            Call cmdOpen
        
        
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
    Ini_Scr
    
    MousePointer = vbDefault
  
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 6600
        Me.Width = 10005
    End If
End Sub

'-- Set toolbar buttons status in different mode, Default, AddEdit, None.
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
                .Buttons(tcFind).Enabled = True
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
Public Sub SetFieldStatus(ByVal SSTATUS As String)
Dim i As Integer
    Select Case SSTATUS
        Case "Default"
            Me.cboItmCode.Enabled = False
            Me.txtItmCode.Enabled = False
            Me.txtItmBarCode.Enabled = False
            Me.txtItmGrpChiName.Enabled = False
            Me.txtItmGrpEngName.Enabled = False
            Me.txtItmChiName.Enabled = False
            Me.txtItmEngName.Enabled = False
            
            'Tab 0 fields
            Me.txtItmAuthor.Enabled = False
            Me.txtItmTranslator.Enabled = False
            Me.txtEditor.Enabled = False
            Me.cboItmTypeCode.Enabled = False
            Me.cboItmLevelCode.Enabled = False
            Me.txtItmSeriesNo.Enabled = False
            Me.cboItmUOMCode.Enabled = False
            Me.txtDraw.Enabled = False
            Me.txtArt.Enabled = False
            Me.txtPhoto.Enabled = False
            Me.cboItmCatCode.Enabled = False
            Me.cboItmLangCode.Enabled = False
            Me.txtItmVolume.Enabled = False
            
            'Tab 1 fields
            Me.txtItmDir.Enabled = False
            Me.btnItmDir.Enabled = False
            Me.txtItmTextDir.Enabled = False
            Me.btnItmTextDir.Enabled = False
            Me.cboItmPackTypeCode.Enabled = False
            Me.cboItmPrintSizeCode.Enabled = False
            Me.txtItmSize.Enabled = False
            Me.txtItmWidth.Enabled = False
            Me.txtItmWeight.Enabled = False
            Me.txtItmPage.Enabled = False
            Me.txtItmBinNo.Enabled = False
            
            'Tab 2 fields
            Me.txtItmPublisher.Enabled = False
            Me.medItmPrintDate.Enabled = False
            Me.txtItmVersion.Enabled = False
            Me.txtItmPrint.Enabled = False
            Me.cboItmPackUOMCode.Enabled = False
            Me.txtItmExtWidth.Enabled = False
            Me.txtItmExtHeight.Enabled = False
            Me.txtItmExtLength.Enabled = False
            Me.txtItmPackQty.Enabled = False
            Me.txtItmIntWidth.Enabled = False
            Me.txtItmIntHeight.Enabled = False
            Me.txtItmIntLength.Enabled = False
            
            'Tab 3 fields
            Me.cboItmCurr.Enabled = False
            Me.txtItmDefaultPrice.Enabled = False
            Me.txtItmBottomPrice.Enabled = False
            Me.cboItmCDisCode.Enabled = False
            Me.cboItmAccTypeCode.Enabled = False
            Me.btnItemPrice.Enabled = False
            Me.btnPriceChange.Enabled = False
            Me.chkItmInActive.Enabled = False
            Me.chkItmInvItemFlg.Enabled = False
            Me.chkItmTaxFlg.Enabled = False
            Me.chkItmReorderFlg.Enabled = False
            Me.txtItmReorderQty.Enabled = False
            Me.chkItmOwnEdition.Enabled = False
            
            For i = 0 To 11
                Me.txtAdj(i).Enabled = False
                Me.txtAdj(i) = 0
            Next i
            
            Me.chkItmReorderInd.Enabled = False
            Me.txtItmPORepuQty.Enabled = False
            Me.txtItmPORtnDate.Enabled = False
            
            Me.txtItmVersion = 0
            Me.txtItmVolume = 0
            Me.txtItmDefaultPrice = 0
            Me.txtItmBottomPrice = 0
            Me.txtItmReorderQty = 0
            Me.txtItmPORepuQty = 0
            Me.txtItmPackQty = 0
            Me.txtItmSize = 0
            Me.txtItmWidth = 0
            Me.txtItmExtWidth = 0
            Me.txtItmExtHeight = 0
            Me.txtItmExtLength = 0
            Me.txtItmPage = 0
            Me.txtItmWeight = 0
            Me.txtItmIntWidth = 0
            Me.txtItmIntHeight = 0
            Me.txtItmIntLength = 0
            
        Case "AfrActAdd"
            Me.cboItmCode.Enabled = False
            Me.cboItmCode.Visible = False
            
            Me.txtItmCode.Enabled = True
            Me.txtItmCode.Visible = True
            
       Case "AfrActEdit"
            Me.cboItmCode.Enabled = True
            Me.cboItmCode.Visible = True
            
            Me.txtItmCode.Enabled = False
            Me.txtItmCode.Visible = False
            
        Case "AfrKey"
            Me.txtItmCode.Enabled = False
            Me.cboItmCode.Enabled = False
            
            Me.txtItmBarCode.Enabled = True
            Me.txtItmGrpChiName.Enabled = True
            Me.txtItmGrpEngName.Enabled = True
            Me.txtItmChiName.Enabled = True
            Me.txtItmEngName.Enabled = True
            
            'Tab 0 fields
            Me.txtItmAuthor.Enabled = True
            Me.txtItmTranslator.Enabled = True
            Me.txtEditor.Enabled = True
            Me.cboItmTypeCode.Enabled = True
            Me.cboItmLevelCode.Enabled = True
            Me.txtItmSeriesNo.Enabled = True
            Me.cboItmUOMCode.Enabled = True
            Me.txtDraw.Enabled = True
            Me.txtArt.Enabled = True
            Me.txtPhoto.Enabled = True
            Me.cboItmCatCode.Enabled = True
            Me.cboItmLangCode.Enabled = True
            Me.txtItmVolume.Enabled = True
            
            'Tab 1 fields
            Me.txtItmDir.Enabled = True
            Me.btnItmDir.Enabled = True
            Me.txtItmTextDir.Enabled = True
            Me.btnItmTextDir.Enabled = True
            Me.cboItmPackTypeCode.Enabled = True
            Me.cboItmPrintSizeCode.Enabled = True
            Me.txtItmSize.Enabled = True
            Me.txtItmWidth.Enabled = True
            Me.txtItmWeight.Enabled = True
            Me.txtItmPage.Enabled = True
            Me.txtItmBinNo.Enabled = True
            
            'Tab 2 fields
            Me.txtItmPublisher.Enabled = True
            Me.medItmPrintDate.Enabled = True
            Me.txtItmVersion.Enabled = True
            Me.txtItmPrint.Enabled = True
            Me.cboItmPackUOMCode.Enabled = True
            Me.txtItmExtWidth.Enabled = True
            Me.txtItmExtHeight.Enabled = True
            Me.txtItmExtLength.Enabled = True
            Me.txtItmPackQty.Enabled = True
            Me.txtItmIntWidth.Enabled = True
            Me.txtItmIntHeight.Enabled = True
            Me.txtItmIntLength.Enabled = True
            
            'Tab 3 fields
            Me.cboItmCurr.Enabled = True
            Me.txtItmDefaultPrice.Enabled = True
            Me.txtItmBottomPrice.Enabled = True
            Me.cboItmCDisCode.Enabled = True
            Me.cboItmAccTypeCode.Enabled = True
            Me.chkItmInActive.Enabled = True
            Me.chkItmInvItemFlg.Enabled = True
            Me.chkItmTaxFlg.Enabled = True
          '  Me.chkItmReorderFlg.Enabled = True
            Me.chkItmOwnEdition.Enabled = True
            Me.txtItmReorderQty.Enabled = True
            
            For i = 0 To 11
                Me.txtAdj(i).Enabled = True
                Me.txtAdj(i) = 0
            Next i
            
            Me.chkItmReorderInd.Enabled = True
           ' Me.txtItmPORepuQty.Enabled = True
           ' Me.txtItmPORtnDate.Enabled = True
        
            If wiAction = CorRec Then
                Me.btnItemPrice.Enabled = True
                Me.btnPriceChange.Enabled = True
            End If
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
        
    InputValidation = False
    
    
    If Chk_txtItmChiName = False Then
        Exit Function
    End If
    
    If Chk_txtItmAuthor = False Then
        Exit Function
    End If
    
    If Chk_cboItmTypeCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmCatCode = False Then
        Exit Function
    End If
    
    'If Chk_cboItmStoreCode = False Then
    '    Exit Function
    'End If
    
    If Chk_cboItmLevelCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmLangCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmPackTypeCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmUOMCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmPackUOMCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmPrintSizeCode = False Then
        Exit Function
    End If
    
    If Chk_medItmPrintDate = False Then
        Exit Function
    End If
    
    If Chk_cboItmCurr = False Then
        Exit Function
    End If
    
    If Chk_cboItmCDisCode = False Then
        Exit Function
    End If
    
    If Chk_cboItmAccTypeCode() = False Then
        Exit Function
    End If
    
    InputValidation = True
    
End Function

Public Function LoadRecord() As Boolean
    Dim wsSql As String
    Dim rsRcd As New ADODB.Recordset
    
        
    wsSql = "SELECT * "
    wsSql = wsSql + "From MstItem "
    wsSql = wsSql + "WHERE (((MstItem.ItmCode)='" + Set_Quote(cboItmCode) + "') AND ((MstItem.ItmStatus)='1'));"

    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "ItmID")
        
        Me.txtItmBarCode = ReadRs(rsRcd, "ItmBarCode")
        Me.txtItmGrpChiName = ReadRs(rsRcd, "ItmGrpChiName")
        Me.txtItmGrpEngName = ReadRs(rsRcd, "ItmGrpEngName")
        Me.txtItmChiName = ReadRs(rsRcd, "ItmChiName")
        Me.txtItmEngName = ReadRs(rsRcd, "ItmEngName")
        
        'Tab 0
        Me.txtItmAuthor = ReadRs(rsRcd, "ItmAuthor")
        Me.txtItmTranslator = ReadRs(rsRcd, "ItmTranslator")
        Me.txtEditor = ReadRs(rsRcd, "ItmEditor")
        Me.cboItmTypeCode = ReadRs(rsRcd, "ItmItmTypeCode")
        lblDspItmTypeDesc = LoadDescByCode("MstItemType", "ItmTypeCode", "ItmTypeChiDesc", cboItmTypeCode, True)
        Me.cboItmLevelCode = ReadRs(rsRcd, "ItmLevelCode")
        lblDspItmLevelDesc = LoadDescByCode("MstLevel", "LevelCode", "LevelDesc", cboItmLevelCode, True)
        Me.txtItmSeriesNo = ReadRs(rsRcd, "ItmSeriesNo")
        Me.cboItmUOMCode = ReadRs(rsRcd, "ItmUOMCode")
        Me.txtDraw = ReadRs(rsRcd, "ItmDraw")
        Me.txtArt = ReadRs(rsRcd, "ItmArt")
        Me.txtPhoto = ReadRs(rsRcd, "ItmPhoto")
        Me.cboItmCatCode = ReadRs(rsRcd, "ItmCatCode")
        lblDspItmCatDesc = LoadDescByCode("MstCategory", "CatCode", "CatDesc", cboItmCatCode, True)
        Me.cboItmLangCode = ReadRs(rsRcd, "ItmLangCode")
        lblDspItmLangDesc = LoadDescByCode("MstLanguage", "LangCode", "LangDesc", cboItmLangCode, True)
        Me.txtItmVolume = ReadRs(rsRcd, "ItmVolume")
        Me.lblDspItmLastUpd = ReadRs(rsRcd, "ItmLastUpd")
        Me.lblDspItmLastUpdDate = ReadRs(rsRcd, "ItmLastUpdDate")
        
        Me.lblDspIctrnQty = Get_StockAvailable(wlKey, "", "")
        
        'Tab 1
        Me.txtItmDir = ReadRs(rsRcd, "ItmDir")
        Load_Cover txtItmDir
        Me.txtItmTextDir = ReadRs(rsRcd, "ItmTextDir")
        Load_Content txtItmTextDir
        Me.cboItmPackTypeCode = ReadRs(rsRcd, "ItmPackTypeCode")
        lblDspItmPackTypeDesc = LoadDescByCode("MstPackingType", "PackTypeCode", "PackTypeDesc", cboItmPackTypeCode, True)
        Me.cboItmPrintSizeCode = ReadRs(rsRcd, "ItmPrintSizeCode")
        lblDspItmPrintSizeDesc = LoadDescByCode("MstPrintSize", "PrintSizeCode", "PrintSizeDesc", cboItmPrintSizeCode, True)
        Me.txtItmSize = Format(To_Value(ReadRs(rsRcd, "ItmSize")), gsQtyFmt)
        Me.txtItmWidth = Format(To_Value(ReadRs(rsRcd, "ItmWidth")), gsAmtFmt)
        Me.txtItmWeight = Format(To_Value(ReadRs(rsRcd, "ItmWeight")), gsAmtFmt)
        Me.txtItmPage = Format(To_Value(ReadRs(rsRcd, "ItmPage")), gsQtyFmt)
        Me.txtItmBinNo = ReadRs(rsRcd, "ItmBinNo")
        
        'Tab 2
        Me.txtItmPublisher = ReadRs(rsRcd, "ItmPublisher")
        Me.medItmPrintDate = Dsp_PeriodDate(ReadRs(rsRcd, "ItmPrintDate"))
        Me.txtItmVersion = Format(To_Value(ReadRs(rsRcd, "ItmVersion")), gsQtyFmt)
        Me.txtItmPrint = Format(To_Value(ReadRs(rsRcd, "ItmPrint")), gsQtyFmt)
        Me.cboItmPackUOMCode = ReadRs(rsRcd, "ItmPackUOMCode")
        lblDspItmPackUOMDesc = LoadDescByCode("MstUOM", "UomCode", "UomDesc", cboItmPackUOMCode, True)
        Me.txtItmExtWidth = Format(To_Value(ReadRs(rsRcd, "ItmExtWidth")), gsAmtFmt)
        Me.txtItmExtHeight = Format(To_Value(ReadRs(rsRcd, "ItmExtHeight")), gsAmtFmt)
        Me.txtItmExtLength = Format(To_Value(ReadRs(rsRcd, "ItmExtLength")), gsAmtFmt)
        Me.txtItmPackQty = Format(To_Value(ReadRs(rsRcd, "ItmPackQty")), gsQtyFmt)
        Me.txtItmIntWidth = Format(To_Value(ReadRs(rsRcd, "ItmIntWidth")), gsAmtFmt)
        Me.txtItmIntHeight = Format(To_Value(ReadRs(rsRcd, "ItmIntHeight")), gsAmtFmt)
        Me.txtItmIntLength = Format(To_Value(ReadRs(rsRcd, "ItmIntLength")), gsAmtFmt)
        
        'Tab 3
        Me.cboItmCurr = ReadRs(rsRcd, "ItmCurr")
        Me.txtItmDefaultPrice = Format(To_Value(ReadRs(rsRcd, "ItmDefaultPrice")), gsAmtFmt)
        Me.txtItmBottomPrice = Format(To_Value(ReadRs(rsRcd, "ItmBottomPrice")), gsAmtFmt)
        Me.cboItmCDisCode = ReadRs(rsRcd, "ItmCDisCode")
        lblDspItmCDisDesc = LoadDescByCode("MstCategoryDiscount", "CDisCode", "CDisDesc", cboItmCDisCode, True)
        Me.cboItmAccTypeCode = ReadRs(rsRcd, "ItmAccTypeCode")
        lblDspItmAccTypeDesc = LoadDescByCode("MstAccountType", "AccTypeCode", "AccTypeDesc", cboItmAccTypeCode, True)
        Call Set_CheckValue(chkItmInActive, ReadRs(rsRcd, "ItmInActive"))
        Call Set_CheckValue(chkItmInvItemFlg, ReadRs(rsRcd, "ItmInvItemFlg"))
        Call Set_CheckValue(chkItmTaxFlg, ReadRs(rsRcd, "ItmTaxFlg"))
        Call Set_CheckValue(chkItmReorderFlg, ReadRs(rsRcd, "ItmReorderFlg"))
        Me.txtItmReorderQty = Format(To_Value(ReadRs(rsRcd, "ItmReorderQty")), gsQtyFmt)
        Me.txtAdj(0) = ReadRs(rsRcd, "ItmAdjJan")
        Me.txtAdj(1) = ReadRs(rsRcd, "ItmAdjFeb")
        Me.txtAdj(2) = ReadRs(rsRcd, "ItmAdjMar")
        Me.txtAdj(3) = ReadRs(rsRcd, "ItmAdjApr")
        Me.txtAdj(4) = ReadRs(rsRcd, "ItmAdjMay")
        Me.txtAdj(5) = ReadRs(rsRcd, "ItmAdjJun")
        Me.txtAdj(6) = ReadRs(rsRcd, "ItmAdjJul")
        Me.txtAdj(7) = ReadRs(rsRcd, "ItmAdjAug")
        Me.txtAdj(8) = ReadRs(rsRcd, "ItmAdjSep")
        Me.txtAdj(9) = ReadRs(rsRcd, "ItmAdjOct")
        Me.txtAdj(10) = ReadRs(rsRcd, "ItmAdjNov")
        Me.txtAdj(11) = ReadRs(rsRcd, "ItmAdjDec")
        Call Set_CheckValue(chkItmReorderInd, ReadRs(rsRcd, "ItmReorderInd"))
        Call Set_CheckValue(chkItmOwnEdition, ReadRs(rsRcd, "ItmOwnEdition"))
        Me.txtItmPORepuQty = Format(To_Value(ReadRs(rsRcd, "ItmPORepuQty")), gsQtyFmt)
        Me.txtItmPORtnDate = ReadRs(rsRcd, "ItmPORtnDate")
        
        lblDspUnitPrice = txtItmDefaultPrice
        
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
   ' Set waResult = Nothing
    Set waScrItm = Nothing
    Set waScrToolTip = Nothing
   ' Set waPgmItm = Nothing
    Set frmB001 = Nothing

End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        If txtItmAuthor.Enabled = True Then
            txtItmAuthor.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 1 Then
        If txtItmDir.Enabled = True Then
            txtItmDir.SetFocus
        End If
    
    ElseIf tabDetailInfo.Tab = 2 Then
        If txtItmPublisher.Enabled = True Then
            txtItmPublisher.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 3 Then
        If cboItmCurr.Enabled = True Then
            cboItmCurr.SetFocus
        End If
    End If
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
                If MsgBox("你是否確定要放棄現時之作業?", vbYesNo, gsTitle) = vbYes Then
                    Call cmdCancel
                End If
            Else
                Call cmdCancel
            End If
        
        Case tcFind
            
            Call OpenPromptForm
            
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
    wsFormID = "B001"
    wsTrnCd = ""
End Sub


Private Sub Ini_Caption()
Dim i As Integer
On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    fraCustomerInfo.Caption = Get_Caption(waScrItm, "FRACUSTOMERINFO")
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO0")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO1")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO2")
    tabDetailInfo.TabCaption(3) = Get_Caption(waScrItm, "TABDETAILINFO3")
    
    fraCover = Get_Caption(waScrItm, "COVER")
    fraContent = Get_Caption(waScrItm, "CONTENT")
    
    'not added to insert script
    lblItmCode.Caption = Get_Caption(waScrItm, "ITMCODE")
    lblItmBarCode.Caption = Get_Caption(waScrItm, "ITMBARCODE")
    lblItmGrpChiName.Caption = Get_Caption(waScrItm, "ITMGRPCHINAME")
    lblItmGrpEngName.Caption = Get_Caption(waScrItm, "ITMGRPENGNAME")
    lblItmChiName.Caption = Get_Caption(waScrItm, "ITMCHINAME")
    lblItmAuthor.Caption = Get_Caption(waScrItm, "ITMAUTHOR")
    lblItmTranslator.Caption = Get_Caption(waScrItm, "ITMTRANSLATOR")
    lblItmTypeCode.Caption = Get_Caption(waScrItm, "ITMTYPECODE")
    lblItmCatCode.Caption = Get_Caption(waScrItm, "ITMCATCODE")
    lblItmLevelCode.Caption = Get_Caption(waScrItm, "ITMLEVELCODE")
    lblItmLangCode.Caption = Get_Caption(waScrItm, "ITMLANGCODE")
    lblItmLastUpd.Caption = Get_Caption(waScrItm, "ITMLASTUPD")
    lblItmLastUpdDate.Caption = Get_Caption(waScrItm, "ITMLASTUPDDATE")
    
    lblItmPackTypeCode.Caption = Get_Caption(waScrItm, "ITMPACKTYPECODE")
    lblItmPackUOMCode.Caption = Get_Caption(waScrItm, "ITMPACKUOMCODE")
    lblItmPrintSizeCode.Caption = Get_Caption(waScrItm, "ITMPRINTSIZECODE")
    lblItmPackQty.Caption = Get_Caption(waScrItm, "ITMPACKQTY")
    'lblItmSize.Caption = Get_Caption(waScrItm, "ITMSIZE")
    lblItmHeight.Caption = Get_Caption(waScrItm, "ITMHEIGHT")
    lblItmWidth.Caption = Get_Caption(waScrItm, "ITMWIDTH")
    lblItmExternal.Caption = Get_Caption(waScrItm, "ITMEXTERNAL")
    lblWidth.Caption = Get_Caption(waScrItm, "WIDTH")
    lblHeight.Caption = Get_Caption(waScrItm, "HEIGHT")
    lblLength.Caption = Get_Caption(waScrItm, "LENGTH")
    lblItmInternal.Caption = Get_Caption(waScrItm, "ITMINTERNAL")
    lblItmPage.Caption = Get_Caption(waScrItm, "ITMPAGE")
    lblItmWeight.Caption = Get_Caption(waScrItm, "ITMWEIGHT")
    
    lblItmPrintDate.Caption = Get_Caption(waScrItm, "ITMPRINTDATE")
    lblItmVersion.Caption = Get_Caption(waScrItm, "ITMVERSION")
    lblItmVolume.Caption = Get_Caption(waScrItm, "ITMVOLUME")
    lblItmPrint.Caption = Get_Caption(waScrItm, "ITMPRINT")
    lblItmPublisher.Caption = Get_Caption(waScrItm, "ITMPUBLISHER")
    lblItmSeriesNo.Caption = Get_Caption(waScrItm, "ITMSERIESNO")
    chkItmInActive.Caption = Get_Caption(waScrItm, "ITMINACTIVE")
    chkItmInvItemFlg.Caption = Get_Caption(waScrItm, "ITMINVITEMFLG")
    chkItmTaxFlg.Caption = Get_Caption(waScrItm, "ITMTAXFLG")
    
    lblItmCurrCode.Caption = Get_Caption(waScrItm, "ITMCURRCODE")
    lblItmDefaultPrice.Caption = Get_Caption(waScrItm, "ITMDEFAULTPRICE")
    lblItmBottomPrice.Caption = Get_Caption(waScrItm, "ITMBOTTOMPRICE")
    lblItmCDisCode.Caption = Get_Caption(waScrItm, "ITMCDISCODE")
    lblItmReorderQty.Caption = Get_Caption(waScrItm, "ITMREORDERQTY")
    chkItmReorderFlg.Caption = Get_Caption(waScrItm, "ITMREORDERFLG")
    chkItmReorderInd.Caption = Get_Caption(waScrItm, "ITMREORDERIND")
    chkItmOwnEdition.Caption = Get_Caption(waScrItm, "ITMOWNEDITION")
    lblItmPORtnDate.Caption = Get_Caption(waScrItm, "ITMPORTNDATE")
    lblItmPORepuQty.Caption = Get_Caption(waScrItm, "ITMPOREPUQTY")
    lblItmAccTypeCode.Caption = Get_Caption(waScrItm, "ITMEDITIONCODE")
    lblUnitPrice.Caption = Get_Caption(waScrItm, "UNITPRICE")
    lblEditor.Caption = Get_Caption(waScrItm, "EDITOR")
    lblIctrnQty.Caption = Get_Caption(waScrItm, "ICTRNQTY")
    lblDraw.Caption = Get_Caption(waScrItm, "DRAW")
    lblArt.Caption = Get_Caption(waScrItm, "ART")
    lblPhoto.Caption = Get_Caption(waScrItm, "PHOTO")
    lblItmBinNo.Caption = Get_Caption(waScrItm, "BINNO")
    lblPackInfo.Caption = Get_Caption(waScrItm, "PACKINFO")
    lblItmHeight.Caption = Get_Caption(waScrItm, "ITMHEIGHT")
    lblItmWidth.Caption = Get_Caption(waScrItm, "ITMWIDTH")
    lblWeight.Caption = Get_Caption(waScrItm, "WEIGHT")
    lblPages.Caption = Get_Caption(waScrItm, "PAGES")
    
    btnPriceChange.Caption = Get_Caption(waScrItm, "PRICECHANGE")
        
    For i = 0 To 11
        lblAdj(i).Caption = Get_Caption(waScrItm, "ADJ" & i)
    Next i
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    wsActNam(1) = Get_Caption(waScrItm, "BADD")
    wsActNam(2) = Get_Caption(waScrItm, "BEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "BDELETE")
    
    
    btnItemPrice.Caption = Get_Caption(waScrItm, "ITMPRICE")
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
    wdOldPrice = 0
    lblDspUnitPrice = 0
    
    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    
    'Call SetDateMask(medItmPrintDate)
    Call SetPeriodMask(medItmPrintDate)
    
    
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
        cboItmCode.SetFocus
    
    Case DelRec
    
        Me.Caption = wsFormCaption + " - DELETE"
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboItmCode.SetFocus
    
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
            If RowLock(wsConnTime, wsKeyType, cboItmCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
                
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtItmBarCode.SetFocus
End Sub



Private Function Chk_txtItmCode() As Boolean
Dim wsStatus As String

    Chk_txtItmCode = False
    
        If Trim(txtItmCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            txtItmCode.SetFocus
            Exit Function
        End If
    
        If Chk_ItmCode(txtItmCode.Text, wsStatus) = True Then
        
        If wsStatus = "2" Then
            gsMsg = "書本已存在但已無效!"
        Else
            gsMsg = "書本已存在!"
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
    
        If Trim(cboItmCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboItmCode.SetFocus
            Exit Function
        End If
    
        If Chk_ItmCode(cboItmCode.Text, wsStatus) = False Then
            gsMsg = "會計版別不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboItmCode.SetFocus
            Exit Function
        Else
        If wsStatus = "2" Then
            gsMsg = "書本已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboItmCode.SetFocus
            Exit Function
        End If
        End If
    
    Chk_cboItmCode = True
End Function
Private Sub cmdOpen()

    Dim newForm As New frmB001
    
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
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboItmCode, wsFormID) Then
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
    
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_B001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, txtItmCode.Text, cboItmCode.Text))
    Call SetSPPara(adcmdSave, 4, txtItmBarCode.Text)
    Call SetSPPara(adcmdSave, 5, txtItmGrpChiName)
    Call SetSPPara(adcmdSave, 6, txtItmGrpEngName)
    Call SetSPPara(adcmdSave, 7, txtItmChiName)
    Call SetSPPara(adcmdSave, 8, txtItmAuthor)
    Call SetSPPara(adcmdSave, 9, txtItmTranslator)
    Call SetSPPara(adcmdSave, 10, txtEditor)
    Call SetSPPara(adcmdSave, 11, cboItmTypeCode)
    Call SetSPPara(adcmdSave, 12, cboItmLevelCode)
    Call SetSPPara(adcmdSave, 13, txtItmSeriesNo)
    Call SetSPPara(adcmdSave, 14, txtDraw)
    Call SetSPPara(adcmdSave, 15, txtArt)
    Call SetSPPara(adcmdSave, 16, txtPhoto)
    Call SetSPPara(adcmdSave, 17, cboItmCatCode)
    Call SetSPPara(adcmdSave, 18, cboItmLangCode)
    Call SetSPPara(adcmdSave, 19, txtItmVolume)
    Call SetSPPara(adcmdSave, 20, gsUserID)
    Call SetSPPara(adcmdSave, 21, wsGenDte)
    Call SetSPPara(adcmdSave, 22, txtItmDir)
    Call SetSPPara(adcmdSave, 23, txtItmTextDir)
    Call SetSPPara(adcmdSave, 24, cboItmPackTypeCode)
    Call SetSPPara(adcmdSave, 25, cboItmPrintSizeCode)
    Call SetSPPara(adcmdSave, 26, txtItmSize)
    Call SetSPPara(adcmdSave, 27, txtItmWidth)
    Call SetSPPara(adcmdSave, 28, txtItmWeight)
    Call SetSPPara(adcmdSave, 29, txtItmPage)
    Call SetSPPara(adcmdSave, 30, txtItmBinNo)
    Call SetSPPara(adcmdSave, 31, txtItmPublisher)
    Call SetSPPara(adcmdSave, 32, Set_MedDate(medItmPrintDate))
    Call SetSPPara(adcmdSave, 33, txtItmVersion)
    Call SetSPPara(adcmdSave, 34, txtItmPrint)
    Call SetSPPara(adcmdSave, 35, cboItmPackUOMCode)
    Call SetSPPara(adcmdSave, 36, txtItmExtWidth)
    Call SetSPPara(adcmdSave, 37, txtItmExtHeight)
    Call SetSPPara(adcmdSave, 38, txtItmExtLength)
    Call SetSPPara(adcmdSave, 39, txtItmPackQty)
    Call SetSPPara(adcmdSave, 40, txtItmIntWidth)
    Call SetSPPara(adcmdSave, 41, txtItmIntHeight)
    Call SetSPPara(adcmdSave, 42, txtItmIntLength)
    Call SetSPPara(adcmdSave, 43, cboItmCurr)
    Call SetSPPara(adcmdSave, 44, txtItmDefaultPrice)
    Call SetSPPara(adcmdSave, 45, txtItmBottomPrice)
    Call SetSPPara(adcmdSave, 46, cboItmCDisCode)
    Call SetSPPara(adcmdSave, 47, cboItmAccTypeCode)
    Call SetSPPara(adcmdSave, 48, Get_CheckValue(chkItmInActive))
    Call SetSPPara(adcmdSave, 49, Get_CheckValue(chkItmInvItemFlg))
    Call SetSPPara(adcmdSave, 50, Get_CheckValue(chkItmTaxFlg))
    Call SetSPPara(adcmdSave, 51, Get_CheckValue(chkItmReorderFlg))
    Call SetSPPara(adcmdSave, 52, txtItmReorderQty)
    
    For i = 0 To 11
        Call SetSPPara(adcmdSave, 53 + i, txtAdj(i))
    Next i
    
    Call SetSPPara(adcmdSave, 65, Get_CheckValue(chkItmReorderInd))
    Call SetSPPara(adcmdSave, 66, txtItmPORepuQty)
    Call SetSPPara(adcmdSave, 67, txtItmPORtnDate)
    Call SetSPPara(adcmdSave, 68, Get_CheckValue(chkItmOwnEdition))
    Call SetSPPara(adcmdSave, 69, cboItmUOMCode)
    Call SetSPPara(adcmdSave, 70, txtItmEngName)
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 71)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - B001!"
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
       If MsgBox("你是否確定不儲存現時之變更而離開?", vbYesNo, gsTitle) = vbYes Then
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
    Dim wsSql As String
    
    ReDim vFilterAry(5, 2)
    vFilterAry(1, 1) = "國際書號"
    vFilterAry(1, 2) = "ItmCode"
    
    vFilterAry(2, 1) = "條碼"
    vFilterAry(2, 2) = "ItmBarCode"
    
    vFilterAry(3, 1) = "書名"
    vFilterAry(3, 2) = "ItmChiName"
    
    vFilterAry(4, 1) = "杜威分類"
    vFilterAry(4, 2) = "CatDesc"
    
    vFilterAry(5, 1) = "作者"
    vFilterAry(5, 2) = "ItmAuthor"
    
    
    ReDim vAry(6, 3)
    vAry(1, 1) = "ISBN"
    vAry(1, 2) = "ItmCode"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "條碼"
    vAry(2, 2) = "ItmBarCode"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "書名"
    vAry(3, 2) = "ItmChiName"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "杜威分類"
    vAry(4, 2) = "CatDesc"
    vAry(4, 3) = "1500"
    
    vAry(5, 1) = "作者"
    vAry(5, 2) = "ItmAuthor"
    vAry(5, 3) = "1000"
    
    
    'frmShareSearch.Show vbModal
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
         wsSql = "SELECT MstItem.ItmCode, MstItem.ItmBarCode, "
        wsSql = wsSql + "MstItem.ItmChiName, MstCategory.CatDesc, MstItem.ItmAuthor "
        wsSql = wsSql + "FROM MstItem, MstCategory "
        .sBindSQL = wsSql
        .sBindWhereSQL = "WHERE MstItem.ItmStatus = '1' AND MstItem.ItmCatCode = MstCategory.CatCode "
        .sBindOrderSQL = "ORDER BY MstItem.ItmCode"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboItmCode Then
        cboItmCode = Trim(frmShareSearch.Tag)
       If cboItmCode.Enabled = False Then
        LoadRecord
        txtItmBarCode.Text = ""
        txtItmCode.SetFocus
       Else
        cboItmCode.SetFocus
        SendKeys "{Enter}"
       End If
    End If

    
End Sub







Private Sub txtAdj_GotFocus(Index As Integer)
    FocusMe txtAdj(Index)
End Sub

Private Sub txtAdj_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtAdj(Index), False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        
        If Index < 11 Then
            txtAdj(Index + 1).SetFocus
        Else
            chkItmReorderInd.SetFocus
        End If
    End If
End Sub

Private Sub txtAdj_LostFocus(Index As Integer)
    txtAdj(Index) = Format(txtAdj(Index), gsQtyFmt)
    FocusMe txtAdj(Index), True
End Sub

Private Sub txtArt_GotFocus()
    FocusMe txtArt
End Sub

Private Sub txtArt_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtArt, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtPhoto.SetFocus
    End If
End Sub

Private Sub txtArt_LostFocus()
    FocusMe txtArt, True
End Sub

Private Sub txtDraw_GotFocus()
    FocusMe txtDraw
End Sub

Private Sub txtDraw_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtDraw, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtArt.SetFocus
    End If
End Sub

Private Sub txtDraw_LostFocus()
    FocusMe txtDraw, True
End Sub

Private Sub txtEditor_GotFocus()
    FocusMe txtEditor
End Sub

Private Sub txtEditor_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtEditor, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        cboItmTypeCode.SetFocus
    End If
End Sub

Private Sub txtEditor_LostFocus()
    FocusMe txtEditor, True
End Sub

Private Sub txtItmAuthor_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmAuthor, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtItmAuthor() = False Then
            Exit Sub
        End If
        
        'txtItmCallNo = getCallNo(txtItmAuthor)
        
        tabDetailInfo.Tab = 0
        txtItmTranslator.SetFocus
    End If
End Sub

Private Sub txtItmAuthor_LostFocus()
    FocusMe txtItmAuthor, True
End Sub



Private Sub txtItmBarCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmBarCode, 13, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtItmGrpChiName.SetFocus
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
        
        tabDetailInfo.Tab = 2
        txtItmPublisher.SetFocus
    End If
End Sub

Private Sub txtItmBinNo_LostFocus()
    FocusMe txtItmBinNo, True
End Sub

Private Sub txtItmBottomPrice_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmBottomPrice, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 3
        cboItmCDisCode.SetFocus
        
        'If btnItemPrice.Enabled = True Then
        '    btnItemPrice.SetFocus
        'Else
        '    tabDetailInfo.Tab = 0
        '    txtItmAuthor.SetFocus
        'End If
    End If
End Sub

Private Sub txtItmBottomPrice_LostFocus()
    txtItmBottomPrice = Format(txtItmBottomPrice, gsAmtFmt)
    FocusMe txtItmBottomPrice, True
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
    Call chk_InpLen(txtItmCode, 13, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtItmCode() = True Then
            Call Ini_Scr_AfrKey
        End If
        
    End If
End Sub

Private Sub txtItmCode_LostFocus()
    FocusMe txtItmCode, True
End Sub

Private Sub txtItmDefaultPrice_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmDefaultPrice, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        txtItmBottomPrice.SetFocus
    End If
End Sub

Private Sub txtItmDefaultPrice_LostFocus()
    txtItmDefaultPrice = Format(txtItmDefaultPrice, gsAmtFmt)
    FocusMe txtItmDefaultPrice, True
End Sub

Private Sub txtItmDir_GotFocus()
    FocusMe txtItmDir
End Sub

Private Sub txtItmDir_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmDir, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        If Trim(txtItmDir) = "" Then
            Clear_Cover
            btnItmDir.SetFocus
        Else
            If Chk_Load_Cover(txtItmDir) Then
                btnItmDir.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtItmDir_LostFocus()
    FocusMe txtItmDir, True
End Sub

Private Sub txtItmEngName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmEngName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtItmAuthor.SetFocus
    End If
End Sub

Private Sub txtItmEngName_LostFocus()
    FocusMe txtItmEngName, True
End Sub

Private Sub txtItmExtHeight_GotFocus()
    FocusMe txtItmExtHeight
End Sub

Private Sub txtItmExtHeight_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmExtHeight, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        txtItmExtLength.SetFocus
    End If
End Sub

Private Sub txtItmExtHeight_LostFocus()
    txtItmExtHeight = Format(txtItmExtHeight, gsAmtFmt)
    FocusMe txtItmExtHeight, True
End Sub

Private Sub txtItmExtLength_GotFocus()
    FocusMe txtItmExtLength
End Sub

Private Sub txtItmExtLength_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmExtLength, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtItmPackQty.SetFocus
    End If
End Sub

Private Sub txtItmExtLength_LostFocus()
    txtItmExtLength = Format(txtItmExtLength, gsAmtFmt)
    FocusMe txtItmExtLength, True
End Sub

Private Sub txtItmExtWidth_GotFocus()
    FocusMe txtItmExtWidth
End Sub

Private Sub txtItmExtWidth_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmExtWidth, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        txtItmExtHeight.SetFocus
    End If
End Sub

Private Sub txtItmExtWidth_LostFocus()
    txtItmExtWidth = Format(txtItmExtWidth, gsAmtFmt)
    FocusMe txtItmExtWidth, True
End Sub

Private Sub txtItmGrpChiName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmGrpChiName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
            
        txtItmGrpEngName.SetFocus
    End If
End Sub

Private Sub txtItmGrpChiName_LostFocus()
    FocusMe txtItmGrpChiName, True
End Sub

Private Sub txtItmGrpEngName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmGrpEngName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtItmChiName.SetFocus
    End If
End Sub

Private Sub txtItmGrpEngName_LostFocus()
    FocusMe txtItmGrpEngName, True
End Sub

Private Sub txtItmIntHeight_GotFocus()
    FocusMe txtItmIntHeight
End Sub

Private Sub txtItmIntHeight_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmIntHeight, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        txtItmIntLength.SetFocus
    End If
End Sub

Private Sub txtItmIntHeight_LostFocus()
    txtItmIntHeight = Format(txtItmIntHeight, gsAmtFmt)
    FocusMe txtItmIntHeight, True
End Sub

Private Sub txtItmIntLength_GotFocus()
    FocusMe txtItmIntLength
End Sub

Private Sub txtItmIntLength_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmIntLength, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 3
        cboItmCurr.SetFocus
    End If
End Sub

Private Sub txtItmIntLength_LostFocus()
    txtItmIntLength = Format(txtItmIntLength, gsAmtFmt)
    FocusMe txtItmIntLength, True
End Sub

Private Sub txtItmIntWidth_GotFocus()
    txtItmIntWidth = Format(txtItmIntWidth, gsAmtFmt)
    FocusMe txtItmIntWidth
End Sub

Private Sub txtItmIntWidth_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmIntWidth, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        txtItmIntHeight.SetFocus
    End If
End Sub

Private Sub txtItmIntWidth_LostFocus()
    txtItmIntWidth = Format(txtItmIntWidth, gsAmtFmt)
    FocusMe txtItmIntWidth, True
End Sub

Private Sub txtItmPackQty_GotFocus()
    FocusMe txtItmPackQty
End Sub

Private Sub txtItmPackQty_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmPackQty, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtItmIntWidth.SetFocus
    End If
End Sub

Private Sub txtItmPackQty_LostFocus()
    txtItmPackQty = Format(txtItmPackQty, gsQtyFmt)
    FocusMe txtItmPackQty, True
    
End Sub

Private Sub txtItmPage_GotFocus()
    FocusMe txtItmPage
End Sub

Private Sub txtItmPage_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmPage, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        txtItmBinNo.SetFocus
    End If
End Sub

Private Sub txtItmPage_LostFocus()
    txtItmPage = Format(txtItmPage, gsQtyFmt)
    FocusMe txtItmPage, True
End Sub

Private Sub txtItmPORepuQty_GotFocus()
    If tabDetailInfo.Tab <> 3 Then tabDetailInfo.Tab = 3
    FocusMe txtItmPORepuQty
End Sub

Private Sub txtItmPORepuQty_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmPORepuQty, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        txtItmPORtnDate.SetFocus
    End If
End Sub

Private Sub txtItmPORepuQty_LostFocus()
    txtItmPORepuQty = Format(txtItmPORepuQty, gsQtyFmt)
    FocusMe txtItmPORepuQty, True
End Sub

Private Sub txtItmPORtnDate_GotFocus()
    FocusMe txtItmPORtnDate
End Sub

Private Sub txtItmPORtnDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 3
        txtItmBarCode.SetFocus
    End If
End Sub

Private Sub txtItmPORtnDate_LostFocus()
    FocusMe txtItmPORtnDate, True
End Sub

Private Sub txtItmPrint_GotFocus()
    FocusMe txtItmPrint
End Sub

Private Sub txtItmPrint_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmPrint, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        cboItmPackUOMCode.SetFocus
    End If
End Sub

Private Sub txtItmPrint_LostFocus()
    txtItmPrint = Format(txtItmPrint, gsQtyFmt)
    FocusMe txtItmPrint, True
End Sub

Private Sub medItmPrintDate_GotFocus()
    FocusMe medItmPrintDate
End Sub

Private Sub medItmPrintDate_KeyPress(KeyAscii As Integer)
   Call Chk_InpNum(KeyAscii, medItmPrintDate, False, False)
    
   If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_medItmPrintDate = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 2
        txtItmVersion.SetFocus
    End If
End Sub

Private Sub medItmPrintDate_LostFocus()
    FocusMe medItmPrintDate, True
End Sub

Private Sub txtItmPublisher_GotFocus()
    FocusMe txtItmPublisher
End Sub

Private Sub txtItmPublisher_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmPublisher, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        medItmPrintDate.SetFocus
    End If
End Sub

Private Sub txtItmPublisher_LostFocus()
    FocusMe txtItmPublisher, True
End Sub

Private Sub txtItmReorderQty_GotFocus()
    FocusMe txtItmReorderQty
End Sub

Private Sub txtItmReorderQty_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmReorderQty, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 3
        txtAdj(0).SetFocus
    End If
End Sub

Private Sub txtItmReorderQty_LostFocus()
    txtItmReorderQty = Format(txtItmReorderQty, gsQtyFmt)
    FocusMe txtItmReorderQty, True
End Sub

Private Sub txtItmSeriesNo_GotFocus()
    FocusMe txtItmSeriesNo
End Sub

Private Sub txtItmSeriesNo_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmSeriesNo, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 0
        cboItmUOMCode.SetFocus
    End If
End Sub

Private Sub txtItmSeriesNo_LostFocus()
    FocusMe txtItmSeriesNo, True
End Sub

Private Sub txtItmSize_GotFocus()
    txtItmSize = Format(txtItmSize, gsQtyFmt)
    FocusMe txtItmSize
End Sub

Private Sub txtItmSize_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmSize, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 1
        txtItmWidth.SetFocus
    End If
End Sub

Private Sub txtItmSize_LostFocus()
    FocusMe txtItmSize, True
End Sub

Private Sub txtItmTextDir_GotFocus()
    FocusMe txtItmTextDir
End Sub

Private Sub txtItmTextDir_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmTextDir, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        
        If Trim(txtItmTextDir) = "" Then
            Clear_Content
            btnItmTextDir.SetFocus
        Else
            If Chk_Load_Content(txtItmTextDir) Then
                btnItmTextDir.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtItmTextDir_LostFocus()
    FocusMe txtItmTextDir, True
End Sub

Private Sub txtItmTranslator_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtItmTranslator, 40, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtEditor.SetFocus
    End If
End Sub

Private Sub txtItmAuthor_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtItmAuthor
End Sub

Private Sub txtItmBarCode_GotFocus()
    FocusMe txtItmBarCode
End Sub

Private Sub txtItmBottomPrice_GotFocus()
    FocusMe txtItmBottomPrice
End Sub

Private Sub txtItmChiName_GotFocus()
    FocusMe txtItmChiName
End Sub

Private Sub txtItmCode_GotFocus()
    FocusMe txtItmCode
End Sub

Private Sub txtItmDefaultPrice_GotFocus()
    FocusMe txtItmDefaultPrice
End Sub

Private Sub txtItmEngName_GotFocus()
    FocusMe txtItmEngName
End Sub

Private Sub txtItmGrpChiName_GotFocus()
    FocusMe txtItmGrpChiName
End Sub

Private Sub txtItmGrpEngName_GotFocus()
    FocusMe txtItmGrpEngName
End Sub

Private Sub txtItmTranslator_GotFocus()
    FocusMe txtItmTranslator
End Sub

Private Sub txtItmTranslator_LostFocus()
    FocusMe txtItmTranslator, True
End Sub

Private Sub txtItmVersion_GotFocus()
    FocusMe txtItmVersion
End Sub

Private Sub txtItmVersion_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmVersion, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtItmPrint.SetFocus
    End If
End Sub

Private Sub txtItmVersion_LostFocus()
    txtItmVersion = Format(txtItmVersion, gsQtyFmt)
    FocusMe txtItmVersion, True
End Sub

Private Sub txtItmVolume_GotFocus()
    FocusMe txtItmVolume
End Sub

Private Sub txtItmVolume_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmVolume, False, False)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 1
        txtItmDir.SetFocus
    End If
End Sub

Private Sub txtItmVolume_LostFocus()
    FocusMe txtItmVolume, True
End Sub

Private Sub txtItmWeight_GotFocus()
    FocusMe txtItmWeight
End Sub

Private Sub txtItmWeight_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmWeight, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        txtItmPage.SetFocus
    End If
End Sub

Private Sub txtItmWeight_LostFocus()
    txtItmWeight = Format(txtItmWeight, gsAmtFmt)
    FocusMe txtItmWeight, True
End Sub

Private Sub txtItmWidth_GotFocus()
    FocusMe txtItmWidth
End Sub

Private Sub txtItmWidth_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtItmWidth, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 1
        txtItmWeight.SetFocus
    End If
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
    
    If Trim(cboItmCurr.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        Me.tabDetailInfo.Tab = 3
        cboItmCurr.SetFocus
        Exit Function
    End If
    
    If Chk_ItmCurr() = False Then
            gsMsg = "貨幣不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 3
            cboItmCurr.SetFocus
            Exit Function
    End If

    
    Chk_cboItmCurr = True
End Function


Private Function Chk_cboItmTypeCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

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
            gsMsg = "圖書分類編碼不存在!"
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
    Dim wsSql As String
        
    Chk_ItmTypeCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT ItmTypeChiDesc "
    wsSql = wsSql & " FROM MstItemType WHERE MstItemType.ItmTypeCode = '" & Set_Quote(inCode) & "'  And ItmTypeStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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

Private Function Chk_cboItmCatCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmCatCode = False
    
        If Trim(cboItmCatCode.Text) = "" Then
            gsMsg = "沒有輸入須要之資料!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboItmCatCode.SetFocus
            Exit Function
        End If

        
        If Chk_ItmCatCode(cboItmCatCode.Text, wsRetName) = False Then
            lblDspItmCatDesc = ""
            gsMsg = "杜威分類編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboItmCatCode.SetFocus
            Exit Function
        Else
            lblDspItmCatDesc = wsRetName
        End If
            
    Chk_cboItmCatCode = True
End Function

Private Function Chk_ItmCatCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmCatCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT CatDesc "
    wsSql = wsSql & " FROM MstCategory WHERE MstCategory.CatCode= '" & Set_Quote(inCode) & "' And CatStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "CatDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmCatCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboItmLevelCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmLevelCode = False
    
    If Trim(cboItmLevelCode.Text) = "" Then
        lblDspItmLevelDesc = ""
         gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboItmLevelCode.SetFocus
        Exit Function
    End If
    
    If Chk_ItmLevelCode(cboItmLevelCode.Text, wsRetName) = False Then
            gsMsg = "程度編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboItmLevelCode.SetFocus
            Exit Function
    Else
            lblDspItmLevelDesc = wsRetName
    End If

    
    Chk_cboItmLevelCode = True
End Function

Private Function Chk_ItmLevelCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmLevelCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT LevelDesc "
    wsSql = wsSql & " FROM MstLevel WHERE MstLevel.LevelCode = '" & Set_Quote(inCode) & "' And LevelStatus = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "LevelDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmLevelCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboItmLangCode() As Boolean

    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmLangCode = False
    
    If Trim(cboItmLangCode.Text) = "" Then
        lblDspItmLangDesc = ""
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboItmLangCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_ItmLangCode(cboItmLangCode.Text, wsRetName) = False Then
            gsMsg = "語系編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboItmLangCode.SetFocus
            Exit Function
    Else
            lblDspItmLangDesc = wsRetName
    End If

    
    Chk_cboItmLangCode = True
End Function

Private Function Chk_ItmLangCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmLangCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT LangDesc "
    wsSql = wsSql & " FROM MstLanguage WHERE MstLanguage.LangCode = '" & Set_Quote(inCode) & "'  And LangStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "LangDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmLangCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboItmPackTypeCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmPackTypeCode = False
    
    If Trim(cboItmPackTypeCode.Text) = "" Then
        lblDspItmPackTypeDesc = ""
        Chk_cboItmPackTypeCode = True
        Exit Function
    End If
    
    If Chk_ItmPackTypeCode(cboItmPackTypeCode.Text, wsRetName) = False Then
            lblDspItmPackTypeDesc = ""
            gsMsg = "裝幀編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 2
            cboItmPackTypeCode.SetFocus
            Exit Function
    Else
            lblDspItmPackTypeDesc = wsRetName
    End If

    
    Chk_cboItmPackTypeCode = True
End Function

Private Function Chk_ItmPackTypeCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmPackTypeCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT PackTypeDesc "
    wsSql = wsSql & " FROM MstPackingType WHERE MstPackingType.PackTypeCode = '" & Set_Quote(inCode) & "' And PackTypeStatus = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "PackTypeDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
        
    End If
    
    Chk_ItmPackTypeCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_medItmPrintDate() As Boolean

    
    Chk_medItmPrintDate = False
    
    If Trim(medItmPrintDate.Text) = "/" Then
        Chk_medItmPrintDate = True
        Exit Function
    End If
    
   ' If Chk_Period(medItmPrintDate) = False Then
   '     gsMsg = "日子不正確!"
   '     MsgBox gsMsg, vbOKOnly, gsTitle
   '     medItmPrintDate.SetFocus
   '     Exit Function
   ' End If
    
    
    Chk_medItmPrintDate = True

End Function

Private Function Chk_ItmUOMCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmUOMCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT UomDesc "
    wsSql = wsSql & " FROM MstUOM WHERE MstUOM.UomCode = '" & Set_Quote(inCode) & "' And UomStatus = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
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

Private Function Chk_cboItmPackUOMCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmPackUOMCode = False
    
    If Trim(cboItmPackUOMCode.Text) = "" Then
        lblDspItmPackUOMDesc = ""
        Chk_cboItmPackUOMCode = True
        Exit Function
    End If
    
        If Chk_ItmPackUOMCode(cboItmPackUOMCode.Text, wsRetName) = False Then
            gsMsg = "包裝類別編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 2
            cboItmPackUOMCode.SetFocus
            Exit Function
        Else
            lblDspItmPackUOMDesc = wsRetName
        End If

    
    Chk_cboItmPackUOMCode = True
End Function

Private Function Chk_cboItmUOMCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmUOMCode = False
    
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
        lblDspItmPackUOMDesc = wsRetName
    End If

    
    Chk_cboItmUOMCode = True
End Function

Private Function Chk_ItmPackUOMCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmPackUOMCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT UomDesc "
    wsSql = wsSql & " FROM MstUOM WHERE MstUOM.UomCode = '" & Set_Quote(inCode) & "'  And UOMStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "UomDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmPackUOMCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboItmPrintSizeCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmPrintSizeCode = False
    
    If Trim(cboItmPrintSizeCode.Text) = "" Then
        lblDspItmPrintSizeDesc = ""
        Chk_cboItmPrintSizeCode = True
        Exit Function
    End If
    
    
    If Chk_ItmPrintSizeCode(cboItmPrintSizeCode.Text, wsRetName) = False Then
            gsMsg = "開度編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 2
            cboItmPrintSizeCode.SetFocus
            Exit Function
    Else
            lblDspItmPrintSizeDesc = wsRetName
    End If

    
    Chk_cboItmPrintSizeCode = True
End Function

Private Function Chk_ItmPrintSizeCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmPrintSizeCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT PrintSizeDesc "
    wsSql = wsSql & " FROM MstPrintSize WHERE MstPrintSize.PrintSizeCode = '" & Set_Quote(inCode) & "' And PrintSizeStatus = '1' "
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "PrintSizeDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmPrintSizeCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboItmCDisCode() As Boolean
    Dim wsRetName As String

    wsRetName = ""

    Chk_cboItmCDisCode = False
    
    If Trim(cboItmCDisCode.Text) = "" Then
        lblDspItmCDisDesc = ""
        Chk_cboItmCDisCode = True
        Exit Function
    End If
    
    If Chk_ItmCDisCode(cboItmCDisCode.Text, wsRetName) = False Then
            gsMsg = "折扣類別編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            tabDetailInfo.Tab = 1
            cboItmCDisCode.SetFocus
            Exit Function
    Else
            lblDspItmCDisDesc = wsRetName
    End If

    
    Chk_cboItmCDisCode = True
End Function

Private Function Chk_ItmCDisCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
        
    Chk_ItmCDisCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSql = "SELECT CDisDesc "
    wsSql = wsSql & " FROM MstCategoryDiscount WHERE MstCategoryDiscount.CDisCode = '" & Set_Quote(inCode) & "'  And CDisStatus = '1'"
    
    rsRcd.Open wsSql, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "CDisDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_ItmCDisCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
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

Private Function Chk_txtItmAuthor() As Boolean
     
    Chk_txtItmAuthor = False
    
    If Trim(txtItmAuthor.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtItmAuthor.SetFocus
        Exit Function
    End If
    
    Chk_txtItmAuthor = True
End Function


Private Sub tblCommon_DblClick()
    wcCombo.Text = tblCommon.Columns(0).Text
    wcCombo.SetFocus
    tblCommon.Visible = False
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

Private Sub cboItmCode_KeyPress(KeyAscii As Integer)

    
    Call chk_InpLen(cboItmCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmCode() = False Then
            cboItmCode.SetFocus
        Else
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboItmCode_DropDown()
    
    Dim wsSql As String


    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmCode
    
    wsSql = "SELECT ItmCode, ItmBarCode, ItmChiName FROM MstItem WHERE ItmStatus = '1'"
    wsSql = wsSql & " AND ItmCode LIKE '%" & IIf(cboItmCode.SelLength > 0, "", Set_Quote(cboItmCode.Text)) & "%' "
    wsSql = wsSql & "ORDER BY ItmCode "
    Call Ini_Combo(3, wsSql, cboItmCode.Left, cboItmCode.Top + cboItmCode.Height, tblCommon, "B001", "TBLB", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmCode_GotFocus()

    FocusMe cboItmCode
End Sub

Private Sub cboItmCode_LostFocus()
    FocusMe cboItmCode, True
End Sub

Private Sub cboItmTypeCode_DropDown()
    
    Dim wsSql As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmTypeCode
    
    wsSql = "SELECT ItmTypeCode, ItmTypeChiDesc FROM MstItemType WHERE ItmTypeStatus = '1'"
    wsSql = wsSql & "ORDER BY ItmTypeCode "
    Call Ini_Combo(2, wsSql, cboItmTypeCode.Left + tabDetailInfo.Left, cboItmTypeCode.Top + cboItmTypeCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLIT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Sub cboItmTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmTypeCode() = False Then
            Exit Sub
        End If
            
        tabDetailInfo.Tab = 0
        cboItmLevelCode.SetFocus
    End If
End Sub

Private Sub cboItmTypeCode_GotFocus()

    FocusMe cboItmTypeCode
End Sub

Private Sub cboItmTypeCode_LostFocus()
    FocusMe cboItmTypeCode, True
End Sub

Private Sub cboItmCatCode_DropDown()
    
    Dim wsSql As String
 
    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmCatCode
    
    wsSql = "SELECT CatCode, CatDesc FROM MstCategory WHERE CatStatus = '1'"
    wsSql = wsSql & "ORDER BY CatCode "
    Call Ini_Combo(2, wsSql, cboItmCatCode.Left + tabDetailInfo.Left, cboItmCatCode.Top + cboItmCatCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLCAT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmCatCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCatCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
         
        If Chk_cboItmCatCode() = False Then
           Exit Sub
        End If
            
        tabDetailInfo.Tab = 0
        cboItmLangCode.SetFocus
    End If
End Sub

Private Sub cboItmCatCode_GotFocus()
    FocusMe cboItmCatCode
End Sub

Private Sub cboItmCatCode_LostFocus()
    FocusMe cboItmCatCode, True
End Sub

Private Sub cboItmLevelCode_DropDown()
    
    Dim wsSql As String
   
    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmLevelCode
    
    wsSql = "SELECT LevelCode, LevelDesc FROM MstLevel WHERE LevelStatus = '1'"
    wsSql = wsSql & "ORDER BY LevelCode "
    Call Ini_Combo(2, wsSql, cboItmLevelCode.Left + tabDetailInfo.Left, cboItmLevelCode.Top + cboItmLevelCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLLVL", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmLevelCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmLevelCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmLevelCode() = False Then
            Exit Sub
        End If
            
        tabDetailInfo.Tab = 0
        txtItmSeriesNo.SetFocus
    End If
End Sub

Private Sub cboItmLevelCode_GotFocus()

    FocusMe cboItmLevelCode
End Sub

Private Sub cboItmLevelCode_LostFocus()
    FocusMe cboItmLevelCode, True
End Sub




Private Sub cboItmLangCode_DropDown()
    
    Dim wsSql As String
  
    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmLangCode
    
    wsSql = "SELECT LangCode, LangDesc FROM MstLanguage WHERE LangStatus = '1'"
    wsSql = wsSql & "ORDER BY LangCode "
    Call Ini_Combo(2, wsSql, cboItmLangCode.Left + tabDetailInfo.Left, cboItmLangCode.Top + cboItmLangCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLL", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmLangCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmLangCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmLangCode() = False Then
            Exit Sub
        End If
        
        Me.tabDetailInfo.Tab = 0
        txtItmVolume.SetFocus
    End If
End Sub

Private Sub cboItmLangCode_GotFocus()
    FocusMe cboItmLangCode
End Sub

Private Sub cboItmLangCode_LostFocus()
    FocusMe cboItmLangCode, True
End Sub

Private Sub cboItmPackTypeCode_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmPackTypeCode
    
    wsSql = "SELECT PackTypeCode, PackTypeDesc FROM MstPackingType WHERE PackTypeStatus = '1'"
    wsSql = wsSql & "ORDER BY PackTypeCode "
    Call Ini_Combo(2, wsSql, cboItmPackTypeCode.Left + tabDetailInfo.Left, cboItmPackTypeCode.Top + cboItmPackTypeCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLPT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmPackTypeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmPackTypeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmPackTypeCode() = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 1
        cboItmPrintSizeCode.SetFocus
    End If
End Sub

Private Sub cboItmPackTypeCode_GotFocus()
    FocusMe cboItmPackTypeCode
End Sub

Private Sub cboItmPackTypeCode_LostFocus()
    FocusMe cboItmPackTypeCode, True
End Sub

Private Sub cboItmPackUOMCode_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmPackUOMCode
    
    wsSql = "SELECT UOMCode, UOMDesc FROM MstUOM WHERE UOMStatus = '1'"
    wsSql = wsSql & "ORDER BY UOMCode "
    Call Ini_Combo(2, wsSql, cboItmPackUOMCode.Left + tabDetailInfo.Left, cboItmPackUOMCode.Top + cboItmPackUOMCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLUOM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmPackUOMCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmPackUOMCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmPackUOMCode() = False Then
            Exit Sub
        End If
           
        tabDetailInfo.Tab = 2
        txtItmExtWidth.SetFocus
    End If
End Sub

Private Sub cboItmPackUOMCode_GotFocus()
    FocusMe cboItmPackUOMCode
End Sub

Private Sub cboItmPackUOMCode_LostFocus()
    FocusMe cboItmPackUOMCode, True
End Sub

Private Sub cboItmPrintSizeCode_DropDown()
    
    Dim wsSql As String
 
    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmPrintSizeCode
    
    wsSql = "SELECT PrintSizeCode, PrintSizeDesc FROM MstPrintSize WHERE PrintSizeStatus = '1'"
    wsSql = wsSql & "ORDER BY PrintSizeCode "
    Call Ini_Combo(2, wsSql, cboItmPrintSizeCode.Left + tabDetailInfo.Left, cboItmPrintSizeCode.Top + cboItmPrintSizeCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLPS", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmPrintSizeCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmPrintSizeCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmPrintSizeCode() = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 1
        txtItmSize.SetFocus
        
    End If
End Sub

Private Sub cboItmPrintSizeCode_GotFocus()
    FocusMe cboItmPrintSizeCode
End Sub

Private Sub cboItmPrintSizeCode_LostFocus()
    FocusMe cboItmPrintSizeCode, True
End Sub

Private Sub cboItmCDisCode_DropDown()
    
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmCDisCode
    
    wsSql = "SELECT CDisCode, CDisDesc FROM MstCategoryDiscount WHERE CDisStatus = '1'"
    wsSql = wsSql & "ORDER BY CDisCode "
    Call Ini_Combo(2, wsSql, cboItmCDisCode.Left + tabDetailInfo.Left, cboItmCDisCode.Top + cboItmCDisCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLCD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboItmCDisCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboItmCDisCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboItmCDisCode() = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 3
        cboItmAccTypeCode.SetFocus
    End If
End Sub

Private Sub cboItmCDisCode_GotFocus()
    FocusMe cboItmCDisCode
End Sub

Private Sub cboItmCDisCode_LostFocus()
    FocusMe cboItmCDisCode, True
End Sub

Private Sub cboItmCurr_DropDown()
    
    Dim wsSql As String


    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmCurr
    
    wsSql = "SELECT DISTINCT ExcCurr FROM MstExchangeRate WHERE ExcStatus = '1'"
    wsSql = wsSql & "ORDER BY ExcCurr "
    Call Ini_Combo(1, wsSql, cboItmCurr.Left + tabDetailInfo.Left, cboItmCurr.Top + cboItmCurr.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLCURR", Me.Width, Me.Height)
    
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
            
        tabDetailInfo.Tab = 3
        txtItmDefaultPrice.SetFocus
    End If
End Sub

Private Sub cboItmCurr_GotFocus()
    FocusMe cboItmCurr
End Sub

Private Sub cboItmCurr_LostFocus()
    FocusMe cboItmCurr, True
End Sub

Private Sub txtItmWidth_LostFocus()
    txtItmWidth = Format(txtItmWidth, gsAmtFmt)
    FocusMe txtItmWidth, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSql As String
    
    wsSql = "SELECT ItmStatus FROM MstItem WHERE ItmCode = '" & Set_Quote(txtItmCode) & "'"
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
    Dim wsSql As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboItmAccTypeCode
    
    wsSql = "SELECT AccTypeCode, AccTypeDesc FROM MstAccountType WHERE AccTypeStatus = '1'"
    wsSql = wsSql & "ORDER BY AccTypeCode "
    Call Ini_Combo(2, wsSql, cboItmAccTypeCode.Left + tabDetailInfo.Left, cboItmAccTypeCode.Top + cboItmAccTypeCode.Height + tabDetailInfo.Top, tblCommon, "B001", "TBLACCTYPE", Me.Width, Me.Height)
    
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
            
        tabDetailInfo.Tab = 3
        If btnItemPrice.Enabled = True Then
            btnItemPrice.SetFocus
        Else
            chkItmOwnEdition.SetFocus
        End If
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

    If Trim(cboItmAccTypeCode.Text) = "" Then
        Chk_cboItmAccTypeCode = True
        Exit Function
    End If
    
    If Chk_ItmAccTypeCode(wsDesc) = False Then
        gsMsg = "會計分類不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 1
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

Private Sub txtPhoto_GotFocus()
    FocusMe txtPhoto
End Sub

Private Sub txtPhoto_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtPhoto, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboItmCatCode.SetFocus
    End If
End Sub

Private Sub txtPhoto_LostFocus()
    FocusMe txtPhoto, True
End Sub

Private Function Chk_Load_Cover(inPath As String) As Boolean
    Chk_Load_Cover = False
    
    If Load_Cover(inPath) = False Then
        gsMsg = "封面圖象不存在或錯誤!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 1
        txtItmDir.SetFocus
        Exit Function
    End If
    
    Chk_Load_Cover = True
End Function

Private Function Load_Cover(inPath As String) As Boolean
On Error GoTo Load_Cover_Err

    Load_Cover = False
    imgCover.Picture = LoadPicture(inPath)
    Load_Cover = True
    Exit Function
    
Load_Cover_Err:
    Load_Cover = False
End Function

Private Function Load_Content(inPath As String) As Boolean
On Error GoTo Load_Content_Err

    Load_Content = False
    rtContent.LoadFile inPath, rtfText
    Load_Content = True
    Exit Function
    
Load_Content_Err:
    Load_Content = False
End Function

Private Function Chk_Load_Content(inFilePath As String) As Boolean
    Chk_Load_Content = False
    
    If Load_Content(inFilePath) = False Then
        gsMsg = "內容簡介文字不存在或錯誤!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        tabDetailInfo.Tab = 1
        txtItmTextDir.SetFocus
        Exit Function
    End If
    
    Chk_Load_Content = True
End Function

Private Function Clear_Cover() As Boolean
On Error GoTo Clear_Cover_Err

    Clear_Cover = False
    imgCover.Picture = LoadPicture()
    Clear_Cover = True
    Exit Function
    
Clear_Cover_Err:
    Clear_Cover = False
End Function

Private Function Clear_Content() As Boolean
On Error GoTo Clear_Content_Err

    Clear_Content = False
    rtContent.Text = ""
    Clear_Content = True
    Exit Function
    
Clear_Content_Err:
    Clear_Content = False
End Function

