VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmC001 
   BackColor       =   &H8000000A&
   Caption         =   "客戶資料"
   ClientHeight    =   6075
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   9945
   Icon            =   "C001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9945
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11760
      OleObjectBlob   =   "C001.frx":08CA
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboCusCode 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   2385
   End
   Begin VB.TextBox txtSaleID 
      Height          =   270
      Left            =   9120
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList iglProcess 
      Left            =   0
      Top             =   360
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
            Picture         =   "C001.frx":2FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":38A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":4181
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":45D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":4A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":4D3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":5191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":55E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":58FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":5C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":6069
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "C001.frx":6945
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
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
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   3375
      Left            =   120
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   2520
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5953
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "附加通訊資料"
      TabPicture(0)   =   "C001.frx":6C6D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblCusAddress1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDspCusRgnDesc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCusRgnCode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCusAddress1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCusAddress2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCusAddress3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCusAddress4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboCusRgnCode"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "貨運資料"
      TabPicture(1)   =   "C001.frx":6C89
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCusShipAddr2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraCusShipAddr1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboCusShipTerrCode2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboCusShipTerrCode"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "其他資料"
      TabPicture(2)   =   "C001.frx":6CA5
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblCusSpecDis"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblCusCurr"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblCusPayCode"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblCusCreditLimit"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblCusSaleName"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblCusMLCode"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblDspCusMLDesc"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblCusRemark"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblDspCusLastUpdDate"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblDspCusLastUpd"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "lblCusLastUpdDate"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "lblCusLastUpd"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "lblDspCusSaleName"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cboCusPayCode"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "cboCusCurr"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cboCusSaleCode"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtCusSpecDis"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtCusPayDesc"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtCusCreditLimit"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "cboCusMLCode"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtCusRemark"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "備註"
      TabPicture(3)   =   "C001.frx":6CC1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tblDetail"
      Tab(3).Control(1)=   "lblOpenBal"
      Tab(3).Control(2)=   "lblDspOpenBal"
      Tab(3).Control(3)=   "lblDspARBal"
      Tab(3).Control(4)=   "lblARBal"
      Tab(3).Control(5)=   "lblDspCloseBal"
      Tab(3).Control(6)=   "lblCloseBal"
      Tab(3).Control(7)=   "lblAcmMnSale"
      Tab(3).Control(8)=   "lblDspAcmMnSaleNet"
      Tab(3).Control(9)=   "lblDspAcmMnSaleAmt"
      Tab(3).Control(10)=   "lblDspAcmMnSaleQty"
      Tab(3).Control(11)=   "lblAcmYrSale"
      Tab(3).Control(12)=   "lblDspAcmYrSaleNet"
      Tab(3).Control(13)=   "lblDspAcmYrSaleAmt"
      Tab(3).Control(14)=   "lblDspAcmYrSaleQty"
      Tab(3).Control(15)=   "lblAcmSale"
      Tab(3).Control(16)=   "lblDspAcmSaleNet"
      Tab(3).Control(17)=   "lblDspAcmSaleAmt"
      Tab(3).Control(18)=   "lblNet"
      Tab(3).Control(19)=   "lblAmt"
      Tab(3).Control(20)=   "lblDspAcmSaleQty"
      Tab(3).Control(21)=   "lblQty"
      Tab(3).Control(22)=   "lblCusCrtDate"
      Tab(3).Control(23)=   "lblDspCrtDate"
      Tab(3).ControlCount=   24
      Begin VB.TextBox txtCusRemark 
         Enabled         =   0   'False
         Height          =   1020
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   1320
         Width           =   7665
      End
      Begin VB.ComboBox cboCusRgnCode 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "C001.frx":6CDD
         Left            =   -73800
         List            =   "C001.frx":6CDF
         TabIndex        =   14
         Top             =   1740
         Width           =   2355
      End
      Begin VB.ComboBox cboCusMLCode 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "C001.frx":6CE1
         Left            =   1680
         List            =   "C001.frx":6CE3
         TabIndex        =   39
         Top             =   960
         Width           =   1875
      End
      Begin VB.TextBox txtCusCreditLimit 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3600
         TabIndex        =   37
         Top             =   600
         Width           =   1155
      End
      Begin VB.TextBox txtCusPayDesc 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3000
         TabIndex        =   34
         Top             =   240
         Width           =   1740
      End
      Begin VB.TextBox txtCusAddress4 
         Height          =   300
         Left            =   -73800
         TabIndex        =   13
         Top             =   1380
         Width           =   7695
      End
      Begin VB.TextBox txtCusAddress3 
         Height          =   300
         Left            =   -73800
         TabIndex        =   12
         Top             =   1020
         Width           =   7695
      End
      Begin VB.TextBox txtCusAddress2 
         Height          =   300
         Left            =   -73800
         TabIndex        =   11
         Top             =   660
         Width           =   7695
      End
      Begin VB.TextBox txtCusAddress1 
         Height          =   300
         Left            =   -73800
         TabIndex        =   10
         Top             =   300
         Width           =   7695
      End
      Begin VB.TextBox txtCusSpecDis 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6120
         TabIndex        =   38
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboCusSaleCode 
         Height          =   300
         Left            =   6120
         TabIndex        =   35
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox cboCusCurr 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "C001.frx":6CE5
         Left            =   1680
         List            =   "C001.frx":6CE7
         TabIndex        =   36
         Top             =   600
         Width           =   915
      End
      Begin VB.ComboBox cboCusPayCode 
         Height          =   300
         Left            =   1680
         TabIndex        =   33
         Top             =   240
         Width           =   1275
      End
      Begin VB.ComboBox cboCusShipTerrCode 
         Height          =   300
         Left            =   -73680
         TabIndex        =   22
         Top             =   2520
         Width           =   1155
      End
      Begin VB.ComboBox cboCusShipTerrCode2 
         Height          =   300
         Left            =   -69240
         TabIndex        =   31
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Frame fraCusShipAddr1 
         Caption         =   "地址 (一)"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   52
         Top             =   120
         Width           =   4455
         Begin VB.TextBox txtCusShipTerrName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   23
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox txtCusShipFax 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2760
            TabIndex        =   21
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtCusShipAdd2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   16
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipTel 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   20
            Top             =   2040
            Width           =   1425
         End
         Begin VB.TextBox txtCusShipAdd4 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   18
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipAdd3 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   17
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipContactPerson 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   19
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipAdd1 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   15
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label lblCusShipTerrCode 
            Caption         =   "分區代號 :"
            Height          =   240
            Left            =   120
            TabIndex        =   56
            Top             =   2460
            Width           =   1020
         End
         Begin VB.Label lblCusShipTel 
            Caption         =   "電話 / 傳真 :"
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   2100
            Width           =   1020
         End
         Begin VB.Label lblCusShipContactPerson 
            Caption         =   "聯絡人 :"
            Height          =   240
            Left            =   120
            TabIndex        =   54
            Top             =   1725
            Width           =   900
         End
         Begin VB.Label lblCusShipAdd1 
            Caption         =   "送貨地址 :"
            Height          =   600
            Left            =   120
            TabIndex        =   53
            Top             =   300
            Width           =   900
         End
      End
      Begin VB.Frame fraCusShipAddr2 
         Caption         =   "地址 (二)"
         Height          =   2775
         Left            =   -70440
         TabIndex        =   57
         Top             =   120
         Width           =   4455
         Begin VB.TextBox txtCusShipAdd12 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   24
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipContactPerson2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   28
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipAdd32 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   26
            Top             =   960
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipAdd42 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   27
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipTel2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   29
            Top             =   2040
            Width           =   1425
         End
         Begin VB.TextBox txtCusShipAdd22 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            TabIndex        =   25
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtCusShipFax2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2760
            TabIndex        =   30
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtCusShipTerrName2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2400
            TabIndex        =   32
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label lblCusShipAdd2 
            Caption         =   "送貨地址 :"
            Height          =   600
            Left            =   120
            TabIndex        =   61
            Top             =   300
            Width           =   900
         End
         Begin VB.Label lblCusShipContactPerson2 
            Caption         =   "聯絡人 :"
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   1725
            Width           =   900
         End
         Begin VB.Label lblCusShipTel2 
            Caption         =   "電話 / 傳真 :"
            Height          =   240
            Left            =   120
            TabIndex        =   59
            Top             =   2100
            Width           =   1020
         End
         Begin VB.Label lblCusShipTerrCode2 
            Caption         =   "分區代號 :"
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   2460
            Width           =   1020
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   2655
         Left            =   -69840
         OleObjectBlob   =   "C001.frx":6CE9
         TabIndex        =   94
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label lblDspCusSaleName 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   7440
         TabIndex        =   101
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label lblOpenBal 
         Caption         =   "OPENBAL"
         Height          =   240
         Left            =   -74760
         TabIndex        =   100
         Top             =   1980
         Width           =   2115
      End
      Begin VB.Label lblDspOpenBal 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71520
         TabIndex        =   99
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label lblCusLastUpd 
         Caption         =   "最後修改人 :"
         Height          =   240
         Left            =   240
         TabIndex        =   98
         Top             =   2715
         Width           =   2145
      End
      Begin VB.Label lblCusLastUpdDate 
         Caption         =   "最後修改日期 :"
         Height          =   240
         Left            =   4560
         TabIndex        =   97
         Top             =   2715
         Width           =   2580
      End
      Begin VB.Label lblDspCusLastUpd 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   2520
         TabIndex        =   96
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblDspCusLastUpdDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   7440
         TabIndex        =   95
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblDspARBal 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71520
         TabIndex        =   93
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Label lblARBal 
         Caption         =   "ARBAL"
         Height          =   240
         Left            =   -74760
         TabIndex        =   92
         Top             =   2700
         Width           =   2115
      End
      Begin VB.Label lblDspCloseBal 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71520
         TabIndex        =   91
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label lblCloseBal 
         Caption         =   "CLOSEBAL"
         Height          =   240
         Left            =   -74760
         TabIndex        =   90
         Top             =   2340
         Width           =   2115
      End
      Begin VB.Label lblAcmMnSale 
         Caption         =   "ACMMNSALE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   89
         Top             =   1620
         Width           =   1515
      End
      Begin VB.Label lblDspAcmMnSaleNet 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71040
         TabIndex        =   88
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblDspAcmMnSaleAmt 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -72120
         TabIndex        =   87
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblDspAcmMnSaleQty 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -73200
         TabIndex        =   86
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblAcmYrSale 
         Caption         =   "ACMYRSALE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   85
         Top             =   1260
         Width           =   1515
      End
      Begin VB.Label lblDspAcmYrSaleNet 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71040
         TabIndex        =   84
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblDspAcmYrSaleAmt 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -72120
         TabIndex        =   83
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblDspAcmYrSaleQty 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -73200
         TabIndex        =   82
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblAcmSale 
         Caption         =   "ACMSALE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   81
         Top             =   900
         Width           =   1515
      End
      Begin VB.Label lblDspAcmSaleNet 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71040
         TabIndex        =   80
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblDspAcmSaleAmt 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -72120
         TabIndex        =   79
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblNet 
         Caption         =   "NET"
         Height          =   240
         Left            =   -70680
         TabIndex        =   78
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblAmt 
         Caption         =   "AMT"
         Height          =   240
         Left            =   -71760
         TabIndex        =   77
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblDspAcmSaleQty 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -73200
         TabIndex        =   76
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblQty 
         Caption         =   "QTY"
         Height          =   240
         Left            =   -72840
         TabIndex        =   75
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblCusCrtDate 
         Caption         =   "CUSCRTDATE"
         Height          =   240
         Left            =   -74760
         TabIndex        =   74
         Top             =   300
         Width           =   915
      End
      Begin VB.Label lblDspCrtDate 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -73200
         TabIndex        =   73
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblCusRemark 
         Caption         =   "備註 :"
         Height          =   240
         Left            =   240
         TabIndex        =   72
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblCusRgnCode 
         Caption         =   "CUSRGNCODE"
         Height          =   240
         Left            =   -74880
         TabIndex        =   71
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label lblDspCusRgnDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   -71400
         TabIndex        =   70
         Top             =   1740
         Width           =   5265
      End
      Begin VB.Label lblDspCusMLDesc 
         BorderStyle     =   1  '單線固定
         Height          =   300
         Left            =   3600
         TabIndex        =   69
         Top             =   960
         Width           =   5760
      End
      Begin VB.Label lblCusMLCode 
         Caption         =   "CUSMLCODE"
         Height          =   240
         Left            =   240
         TabIndex        =   68
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblCusSaleName 
         Caption         =   "營業員 :"
         Height          =   240
         Left            =   4800
         TabIndex        =   67
         Top             =   315
         Width           =   945
      End
      Begin VB.Label lblCusCreditLimit 
         Caption         =   "信用限度 :"
         Height          =   240
         Left            =   2640
         TabIndex        =   66
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label lblCusPayCode 
         Caption         =   "付款條款 :"
         Height          =   240
         Left            =   240
         TabIndex        =   65
         Top             =   315
         Width           =   1260
      End
      Begin VB.Label lblCusAddress1 
         Caption         =   "發票地址 :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCusCurr 
         Caption         =   "貨幣 :"
         Height          =   240
         Left            =   240
         TabIndex        =   63
         Top             =   660
         Width           =   915
      End
      Begin VB.Label lblCusSpecDis 
         Caption         =   "特別折扣 :"
         Height          =   240
         Left            =   4800
         TabIndex        =   62
         Top             =   660
         Width           =   1740
      End
   End
   Begin VB.Frame fraCustomerInfo 
      Caption         =   "客戶資料"
      Height          =   1695
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Width           =   9495
      Begin VB.TextBox txtCusContactPerson1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5280
         TabIndex        =   7
         Top             =   1320
         Width           =   2265
      End
      Begin VB.TextBox txtCusName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   3
         Top             =   600
         Width           =   6345
      End
      Begin VB.TextBox txtCusTel 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   2385
      End
      Begin VB.TextBox txtCusFax 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5280
         TabIndex        =   5
         Top             =   960
         Width           =   2265
      End
      Begin VB.CheckBox chkBadList 
         Alignment       =   1  '靠右對齊
         Caption         =   "黑名單 :"
         Height          =   225
         Left            =   7680
         TabIndex        =   8
         Top             =   640
         Width           =   1215
      End
      Begin VB.CheckBox chkInActive 
         Alignment       =   1  '靠右對齊
         Caption         =   "有效 :"
         Height          =   180
         Left            =   7680
         TabIndex        =   9
         Top             =   1000
         Width           =   1215
      End
      Begin VB.TextBox txtCusContactPerson 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   2385
      End
      Begin VB.TextBox txtCusEmail 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5280
         TabIndex        =   2
         Top             =   240
         Width           =   3585
      End
      Begin VB.TextBox txtCusCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2385
      End
      Begin VB.Label lblCusContactPerson1 
         Caption         =   "CUSCONTACTPERSON1"
         Height          =   240
         Left            =   3720
         TabIndex        =   102
         Top             =   1380
         Width           =   1500
      End
      Begin VB.Label lblCusCode 
         Caption         =   "編號 :"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label lblCusName 
         Caption         =   "名稱 :"
         Height          =   240
         Left            =   240
         TabIndex        =   49
         Top             =   660
         Width           =   900
      End
      Begin VB.Label lblCusTel 
         Caption         =   "電話 :"
         Height          =   240
         Left            =   240
         TabIndex        =   48
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label lblCusFax 
         Caption         =   "傳真 :"
         Height          =   240
         Left            =   3720
         TabIndex        =   47
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label lblCusContactPerson 
         Caption         =   "聯絡人 :"
         Height          =   240
         Left            =   240
         TabIndex        =   46
         Top             =   1380
         Width           =   900
      End
      Begin VB.Label lblCusEmail 
         Caption         =   "電郵 :"
         Height          =   240
         Left            =   3720
         TabIndex        =   45
         Top             =   300
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmC001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wlSalesmanID As Long
Private wlCusTyp As Long
Private wsFormCaption As String
Private waResult As New XArrayDB
Private waScrItm  As New XArrayDB
Private waScrToolTip As New XArrayDB
 
Private Const tcOpen = "Open"
Private Const tcAdd = "Add"
Private Const tcEdit = "Edit"
Private Const tcDelete = "Delete"
Private Const tcSave = "Save"
Private Const tcCancel = "Cancel"
Private Const tcFind = "Find"
Private Const tcExit = "Exit"

Private Const PERIOD = 0
Private Const SALES = 1
Private Const DEPOSIT = 2
Private Const BALID = 3

Private wiAction As Integer

Private wsActNam(4) As String

Private wlKey As Long
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control

Private wbErr As Boolean

Private Const wsKeyType = "MstCustomer"
Private wsTrnCd As String
Private wsUsrId As String

Private wsOldTerr As String
Private wsOldShipTerr As String
Private wsOldShipTerr2 As String
Private wsOldPayCode As String
Private wsOldSaleCode As String



Private Sub cboCusMLCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCusMLCode.Left + tabDetailInfo.Left, cboCusMLCode.Top + cboCusMLCode.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLCUSML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusMLCode_GotFocus()
    FocusMe cboCusMLCode
End Sub

Private Sub cboCusMLCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboCusMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCusMLCode = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 2
        txtCusRemark.SetFocus
    End If
End Sub

Private Sub cboCusMLCode_LostFocus()
    FocusMe cboCusMLCode, True
End Sub


Private Sub cboCusRgnCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusRgnCode
    
    wsSQL = "SELECT RgnCode, RgnDesc FROM MstRegion WHERE RgnStatus = '1'"
    wsSQL = wsSQL & "ORDER BY RgnCode "
    Call Ini_Combo(2, wsSQL, cboCusRgnCode.Left + tabDetailInfo.Left, cboCusRgnCode.Top + cboCusRgnCode.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLCUSRGN", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusRgnCode_GotFocus()
    FocusMe cboCusRgnCode
End Sub


Private Sub cboCusRgnCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboCusRgnCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCusRgnCode = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 1
        txtCusShipAdd1.SetFocus
            
    End If
End Sub

Private Sub cboCusRgnCode_LostFocus()
    FocusMe cboCusRgnCode, True
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
    MousePointer = vbHourglass
  
    Call IniForm
    Call Ini_Grid
    Call Ini_Caption
    Call Ini_Scr
    
    MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    '-- Resize, not maximum and minimax.
    If Me.WindowState = 0 Then
        Me.Height = 6660
        Me.Width = 10020
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
    Select Case sStatus
        Case "Default"
            Me.cboCusCode.Enabled = False
            Me.txtCusCode.Enabled = False
            
            Me.txtCusName.Enabled = False
            Me.chkBadList.Enabled = False
            Me.chkInActive.Enabled = False
            Me.txtCusTel.Enabled = False
            Me.txtCusFax.Enabled = False
            Me.txtCusContactPerson.Enabled = False
            Me.txtCusContactPerson1.Enabled = False
            Me.txtCusEmail.Enabled = False
            Me.txtCusAddress1.Enabled = False
            Me.txtCusAddress2.Enabled = False
            Me.txtCusAddress3.Enabled = False
            Me.txtCusAddress4.Enabled = False
            
            Me.txtCusShipAdd1.Enabled = False
            Me.txtCusShipAdd2.Enabled = False
            Me.txtCusShipAdd3.Enabled = False
            Me.txtCusShipAdd4.Enabled = False
            Me.txtCusShipContactPerson.Enabled = False
            Me.txtCusShipTel.Enabled = False
            Me.txtCusShipFax.Enabled = False
            Me.txtCusShipTerrName.Enabled = False
            
            Me.txtCusShipAdd12.Enabled = False
            Me.txtCusShipAdd22.Enabled = False
            Me.txtCusShipAdd32.Enabled = False
            Me.txtCusShipAdd42.Enabled = False
            Me.txtCusShipContactPerson2.Enabled = False
            Me.txtCusShipTel2.Enabled = False
            Me.txtCusShipFax2.Enabled = False
            Me.txtCusShipTerrName2.Enabled = False
            
            Me.txtCusPayDesc.Enabled = False
            Me.txtCusCreditLimit.Enabled = False
            Me.txtCusSpecDis.Enabled = False
            Me.txtCusRemark.Enabled = False
            
            Me.cboCusCurr.Enabled = False
            Me.cboCusShipTerrCode.Enabled = False
            Me.cboCusShipTerrCode2.Enabled = False
            Me.cboCusPayCode.Enabled = False
            Me.cboCusSaleCode.Enabled = False
            
            Me.cboCusMLCode.Enabled = False
            Me.cboCusRgnCode.Enabled = False

            
        Case "AfrActAdd"
            Me.cboCusCode.Enabled = False
            Me.cboCusCode.Visible = False
            
            Me.txtCusCode.Enabled = True
            Me.txtCusCode.Visible = True
            
       Case "AfrActEdit"
            Me.cboCusCode.Enabled = True
            Me.cboCusCode.Visible = True
            
            Me.txtCusCode.Enabled = False
            Me.txtCusCode.Visible = False
            
            
        Case "AfrKey"
            Me.cboCusCode.Enabled = False
            Me.txtCusCode.Enabled = False
            
            Me.txtCusName.Enabled = True
            Me.chkBadList.Enabled = True
            Me.chkInActive.Enabled = True
            Me.txtCusTel.Enabled = True
            Me.txtCusFax.Enabled = True
            Me.txtCusContactPerson.Enabled = True
            Me.txtCusContactPerson1.Enabled = True
            Me.txtCusEmail.Enabled = True
           
            Me.txtCusAddress1.Enabled = True
            Me.txtCusAddress2.Enabled = True
            Me.txtCusAddress3.Enabled = True
            Me.txtCusAddress4.Enabled = True
            Me.txtCusShipAdd1.Enabled = True
            Me.txtCusShipAdd2.Enabled = True
            Me.txtCusShipAdd3.Enabled = True
            Me.txtCusShipAdd4.Enabled = True
            Me.txtCusShipContactPerson.Enabled = True
            Me.txtCusShipTel.Enabled = True
            Me.txtCusShipFax.Enabled = True
            Me.txtCusShipTerrName.Enabled = True
            
            Me.txtCusShipAdd12.Enabled = True
            Me.txtCusShipAdd22.Enabled = True
            Me.txtCusShipAdd32.Enabled = True
            Me.txtCusShipAdd42.Enabled = True
            Me.txtCusShipContactPerson2.Enabled = True
            Me.txtCusShipTel2.Enabled = True
            Me.txtCusShipFax2.Enabled = True
            Me.txtCusShipTerrName2.Enabled = True
            Me.txtCusPayDesc.Enabled = True
            Me.txtCusCreditLimit.Enabled = True
            
            Me.txtCusSpecDis.Enabled = True
            Me.txtCusRemark.Enabled = True
            
            Me.cboCusCurr.Enabled = True
            Me.cboCusShipTerrCode.Enabled = True
            Me.cboCusShipTerrCode2.Enabled = True
            Me.cboCusPayCode.Enabled = True
            Me.cboCusSaleCode.Enabled = True
            
            Me.cboCusMLCode.Enabled = True
            Me.cboCusRgnCode.Enabled = True

    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    Dim sTmp As String
    
    InputValidation = False
    
    If Chk_txtCusName = False Then
        Exit Function
    End If

    If Chk_cboCusShipTerrCode(sTmp) = False Then
        Exit Function
    End If

    'If Chk_cboCusShipTerrCode2(sTmp) = False Then
    '    Exit Function
    'End If

    If Chk_cboCusSaleCode(sTmp) = False Then
        Exit Function
    End If
    
    If Chk_cboCusPayCode(sTmp) = False Then
        Exit Function
    End If
    
    If Chk_cboCusCurr = False Then
        Exit Function
    End If
    
    If Chk_cboCusMLCode = False Then
        Exit Function
    End If
    
    If Chk_cboCusRgnCode = False Then
        Exit Function
    End If
    

    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim wsSaleName As String
    Dim wsSaleCode As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT MstCustomer.* "
    wsSQL = wsSQL + "From MstCustomer "
    wsSQL = wsSQL + "WHERE (((MstCustomer.CusCode)='" + cboCusCode + "') AND ((MstCustomer.CusStatus)='1'));"

    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "CusID")
        
        Me.txtCusName = ReadRs(rsRcd, "CusName")
        Call Set_CheckValue(chkBadList, ReadRs(rsRcd, "CusBadList"))
        Call Set_CheckValue(chkInActive, ReadRs(rsRcd, "CusInActive"))
        
        Me.txtCusTel = ReadRs(rsRcd, "CusTel")
        Me.txtCusFax = ReadRs(rsRcd, "CusFax")
        Me.txtCusContactPerson = ReadRs(rsRcd, "CusContactPerson")
        Me.txtCusEmail = ReadRs(rsRcd, "CusEmail")
        
        Me.txtCusContactPerson1 = ReadRs(rsRcd, "CusContactPerson1")
        Me.txtCusAddress1 = ReadRs(rsRcd, "CusAddress1")
        Me.txtCusAddress2 = ReadRs(rsRcd, "CusAddress2")
        Me.txtCusAddress3 = ReadRs(rsRcd, "CusAddress3")
        Me.txtCusAddress4 = ReadRs(rsRcd, "CusAddress4")
        Me.cboCusRgnCode = ReadRs(rsRcd, "CusRgnCode")

        Me.txtCusShipAdd1 = ReadRs(rsRcd, "CusShipAdd1")
        Me.txtCusShipAdd2 = ReadRs(rsRcd, "CusShipAdd2")
        Me.txtCusShipAdd3 = ReadRs(rsRcd, "CusShipAdd3")
        Me.txtCusShipAdd4 = ReadRs(rsRcd, "CusShipAdd4")
        Me.txtCusShipContactPerson = ReadRs(rsRcd, "CusShipContactPerson")
        Me.txtCusShipTel = ReadRs(rsRcd, "CusShipTel")
        Me.txtCusShipFax = ReadRs(rsRcd, "CusShipFax")
        Me.cboCusShipTerrCode = ReadRs(rsRcd, "CusShipTerrCode")
        Me.txtCusShipTerrName = ReadRs(rsRcd, "CusShipTerrName")
        
        Me.txtCusShipAdd12 = ReadRs(rsRcd, "CusShipAdd12")
        Me.txtCusShipAdd22 = ReadRs(rsRcd, "CusShipAdd22")
        Me.txtCusShipAdd32 = ReadRs(rsRcd, "CusShipAdd32")
        Me.txtCusShipAdd42 = ReadRs(rsRcd, "CusShipAdd42")
        Me.txtCusShipContactPerson2 = ReadRs(rsRcd, "CusShipContactPerson2")
        Me.txtCusShipTel2 = ReadRs(rsRcd, "CusShipTel2")
        Me.txtCusShipFax2 = ReadRs(rsRcd, "CusShipFax2")
        Me.cboCusShipTerrCode2 = ReadRs(rsRcd, "CusShipTerrCode2")
        Me.txtCusShipTerrName2 = ReadRs(rsRcd, "CusShipTerrName2")
        
        
        Me.cboCusPayCode = ReadRs(rsRcd, "CusPayCode")
        Me.txtCusPayDesc = ReadRs(rsRcd, "CusPayTerm")
        wlSalesmanID = ReadRs(rsRcd, "CusSaleID")
        Me.cboCusSaleCode = LoadSaleCodeByID(wlSalesmanID)
        Me.cboCusCurr = ReadRs(rsRcd, "CusCurr")
        Me.txtCusCreditLimit = Format(To_Value(ReadRs(rsRcd, "CusCreditLimit")), gsAmtFmt)
        Me.cboCusMLCode = ReadRs(rsRcd, "CusMLCode")
        Me.txtCusSpecDis = Format(ReadRs(rsRcd, "CusSpecDis"), gsAmtFmt)
        Me.txtCusRemark = ReadRs(rsRcd, "CusRemark")
        
        lblDspCusMLDesc = LoadDescByCode("MstMerchClass", "MLCode", "MLDesc", cboCusMLCode, True)
        lblDspCusRgnDesc = LoadDescByCode("MstRegion", "RgnCode", "RgnDesc", cboCusRgnCode, True)
        
        lblDspCrtDate = Dsp_Date(ReadRs(rsRcd, "CusCrtDate"))
        
        LoadSaleByID wsSaleCode, wsSaleName, wlSalesmanID
        cboCusSaleCode = wsSaleCode
        lblDspCusSaleName = wsSaleName
        
        wsOldShipTerr = cboCusShipTerrCode.Text
        wsOldShipTerr2 = cboCusShipTerrCode2.Text
        wsOldPayCode = cboCusPayCode
        wsOldSaleCode = cboCusSaleCode
        
        Call LoadSaleInfo
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
    Set frmC001 = Nothing

End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        If txtCusAddress1.Enabled And txtCusAddress1.Visible Then
            txtCusAddress1.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 1 Then
        If txtCusShipAdd1.Enabled And txtCusShipAdd1.Visible Then
            txtCusShipAdd1.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 2 Then
        If cboCusPayCode.Enabled And cboCusPayCode.Visible Then
            cboCusPayCode.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 3 Then
        'If txtCusRemark.Enabled And txtCusRemark.Visible Then
        '    tblDetail.SetFocus
        'End If
    End If
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

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            'Case SONO
            '    If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
            '        GoTo Tbl_BeforeColUpdate_Err
            '    End If
            '
            '    If Chk_grdSoNo(.Columns(ColIndex).Text) = False Then
            '       GoTo Tbl_BeforeColUpdate_Err
            '    End If
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

Private Sub tblDetail_BeforeRowColChange(Cancel As Integer)

    On Error GoTo tblDetail_BeforeRowColChange_Err
    With tblDetail
      '  If .Bookmark <> .DestinationRow Then
      '      If Chk_GrdRow(To_Value(.Bookmark)) = False Then
      '          Cancel = True
      '          Exit Sub
      '      End If
      '  End If
    End With
    
    Exit Sub
    
tblDetail_BeforeRowColChange_Err:
    
    MsgBox "Check tblDeiail BeforeRowColChange!"
    Cancel = True

End Sub

Private Sub tblDetail_ButtonClick(ByVal ColIndex As Integer)
    
    Dim wsSQL As String
    Dim wiTop As Long
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
            'Case SONO
            '
            '    wsSql = "SELECT SOHDDOCNO, SOHDDOCDATE FROM soaSOHD "
            '    wsSql = wsSql & " WHERE SOHDSTATUS = '1' "
            '    wsSql = wsSql & " AND SOHDDOCNO LIKE '%" & Set_Quote(.Columns(SONO).Text) & "%' "
            '    wsSql = wsSql & " AND SOHDCUSID = " & wlCusID
            '    wsSql = wsSql & " ORDER BY SOHDDOCNO "
            '
            '    Call Ini_Combo(2, wsSql, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSONO", Me.Width, Me.Height)
            '    tblCommon.Visible = True
            '    tblCommon.SetFocus
            '    Set wcCombo = tblDetail
            '
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
                'Case sono
                '    KeyCode = vbDefault
                '       .Col = BOOKCODE
                'Case SALES
                '    KeyCode = vbDefault
                '       .Col = WhsCode
                   ' KeyCode = vbKeyDown
                   ' .Col = BOOKCODE
                'Case BOOKNAME, BARCODE, WhsCode, LOTNO, PUBLISHER, Qty, DisPer, Amt, Dis
                '    KeyCode = vbDefault
                '    .Col = .Col + 1
                'Case Price, Net, Amtl
                '    KeyCode = vbKeyDown
                '    .Col = BOOKCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> PERIOD Then
                   .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> DEPOSIT Then
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
        
    End Select
End Sub

Private Sub tblDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Button = 2 Then
    '    PopupMenu mnuPopUp
    'End If
End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = PERIOD
        End If
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                'Case SONO
                '    Call Chk_grdSoNo(.Columns(SONO).Text)
                'Case BOOKCODE
                '    Call Chk_grdBookCode(.Columns(SONO).Text, .Columns(BOOKCODE).Text, "", "", "", "", "", 0, 0, "", "", 0)
                'Case WhsCode
                '    Call Chk_grdWhsCode(.Columns(WhsCode).Text)
                ' Case LOTNO
                '    Call Chk_grdLotNo(.Columns(LOTNO).Text)
                'Case Qty
                '    Call Chk_grdQty(.Columns(Qty).Text)
                'Case DisPer
                '    Call Chk_grdDisPer(.Columns(DisPer).Text)
            
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
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
                gsMsg = "你是否確定儲存現時之變更而離開?"
                If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
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
    wsFormID = "C001"
    wsTrnCd = ""
End Sub

Private Sub Ini_Scr()
    Dim MyControl As Control
    
    waResult.ReDim 0, -1, PERIOD, BALID
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
    wlSalesmanID = 0
    
    wsTrnCd = "CUS"
    
    
    wsOldTerr = ""
    wsOldShipTerr = ""
    wsOldShipTerr2 = ""
    wsOldPayCode = ""
    wsOldSaleCode = ""
    
    Call SetFieldStatus("Default")
    Call SetButtonStatus("Default")
    tblCommon.Visible = False
    Me.tabDetailInfo.Tab = 0
    Me.Caption = wsFormCaption
End Sub

Private Sub Ini_Scr_AfrAct()
    Select Case wiAction
    Case AddRec
        Call SetFieldStatus("AfrActAdd")
        Call SetButtonStatus("AfrActAdd")
        txtCusCode.SetFocus
       
    Case CorRec
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCusCode.SetFocus
    
    Case DelRec
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboCusCode.SetFocus
    
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub
Private Sub Ini_Scr_AfrKey()
    Dim Ctrl As Control
    
    Select Case wiAction
    
    Case CorRec, DelRec

        If LoadRecord() = False Then
            gsMsg = "存取檔案失敗! 請聯絡系統管理員或無限系統顧問!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
            Exit Sub
        Else
            If RowLock(wsConnTime, wsKeyType, cboCusCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    txtCusEmail.SetFocus
End Sub

Private Function Chk_CusCode(ByVal inCode As String, ByRef outCode As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_CusCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT CusStatus "
    wsSQL = wsSQL & " FROM MstCustomer WHERE CusCode = '" & Set_Quote(inCode) & "' And CusStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        outCode = ""
    
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    outCode = ReadRs(rsRcd, "CusStatus")
    
    Chk_CusCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Function Chk_CusSaleCode(ByVal inCode As String, ByRef OutID As Long, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_CusSaleCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT SaleID, SaleName "
    wsSQL = wsSQL & " FROM MstSalesman WHERE SaleCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & " AND SaleStatus = '1' "
        
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "SaleName")
        OutID = ReadRs(rsRcd, "SaleID")
    Else
        OutName = ""
        OutID = 0
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_CusSaleCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Function Chk_CusPayCode(ByVal inCode As String, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_CusPayCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT PayDesc "
    wsSQL = wsSQL & " FROM MstPayTerm WHERE PayCode = '" & Set_Quote(inCode) & "' "
    wsSQL = wsSQL & " AND PayStatus = '1' "
        
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "PayDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_CusPayCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Function chk_cboCusCode() As Boolean
    Dim wsStatus As String
    
    chk_cboCusCode = False
    
    If Trim(cboCusCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    End If
        
    If Chk_CusCode(cboCusCode.Text, wsStatus) = False Then
        gsMsg = "客戶編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "客戶編碼已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCusCode.SetFocus
            Exit Function
        End If
    End If
    
    chk_cboCusCode = True
End Function

Private Function Chk_txtCusCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtCusCode = False
    
    If Trim(txtCusCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCusCode.SetFocus
        Exit Function
    End If
        
    If Chk_CusCode(txtCusCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "客戶編碼已存在但已無效!"
        Else
            gsMsg = "客戶編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCusCode.SetFocus
        Exit Function
    End If
    
    Chk_txtCusCode = True
End Function


Private Function Chk_txtCusName() As Boolean
    Chk_txtCusName = False
    
    If Trim(txtCusName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtCusName.SetFocus
        Exit Function
    End If
    
    Chk_txtCusName = True

End Function

Private Function Chk_cboCusSaleCode(ByRef sOutName) As Boolean
    Dim sRetName As String
    
    sRetName = ""
    
    Chk_cboCusSaleCode = False
    
    If Trim(cboCusSaleCode.Text) <> "" Then
        If Chk_CusSaleCode(cboCusSaleCode.Text, wlSalesmanID, sRetName) = False Then
            Me.tabDetailInfo.Tab = 2
            gsMsg = "營業員編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCusSaleCode.SetFocus
            Exit Function
        Else
            sOutName = sRetName
            Chk_cboCusSaleCode = True
        End If
    Else
        Me.tabDetailInfo.Tab = 2
        gsMsg = "營業員編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusSaleCode.SetFocus
        Exit Function
    End If
End Function

Private Function Chk_cboCusPayCode(ByRef sOutName) As Boolean
    Dim sRetName As String
    
    sRetName = ""
    
    Chk_cboCusPayCode = False
    
    If Trim(cboCusPayCode.Text) <> "" Then
        If Chk_CusPayCode(cboCusPayCode.Text, sRetName) = False Then
            Me.tabDetailInfo.Tab = 2
            gsMsg = "付款條款編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboCusPayCode.SetFocus
            Exit Function
        Else
            sOutName = sRetName
        End If
    Else
        Me.tabDetailInfo.Tab = 2
        gsMsg = "沒有輸入須要資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusPayCode.SetFocus
        Exit Function
    End If
    
    Chk_cboCusPayCode = True
End Function

Private Function Chk_cboCusShipTerrCode(ByRef sOutName) As Boolean
    Dim sRetName As String
    
    sRetName = ""
    
    Chk_cboCusShipTerrCode = False
    
    If Trim(cboCusShipTerrCode.Text) = "" Then
        Me.tabDetailInfo.Tab = 1
        gsMsg = "地區編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusShipTerrCode.SetFocus
        Exit Function
    End If
    
    If Chk_CusTerrCode(cboCusShipTerrCode.Text, sRetName) = False Then
        Me.tabDetailInfo.Tab = 1
        gsMsg = "地區編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusShipTerrCode.SetFocus
        Exit Function
    Else
        sOutName = sRetName
        Chk_cboCusShipTerrCode = True
    End If
End Function

Private Function Chk_CusTerrCode(ByVal inCode, ByRef OutName As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_CusTerrCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT TerrDesc "
    wsSQL = wsSQL & " FROM MstTerritory WHERE TerrCode = '" & Set_Quote(inCode) & "' AND TerrStatus='1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "TerrDesc")
    Else
        OutName = ""
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    
    Chk_CusTerrCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboCusShipTerrCode2(ByRef OutName) As Boolean
    Dim sRetName As String
    
    sRetName = ""
    
    Chk_cboCusShipTerrCode2 = False
    
    If Trim(cboCusShipTerrCode2.Text) = "" Then
        Me.tabDetailInfo.Tab = 1
        gsMsg = "地區編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusShipTerrCode2.SetFocus
        Exit Function
    End If
    
    If Chk_CusTerrCode(cboCusShipTerrCode2.Text, sRetName) = False Then
        Me.tabDetailInfo.Tab = 1
        gsMsg = "地區編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboCusShipTerrCode2.SetFocus
        Exit Function
    Else
        OutName = sRetName
        Chk_cboCusShipTerrCode2 = True
    End If
End Function

Private Sub cmdOpen()

    Dim newForm As New frmC001
    
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

Private Sub cmdFind()
     Call OpenPromptForm
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

Private Function cmdSave() As Boolean
    Dim wsGenDte As String
    Dim wsNo As String
    Dim adcmdSave As New ADODB.Command
    Dim wsCode As String
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboCusCode, wsFormID) Then
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
        
    adcmdSave.CommandText = "USP_C001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, UCase(txtCusCode.Text), UCase(cboCusCode.Text)))
    Call SetSPPara(adcmdSave, 4, txtCusName.Text)
    Call SetSPPara(adcmdSave, 5, Get_CheckValue(chkInActive))
    Call SetSPPara(adcmdSave, 6, Get_CheckValue(chkBadList))
    Call SetSPPara(adcmdSave, 7, txtCusContactPerson.Text)
    Call SetSPPara(adcmdSave, 8, txtCusTel.Text)
    Call SetSPPara(adcmdSave, 9, txtCusFax.Text)
    Call SetSPPara(adcmdSave, 10, txtCusEmail.Text)
    
    Call SetSPPara(adcmdSave, 11, txtCusContactPerson1.Text)
    Call SetSPPara(adcmdSave, 12, txtCusAddress1.Text)
    Call SetSPPara(adcmdSave, 13, txtCusAddress2.Text)
    Call SetSPPara(adcmdSave, 14, txtCusAddress3.Text)
    Call SetSPPara(adcmdSave, 15, txtCusAddress4.Text)
    Call SetSPPara(adcmdSave, 16, cboCusRgnCode.Text)
    Call SetSPPara(adcmdSave, 17, "")
    Call SetSPPara(adcmdSave, 18, "")
    
    Call SetSPPara(adcmdSave, 19, txtCusShipAdd1.Text)
    Call SetSPPara(adcmdSave, 20, txtCusShipAdd2.Text)
    Call SetSPPara(adcmdSave, 21, txtCusShipAdd3.Text)
    Call SetSPPara(adcmdSave, 22, txtCusShipAdd4.Text)
    Call SetSPPara(adcmdSave, 23, txtCusShipContactPerson.Text)
    Call SetSPPara(adcmdSave, 24, txtCusShipTel.Text)
    Call SetSPPara(adcmdSave, 25, txtCusShipFax.Text)
    Call SetSPPara(adcmdSave, 26, cboCusShipTerrCode.Text)
    Call SetSPPara(adcmdSave, 27, txtCusShipTerrName.Text)
        
    Call SetSPPara(adcmdSave, 28, txtCusShipAdd12.Text)
    Call SetSPPara(adcmdSave, 29, txtCusShipAdd22.Text)
    Call SetSPPara(adcmdSave, 30, txtCusShipAdd32.Text)
    Call SetSPPara(adcmdSave, 31, txtCusShipAdd42.Text)
    Call SetSPPara(adcmdSave, 32, txtCusShipContactPerson2.Text)
    Call SetSPPara(adcmdSave, 33, txtCusShipTel2.Text)
    Call SetSPPara(adcmdSave, 34, txtCusShipFax2.Text)
    Call SetSPPara(adcmdSave, 35, cboCusShipTerrCode2.Text)
    Call SetSPPara(adcmdSave, 36, txtCusShipTerrName2.Text)
    
    Call SetSPPara(adcmdSave, 37, cboCusPayCode.Text)
    Call SetSPPara(adcmdSave, 38, txtCusPayDesc.Text)
    Call SetSPPara(adcmdSave, 39, wlSalesmanID)
    Call SetSPPara(adcmdSave, 40, cboCusCurr.Text)
    Call SetSPPara(adcmdSave, 41, txtCusCreditLimit.Text)
    Call SetSPPara(adcmdSave, 42, cboCusMLCode.Text)
    Call SetSPPara(adcmdSave, 43, txtCusSpecDis.Text)
    Call SetSPPara(adcmdSave, 44, txtCusRemark.Text)
    
    Call SetSPPara(adcmdSave, 45, gsUserID)
    Call SetSPPara(adcmdSave, 46, wsGenDte)
    
    Call SetSPPara(adcmdSave, 47, "CUS")
    Call SetSPPara(adcmdSave, 48, "")
    
    adcmdSave.Execute
    wsCode = GetSPPara(adcmdSave, 49)
    wsNo = GetSPPara(adcmdSave, 50)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - C001!"
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
    SaveData = False
    
    If (wiAction = AddRec Or wiAction = CorRec Or wiAction = DelRec) And tbrProcess.Buttons(tcSave).Enabled = True Then
        gsMsg = "你是否確定要儲存現時之作業?"
       If MsgBox(gsMsg, vbYesNo, gsTitle) = vbNo Then
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

Private Sub txtCusAddress1_LostFocus()
    FocusMe txtCusAddress1, True
End Sub

Private Sub txtCusAddress2_LostFocus()
    FocusMe txtCusAddress2, True
End Sub

Private Sub txtCusAddress3_LostFocus()
    FocusMe txtCusAddress3, True
End Sub

Private Sub txtCusAddress4_LostFocus()
    FocusMe txtCusAddress4, True
End Sub

Private Sub txtCusCode_GotFocus()
    'Call SelObj(txtCusCode)
    FocusMe txtCusCode
End Sub

Private Sub txtCusCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtCusCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtCusCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtCusCode_LostFocus()
    FocusMe txtCusCode, True
End Sub

Private Sub txtCusContactPerson_LostFocus()
    FocusMe txtCusContactPerson, True
End Sub

Private Sub txtCusContactPerson1_GotFocus()
    FocusMe txtCusContactPerson1
End Sub

Private Sub txtCusContactPerson1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusContactPerson1, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        chkBadList.SetFocus
    End If
End Sub

Private Sub txtCusContactPerson1_LostFocus()
    FocusMe txtCusContactPerson1, True
End Sub

Private Sub txtCusCreditLimit_LostFocus()
    txtCusCreditLimit = Format(txtCusCreditLimit, gsAmtFmt)
    FocusMe txtCusContactPerson, True
End Sub

Private Sub txtCusEmail_LostFocus()
    FocusMe txtCusEmail, True
End Sub

Private Sub txtCusFax_LostFocus()
    FocusMe txtCusFax, True
End Sub

Private Sub txtCusName_GotFocus()
    FocusMe txtCusName
End Sub

Private Sub txtCusName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtCusName() = True Then
            txtCusTel.SetFocus
        End If
        
    End If
    
End Sub

Private Sub txtCusName_LostFocus()
    FocusMe txtCusName, True
End Sub

Private Sub txtCusPayDesc_LostFocus()
    FocusMe txtCusPayDesc, True
End Sub

Private Sub txtCusRemark_LostFocus()
    FocusMe txtCusRemark, True
End Sub

Private Sub txtCusShipAdd1_LostFocus()
    FocusMe txtCusShipAdd1, True
End Sub

Private Sub txtCusShipAdd12_LostFocus()
    FocusMe txtCusShipAdd12, True
End Sub

Private Sub txtCusShipAdd2_LostFocus()
    FocusMe txtCusShipAdd2, True
End Sub

Private Sub txtCusShipAdd22_LostFocus()
    FocusMe txtCusShipAdd22, True
End Sub

Private Sub txtCusShipAdd3_LostFocus()
    FocusMe txtCusShipAdd3, True
End Sub

Private Sub txtCusShipAdd32_LostFocus()
    FocusMe txtCusShipAdd32, True
End Sub

Private Sub txtCusShipAdd4_LostFocus()
    FocusMe txtCusShipAdd4, True
End Sub

Private Sub txtCusShipAdd42_LostFocus()
    FocusMe txtCusShipAdd42, True
End Sub

Private Sub txtCusShipContactPerson_LostFocus()
    FocusMe txtCusShipContactPerson, True
End Sub

Private Sub txtCusShipContactPerson2_LostFocus()
    FocusMe txtCusShipContactPerson2, True
End Sub

Private Sub txtCusShipFax_LostFocus()
    FocusMe txtCusShipFax, True
End Sub

Private Sub txtCusShipFax2_LostFocus()
    FocusMe txtCusShipFax2, True
End Sub

Private Sub txtCusShipTel_LostFocus()
    FocusMe txtCusShipTel, True
End Sub

Private Sub txtCusShipTel2_LostFocus()
    FocusMe txtCusShipTel2, True
End Sub

Private Sub txtCusShipTerrName_LostFocus()
    FocusMe txtCusShipTerrName, True
End Sub

Private Sub txtCusShipTerrName2_LostFocus()
    FocusMe txtCusShipTerrName2, True
End Sub

Private Sub txtCusSpecDis_GotFocus()
    FocusMe txtCusSpecDis
End Sub

Private Sub txtCusSpecDis_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtCusSpecDis, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 2
        cboCusMLCode.SetFocus
    End If
End Sub

Private Sub txtCusSpecDis_LostFocus()
    txtCusSpecDis = Format(txtCusSpecDis, gsAmtFmt)
    FocusMe txtCusSpecDis, True
End Sub

Private Sub txtCusTel_GotFocus()
    FocusMe txtCusTel
End Sub

Private Sub txtCusTel_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusTel, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCusFax.SetFocus
    End If
End Sub

Private Sub txtCusFax_GotFocus()
    FocusMe txtCusFax
End Sub




Private Sub txtCusFax_KeyPress(KeyAscii As Integer)
Call chk_InpLen(txtCusFax, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCusContactPerson.SetFocus
        
    End If
    
End Sub


Private Sub txtCusContactPerson_GotFocus()
    'Call SelObj(txtCusContactPerson)
    FocusMe txtCusContactPerson
End Sub

Private Sub txtCusContactPerson_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusContactPerson, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCusContactPerson1.SetFocus
        
    End If
    
End Sub

Private Sub txtCusEmail_GotFocus()
    'Call SelObj(txtCusEmail)
    FocusMe txtCusEmail
End Sub




Private Sub txtCusEmail_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusEmail, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCusName.SetFocus
        
    End If
End Sub

Private Sub chkBadList_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        chkInActive.SetFocus
        
    End If
    
End Sub

Private Sub chkInActive_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 0
        txtCusAddress1.SetFocus
    End If
    
End Sub


Private Sub txtCusAddress1_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtCusAddress1
End Sub




Private Sub txtCusAddress1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusAddress1, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 0
        txtCusAddress2.SetFocus
    End If
End Sub


Private Sub txtCusAddress2_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtCusAddress2
End Sub

Private Sub txtCusAddress2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusAddress2, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 0
        txtCusAddress3.SetFocus
    End If
End Sub


Private Sub txtCusAddress3_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtCusAddress3
End Sub

Private Sub txtCusAddress3_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusAddress3, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 0
        txtCusAddress4.SetFocus
    End If
End Sub

Private Sub txtCusAddress4_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtCusAddress4
End Sub

Private Sub txtCusAddress4_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusAddress4, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 0
        cboCusRgnCode.SetFocus
    End If
End Sub

Private Sub txtCusTel_LostFocus()
    FocusMe txtCusTel, True
End Sub

Private Sub txtCusShipAdd1_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd1
End Sub




Private Sub txtCusShipAdd1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd1, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd2.SetFocus
    End If
End Sub


Private Sub txtCusShipAdd2_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd2
End Sub

Private Sub txtCusShipAdd2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd2, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd3.SetFocus
    End If
End Sub


Private Sub txtCusShipAdd3_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd3
End Sub

Private Sub txtCusShipAdd3_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd3, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd4.SetFocus
    End If
End Sub





Private Sub txtCusShipAdd4_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd4
End Sub




Private Sub txtCusShipAdd4_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd4, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipContactPerson.SetFocus
    End If
End Sub

Private Sub txtCusShipContactPerson_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipContactPerson
End Sub




Private Sub txtCusShipContactPerson_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipContactPerson, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipTel.SetFocus
    End If
End Sub

Private Sub txtCusShipTel_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipTel
End Sub




Private Sub txtCusShipTel_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipTel, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipFax.SetFocus
    End If
End Sub

Private Sub txtCusShipFax_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipFax
End Sub

Private Sub txtCusShipFax_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipFax, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        cboCusShipTerrCode.SetFocus
    End If
End Sub

Private Sub txtCusShipTerrName_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipTerrName
End Sub




Private Sub txtCusShipTerrName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipTerrName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd12.SetFocus
    End If
End Sub

Private Sub txtCusShipAdd12_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd12
End Sub




Private Sub txtCusShipAdd12_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd12, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd22.SetFocus
    End If
End Sub


Private Sub txtCusShipAdd22_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd22
End Sub

Private Sub txtCusShipAdd22_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd22, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd32.SetFocus
    End If
End Sub


Private Sub txtCusShipAdd32_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd32
End Sub

Private Sub txtCusShipAdd32_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd32, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipAdd42.SetFocus
    End If
End Sub





Private Sub txtCusShipAdd42_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipAdd42
End Sub




Private Sub txtCusShipAdd42_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipAdd42, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipContactPerson2.SetFocus
    End If
End Sub

Private Sub txtCusShipContactPerson2_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipContactPerson2
End Sub




Private Sub txtCusShipContactPerson2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipContactPerson2, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipTel2.SetFocus
    End If
End Sub
Private Sub txtCusShipTel2_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipTel2
End Sub

Private Sub txtCusShipTel2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipTel2, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        txtCusShipFax2.SetFocus
    End If
End Sub

Private Sub txtCusShipFax2_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipFax2
End Sub

Private Sub txtCusShipFax2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipFax2, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 1
        cboCusShipTerrCode2.SetFocus
    End If
End Sub

Private Sub txtCusShipTerrName2_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtCusShipTerrName2
End Sub

Private Sub txtCusShipTerrName2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusShipTerrName2, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 2
        cboCusPayCode.SetFocus
    End If
End Sub

Private Sub txtCusPayDesc_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe txtCusPayDesc
End Sub

Private Sub txtCusPayDesc_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusPayDesc, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 2
        cboCusSaleCode.SetFocus
    End If
End Sub

Private Sub txtCusCreditLimit_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe txtCusCreditLimit
End Sub

Private Sub txtCusCreditLimit_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtCusCreditLimit, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        Me.tabDetailInfo.Tab = 2
        txtCusSpecDis.SetFocus
    End If
End Sub

Private Sub txtCusRemark_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe txtCusRemark
End Sub

Private Sub txtCusRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtCusRemark, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtCusEmail.SetFocus
    End If
End Sub

Private Sub txtSaleID_Change()
   
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    If txtSaleID = "" Then
    Exit Sub
    End If
    
    wsSQL = "SELECT MstSalesman.SaleCode "
    wsSQL = wsSQL + "From MstSalesman "
    wsSQL = wsSQL + "WHERE (((MstSalesman.SaleID)=" + To_Value(txtSaleID) + "));"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
       cboCusSaleCode = ReadRs(rsRcd, "SaleCode")
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Sub

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim sSQL As String
    
    ReDim vFilterAry(9, 2)
    vFilterAry(1, 1) = "編碼"
    vFilterAry(1, 2) = "CusCode"
    
    vFilterAry(2, 1) = "名稱"
    vFilterAry(2, 2) = "CusName"
    
    vFilterAry(3, 1) = "無效"
    vFilterAry(3, 2) = "CusInActive"
    
    vFilterAry(4, 1) = "黑名單"
    vFilterAry(4, 2) = "CusBadList"
    
    vFilterAry(5, 1) = "聯絡人"
    vFilterAry(5, 2) = "CusContactPerson"
    
    vFilterAry(6, 1) = "電話"
    vFilterAry(6, 2) = "CusTel"
    
    vFilterAry(7, 1) = "傳真"
    vFilterAry(7, 2) = "CusFax"
    
    vFilterAry(8, 1) = "電郵"
    vFilterAry(8, 2) = "CusEmail"
    
    vFilterAry(9, 1) = "地區"
    vFilterAry(9, 2) = "CusTerritory"
    
    ReDim vAry(9, 3)
    vAry(1, 1) = "編碼"
    vAry(1, 2) = "CusCode"
    vAry(1, 3) = "800"
    
    vAry(2, 1) = "名稱"
    vAry(2, 2) = "CusName"
    vAry(2, 3) = "2000"
    
    vAry(3, 1) = "聯絡人"
    vAry(3, 2) = "CusContactPerson"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "電話"
    vAry(4, 2) = "CusTel"
    vAry(4, 3) = "1000"
    
    vAry(5, 1) = "傳真"
    vAry(5, 2) = "CusFax"
    vAry(5, 3) = "1000"
    
    vAry(6, 1) = "電郵"
    vAry(6, 2) = "CusEmail"
    vAry(6, 3) = "0"
    
    vAry(7, 1) = "地區"
    vAry(7, 2) = "CusTerritory"
    vAry(7, 3) = "1600"
    
    vAry(8, 1) = "無效"
    vAry(8, 2) = "CusInActive"
    vAry(8, 3) = "550"
    
    vAry(9, 1) = "黑名單"
    vAry(9, 2) = "CusBadList"
    vAry(9, 3) = "550"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstCustomer.CusCode, MstCustomer.CusName, "
        sSQL = sSQL + "MstCustomer.CusContactPerson, MstCustomer.CusTel, MstCustomer.CusFax, MstCustomer.CusEmail, "
        sSQL = sSQL + "MstCustomer.CusTerritory, MstCustomer.CusInActive, MstCustomer.CusBadList "
        sSQL = sSQL + "FROM MstCustomer "
        .sBindSQL = sSQL
        .sBindWhereSQL = "WHERE MstCustomer.CusStatus = '1' "
        .sBindOrderSQL = "ORDER BY MstCustomer.CusName"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboCusCode Then
        cboCusCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
    Unload frmShareSearch
End Sub

Public Function Chk_cboCusCurr() As Boolean
    Chk_cboCusCurr = False
    
    If Trim(cboCusCurr) = "" Then
        Me.tabDetailInfo.Tab = 2
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCurr.SetFocus
        Exit Function
    End If
    
    If Chk_CusCurr = False Then
        gsMsg = "貨幣不存在!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboCusCurr.SetFocus
        Exit Function
    End If
    
    Chk_cboCusCurr = True
End Function

Public Function Chk_CusCurr() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT ExcCurr FROM MstExchangeRate WHERE ExcCurr='" & Set_Quote(cboCusCurr.Text) + "' And ExcStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount < 1 Then
        Chk_CusCurr = False
    Else
        Chk_CusCurr = True
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
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

Private Sub cboCusCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusCode
    
    wsSQL = "SELECT CusCode, CusName FROM MstCustomer WHERE CusStatus = '1'"
    wsSQL = wsSQL & " AND CusCode LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "

    wsSQL = wsSQL & "ORDER BY CusCode "
    Call Ini_Combo(2, wsSQL, cboCusCode.Left, cboCusCode.Top + cboCusCode.Height, tblCommon, "C001", "TBLC", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboCusCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboCusCode_GotFocus()
    FocusMe cboCusCode
End Sub

Private Sub cboCusCode_LostFocus()
    FocusMe cboCusCode, True
End Sub

Private Sub cboCusShipTerrCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusShipTerrCode
    
    wsSQL = "SELECT TerrCode, TerrDesc FROM MstTerritory WHERE TerrStatus = '1'"
    wsSQL = wsSQL & "ORDER BY TerrCode "
    Call Ini_Combo(2, wsSQL, cboCusShipTerrCode.Left + tabDetailInfo.Left, cboCusShipTerrCode.Top + cboCusShipTerrCode.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLTERR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusShipTerrCode_KeyPress(KeyAscii As Integer)
    Dim sShipTerritory As String
    
    Call chk_InpLen(cboCusShipTerrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCusShipTerrCode(sShipTerritory) = True Then
            If wsOldShipTerr <> cboCusShipTerrCode.Text Then
                txtCusShipTerrName = sShipTerritory
                wsOldShipTerr = cboCusShipTerrCode.Text
            End If
            Me.tabDetailInfo.Tab = 1
            txtCusShipTerrName.SetFocus
        End If
    End If
End Sub

Private Sub cboCusShipTerrCode_GotFocus()
    FocusMe cboCusShipTerrCode
End Sub

Private Sub cboCusShipTerrCode_LostFocus()
    FocusMe cboCusShipTerrCode, True
End Sub

Private Sub cboCusShipTerrCode2_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusShipTerrCode2
    
    wsSQL = "SELECT TerrCode, TerrDesc FROM MstTerritory WHERE TerrStatus = '1'"
    wsSQL = wsSQL & "ORDER BY TerrCode "
    Call Ini_Combo(2, wsSQL, cboCusShipTerrCode2.Left + tabDetailInfo.Left, cboCusShipTerrCode2.Top + cboCusShipTerrCode2.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLTERR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusShipTerrCode2_KeyPress(KeyAscii As Integer)
    Dim sShipTerritory2 As String
    
    Call chk_InpLen(cboCusShipTerrCode2, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        'If Chk_cboCusShipTerrCode2(sShipTerritory2) = True Then
        If wsOldShipTerr2 <> cboCusShipTerrCode2.Text Then
            txtCusShipTerrName2 = sShipTerritory2
            wsOldShipTerr2 = cboCusShipTerrCode2.Text
        End If
        Me.tabDetailInfo.Tab = 1
        txtCusShipTerrName2.SetFocus
        'End If
    End If
End Sub

Private Sub cboCusShipTerrCode2_GotFocus()
    FocusMe cboCusShipTerrCode2
End Sub

Private Sub cboCusShipTerrCode2_LostFocus()
    FocusMe cboCusShipTerrCode2, True
End Sub

Private Sub cboCusPayCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusPayCode
    
    wsSQL = "SELECT PayCode, PayDesc, PayDay FROM MstPayTerm WHERE PayStatus = '1'"
    wsSQL = wsSQL & "ORDER BY PayCode "
    Call Ini_Combo(3, wsSQL, cboCusPayCode.Left + tabDetailInfo.Left, cboCusPayCode.Top + cboCusPayCode.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLPYT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusPayCode_KeyPress(KeyAscii As Integer)
    Dim sPayDesc As String
    
    Call chk_InpLen(cboCusPayCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCusPayCode(sPayDesc) = True Then
            If wsOldPayCode <> cboCusPayCode.Text Then
                txtCusPayDesc = sPayDesc
                wsOldPayCode = cboCusPayCode.Text
            End If
            Me.tabDetailInfo.Tab = 2
            txtCusPayDesc.SetFocus
        End If
    End If
End Sub

Private Sub cboCusPayCode_GotFocus()
    FocusMe cboCusPayCode
End Sub

Private Sub cboCusPayCode_LostFocus()
    FocusMe cboCusPayCode, True
End Sub

Private Sub cboCusCurr_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusCurr
    
    wsSQL = "SELECT DISTINCT ExcCurr FROM MstExchangeRate WHERE ExcStatus = '1'"
    wsSQL = wsSQL & "ORDER BY ExcCurr "
    Call Ini_Combo(1, wsSQL, cboCusCurr.Left + tabDetailInfo.Left, cboCusCurr.Top + cboCusCurr.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLCUR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboCusCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCusCurr = True Then
            txtCusCreditLimit.SetFocus
        End If
    End If
End Sub

Private Sub cboCusCurr_GotFocus()
    FocusMe cboCusCurr
End Sub

Private Sub cboCusCurr_LostFocus()
    FocusMe cboCusCurr, True
End Sub

Private Sub cboCusSaleCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboCusSaleCode
    
    wsSQL = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleStatus = '1'"
    wsSQL = wsSQL & " and SaleType = 'S' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSQL, cboCusSaleCode.Left + tabDetailInfo.Left, cboCusSaleCode.Top + cboCusSaleCode.Height + tabDetailInfo.Top, tblCommon, "C001", "TBLSLM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboCusSaleCode_KeyPress(KeyAscii As Integer)
    Dim sSalesName As String
    
    Call chk_InpLen(cboCusSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCusSaleCode(sSalesName) = True Then
            If wsOldSaleCode <> cboCusSaleCode.Text Then
                lblDspCusSaleName = sSalesName
                wsOldSaleCode = cboCusSaleCode.Text
            End If
            Me.tabDetailInfo.Tab = 2
            cboCusCurr.SetFocus
        End If
    End If
End Sub

Private Sub cboCusSaleCode_GotFocus()
    FocusMe cboCusSaleCode
End Sub

Private Sub cboCusSaleCode_LostFocus()
    FocusMe cboCusSaleCode, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT CusStatus FROM MstCustomer WHERE CusCode = '" & Set_Quote(txtCusCode) & "'"
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
    
    'Create Selection Criteria
    With Newfrm
        .TableID = wsKeyType
        .TableType = wsTrnCd
        .TableKey = "CusCode"
        .KeyLen = 10
        Set .ctlKey = txtCusCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

 '   Call Get_Scr_Item("V001", waScrItm)
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    fraCustomerInfo.Caption = Get_Caption(waScrItm, "FRACUSTOMERINFO")
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    chkBadList.Caption = Get_Caption(waScrItm, "BADLIST")
    chkInActive.Caption = Get_Caption(waScrItm, "INACTIVE")
    lblCusTel.Caption = Get_Caption(waScrItm, "CUSTEL")
    lblCusFax.Caption = Get_Caption(waScrItm, "CUSFAX")
    lblCusContactPerson.Caption = Get_Caption(waScrItm, "CUSCONTACTPERSON")
    lblCusEMail.Caption = Get_Caption(waScrItm, "CUSEMAIL")
    lblCusAddress1.Caption = Get_Caption(waScrItm, "CUSADDRESS1")
    fraCusShipAddr1.Caption = Get_Caption(waScrItm, "FRACUSSHIPADDR1")
    fraCusShipAddr2.Caption = Get_Caption(waScrItm, "FRACUSSHIPADDR2")
    lblCusShipAdd1.Caption = Get_Caption(waScrItm, "CUSSHIPADD1")
    lblCusShipContactPerson.Caption = Get_Caption(waScrItm, "CUSSHIPCONTACTPERSON")
    lblCusShipTel.Caption = Get_Caption(waScrItm, "CUSSHIPTEL")
    lblCusShipTerrCode.Caption = Get_Caption(waScrItm, "CUSSHIPTERRCODE")
    lblCusShipAdd2.Caption = Get_Caption(waScrItm, "CUSSHIPADD2")
    lblCusShipContactPerson2.Caption = Get_Caption(waScrItm, "CUSSHIPCONTACTPERSON2")
    lblCusShipTel2.Caption = Get_Caption(waScrItm, "CUSSHIPTEL2")
    lblCusShipTerrCode2.Caption = Get_Caption(waScrItm, "CUSSHIPTERRCODE2")
    lblCusPayCode.Caption = Get_Caption(waScrItm, "CUSPAYCODE")
    lblCusCreditLimit.Caption = Get_Caption(waScrItm, "CUSCREDITLIMIT")
    lblCusCurr.Caption = Get_Caption(waScrItm, "CUSCURR")
    lblCusSpecDis.Caption = Get_Caption(waScrItm, "CUSSPECDIS")
    lblCusRemark.Caption = Get_Caption(waScrItm, "CUSREMARK")
    lblCusMLCode.Caption = Get_Caption(waScrItm, "CUSMLCODE")
    lblCusRgnCode.Caption = Get_Caption(waScrItm, "CUSRGNCODE")
    lblCusSaleName.Caption = Get_Caption(waScrItm, "CUSSALENAME")
    
    lblCusContactPerson1.Caption = Get_Caption(waScrItm, "CUSCONTACTPERSON1")
    
    lblCusLastUpd.Caption = Get_Caption(waScrItm, "CUSLASTUPD")
    lblCusLastUpdDate.Caption = Get_Caption(waScrItm, "CUSLASTUPDDATE")
    
    lblCusCrtDate.Caption = Get_Caption(waScrItm, "CUSCRTDATE")
    lblAcmSale.Caption = Get_Caption(waScrItm, "ACMSALE")
    lblAcmYrSale.Caption = Get_Caption(waScrItm, "ACMYRSALE")
    lblAcmMnSale.Caption = Get_Caption(waScrItm, "ACMMNSALE")
    lblOpenBal.Caption = Get_Caption(waScrItm, "OPENBAL")
    lblCloseBal.Caption = Get_Caption(waScrItm, "CLOSEBAL")
    lblARBal.Caption = Get_Caption(waScrItm, "ARBAL")
    lblQty.Caption = Get_Caption(waScrItm, "QTY")
    lblAmt.Caption = Get_Caption(waScrItm, "AMT")
    lblNet.Caption = Get_Caption(waScrItm, "NET")
    
    With tblDetail
        .Columns(PERIOD).Caption = Get_Caption(waScrItm, "PERIOD")
        .Columns(SALES).Caption = Get_Caption(waScrItm, "SALES")
        .Columns(DEPOSIT).Caption = Get_Caption(waScrItm, "DEPOSIT")
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO0")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO1")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO2")
    tabDetailInfo.TabCaption(3) = Get_Caption(waScrItm, "TABDETAILINFO3")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
       
    wsActNam(1) = Get_Caption(waScrItm, "CADD")
    wsActNam(2) = Get_Caption(waScrItm, "CEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "CDELETE")
    
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_cboCusMLCode() As Boolean
    Dim wsDesc As String
    Chk_cboCusMLCode = False
     
    If Trim(cboCusMLCode.Text) = "" Then
        gsMsg = "必須輸入會計號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCusMLCode.SetFocus
        Exit Function
    End If
    
    If Chk_MerchClass(cboCusMLCode, wsDesc) = False Then
        gsMsg = "無此會計號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboCusMLCode.SetFocus
        lblDspCusMLDesc = ""
       Exit Function
    End If
    
    lblDspCusMLDesc = wsDesc
    
    Chk_cboCusMLCode = True
End Function

Private Function Chk_cboCusRgnCode() As Boolean
    Dim wsDesc As String
    Chk_cboCusRgnCode = False
     
    If Trim(cboCusRgnCode.Text) = "" Then
        gsMsg = "必須輸入銷售區域!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboCusRgnCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_Region(cboCusRgnCode, wsDesc) = False Then
        gsMsg = "無此銷售區域!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboCusRgnCode.SetFocus
        lblDspCusRgnDesc = ""
       Exit Function
    End If
    
    lblDspCusRgnDesc = wsDesc
    
    Chk_cboCusRgnCode = True
End Function

Private Sub Ini_Grid()
    
    Dim wiCtr As Integer

    With tblDetail
        .EmptyRows = True
        .MultipleLines = 0
        .AllowAddNew = True
        .AllowUpdate = True
        .AllowDelete = True
        .AlternatingRowStyle = True
        .RecordSelectors = False
        .AllowColMove = False
        .AllowColSelect = False
        
        For wiCtr = PERIOD To BALID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case PERIOD
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                Case SALES
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                Case DEPOSIT
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                Case BALID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
            End Select
        Next
        .Styles("EvenRow").BackColor = &H8000000F
    End With
    
End Sub

Private Function IsEmptyRow(Optional inRow) As Boolean

    IsEmptyRow = True
    
        If IsMissing(inRow) Then
            With tblDetail
                If Trim(.Columns(PERIOD)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, PERIOD)) = "" And _
                   Trim(waResult(inRow, SALES)) = "" And _
                   Trim(waResult(inRow, DEPOSIT)) = "" And _
                   Trim(waResult(inRow, BALID)) = "" Then
                   Exit Function
                End If
            End If
        End If
    
    IsEmptyRow = False
    
End Function

Private Function LoadSaleInfo() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiCtr As Long
    Dim wsYYYY As String
    Dim wsMM As String
    
    Dim wdARBal As Double
    Dim wdOpnBal As Double
    Dim wdTotBal As Double
    Dim wdCMQty As Double
    Dim wdCYQty As Double
    Dim wdTotQty As Double
    Dim wdCMSal As Double
    Dim wdCYSal As Double
    Dim wdTotSal As Double
    Dim wdCMNet As Double
    Dim wdCYNet As Double
    Dim wdTotNet As Double
    Dim wdAmt As Double
    
    
    wsYYYY = Left(gsSystemDate, 4)
    wsMM = Mid(gsSystemDate, 6, 2)
    
    
    Me.MousePointer = vbHourglass
    LoadSaleInfo = False
    
    Call Get_CusSaleInfo(wlKey, wsYYYY, wsMM, 0, 0, wdOpnBal, wdTotBal, wdCMQty, wdCYQty, wdTotQty, wdCMSal, wdCYSal, wdTotSal, wdCMNet, wdCYNet, wdTotNet)
  
    lblDspARBal.Caption = Format(wdTotBal, gsAmtFmt)
    lblDspOpenBal.Caption = Format(wdOpnBal, gsAmtFmt)
    lblDspCloseBal.Caption = Format(wdTotBal, gsAmtFmt)
    
    lblDspAcmSaleQty.Caption = Format(wdTotQty, gsQtyFmt)
    lblDspAcmSaleNet.Caption = Format(wdTotNet, gsAmtFmt)
    lblDspAcmSaleAmt.Caption = Format(wdTotSal, gsAmtFmt)
    
    lblDspAcmYrSaleQty.Caption = Format(wdCYQty, gsQtyFmt)
    lblDspAcmYrSaleNet.Caption = Format(wdCYNet, gsAmtFmt)
    lblDspAcmYrSaleAmt.Caption = Format(wdCYSal, gsAmtFmt)
    
    lblDspAcmMnSaleQty.Caption = Format(wdCMQty, gsQtyFmt)
    lblDspAcmMnSaleNet.Caption = Format(wdCMNet, gsAmtFmt)
    lblDspAcmMnSaleAmt.Caption = Format(wdCMSal, gsAmtFmt)
    
    
    wsSQL = "SELECT SOHDCTLPRD, SUM(SODTNETL) NETL "
    wsSQL = wsSQL & " FROM SOASOHD, SOASODT "
    wsSQL = wsSQL & " WHERE SOHDCUSID = " & wlKey
    wsSQL = wsSQL & " AND SOHDDOCID = SODTDOCID "
    wsSQL = wsSQL & " AND SOHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND SOHDCTLPRD >= '" & wsYYYY & "01" & "'"
    wsSQL = wsSQL & " GROUP BY SOHDCTLPRD "
    wsSQL = wsSQL & " ORDER BY SOHDCTLPRD "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        waResult.ReDim 0, -1, PERIOD, BALID
        tblDetail.ReBind
        tblDetail.Bookmark = 0
        Me.MousePointer = vbNormal
        Exit Function
    End If
    
     
    With waResult
    .ReDim 0, -1, PERIOD, BALID
    rsRcd.MoveFirst
    Do Until rsRcd.EOF
    
    wdAmt = Get_CusCreditAmt(wlKey, ReadRs(rsRcd, "SOHDCTLPRD"))
     
     .AppendRows
        waResult(.UpperBound(1), PERIOD) = ReadRs(rsRcd, "SOHDCTLPRD")
        waResult(.UpperBound(1), BALID) = ReadRs(rsRcd, "SOHDCTLPRD")
        waResult(.UpperBound(1), SALES) = Format(To_Value(ReadRs(rsRcd, "NETL")), gsAmtFmt)
        waResult(.UpperBound(1), DEPOSIT) = Format(wdAmt, gsAmtFmt)
    rsRcd.MoveNext
    Loop
    End With
    
    tblDetail.ReBind
    tblDetail.Bookmark = 0
    
    

    
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    LoadSaleInfo = True
    Me.MousePointer = vbNormal
    
End Function

