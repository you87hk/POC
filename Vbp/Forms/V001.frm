VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmV001 
   BackColor       =   &H8000000A&
   Caption         =   "供應商資料"
   ClientHeight    =   6075
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9
      Charset         =   136
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "V001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9945
   StartUpPosition =   2  '螢幕中央
   Visible         =   0   'False
   Begin VB.ComboBox cboVdrRgnCode 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "V001.frx":08CA
      Left            =   5880
      List            =   "V001.frx":08CC
      TabIndex        =   2
      Top             =   840
      Width           =   915
   End
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11400
      OleObjectBlob   =   "V001.frx":08CE
      TabIndex        =   32
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cboVdrCode 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame fraVdrInfo 
      Caption         =   "供應商資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Width           =   9615
      Begin VB.TextBox txtVdrCode 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Tag             =   "K"
         Top             =   240
         Width           =   2265
      End
      Begin VB.TextBox txtVdrName 
         BackColor       =   &H80000014&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   7185
      End
      Begin VB.TextBox txtVdrTel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtVdrFax 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   6
         Top             =   960
         Width           =   3465
      End
      Begin VB.CheckBox chkInActive 
         Alignment       =   1  '靠右對齊
         Caption         =   "有效 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7920
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtVdrContactName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtVdrEmail 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   8
         Top             =   1320
         Width           =   3465
      End
      Begin VB.Label lblVdrTel 
         Caption         =   "電話 :"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   90
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label lblVdrName 
         Caption         =   "名稱 :"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   89
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label lblDspVdrRgnDesc 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6360
         TabIndex        =   62
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lblVdrRgnCode 
         Caption         =   "VDRRGNCODE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   61
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label lblVdrCode 
         Caption         =   "VDRCODE"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   37
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label lblVdrFax 
         Caption         =   "傳真 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   36
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label lblVdrContactName 
         Caption         =   "聯絡人 :"
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   35
         Top             =   1380
         Width           =   1380
      End
      Begin VB.Label lblVdrEmail 
         Caption         =   "電郵 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   34
         Top             =   1380
         Width           =   900
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
            Picture         =   "V001.frx":2FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":38AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":4185
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":45D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":4A29
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":4D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":5195
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":55E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":5901
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":5C1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":606D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "V001.frx":6949
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   31
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
   Begin TabDlg.SSTab tabDetailInfo 
      Height          =   3375
      Left            =   240
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2400
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5953
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "附加通訊資料"
      TabPicture(0)   =   "V001.frx":6C71
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtVdrContactName1"
      Tab(0).Control(1)=   "txtVdrAddress4"
      Tab(0).Control(2)=   "txtVdrAddress3"
      Tab(0).Control(3)=   "txtVdrAddress2"
      Tab(0).Control(4)=   "txtVdrAddress1"
      Tab(0).Control(5)=   "lblVdrContactName1"
      Tab(0).Control(6)=   "lblVdrAddress1"
      Tab(0).Control(7)=   "lblDspVdrLastUpdDate"
      Tab(0).Control(8)=   "lblDspVdrLastUpd"
      Tab(0).Control(9)=   "lblVdrLastUpdDate"
      Tab(0).Control(10)=   "lblVdrLastUpd"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "貨運資料"
      TabPicture(1)   =   "V001.frx":6C8D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraVdrShipAddr1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "其他資料"
      TabPicture(2)   =   "V001.frx":6CA9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboVdrSaleCode"
      Tab(2).Control(1)=   "cboVdrMLCode"
      Tab(2).Control(2)=   "txtVdrCreditLimit"
      Tab(2).Control(3)=   "txtVdrPayTerm"
      Tab(2).Control(4)=   "txtVdrRemark"
      Tab(2).Control(5)=   "txtVdrSpecDis"
      Tab(2).Control(6)=   "cboVdrCurr"
      Tab(2).Control(7)=   "cboVdrPayCode"
      Tab(2).Control(8)=   "lblVdrSaleName"
      Tab(2).Control(9)=   "lblDspVdrSaleName"
      Tab(2).Control(10)=   "lblVdrMLCode"
      Tab(2).Control(11)=   "lblDspVdrMLDesc"
      Tab(2).Control(12)=   "lblDspVdrOpenBal"
      Tab(2).Control(13)=   "lblVdrCreditLimit"
      Tab(2).Control(14)=   "lblVdrPayCode"
      Tab(2).Control(15)=   "lblVdrRemark"
      Tab(2).Control(16)=   "lblVdrCurr"
      Tab(2).Control(17)=   "lblVdrOpenBal"
      Tab(2).Control(18)=   "lblVdrSpecDis"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "V001.frx":6CC5
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblOpenBal"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblDspOpenBal"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblDspARBal"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblARBal"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblDspCloseBal"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblCloseBal"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblAcmMnSale"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lblDspAcmMnSaleNet"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblDspAcmMnSaleAmt"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblDspAcmMnSaleQty"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lblAcmYrSale"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "lblDspAcmYrSaleNet"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "lblDspAcmYrSaleAmt"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "lblDspAcmYrSaleQty"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "lblAcmSale"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "lblDspAcmSaleNet"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "lblDspAcmSaleAmt"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "lblNet"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "lblAmt"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "lblDspAcmSaleQty"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "lblQty"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "lblVdrCrtDate"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "lblDspCrtDate"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "tblDetail"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).ControlCount=   24
      Begin VB.ComboBox cboVdrSaleCode 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69480
         TabIndex        =   23
         Top             =   360
         Width           =   1275
      End
      Begin VB.ComboBox cboVdrMLCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "V001.frx":6CE1
         Left            =   -73800
         List            =   "V001.frx":6CE3
         TabIndex        =   27
         Top             =   1080
         Width           =   915
      End
      Begin VB.TextBox txtVdrContactName1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtVdrCreditLimit 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69480
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtVdrPayTerm 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72840
         TabIndex        =   25
         Top             =   720
         Width           =   2100
      End
      Begin VB.Frame fraVdrShipAddr1 
         Caption         =   "運送資料"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74880
         TabIndex        =   39
         Top             =   120
         Width           =   8895
         Begin VB.TextBox txtVdrShipAdd2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            TabIndex        =   17
            Top             =   1200
            Width           =   7575
         End
         Begin VB.TextBox txtVdrShipAdd4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            TabIndex        =   19
            Top             =   1920
            Width           =   7575
         End
         Begin VB.TextBox txtVdrShipAdd3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            TabIndex        =   18
            Top             =   1560
            Width           =   7575
         End
         Begin VB.TextBox txtVdrShipAdd1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            TabIndex        =   16
            Top             =   840
            Width           =   7575
         End
         Begin VB.TextBox txtVdrShipName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            TabIndex        =   14
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtVdrShipContactPerson 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5400
            TabIndex        =   15
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtVdrShipTel 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1200
            TabIndex        =   20
            Top             =   2280
            Width           =   1425
         End
         Begin VB.TextBox txtVdrShipFax 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3240
            TabIndex        =   21
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtVdrShipEmail 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5280
            TabIndex        =   22
            Top             =   2280
            Width           =   3495
         End
         Begin VB.Label lblVdrShipAdd 
            Caption         =   "送貨地址 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   900
            Width           =   900
         End
         Begin VB.Label lblVdrShipName 
            Caption         =   "運貨名稱 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   44
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lblVdrShipContactPerson 
            Caption         =   "聯絡人 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4560
            TabIndex        =   43
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lblVdrShipTel 
            Caption         =   "電話 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   2340
            Width           =   1020
         End
         Begin VB.Label lblVdrShipEmail 
            Caption         =   "電郵 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4800
            TabIndex        =   41
            Top             =   2340
            Width           =   660
         End
         Begin VB.Label lblVdrShipFax 
            Caption         =   "傳真 :"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2760
            TabIndex        =   40
            Top             =   2340
            Width           =   1020
         End
      End
      Begin VB.TextBox txtVdrAddress4 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         TabIndex        =   13
         Top             =   1800
         Width           =   7215
      End
      Begin VB.TextBox txtVdrAddress3 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         TabIndex        =   12
         Top             =   1440
         Width           =   7215
      End
      Begin VB.TextBox txtVdrAddress2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         TabIndex        =   11
         Top             =   1080
         Width           =   7215
      End
      Begin VB.TextBox txtVdrAddress1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73320
         TabIndex        =   10
         Top             =   720
         Width           =   7215
      End
      Begin VB.TextBox txtVdrRemark 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   -73800
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   1800
         Width           =   7665
      End
      Begin VB.TextBox txtVdrSpecDis 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69480
         TabIndex        =   29
         Top             =   1440
         Width           =   3435
      End
      Begin VB.ComboBox cboVdrCurr 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "V001.frx":6CE5
         Left            =   -73800
         List            =   "V001.frx":6CE7
         TabIndex        =   28
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox cboVdrPayCode 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73800
         TabIndex        =   24
         Top             =   720
         Width           =   915
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   2655
         Left            =   5280
         OleObjectBlob   =   "V001.frx":6CE9
         TabIndex        =   63
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblVdrSaleName 
         Caption         =   "VDRSALENAME"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70680
         TabIndex        =   88
         Top             =   435
         Width           =   945
      End
      Begin VB.Label lblDspVdrSaleName 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -68160
         TabIndex        =   87
         Top             =   360
         Width           =   2145
      End
      Begin VB.Label lblDspCrtDate 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   86
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label lblVdrCrtDate 
         Caption         =   "VDRCRTDATE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   85
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label lblQty 
         Alignment       =   2  '置中對齊
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1920
         TabIndex        =   84
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblDspAcmSaleQty 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   83
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblAmt 
         Caption         =   "AMT"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   82
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblNet 
         Caption         =   "NET"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4440
         TabIndex        =   81
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblDspAcmSaleAmt 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   80
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblDspAcmSaleNet 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   79
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label lblAcmSale 
         Caption         =   "ACMSALE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   78
         Top             =   900
         Width           =   1635
      End
      Begin VB.Label lblDspAcmYrSaleQty 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   77
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblDspAcmYrSaleAmt 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   76
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblDspAcmYrSaleNet 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   75
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblAcmYrSale 
         Caption         =   "ACMYRSALE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   74
         Top             =   1260
         Width           =   1635
      End
      Begin VB.Label lblDspAcmMnSaleQty 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   73
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblDspAcmMnSaleAmt 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   72
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblDspAcmMnSaleNet 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   71
         Top             =   1560
         Width           =   1065
      End
      Begin VB.Label lblAcmMnSale 
         Caption         =   "ACMMNSALE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   70
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Label lblCloseBal 
         Caption         =   "CLOSEBAL"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   69
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label lblDspCloseBal 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   68
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label lblARBal 
         Caption         =   "ARBAL"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   67
         Top             =   2700
         Width           =   1635
      End
      Begin VB.Label lblDspARBal 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   66
         Top             =   2640
         Width           =   1545
      End
      Begin VB.Label lblDspOpenBal 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   65
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label lblOpenBal 
         Caption         =   "OPENBAL"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   64
         Top             =   1980
         Width           =   1635
      End
      Begin VB.Label lblVdrMLCode 
         Caption         =   "VDRMLCODE"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   60
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label lblDspVdrMLDesc 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72840
         TabIndex        =   59
         Top             =   1080
         Width           =   2100
      End
      Begin VB.Label lblVdrContactName1 
         Caption         =   "VDRCONTACTNAME1"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   58
         Top             =   420
         Width           =   1500
      End
      Begin VB.Label lblDspVdrOpenBal 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69480
         TabIndex        =   57
         Top             =   1080
         Width           =   3465
      End
      Begin VB.Label lblVdrCreditLimit 
         Caption         =   "信用限度 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70680
         TabIndex        =   56
         Top             =   780
         Width           =   1860
      End
      Begin VB.Label lblVdrPayCode 
         Caption         =   "付款條款 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   55
         Top             =   795
         Width           =   1260
      End
      Begin VB.Label lblVdrAddress1 
         Caption         =   "發票地址 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblVdrRemark 
         Caption         =   "備註 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   53
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label lblVdrCurr 
         Caption         =   "貨幣 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   52
         Top             =   1500
         Width           =   915
      End
      Begin VB.Label lblVdrOpenBal 
         Caption         =   "戶口結存 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70680
         TabIndex        =   51
         Top             =   1140
         Width           =   1740
      End
      Begin VB.Label lblVdrSpecDis 
         Caption         =   "特別折扣 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70680
         TabIndex        =   50
         Top             =   1500
         Width           =   1740
      End
      Begin VB.Label lblDspVdrLastUpdDate 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69120
         TabIndex        =   49
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblDspVdrLastUpd 
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73800
         TabIndex        =   48
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Label lblVdrLastUpdDate 
         Caption         =   "最後修改日期 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70440
         TabIndex        =   47
         Top             =   2715
         Width           =   1260
      End
      Begin VB.Label lblVdrLastUpd 
         Caption         =   "最後修改人 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74880
         TabIndex        =   46
         Top             =   2715
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmV001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private wsFormCaption As String
Private waScrItm As New XArrayDB
Private waResult As New XArrayDB
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

Private wbErr As Boolean

Private wsActNam(4) As String

Private wiAction As Integer
Private wlKey As Long
Private wsFormID As String
Private wsConnTime As String
Private wcCombo As Control
Private wsOldSaleCode As String
Private wlSalesmanID As Long

Private Const wsKeyType = "MstVendor"
Private wsUsrId As String
Private wsTrnCd As String

Private wsOldTerrCode As String
Private wsOldPayCode As String

Private Sub cboVdrMLCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrMLCode
    
    wsSQL = "SELECT MLCode, MLDesc FROM MstMerchClass WHERE MLStatus = '1'"
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboVdrMLCode.Left + tabDetailInfo.Left, cboVdrMLCode.Top + cboVdrMLCode.Height + tabDetailInfo.Top, tblCommon, "V001", "TBLVDRML", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrMLCode_GotFocus()
    FocusMe cboVdrMLCode
End Sub

Private Sub cboVdrMLCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboVdrMLCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboVdrMLCode = False Then
            Exit Sub
        End If
        
        tabDetailInfo.Tab = 2
        cboVdrCurr.SetFocus
    End If
End Sub

Private Sub cboVdrMLCode_LostFocus()
    FocusMe cboVdrMLCode, True
End Sub

Private Sub cboVdrRgnCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrRgnCode
    
    wsSQL = "SELECT RgnCode, RgnDesc FROM MstRegion WHERE RgnStatus = '1'"
    wsSQL = wsSQL & "ORDER BY RgnCode "
    Call Ini_Combo(2, wsSQL, cboVdrRgnCode.Left, cboVdrRgnCode.Top + cboVdrRgnCode.Height, tblCommon, "V001", "TBLVDRRGN", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrRgnCode_GotFocus()
    FocusMe cboVdrRgnCode
End Sub

Private Sub cboVdrRgnCode_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboVdrRgnCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboVdrRgnCode = False Then
            Exit Sub
        End If
        
        chkInActive.SetFocus
    End If
End Sub

Private Sub cboVdrRgnCode_LostFocus()
    FocusMe cboVdrRgnCode, True
End Sub

Private Sub cboVdrSaleCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrSaleCode
    
    wsSQL = "SELECT SaleCode, SaleName FROM MstSalesman WHERE SaleStatus = '1'"
    wsSQL = wsSQL & " and SaleType = 'S' "
    wsSQL = wsSQL & "ORDER BY SaleCode "
    Call Ini_Combo(2, wsSQL, cboVdrSaleCode.Left + tabDetailInfo.Left, cboVdrSaleCode.Top + cboVdrSaleCode.Height + tabDetailInfo.Top + tbrProcess.Height, tblCommon, "V001", "TBLSLM", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrSaleCode_GotFocus()
    FocusMe cboVdrSaleCode
End Sub

Private Sub cboVdrSaleCode_KeyPress(KeyAscii As Integer)
    Dim sSalesName As String
    
    Call chk_InpLen(cboVdrSaleCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboVdrSaleCode(sSalesName) = True Then
            If wsOldSaleCode <> cboVdrSaleCode.Text Then
                lblDspVdrSaleName = sSalesName
                wsOldSaleCode = cboVdrSaleCode.Text
            End If
            Me.tabDetailInfo.Tab = 2
            cboVdrPayCode.SetFocus
        End If
    End If
End Sub

Private Sub cboVdrSaleCode_LostFocus()
    FocusMe cboVdrSaleCode, True
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
                If tbrProcess.Buttons(tcSave).Enabled = True Then
                    Call cmdSave
                End If
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
            Me.txtVdrCode.Enabled = False
            Me.cboVdrCode.Enabled = False
            Me.txtVdrName.Enabled = False
            Me.txtVdrContactName.Enabled = False
            Me.txtVdrTel.Enabled = False
            Me.txtVdrFax.Enabled = False
            Me.txtVdrEmail.Enabled = False
            Me.txtVdrAddress1.Enabled = False
            Me.txtVdrAddress2.Enabled = False
            Me.txtVdrAddress3.Enabled = False
            Me.txtVdrAddress4.Enabled = False
            Me.chkInActive.Enabled = False
            Me.txtVdrPayTerm.Enabled = False
            
            Me.cboVdrCurr.Enabled = False
            Me.txtVdrCreditLimit.Enabled = False
            Me.txtVdrShipName.Enabled = False
            Me.txtVdrShipAdd1.Enabled = False
            Me.txtVdrShipAdd2.Enabled = False
            Me.txtVdrShipAdd3.Enabled = False
            Me.txtVdrShipAdd4.Enabled = False
            Me.txtVdrShipContactPerson.Enabled = False
            Me.txtVdrShipTel.Enabled = False
            Me.txtVdrShipFax.Enabled = False
            Me.txtVdrShipEmail.Enabled = False
            Me.txtVdrSpecDis.Enabled = False
            Me.txtVdrRemark.Enabled = False
            
            Me.cboVdrPayCode.Enabled = False
            Me.cboVdrMLCode.Enabled = False
            Me.txtVdrContactName1.Enabled = False
            Me.cboVdrRgnCode.Enabled = False
            Me.cboVdrSaleCode.Enabled = False
            
        Case "AfrActAdd"
            Me.txtVdrCode.Enabled = True
            Me.txtVdrCode.Visible = True
            
            Me.cboVdrCode.Enabled = False
            Me.cboVdrCode.Visible = False
            
       Case "AfrActEdit"
            Me.txtVdrCode.Enabled = False
            Me.txtVdrCode.Visible = False
            
            Me.cboVdrCode.Enabled = True
            Me.cboVdrCode.Visible = True
            
            
        Case "AfrKey"
            Me.txtVdrCode.Enabled = False
            Me.cboVdrCode.Enabled = False
            
            Me.txtVdrName.Enabled = True
            Me.txtVdrContactName.Enabled = True
            Me.txtVdrTel.Enabled = True
            Me.txtVdrFax.Enabled = True
            Me.txtVdrEmail.Enabled = True
            Me.txtVdrAddress1.Enabled = True
            Me.txtVdrAddress2.Enabled = True
            Me.txtVdrAddress3.Enabled = True
            Me.txtVdrAddress4.Enabled = True
            Me.chkInActive.Enabled = True
            
            Me.txtVdrPayTerm.Enabled = True
            Me.cboVdrCurr.Enabled = True
            Me.txtVdrCreditLimit.Enabled = True
            Me.txtVdrShipName.Enabled = True
            Me.txtVdrShipAdd1.Enabled = True
            Me.txtVdrShipAdd2.Enabled = True
            Me.txtVdrShipAdd3.Enabled = True
            Me.txtVdrShipAdd4.Enabled = True
            Me.txtVdrShipContactPerson.Enabled = True
            Me.txtVdrShipTel.Enabled = True
            Me.txtVdrShipFax.Enabled = True
            Me.txtVdrShipEmail.Enabled = True
            Me.txtVdrSpecDis.Enabled = True
            Me.txtVdrRemark.Enabled = True
            
            Me.cboVdrMLCode.Enabled = True
            Me.txtVdrContactName1.Enabled = True
            Me.cboVdrPayCode.Enabled = True
            Me.cboVdrRgnCode.Enabled = True
            Me.cboVdrSaleCode.Enabled = True
    End Select
End Sub

'-- Input validation checking.
Private Function InputValidation() As Boolean
    InputValidation = False
    
    If Chk_txtVdrName = False Then
        Exit Function
    End If
    
    If Chk_cboVdrMLCode = False Then
        Exit Function
    End If
    
    If Chk_cboVdrRgnCode = False Then
        Exit Function
    End If
    
    If Chk_cboVdrSaleCode("") = False Then
        Exit Function
    End If
    
    If Chk_cboVdrCurr = False Then
        Exit Function
    End If
    
    If Chk_cboVdrPayCode("") = False Then
        Exit Function
    End If
    
    InputValidation = True
End Function

Public Function LoadRecord() As Boolean
    Dim wsSQL As String
    Dim wsSaleName As String
    Dim wsSaleCode As String
    Dim rsRcd As New ADODB.Recordset
    
    wsSQL = "SELECT MstVendor.VdrCode, MstVendor.* "
    wsSQL = wsSQL + "From MstVendor "
    wsSQL = wsSQL + "WHERE (((MstVendor.VdrCode)='" + Set_Quote(cboVdrCode) + "') AND ((MstVendor.VdrStatus)='1'));"

    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
        
    If rsRcd.RecordCount = 0 Then
        LoadRecord = False
        wlKey = 0
    Else
        wlKey = ReadRs(rsRcd, "VdrID")
        
        Me.txtVdrName = ReadRs(rsRcd, "VdrName")
        Me.txtVdrContactName = ReadRs(rsRcd, "VdrContactName")
        Me.txtVdrContactName1 = ReadRs(rsRcd, "VdrContactName1")
        Me.txtVdrTel = ReadRs(rsRcd, "VdrTel")
        Me.txtVdrFax = ReadRs(rsRcd, "VdrFax")
        Me.txtVdrEmail = ReadRs(rsRcd, "VdrEmail")
        Me.txtVdrAddress1 = ReadRs(rsRcd, "VdrAddress1")
        Me.txtVdrAddress2 = ReadRs(rsRcd, "VdrAddress2")
        Me.txtVdrAddress3 = ReadRs(rsRcd, "VdrAddress3")
        Me.txtVdrAddress4 = ReadRs(rsRcd, "VdrAddress4")
        Call Set_CheckValue(chkInActive, ReadRs(rsRcd, "VdrInActive"))
        Me.cboVdrPayCode = ReadRs(rsRcd, "VdrPayCode")
        Me.txtVdrPayTerm = ReadRs(rsRcd, "VdrPayTerm")
        Me.cboVdrCurr = ReadRs(rsRcd, "VdrCurr")
        Me.txtVdrCreditLimit = Format(To_Value(ReadRs(rsRcd, "VdrCreditLimit")), gsAmtFmt)
        Me.txtVdrShipName = ReadRs(rsRcd, "VdrShipName")
        Me.txtVdrShipAdd1 = ReadRs(rsRcd, "VdrShipAdd1")
        Me.txtVdrShipAdd2 = ReadRs(rsRcd, "VdrShipAdd2")
        Me.txtVdrShipAdd3 = ReadRs(rsRcd, "VdrShipAdd3")
        Me.txtVdrShipAdd4 = ReadRs(rsRcd, "VdrShipAdd4")
        Me.txtVdrShipContactPerson = ReadRs(rsRcd, "VdrShipContactPerson")
        Me.txtVdrShipTel = ReadRs(rsRcd, "VdrShipTel")
        Me.txtVdrShipFax = ReadRs(rsRcd, "VdrShipFax")
        Me.txtVdrShipEmail = ReadRs(rsRcd, "VdrShipEmail")
        Me.txtVdrSpecDis = Format(To_Value(ReadRs(rsRcd, "VdrSpecDis")), gsAmtFmt)
        Me.txtVdrRemark = ReadRs(rsRcd, "VdrRemark")
        Me.cboVdrRgnCode = ReadRs(rsRcd, "VdrRgnCode")
        Me.cboVdrMLCode = ReadRs(rsRcd, "VdrMLCode")
        Me.lblDspCrtDate = Dsp_Date(ReadRs(rsRcd, "VdrCrtDate"))
        
        Me.lblDspVdrLastUpd = ReadRs(rsRcd, "VdrLastUpd")
        Me.lblDspVdrLastUpdDate = ReadRs(rsRcd, "VdrLastUpdDate")
        Me.lblDspVdrOpenBal = Format(To_Value(ReadRs(rsRcd, "VdrOpenBal")), gsAmtFmt)
        
        wlSalesmanID = ReadRs(rsRcd, "VdrSaleID")
        LoadSaleByID wsSaleCode, wsSaleName, wlSalesmanID
        cboVdrSaleCode = wsSaleCode
        lblDspVdrSaleName = wsSaleName
        
        lblDspVdrRgnDesc = LoadDescByCode("MstRegion", "RgnCode", "RgnDesc", cboVdrRgnCode, True)
        lblDspVdrMLDesc = LoadDescByCode("MstMerchClass", "MLCode", "MLDesc", cboVdrMLCode, True)
       
        wsOldPayCode = cboVdrPayCode.Text
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
    Set frmV001 = Nothing
End Sub

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        If txtVdrAddress1.Enabled And txtVdrAddress1.Visible Then
            txtVdrContactName1.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 1 Then
        If txtVdrShipName.Enabled And txtVdrShipName.Visible Then
            txtVdrShipName.SetFocus
        End If
    ElseIf tabDetailInfo.Tab = 2 Then
        If txtVdrCreditLimit.Enabled And txtVdrCreditLimit.Visible Then
            cboVdrSaleCode.SetFocus
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
 '   Me.Left = 0
 '   Me.Top = 0
 '   Me.Width = Screen.Width
 '   Me.Height = Screen.Height
    
    
    wsConnTime = Dsp_Date(Now, True)
    wsFormID = "V001"
End Sub

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    For Each MyControl In Me.Controls
        Select Case TypeName(MyControl)
            Case "ComboBox"
                MyControl.Clear
            Case "TextBox"
                MyControl.Font = "MS Sans Serif"
                MyControl.Text = ""
            Case "TDBGrid"
                MyControl.ClearFields
            Case "SSTab"
                MyControl.Font = "MS Sans Serif"
            Case "Label"
                MyControl.Font = "MS Sans Serif"
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
    wsTrnCd = "VDR"
    wsOldTerrCode = ""
    wsOldPayCode = ""

    
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
        txtVdrCode.SetFocus
       
    Case CorRec
           
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboVdrCode.SetFocus
       
    Case DelRec
    
        Call SetFieldStatus("AfrActEdit")
        Call SetButtonStatus("AfrActEdit")
        cboVdrCode.SetFocus
    End Select
    
    Me.Caption = wsFormCaption + " - " & wsActNam(wiAction)
End Sub
Private Sub Ini_Scr_AfrKey()
    Dim Ctrl As Control
    
    Select Case wiAction
    
    Case CorRec, DelRec

        If LoadRecord() = False Then
            gsMsg = "存取記錄失敗! 請聯絡系統管理員或無限系統顧問!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            Exit Sub
        Else
            If RowLock(wsConnTime, wsKeyType, cboVdrCode, wsFormID, wsUsrId) = False Then
                gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
                MsgBox gsMsg, vbOKOnly, gsTitle
            End If
        End If
    End Select
    Call SetFieldStatus("AfrKey")
    Call SetButtonStatus("AfrKey")
    cboVdrRgnCode.SetFocus
End Sub

Private Function Chk_VdrCode(ByVal inCode As String, outCode As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    Chk_VdrCode = False
    
    If Trim(inCode) = "" Then
        Exit Function
    End If
    
    wsSQL = "SELECT VdrCode "
    wsSQL = wsSQL & " FROM MstVendor WHERE VdrCode = '" & Set_Quote(inCode) & "' AND VdrStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount <= 0 Then
        rsRcd.Close
        Exit Function
    End If
    
    
    Chk_VdrCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Function Chk_VdrPayCode(ByVal inCode As String, ByRef OutName As String) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_VdrPayCode = False
        
    If Trim(inCode) = "" Then
        Exit Function
    End If
        
    wsSQL = "SELECT PayDesc "
    wsSQL = wsSQL & " FROM MstPayTerm WHERE PayCode = '" & Set_Quote(inCode) & "' AND PayStatus = '1'"
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        OutName = ReadRs(rsRcd, "PayDesc")
    Else
        OutName = ""
        rsRcd.Close
        Exit Function
    End If
    
    Chk_VdrPayCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function

Private Function Chk_cboVdrPayCode(ByRef OutName As String) As Boolean
    Dim sRetName As String
    
    sRetName = ""
    
    Chk_cboVdrPayCode = False
    
    If Trim(cboVdrPayCode.Text) <> "" Then
        If Chk_VdrPayCode(cboVdrPayCode.Text, sRetName) = False Then
            tabDetailInfo.Tab = 2
            gsMsg = "付款條款編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboVdrPayCode.SetFocus
            Exit Function
        Else
            OutName = sRetName
        End If
    End If
    
    Chk_cboVdrPayCode = True
End Function

Private Function chk_cboVdrCode() As Boolean
    Dim wsStatus As String
    
    chk_cboVdrCode = False
    
    If Trim(cboVdrCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVdrCode.SetFocus
        Exit Function
    End If
        
    If Chk_VdrCode(cboVdrCode.Text, wsStatus) = False Then
        gsMsg = "供應商編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVdrCode.SetFocus
        Exit Function
    Else
        If wsStatus = "2" Then
            gsMsg = "供應商編碼已存在但已無效!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboVdrCode.SetFocus
            Exit Function
        End If
    End If
    
    chk_cboVdrCode = True
End Function

Private Function Chk_txtVdrCode() As Boolean
    Dim wsStatus As String
    
    Chk_txtVdrCode = False
    
    If Trim(txtVdrCode.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVdrCode.SetFocus
        Exit Function
    End If
        
    If Chk_VdrCode(txtVdrCode.Text, wsStatus) = True Then
        If wsStatus = "2" Then
            gsMsg = "供應商編碼已存在但已無效!"
        Else
            gsMsg = "供應商編碼已存在!"
        End If
        
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVdrCode.SetFocus
        Exit Function
    End If
    
    Chk_txtVdrCode = True
End Function

Private Function Chk_txtVdrName() As Boolean
    Chk_txtVdrName = False
    
    If Trim(txtVdrName.Text) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVdrName.SetFocus
        Exit Function
    End If
    
    Chk_txtVdrName = True
End Function
Private Function Chk_txtVdrSpecDis() As Boolean
    Chk_txtVdrSpecDis = False
    
    If To_Value(Trim(txtVdrSpecDis.Text)) < 0 Or To_Value(Trim(txtVdrSpecDis.Text)) > 100 Then
        gsMsg = "特別折扣必需為零至一百!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        txtVdrSpecDis.SetFocus
        Exit Function
    End If
    
    Chk_txtVdrSpecDis = True
End Function

Private Sub cmdOpen()
    Dim newForm As New frmV001
    
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
    Dim gsMsg As String
    Dim wsCode As String
    
    On Error GoTo cmdSave_Err
    
    MousePointer = vbHourglass
    wsGenDte = Format(Date, "YYYY/MM/DD")
    
    If wiAction <> AddRec Then
        If ReadOnlyMode(wsConnTime, wsKeyType, cboVdrCode, wsFormID) Then
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
    
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_V001"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wlKey)
    Call SetSPPara(adcmdSave, 3, IIf(wiAction = AddRec, UCase(txtVdrCode), UCase(cboVdrCode)))
    Call SetSPPara(adcmdSave, 4, txtVdrName)
    Call SetSPPara(adcmdSave, 5, txtVdrContactName)
    Call SetSPPara(adcmdSave, 6, txtVdrContactName1)
    Call SetSPPara(adcmdSave, 7, txtVdrTel)
    
    Call SetSPPara(adcmdSave, 8, txtVdrFax)
    Call SetSPPara(adcmdSave, 9, txtVdrEmail)
    Call SetSPPara(adcmdSave, 10, txtVdrAddress1)
    Call SetSPPara(adcmdSave, 11, txtVdrAddress2)
    Call SetSPPara(adcmdSave, 12, txtVdrAddress3)
    Call SetSPPara(adcmdSave, 13, txtVdrAddress4)
    
    Call SetSPPara(adcmdSave, 14, Get_CheckValue(chkInActive))
    Call SetSPPara(adcmdSave, 15, cboVdrPayCode)
    Call SetSPPara(adcmdSave, 16, txtVdrPayTerm)
    Call SetSPPara(adcmdSave, 17, cboVdrRgnCode)
    Call SetSPPara(adcmdSave, 18, cboVdrMLCode)
    
    Call SetSPPara(adcmdSave, 19, cboVdrCurr)
    Call SetSPPara(adcmdSave, 20, txtVdrCreditLimit)
    Call SetSPPara(adcmdSave, 21, lblDspVdrOpenBal.Caption)
    Call SetSPPara(adcmdSave, 22, txtVdrShipName)
    Call SetSPPara(adcmdSave, 23, txtVdrShipAdd1)
    Call SetSPPara(adcmdSave, 24, txtVdrShipAdd2)
    
    Call SetSPPara(adcmdSave, 25, txtVdrShipAdd3)
    Call SetSPPara(adcmdSave, 26, txtVdrShipAdd4)
    Call SetSPPara(adcmdSave, 27, txtVdrShipContactPerson)
    Call SetSPPara(adcmdSave, 28, txtVdrShipTel)
    Call SetSPPara(adcmdSave, 29, txtVdrShipFax)
    Call SetSPPara(adcmdSave, 30, txtVdrShipEmail)
    Call SetSPPara(adcmdSave, 31, txtVdrSpecDis)
    Call SetSPPara(adcmdSave, 32, txtVdrRemark)
    Call SetSPPara(adcmdSave, 33, wlSalesmanID)
    Call SetSPPara(adcmdSave, 34, gsUserID)
    Call SetSPPara(adcmdSave, 35, wsGenDte)
    Call SetSPPara(adcmdSave, 36, wsTrnCd)
    
    
    adcmdSave.Execute
    wsNo = GetSPPara(adcmdSave, 37)
    wsCode = GetSPPara(adcmdSave, 38)
    
    cnCon.CommitTrans
    
    If wiAction = AddRec And Trim(wsNo) = "" Then
        gsMsg = "儲存失敗, 請檢查 Store Procedure - V001!"
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

Private Sub txtVdrAddress1_LostFocus()
    FocusMe txtVdrAddress1, True
End Sub

Private Sub txtVdrAddress2_LostFocus()
    FocusMe txtVdrAddress2, True
End Sub

Private Sub txtVdrAddress3_LostFocus()
    FocusMe txtVdrAddress3, True
End Sub

Private Sub txtVdrAddress4_LostFocus()
    FocusMe txtVdrAddress4, True
End Sub

Private Sub txtVdrCode_GotFocus()
    FocusMe txtVdrCode
End Sub

Private Sub txtVdrCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 And wiAction <> AddRec Then
        KeyCode = 0
    End If
End Sub

Private Sub txtVdrCode_KeyPress(KeyAscii As Integer)
    Call chk_InpLenA(txtVdrCode, 10, KeyAscii, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtVdrCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub txtVdrCode_LostFocus()
    FocusMe txtVdrCode, True
End Sub

Private Sub txtVdrContactName_LostFocus()
    FocusMe txtVdrContactName, True
End Sub

Private Sub txtVdrContactName1_GotFocus()
    FocusMe txtVdrContactName1
End Sub

Private Sub txtVdrContactName1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrContactName1, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtVdrAddress1.SetFocus
    End If
End Sub

Private Sub txtVdrContactName1_LostFocus()
    FocusMe txtVdrContactName1, True
End Sub

Private Sub txtVdrCreditLimit_LostFocus()
    txtVdrCreditLimit = Format(txtVdrCreditLimit, gsAmtFmt)
    FocusMe txtVdrCreditLimit, True
End Sub

Private Sub txtVdrEmail_LostFocus()
    FocusMe txtVdrEmail, True
End Sub

Private Sub txtVdrFax_LostFocus()
    FocusMe txtVdrFax, True
End Sub

Private Sub txtVdrName_GotFocus()
    FocusMe txtVdrName
End Sub

Private Sub txtVdrName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrName, 60, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_txtVdrName() = True Then
            txtVdrTel.SetFocus
        End If
    End If
End Sub

Private Sub txtVdrName_LostFocus()
    FocusMe txtVdrName, True
End Sub






Private Sub txtVdrPayTerm_LostFocus()
    FocusMe txtVdrPayTerm
End Sub

Private Sub txtVdrRemark_LostFocus()
    FocusMe txtVdrRemark, True
End Sub

Private Sub txtVdrShipAdd1_LostFocus()
    FocusMe txtVdrShipAdd1, True
End Sub

Private Sub txtVdrShipAdd2_LostFocus()
    FocusMe txtVdrShipAdd2, True
End Sub

Private Sub txtVdrShipAdd3_LostFocus()
    FocusMe txtVdrShipAdd3, True
End Sub

Private Sub txtVdrShipAdd4_LostFocus()
    FocusMe txtVdrShipAdd4, True
End Sub

Private Sub txtVdrShipContactPerson_LostFocus()
    FocusMe txtVdrShipContactPerson, True
End Sub

Private Sub txtVdrShipEmail_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipEmail
End Sub

Private Sub txtVdrShipEmail_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipEmail, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        cboVdrSaleCode.SetFocus
    End If
End Sub

Private Sub txtVdrShipEmail_LostFocus()
    FocusMe txtVdrShipEmail, True
End Sub

Private Sub txtVdrShipFax_LostFocus()
    FocusMe txtVdrShipFax, True
End Sub

Private Sub txtVdrShipName_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipName
End Sub

Private Sub txtVdrShipName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtVdrShipContactPerson.SetFocus
    End If
End Sub

Private Sub txtVdrShipName_LostFocus()
    FocusMe txtVdrShipName, True
End Sub

Private Sub txtVdrShipTel_LostFocus()
    FocusMe txtVdrShipTel, True
End Sub

Private Sub txtVdrSpecDis_GotFocus()
    FocusMe txtVdrSpecDis
End Sub

Private Sub txtVdrSpecDis_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtVdrSpecDis, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_txtVdrSpecDis Then
        tabDetailInfo.Tab = 2
        txtVdrRemark.SetFocus
        End If
    End If
End Sub

Private Sub txtVdrSpecDis_LostFocus()
    txtVdrSpecDis = Format(txtVdrSpecDis, gsAmtFmt)
    FocusMe txtVdrSpecDis, True
End Sub

Private Sub txtVdrTel_GotFocus()
    FocusMe txtVdrTel
End Sub

Private Sub txtVdrTel_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrTel, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtVdrFax.SetFocus
    End If
End Sub

Private Sub txtVdrFax_GotFocus()
    FocusMe txtVdrFax
End Sub

Private Sub txtVdrFax_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrFax, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtVdrContactName.SetFocus
    End If
End Sub

Private Sub txtVdrContactName_GotFocus()
    FocusMe txtVdrContactName
End Sub

Private Sub txtVdrContactName_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrContactName, 30, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtVdrEmail.SetFocus
    End If
End Sub

Private Sub txtVdrEmail_GotFocus()
    FocusMe txtVdrEmail
End Sub

Private Sub txtVdrEmail_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrEmail, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 0
        txtVdrContactName1.SetFocus
    End If
End Sub

Private Sub chkInActive_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        txtVdrName.SetFocus
    End If
End Sub

Private Sub txtVdrAddress1_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtVdrAddress1
End Sub

Private Sub txtVdrAddress1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrAddress1, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 0
        txtVdrAddress2.SetFocus
    End If
End Sub

Private Sub txtVdrAddress2_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtVdrAddress2
End Sub

Private Sub txtVdrAddress2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrAddress2, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 0
        txtVdrAddress3.SetFocus
    End If
End Sub

Private Sub txtVdrAddress3_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtVdrAddress3
End Sub

Private Sub txtVdrAddress3_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrAddress3, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 0
        txtVdrAddress4.SetFocus
    End If
End Sub

Private Sub txtVdrAddress4_GotFocus()
    If tabDetailInfo.Tab <> 0 Then tabDetailInfo.Tab = 0
    FocusMe txtVdrAddress4
End Sub

Private Sub txtVdrAddress4_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrAddress4, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 0
        txtVdrShipName.SetFocus
    End If
End Sub

Private Sub txtVdrTel_LostFocus()
    FocusMe txtVdrTel, True
End Sub

Private Sub txtVdrShipAdd1_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipAdd1
End Sub

Private Sub txtVdrShipAdd1_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipAdd1, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 1
        txtVdrShipAdd2.SetFocus
    End If
End Sub

Private Sub txtVdrShipAdd2_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipAdd2
End Sub

Private Sub txtVdrShipAdd2_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipAdd2, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 1
        txtVdrShipAdd3.SetFocus
    End If
End Sub

Private Sub txtVdrShipAdd3_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipAdd3
End Sub

Private Sub txtVdrShipAdd3_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipAdd3, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 1
        txtVdrShipAdd4.SetFocus
    End If
End Sub

Private Sub txtVdrShipAdd4_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipAdd4
End Sub

Private Sub txtVdrShipAdd4_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipAdd4, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtVdrShipTel.SetFocus
    End If
End Sub

Private Sub txtVdrShipContactPerson_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipContactPerson
End Sub

Private Sub txtVdrShipContactPerson_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipContactPerson, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 1
        txtVdrShipAdd1.SetFocus
    End If
End Sub

Private Sub txtVdrShipTel_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipTel
End Sub

Private Sub txtVdrShipTel_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipTel, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtVdrShipFax.SetFocus
    End If
End Sub

Private Sub txtVdrShipFax_GotFocus()
    If tabDetailInfo.Tab <> 1 Then tabDetailInfo.Tab = 1
    FocusMe txtVdrShipFax
End Sub

Private Sub txtVdrShipFax_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrShipFax, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        txtVdrShipEmail.SetFocus
    End If
End Sub

Private Sub txtVdrPayTerm_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe txtVdrPayTerm
End Sub

Private Sub txtVdrPayTerm_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrPayTerm, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        txtVdrCreditLimit.SetFocus
    End If
End Sub

Private Sub txtVdrCreditLimit_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe txtVdrCreditLimit
End Sub

Private Sub txtVdrCreditLimit_KeyPress(KeyAscii As Integer)
    Call Chk_InpNum(KeyAscii, txtVdrCreditLimit, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        tabDetailInfo.Tab = 2
        cboVdrMLCode.SetFocus
    End If
End Sub

Private Sub txtVdrRemark_GotFocus()
    If tabDetailInfo.Tab <> 2 Then tabDetailInfo.Tab = 2
    FocusMe txtVdrRemark
End Sub

Private Sub txtVdrRemark_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(txtVdrRemark, 100, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        cboVdrRgnCode.SetFocus
    End If
End Sub

Private Sub OpenPromptForm()
    Dim wsOutCode As String
    Dim sSQL As String
    
    ReDim vFilterAry(8, 2)
    vFilterAry(1, 1) = "編碼"
    vFilterAry(1, 2) = "VdrCode"
    
    vFilterAry(2, 1) = "名稱"
    vFilterAry(2, 2) = "VdrName"
    
    vFilterAry(3, 1) = "有效"
    vFilterAry(3, 2) = "VdrInActive"
    
    vFilterAry(4, 1) = "聯絡人"
    vFilterAry(4, 2) = "VdrContactName"
    
    vFilterAry(5, 1) = "電話"
    vFilterAry(5, 2) = "VdrTel"
    
    vFilterAry(6, 1) = "傳真"
    vFilterAry(6, 2) = "VdrFax"
    
    vFilterAry(7, 1) = "電郵"
    vFilterAry(7, 2) = "VdrEmail"
    
    vFilterAry(8, 1) = "地區"
    vFilterAry(8, 2) = "VdrTerritory"
    
    
    ReDim vAry(8, 3)
    vAry(1, 1) = "編碼"
    vAry(1, 2) = "VdrCode"
    vAry(1, 3) = "800"
    
    vAry(2, 1) = "名稱"
    vAry(2, 2) = "VdrName"
    vAry(2, 3) = "2000"
    
    vAry(3, 1) = "有效"
    vAry(3, 2) = "VdrInActive"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "聯絡人"
    vAry(4, 2) = "VdrContactName"
    vAry(4, 3) = "1000"
    
    vAry(5, 1) = "電話"
    vAry(5, 2) = "VdrTel"
    vAry(5, 3) = "1000"
    
    vAry(6, 1) = "傳真"
    vAry(6, 2) = "VdrFax"
    vAry(6, 3) = "0"
    
    vAry(7, 1) = "電郵"
    vAry(7, 2) = "VdrEmail"
    vAry(7, 3) = "1500"
    
    vAry(8, 1) = "地區"
    vAry(8, 2) = "VdrTerritory"
    vAry(8, 3) = "1600"
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        sSQL = "SELECT MstVendor.VdrCode, MstVendor.VdrName, "
        sSQL = sSQL + "MstVendor.VdrContactName, MstVendor.VdrTel, MstVendor.VdrFax, MstVendor.VdrEmail, "
        sSQL = sSQL + "MstVendor.VdrTerritory, MstVendor.VdrInActive "
        sSQL = sSQL + "FROM MstVendor WHERE VdrStatus='1'"
        .sBindSQL = sSQL
        .sBindWhereSQL = ""
        .sBindOrderSQL = "ORDER BY MstVendor.VdrName"
        .vHeadDataAry = vAry
        .vFilterAry = vFilterAry
        .Show vbModal
    End With
    Me.MousePointer = vbNormal
    If Trim(frmShareSearch.Tag) <> "" And Trim(frmShareSearch.Tag) <> cboVdrCode Then
        cboVdrCode = Trim(frmShareSearch.Tag)
        SendKeys "{ENTER}"
    End If
    Unload frmShareSearch
End Sub

Public Function Chk_cboVdrCurr() As Boolean
    Chk_cboVdrCurr = False
    
    If Trim(cboVdrCurr) = "" Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg
        tabDetailInfo.Tab = 2
        cboVdrCurr.SetFocus
        Exit Function
    End If
    
    If Chk_VdrCurr = False Then
        gsMsg = "沒有輸入須要之資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboVdrCurr.SetFocus
        Exit Function
    End If
    
    Chk_cboVdrCurr = True
End Function

Private Function Chk_VdrCurr() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT ExcCurr FROM MstExchangeRate WHERE ExcCurr='" & Set_Quote(cboVdrCurr.Text) + "' And ExcStatus = '1'"

    rsRcd.Open sSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount < 1 Then
        Chk_VdrCurr = False
    Else
        Chk_VdrCurr = True
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

Private Sub cboVdrCode_DropDown()
    
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrCode
    
    wsSQL = "SELECT VdrCode, VdrName FROM MstVendor WHERE VdrStatus = '1'"
    wsSQL = wsSQL & " AND VdrInactive = 'N' "
    wsSQL = wsSQL & " AND VdrCode LIKE '%" & IIf(cboVdrCode.SelLength > 0, "", Set_Quote(cboVdrCode.Text)) & "%' "
   
    wsSQL = wsSQL & "ORDER BY VdrCode "
    Call Ini_Combo(2, wsSQL, cboVdrCode.Left, cboVdrCode.Top + cboVdrCode.Height, tblCommon, wsFormID, "TBLV", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrCode_KeyPress(KeyAscii As Integer)
    Dim wsExcRate As String
    Dim wsExcDesc As String
    
    Call chk_InpLen(cboVdrCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If chk_cboVdrCode() = True Then
            Call Ini_Scr_AfrKey
        End If
    End If
End Sub

Private Sub cboVdrCode_GotFocus()
    FocusMe cboVdrCode
End Sub

Private Sub cboVdrCode_LostFocus()
    FocusMe cboVdrCode, True
End Sub

Private Sub cboVdrCurr_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrCurr
    
    wsSQL = "SELECT DISTINCT ExcCurr FROM MstExchangeRate WHERE ExcStatus = '1'"
    wsSQL = wsSQL & "ORDER BY ExcCurr "
    Call Ini_Combo(1, wsSQL, cboVdrCurr.Left + tabDetailInfo.Left, cboVdrCurr.Top + cboVdrCurr.Height + tabDetailInfo.Top + tbrProcess.Height, tblCommon, wsFormID, "TBLCURR", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrCurr_KeyPress(KeyAscii As Integer)
    Call chk_InpLen(cboVdrCurr, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If Chk_cboVdrCurr() = True Then
            tabDetailInfo.Tab = 2
            txtVdrSpecDis.SetFocus
        End If
    End If
End Sub

Private Sub cboVdrCurr_GotFocus()
    FocusMe cboVdrCurr
End Sub

Private Sub cboVdrCurr_LostFocus()
    FocusMe cboVdrCurr, True
End Sub

Private Sub cboVdrPayCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass

    Set wcCombo = cboVdrPayCode
    
    wsSQL = "SELECT PayCode, PayDesc, PayDay FROM MstPayTerm WHERE PayStatus = '1'"
    wsSQL = wsSQL & "ORDER BY PayCode "
    Call Ini_Combo(3, wsSQL, cboVdrPayCode.Left + tabDetailInfo.Left, cboVdrPayCode.Top + cboVdrPayCode.Height + tabDetailInfo.Top + tbrProcess.Height, tblCommon, wsFormID, "TBLPYT", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboVdrPayCode_KeyPress(KeyAscii As Integer)
    Dim sPayTerm As String
    
    Call chk_InpLen(cboVdrPayCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboVdrPayCode(sPayTerm) = True Then
            If wsOldPayCode <> cboVdrPayCode.Text Then
                txtVdrPayTerm = sPayTerm
                wsOldPayCode = cboVdrPayCode.Text
            End If
            tabDetailInfo.Tab = 2
            txtVdrPayTerm.SetFocus
        End If
    End If
End Sub

Private Sub cboVdrPayCode_GotFocus()
    FocusMe cboVdrPayCode
End Sub

Private Sub cboVdrPayCode_LostFocus()
    FocusMe cboVdrPayCode, True
End Sub

Private Function Chk_KeyExist() As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT VdrStatus FROM MstVendor WHERE VdrCode = '" & Set_Quote(txtVdrCode) & "'"
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
        .TableKey = "VdrCode"
        .KeyLen = 10
        Set .ctlKey = txtVdrCode
        .Show vbModal
    End With
    
    Set Newfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err
    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
    
    fraVdrInfo.Caption = Get_Caption(waScrItm, "FRAVDRINFO")
    
    chkInActive.Caption = Get_Caption(waScrItm, "INACTIVE")
    lblVdrCode.Caption = Get_Caption(waScrItm, "VDRCODE")
    lblVdrName.Caption = Get_Caption(waScrItm, "VDRNAME")
    lblVdrTel.Caption = Get_Caption(waScrItm, "VDRTEL")
    lblVdrFax.Caption = Get_Caption(waScrItm, "VDRFAX")
    lblVdrContactName.Caption = Get_Caption(waScrItm, "VDRCONTACTNAME")
    lblVdrEmail.Caption = Get_Caption(waScrItm, "VDREMAIL")
    lblVdrCreditLimit.Caption = Get_Caption(waScrItm, "VDRCREDITLIMIT")
    lblVdrOpenBal.Caption = Get_Caption(waScrItm, "VDROPENBAL")
    lblVdrCurr.Caption = Get_Caption(waScrItm, "VDRCURR")
    lblVdrSpecDis.Caption = Get_Caption(waScrItm, "VDRSPECDIS")
    lblVdrPayCode.Caption = Get_Caption(waScrItm, "VDRPAYCODE")
    lblVdrRemark.Caption = Get_Caption(waScrItm, "VDRREMARK")
    lblVdrAddress1.Caption = Get_Caption(waScrItm, "VDRADDRESS1")
    lblVdrLastUpd.Caption = Get_Caption(waScrItm, "VDRLASTUPD")
    lblVdrLastUpdDate.Caption = Get_Caption(waScrItm, "VDRLASTUPDDATE")
    lblVdrMLCode.Caption = Get_Caption(waScrItm, "VDRMLCODE")
    lblVdrContactName1.Caption = Get_Caption(waScrItm, "VDRCONTACTPERSON1")
    lblVdrRgnCode.Caption = Get_Caption(waScrItm, "VDRRGNCODE")
    
    tbrProcess.Buttons(tcOpen).ToolTipText = Get_Caption(waScrToolTip, tcOpen) & "(F6)"
    tbrProcess.Buttons(tcAdd).ToolTipText = Get_Caption(waScrToolTip, tcAdd) & "(F2)"
    tbrProcess.Buttons(tcEdit).ToolTipText = Get_Caption(waScrToolTip, tcEdit) & "(F5)"
    tbrProcess.Buttons(tcDelete).ToolTipText = Get_Caption(waScrToolTip, tcDelete) & "(F3)"
    tbrProcess.Buttons(tcSave).ToolTipText = Get_Caption(waScrToolTip, tcSave) & "(F10)"
    tbrProcess.Buttons(tcCancel).ToolTipText = Get_Caption(waScrToolTip, tcCancel) & "(F11)"
    tbrProcess.Buttons(tcFind).ToolTipText = Get_Caption(waScrToolTip, tcFind) & "(F9)"
    tbrProcess.Buttons(tcExit).ToolTipText = Get_Caption(waScrToolTip, tcExit) & "(F12)"
    
    fraVdrShipAddr1.Caption = Get_Caption(waScrItm, "FRAVDRSHIPADDR1")
    
    lblVdrShipName.Caption = Get_Caption(waScrItm, "VDRSHIPNAME")
    lblVdrShipContactPerson.Caption = Get_Caption(waScrItm, "VDRSHIPCONTACTPERSON")
    lblVdrShipAdd.Caption = Get_Caption(waScrItm, "VDRSHIPADD")
    lblVdrSaleName.Caption = Get_Caption(waScrItm, "VDRSALENAME")
    
    lblVdrShipTel.Caption = Get_Caption(waScrItm, "VDRSHIPTEL")
    lblVdrShipFax.Caption = Get_Caption(waScrItm, "VDRSHIPFAX")
    lblVdrShipEmail.Caption = Get_Caption(waScrItm, "VDRSHIPEMAIL")
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO0")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO1")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO2")
    tabDetailInfo.TabCaption(3) = Get_Caption(waScrItm, "TABDETAILINFO3")
    
    lblAcmSale.Caption = Get_Caption(waScrItm, "ACMSALE")
    lblAcmYrSale.Caption = Get_Caption(waScrItm, "ACMYRSALE")
    lblAcmMnSale.Caption = Get_Caption(waScrItm, "ACMMNSALE")
    lblOpenBal.Caption = Get_Caption(waScrItm, "OPENBAL")
    lblCloseBal.Caption = Get_Caption(waScrItm, "CLOSEBAL")
    lblARBal.Caption = Get_Caption(waScrItm, "ARBAL")
    lblQty.Caption = Get_Caption(waScrItm, "QTY")
    lblAmt.Caption = Get_Caption(waScrItm, "AMT")
    lblNet.Caption = Get_Caption(waScrItm, "NET")
    lblVdrCrtDate.Caption = Get_Caption(waScrItm, "VDRCRTDATE")
    
    With tblDetail
        .Columns(PERIOD).Caption = Get_Caption(waScrItm, "PERIOD")
        .Columns(SALES).Caption = Get_Caption(waScrItm, "SALES")
        .Columns(DEPOSIT).Caption = Get_Caption(waScrItm, "DEPOSIT")
    End With
    
    wsActNam(1) = Get_Caption(waScrItm, "VADD")
    wsActNam(2) = Get_Caption(waScrItm, "VEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "VDELETE")
Exit Sub

Ini_Caption_Err:

MsgBox "Please Check ini_Caption!"

End Sub

Private Function Chk_cboVdrMLCode() As Boolean
    Dim wsDesc As String
    Chk_cboVdrMLCode = False
     
    If Trim(cboVdrMLCode.Text) = "" Then
        gsMsg = "必須輸入會計號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboVdrMLCode.SetFocus
        Exit Function
    End If
    
    
    If Chk_MerchClass(cboVdrMLCode, wsDesc) = False Then
        gsMsg = "無此會計號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboVdrMLCode.SetFocus
        lblDspVdrMLDesc = ""
       Exit Function
    End If
    
    lblDspVdrMLDesc = wsDesc
    
    Chk_cboVdrMLCode = True
End Function

Private Function Chk_cboVdrRgnCode() As Boolean
    Dim wsDesc As String
    Chk_cboVdrRgnCode = False
     
    If Trim(cboVdrRgnCode.Text) = "" Then
        gsMsg = "必須輸入採購區域!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboVdrRgnCode.SetFocus
        Exit Function
    End If
    
    If Chk_Region(cboVdrRgnCode, wsDesc) = False Then
        gsMsg = "無此採購區域!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        cboVdrRgnCode.SetFocus
        lblDspVdrRgnDesc = ""
       Exit Function
    End If
    
    lblDspVdrRgnDesc = wsDesc
    
    Chk_cboVdrRgnCode = True
End Function

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

Private Function Chk_cboVdrSaleCode(ByRef sOutName) As Boolean
    Dim sRetName As String
    
    sRetName = ""
    
    Chk_cboVdrSaleCode = False
    
    If Trim(cboVdrSaleCode.Text) <> "" Then
        If Chk_VdrSaleCode(cboVdrSaleCode.Text, wlSalesmanID, sRetName) = False Then
            Me.tabDetailInfo.Tab = 2
            gsMsg = "採購員編碼不存在!"
            MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
            cboVdrSaleCode.SetFocus
            Exit Function
        Else
            sOutName = sRetName
            Chk_cboVdrSaleCode = True
        End If
    Else
        Me.tabDetailInfo.Tab = 2
        gsMsg = "採購員編碼不存在!"
        MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
        cboVdrSaleCode.SetFocus
        Exit Function
    End If
    
End Function

Private Function Chk_VdrSaleCode(ByVal inCode As String, ByRef OutID As Long, ByRef OutName As String) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
        
    Chk_VdrSaleCode = False
        
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
    
    Chk_VdrSaleCode = True
    
    rsRcd.Close
    Set rsRcd = Nothing
    
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
    
    Call Get_VdrSaleInfo(wlKey, wsYYYY, wsMM, 0, 0, wdOpnBal, wdTotBal, wdCMQty, wdCYQty, wdTotQty, wdCMSal, wdCYSal, wdTotSal, wdCMNet, wdCYNet, wdTotNet)
  
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
    
    
    wsSQL = "SELECT POHDCTLPRD, SUM(PODTNETL) NETL "
    wsSQL = wsSQL & " FROM POPPOHD, POPPODT "
    wsSQL = wsSQL & " WHERE POHDVDRID = " & wlKey
    wsSQL = wsSQL & " AND POHDDOCID = PODTDOCID "
    wsSQL = wsSQL & " AND POHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND POHDCTLPRD >= '" & wsYYYY & "01" & "'"
    wsSQL = wsSQL & " GROUP BY POHDCTLPRD "
    wsSQL = wsSQL & " ORDER BY POHDCTLPRD "
    
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
    
    wdAmt = Get_VdrDebitAmt(wlKey, ReadRs(rsRcd, "POHDCTLPRD"))
     
     .AppendRows
        waResult(.UpperBound(1), PERIOD) = ReadRs(rsRcd, "POHDCTLPRD")
        waResult(.UpperBound(1), BALID) = ReadRs(rsRcd, "POHDCTLPRD")
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


