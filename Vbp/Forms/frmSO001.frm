VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSO001 
   Caption         =   "訂貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmSO001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   10920
      OleObjectBlob   =   "frmSO001.frx":030A
      TabIndex        =   41
      Top             =   2520
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
            Picture         =   "frmSO001.frx":2A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":32E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":3BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":4013
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":4465
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":477F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":4BD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":5023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":533D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":5657
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":5AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":6385
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":66AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSO001.frx":69C9
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
            Key             =   "Appendix"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   43
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
      TabPicture(0)   =   "frmSO001.frx":6CE5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboJobNo"
      Tab(0).Control(1)=   "cboCRML"
      Tab(0).Control(2)=   "cboCurr"
      Tab(0).Control(3)=   "txtDisAmt"
      Tab(0).Control(4)=   "btnGetDisAmt"
      Tab(0).Control(5)=   "txtSpecDis"
      Tab(0).Control(6)=   "cboDocNo"
      Tab(0).Control(7)=   "cboCusCode"
      Tab(0).Control(8)=   "cboPayCode"
      Tab(0).Control(9)=   "cboPrcCode"
      Tab(0).Control(10)=   "cboMLCode"
      Tab(0).Control(11)=   "cboSaleCode"
      Tab(0).Control(12)=   "FraDate"
      Tab(0).Control(13)=   "fraInfo"
      Tab(0).Control(14)=   "fraCode"
      Tab(0).Control(15)=   "fraKey"
      Tab(0).Control(16)=   "lblDisAmt"
      Tab(0).Control(17)=   "lblSpecDis"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmSO001.frx":6D01
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblDspNetAmtOrg"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDspGrsAmtOrg"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblDspTotalQty"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblNetAmtOrg"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblGrsAmtOrg"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblTotalQty"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblDisAmtOrg"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblDspDisAmtOrg"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblDspCstAmtOrg"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblCol(9)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblCol(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblCol(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblCol(4)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblCol(6)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblCol(10)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblCol(8)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblCol(7)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lblCol(5)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lblCol(3)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lblCol(0)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "tblDetail"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Frame1"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Item Information"
      TabPicture(2)   =   "frmSO001.frx":6D1D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboShipCode"
      Tab(2).Control(1)=   "fraShip"
      Tab(2).Control(2)=   "cboRmkCode"
      Tab(2).Control(3)=   "fraRmk"
      Tab(2).ControlCount=   4
      Begin VB.ComboBox cboJobNo 
         Height          =   300
         Left            =   -68760
         TabIndex        =   14
         Top             =   4800
         Width           =   2010
      End
      Begin VB.ComboBox cboCRML 
         Height          =   300
         Left            =   -73200
         TabIndex        =   8
         Top             =   3960
         Width           =   2370
      End
      Begin VB.ComboBox cboCurr 
         Height          =   300
         Left            =   -65520
         TabIndex        =   106
         Top             =   1300
         Width           =   1335
      End
      Begin VB.TextBox txtDisAmt 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   -73080
         MaxLength       =   20
         TabIndex        =   12
         Top             =   6000
         Width           =   2055
      End
      Begin VB.CommandButton btnGetDisAmt 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -73080
         Picture         =   "frmSO001.frx":6D39
         TabIndex        =   13
         Top             =   6360
         Width           =   1935
      End
      Begin VB.TextBox txtSpecDis 
         Alignment       =   1  '靠右對齊
         Height          =   288
         Left            =   -73080
         MaxLength       =   20
         TabIndex        =   11
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   450
         Left            =   120
         TabIndex        =   91
         Top             =   7200
         Width           =   6135
         Begin VB.Label lblDeleteLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   4800
            TabIndex        =   95
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblInsertLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   3360
            TabIndex        =   94
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblComboPrompt 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   1920
            TabIndex        =   93
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblKeyDesc 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   360
            TabIndex        =   92
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   23
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
               TabIndex        =   29
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1080
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr3 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   28
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   720
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr2 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   27
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   5865
            End
            Begin VB.TextBox txtShipAdr1 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Left            =   0
               TabIndex        =   26
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   5865
            End
         End
         Begin VB.TextBox txtShipName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   25
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   1080
            Width           =   4305
         End
         Begin VB.TextBox txtShipPer 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1680
            TabIndex        =   24
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
         TabIndex        =   30
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
         Left            =   -73200
         TabIndex        =   2
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox cboPayCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   5
         Top             =   2880
         Width           =   2370
      End
      Begin VB.ComboBox cboPrcCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   6
         Top             =   3240
         Width           =   2370
      End
      Begin VB.ComboBox cboMLCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   7
         Top             =   3600
         Width           =   2370
      End
      Begin VB.ComboBox cboSaleCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   4
         Top             =   2520
         Width           =   2370
      End
      Begin VB.Frame FraDate 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   51
         Top             =   4440
         Width           =   3975
         Begin MSMask.MaskEdBox medReserveDate 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medExpiryDate 
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblExpiryDate 
            Caption         =   "ETADATE"
            Height          =   255
            Left            =   240
            TabIndex        =   53
            Top             =   660
            Width           =   1440
         End
         Begin VB.Label lblReserveDate 
            Caption         =   "ONDATE"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   300
            Width           =   1440
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   3135
         Left            =   -70800
         TabIndex        =   44
         Top             =   4440
         Width           =   7575
         Begin VB.TextBox txtJobCost 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtShipFrom 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   16
            Text            =   "0123456789012345789"
            Top             =   840
            Width           =   5265
         End
         Begin VB.TextBox txtShipVia 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   18
            Text            =   "0123456789012345789"
            Top             =   1440
            Width           =   5265
         End
         Begin VB.TextBox txtShipTo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   17
            Text            =   "0123456789012345789"
            Top             =   1140
            Width           =   5265
         End
         Begin VB.TextBox txtLcNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   20
            Text            =   "0123456789012345789"
            Top             =   2280
            Width           =   5265
         End
         Begin VB.TextBox txtPortNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   21
            Text            =   "0123456789012345789"
            Top             =   2640
            Width           =   5265
         End
         Begin VB.TextBox txtCusPo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   19
            Text            =   "0123456789012345789"
            Top             =   1920
            Width           =   5265
         End
         Begin VB.Label lblJobCost 
            Caption         =   "PORTNO"
            Height          =   240
            Left            =   4200
            TabIndex        =   122
            Top             =   360
            Width           =   1380
         End
         Begin VB.Label lblJobNo 
            Caption         =   "MLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   121
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label lblLcNo 
            Caption         =   "LCNO"
            Height          =   240
            Left            =   120
            TabIndex        =   50
            Top             =   2280
            Width           =   2100
         End
         Begin VB.Label lblPortNo 
            Caption         =   "PORTNO"
            Height          =   240
            Left            =   120
            TabIndex        =   49
            Top             =   2640
            Width           =   2100
         End
         Begin VB.Label lblCusPo 
            Caption         =   "CUSPO"
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   1920
            Width           =   2100
         End
         Begin VB.Label lblShipTo 
            Caption         =   "SHIPTO"
            Height          =   240
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   2100
         End
         Begin VB.Label lblShipVia 
            Caption         =   "SHIPVIA"
            Height          =   240
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Width           =   2100
         End
         Begin VB.Label lblShipFrom 
            Caption         =   "SHIPFROM"
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   2100
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   5895
         Left            =   120
         OleObjectBlob   =   "frmSO001.frx":717B
         TabIndex        =   22
         Top             =   1260
         Width           =   11535
      End
      Begin VB.Frame fraCode 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   60
         Top             =   2280
         Width           =   11655
         Begin VB.Label lblDspCRMLDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   108
            Top             =   1680
            Width           =   7335
         End
         Begin VB.Label lblCRMl 
            Caption         =   "MLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   107
            Top             =   1740
            Width           =   1545
         End
         Begin VB.Label lblMlCode 
            Caption         =   "MLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   68
            Top             =   1380
            Width           =   1545
         End
         Begin VB.Label lblDspMLDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   67
            Top             =   1320
            Width           =   7335
         End
         Begin VB.Label lblPrcCode 
            Caption         =   "PRCCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   66
            Top             =   1020
            Width           =   1545
         End
         Begin VB.Label lblDspPrcDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   65
            Top             =   960
            Width           =   7335
         End
         Begin VB.Label lblPayCode 
            Caption         =   "PAYCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   64
            Top             =   660
            Width           =   1545
         End
         Begin VB.Label lblDspPayDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   63
            Top             =   600
            Width           =   7335
         End
         Begin VB.Label lblSaleCode 
            Caption         =   "SALECODE"
            Height          =   240
            Left            =   120
            TabIndex        =   62
            Top             =   300
            Width           =   1545
         End
         Begin VB.Label lblDspSaleDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   61
            Top             =   240
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
               TabIndex        =   32
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   31
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   33
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   690
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   36
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1740
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   34
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1035
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   35
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1395
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   37
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2085
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   8
               Left            =   0
               TabIndex        =   38
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2430
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   9
               Left            =   0
               TabIndex        =   39
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2775
               Width           =   7545
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   10
               Left            =   0
               TabIndex        =   40
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
         Height          =   2175
         Left            =   -74880
         TabIndex        =   79
         Top             =   120
         Width           =   10935
         Begin VB.TextBox txtExcr 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   9360
            MaxLength       =   20
            TabIndex        =   3
            Top             =   1560
            Width           =   1335
         End
         Begin MSMask.MaskEdBox medDocDate 
            Height          =   285
            Left            =   9360
            TabIndex        =   1
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblRefNo 
            Caption         =   "CUSTEL"
            Height          =   255
            Left            =   6960
            TabIndex        =   105
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblDspRefNo 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   8760
            TabIndex        =   104
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDspCusEMail 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   102
            Top             =   1800
            Width           =   5535
         End
         Begin VB.Label lblCusEMail 
            Caption         =   "CUSNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   1860
            Width           =   1575
         End
         Begin VB.Label lblRevNo 
            Caption         =   "CUSFAX"
            Height          =   255
            Left            =   3720
            TabIndex        =   100
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblDspRevNo 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4680
            TabIndex        =   99
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblCusCode 
            Caption         =   "CUSCODE"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   720
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
            Top             =   840
            Width           =   1680
         End
         Begin VB.Label lblDspCusName 
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
            Top             =   1200
            Width           =   1680
         End
         Begin VB.Label lblExcr 
            Caption         =   "EXCR"
            Height          =   255
            Left            =   7365
            TabIndex        =   85
            Top             =   1560
            Width           =   1800
         End
         Begin VB.Label lblDspCusTel 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   84
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lblCusName 
            Caption         =   "CUSNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblDspCusFax 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   5160
            TabIndex        =   82
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lblCusFax 
            Caption         =   "CUSFAX"
            Height          =   255
            Left            =   3840
            TabIndex        =   81
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblCusTel 
            Caption         =   "CUSTEL"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1440
            Width           =   1575
         End
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   120
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   4605
         TabIndex        =   119
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   7695
         TabIndex        =   118
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   10290
         TabIndex        =   117
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   615
         TabIndex        =   116
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   10
         Left            =   6195
         TabIndex        =   115
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   9285
         TabIndex        =   114
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   6690
         TabIndex        =   113
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   8490
         TabIndex        =   112
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   5400
         TabIndex        =   111
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblCol 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000004&
         BackStyle       =   0  '透明
         BorderStyle     =   1  '單線固定
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   9
         Left            =   4110
         TabIndex        =   110
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblDspCstAmtOrg 
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
         Left            =   240
         TabIndex        =   109
         Top             =   480
         Width           =   2490
      End
      Begin VB.Label lblDisAmt 
         Caption         =   "EXCR"
         Height          =   495
         Left            =   -74640
         TabIndex        =   103
         Top             =   6000
         Width           =   1440
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
         Left            =   6720
         TabIndex        =   98
         Top             =   540
         Width           =   2490
      End
      Begin VB.Label lblDisAmtOrg 
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
         Left            =   6720
         TabIndex        =   97
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label lblSpecDis 
         Caption         =   "EXCR"
         Height          =   255
         Left            =   -74640
         TabIndex        =   96
         Top             =   5640
         Width           =   1440
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
         Left            =   2880
         TabIndex        =   59
         Top             =   240
         Width           =   1275
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
         Left            =   4200
         TabIndex        =   58
         Top             =   240
         Width           =   2475
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
         Left            =   9240
         TabIndex        =   57
         Top             =   240
         Width           =   2475
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
         Left            =   2760
         TabIndex        =   56
         Top             =   540
         Width           =   1410
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
         Left            =   4200
         TabIndex        =   55
         Top             =   540
         Width           =   2490
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
         Left            =   9240
         TabIndex        =   54
         Top             =   540
         Width           =   2490
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
Attribute VB_Name = "frmSO001"
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
Private wsOldCurCd As String
Private wsOldShipCd As String
Private wsOldRmkCd As String
Private wsOldPayCd As String
Private wbReadOnly As Boolean
Private wgsTitle As String
Private wbUpdate As Boolean

Private Const GLINENO = 0
Private Const GDESC1 = 1
Private Const GMORE = 2
Private Const GQTY = 3
Private Const GUOM = 4
Private Const GBOM = 5
Private Const GPRICE = 6
Private Const GDISPER = 7
Private Const GMARKUP = 8
Private Const GAMT = 9
Private Const GNET = 10
Private Const GPRINT = 11
Private Const GUCST = 12
Private Const GCST = 13
Private Const GDRMKID = 14
Private Const GPTJID = 15



Private Const SLINENO = 0
Private Const SLN = 1
Private Const SINDENT = 2
Private Const SITMTYPE = 3
Private Const SITMCODE = 4
Private Const SVDRCODE = 5
Private Const SITMNAME = 6
Private Const SQTY = 7
Private Const SUCST = 8
Private Const SCST = 9
Private Const SUNITPRICE = 10
Private Const SDISPER = 11
Private Const SAMT = 12
Private Const SNET = 13
Private Const SDTID = 14
Private Const SVDRID = 15
Private Const SITMID = 16
Private Const SSOID = 17


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
Private Const tcAppendix = "Appendix"


Private wiOpenDoc As Integer
Private wiAction As Integer
Private wiRevNo As Integer
Private wlCusID As Long
Private wlSaleID As Long
Private wlLineNo As Long

Private wlKey As Long
Private wsActNam(4) As String


Private wsConnTime As String
Private Const wsKeyType = "SOASOHD"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String


Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, GLINENO, GPTJID
    waItem.ReDim 0, -1, SLINENO, SSOID
    
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
    Call SetDateMask(medReserveDate)
    Call SetDateMask(medExpiryDate)
      
    
    wsOldCusNo = ""
    wsOldCurCd = ""
    wsOldShipCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""

    
    wlKey = 0
    wlCusID = 0
    wlSaleID = 0
    wlLineNo = 1
    txtSpecDis.Text = Format("0", gsAmtFmt)
    txtDisAmt.Text = Format("0", gsAmtFmt)
    wbReadOnly = False
    
    wbUpdate = False
    
    wiRevNo = Format(0, "##0")
    tblCommon.Visible = False
    
    lblRevNo.Visible = False
    lblDspRevNo.Visible = False
    lblDspCstAmtOrg.Visible = False
    
    Me.Caption = wsFormCaption
    
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
  
    wsSQL = "SELECT SOHDDOCNO, CUSCODE, SOHDSHIPFROM, CUSNAME, SOHDDOCDATE "
    wsSQL = wsSQL & " FROM SOASOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND SOHDCUSID  = CUSID "
    wsSQL = wsSQL & " AND SOHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND SOHDCTLPRD BETWEEN '" & Str(Val(Left(gsSystemDate, 4)) - 1) + "01" & "' AND '" & Left(gsSystemDate, 4) + "12" & "'"
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO DESC "
    Call Ini_Combo(5, wsSQL, cboDocNo.Left + tabDetailInfo.Left, cboDocNo.Top + cboDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub



Private Sub cboDocNo_LostFocus()
FocusMe cboDocNo, True
End Sub

Private Sub cboDocNo_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLenA(cboDocNo, 15, KeyAscii, True)
    
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
        
        If Chk_SoReady(cboDocNo.Text) = True Then
            gsMsg = "文件已封存(Ready), 現在以唯讀模式開啟!請以密碼解封"
            MsgBox gsMsg, vbOKOnly, gsTitle
            wbReadOnly = True
        End If
        
        
    End If
    
    
    Chk_cboDocNo = True

End Function




Private Sub Ini_Scr_AfrKey()
    
    
    
    If LoadRecord() = False Then
        wiAction = AddRec
        lblDspRevNo.Caption = Format(0, "##0")
      '  txtRevNo.Enabled = False
        medDocDate.Text = Dsp_Date(Now)
        medReserveDate.Text = medDocDate.Text
        medExpiryDate.Text = Format(DateAdd("ww", 6, CDate(medDocDate.Text)), "yyyy/mm/dd")

        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
     '   txtRevNo.Enabled = True
        wsOldCusNo = cboCusCode.Text
        wsOldCurCd = cboCurr.Text
        wsOldShipCd = cboShipCode.Text
        wsOldRmkCd = cboRmkCode.Text
        wsOldPayCd = cboPayCode.Text
        
        
        Call SetButtonStatus("AfrKeyEdit")
    End If
    
     Me.Caption = wsFormCaption & " - " & wsActNam(wiAction)
    
    
    Call SetFieldStatus("AfrKey")
    
    If UCase(cboCurr) = UCase(wsBaseCurCd) Then
            txtExcr.Text = Format("1", gsExrFmt)
            txtExcr.Enabled = False
        Else
            txtExcr.Enabled = True
    End If

    
    tabDetailInfo.Tab = 0
    cboCusCode.SetFocus
        
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
    
    wsSQL = "SELECT SOHDDOCID, SOHDDOCNO, SOHDCUSID, CUSID, CUSCODE, CUSNAME, CUSTEL, CUSFAX, CUSEMAIL, SOPTJDOCLINE,"
    wsSQL = wsSQL & "SOHDDOCDATE, SOHDREVNO, SOHDCURR, SOHDEXCR, SOHDCORRNO, "
    wsSQL = wsSQL & "SOHDRESERVEDATE, SOHDEXPIRYDATE, SOHDPAYCODE, SOHDPRCCODE, SOHDSALEID, SOHDMLCODE, SOHDSPECDIS, SOHDNATURECODE, "
    wsSQL = wsSQL & "SOHDCUSPO, SOHDLCNO, SOHDPORTNO, SOHDSHIPPER, SOHDSHIPFROM, SOHDSHIPTO, SOHDSHIPVIA, SOHDSHIPNAME, "
    wsSQL = wsSQL & "SOHDSHIPCODE, SOHDSHIPADR1,  SOHDSHIPADR2,  SOHDSHIPADR3,  SOHDSHIPADR4, "
    wsSQL = wsSQL & "SOHDRMKCODE, SOHDRMK1,  SOHDRMK2,  SOHDRMK3,  SOHDRMK4, SOHDRMK5, SOHDAPRFLG, "
    wsSQL = wsSQL & "SOHDRMK6,  SOHDRMK7,  SOHDRMK8,  SOHDRMK9, SOHDRMK10, "
    wsSQL = wsSQL & "SOHDGRSAMT , SOHDGRSAMTL, SOHDDISAMT, SOHDDISAMTL, SOHDNETAMT, SOHDNETAMTL, "
    wsSQL = wsSQL & "SOPTJID, SOPTJLNTYPE, SOPTJITEMID, SOPTJDESC1, SOPTJDESC2, SOPTJDESC3, SOPTJDESC4, SOPTJQTY, SOPTJUPRICE, SOPTJUCST, SOPTJDISPER, "
    wsSQL = wsSQL & "SOPTJAMT, SOPTJAMTL, SOPTJDIS, SOPTJDISL, SOPTJNET, SOPTJNETL, SOPTJCST, SOPTJCSTL, SOPTJMARKUP, SOPTJUOM, SOPTJDPRICE, SOPTJDRMKID, SOHDJOBNO, JOBCOST, SOPTJPRTFLG "
    wsSQL = wsSQL & "FROM  SOASOHD, SOASOPTJ, mstCUSTOMER, MSTJOBNO "
    wsSQL = wsSQL & "WHERE SOHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
    wsSQL = wsSQL & "AND SOHDDOCID = SOPTJDOCID "
    wsSQL = wsSQL & "AND SOHDCUSID = CUSID "
    wsSQL = wsSQL & "AND SOHDJOBNO *= JOBNO "
    wsSQL = wsSQL & "ORDER BY SOPTJDOCLINE "
    
    rsInvoice.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount <= 0 Then
        rsInvoice.Close
        Set rsInvoice = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsInvoice, "SOHDDOCID")
    lblDspRevNo.Caption = Format(ReadRs(rsInvoice, "SOHDREVNO"), "##0")
   ' wiRevNo = To_Value(ReadRs(rsInvoice, "SOHDREVNO"))
    medDocDate.Text = ReadRs(rsInvoice, "SOHDDOCDATE")
    wlCusID = ReadRs(rsInvoice, "CUSID")
    cboCusCode.Text = ReadRs(rsInvoice, "CUSCODE")
    lblDspCusName.Caption = ReadRs(rsInvoice, "CUSNAME")
    lblDspCusTel.Caption = ReadRs(rsInvoice, "CUSTEL")
    lblDspCusFax.Caption = ReadRs(rsInvoice, "CUSFAX")
    lblDspCusEMail.Caption = ReadRs(rsInvoice, "CUSEMAIL")
    
    cboCurr.Text = ReadRs(rsInvoice, "SOHDCURR")
    txtExcr.Text = Format(ReadRs(rsInvoice, "SOHDEXCR"), gsExrFmt)
    
    medReserveDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "SOHDReserveDate"))
    medExpiryDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "SOHDExpiryDate"))
    
    wlSaleID = To_Value(ReadRs(rsInvoice, "SOHDSALEID"))
    
    cboPayCode = ReadRs(rsInvoice, "SOHDPAYCODE")
    cboPrcCode = ReadRs(rsInvoice, "SOHDPRCCODE")
    cboMLCode = ReadRs(rsInvoice, "SOHDMLCODE")
    cboCRML = ReadRs(rsInvoice, "SOHDNATURECODE")
    

    cboShipCode = ReadRs(rsInvoice, "SOHDSHIPCODE")
    cboRmkCode = ReadRs(rsInvoice, "SOHDRMKCODE")
    
    txtCusPo = ReadRs(rsInvoice, "SOHDCUSPO")
    txtLcNo = ReadRs(rsInvoice, "SOHDLCNO")
    txtPortNo = ReadRs(rsInvoice, "SOHDPORTNO")
    txtSpecDis.Text = Format(ReadRs(rsInvoice, "SOHDSPECDIS"), gsAmtFmt)
    
    txtShipFrom = ReadRs(rsInvoice, "SOHDSHIPFROM")
    txtShipTo = ReadRs(rsInvoice, "SOHDSHIPTO")
    txtShipVia = ReadRs(rsInvoice, "SOHDSHIPVIA")
    txtShipName = ReadRs(rsInvoice, "SOHDSHIPNAME")
    txtShipPer = ReadRs(rsInvoice, "SOHDSHIPPER")
    txtShipAdr1 = ReadRs(rsInvoice, "SOHDSHIPADR1")
    txtShipAdr2 = ReadRs(rsInvoice, "SOHDSHIPADR2")
    txtShipAdr3 = ReadRs(rsInvoice, "SOHDSHIPADR3")
    txtShipAdr4 = ReadRs(rsInvoice, "SOHDSHIPADR4")
    
    
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsInvoice, "SOHDRMK" & i)
    Next i
    
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    lblDspCRMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboCRML.Text) & "'", "MLDESC")
    
    lblDspGrsAmtOrg.Caption = Format(To_Value(ReadRs(rsInvoice, "SOHDGRSAMT")), gsAmtFmt)
    lblDspDisAmtOrg.Caption = Format(To_Value(ReadRs(rsInvoice, "SOHDDISAMT")), gsAmtFmt)
    lblDspNetAmtOrg.Caption = Format(To_Value(ReadRs(rsInvoice, "SOHDNETAMT")), gsAmtFmt)
    'lblDspCstAmtOrg.Caption = Format(To_Value(ReadRs(rsInvoice, "SOHDCSTAMT")), gsAmtFmt)
    
    txtDisAmt.Text = Format(To_Value(ReadRs(rsInvoice, "SOHDDISAMT")), gsAmtFmt)
    lblDspRefNo.Caption = ReadRs(rsInvoice, "SOHDCORRNO")
    
    
    cboJobNo.Text = ReadRs(rsInvoice, "SOHDJOBNO")
    txtJobCost.Text = Format(To_Value(ReadRs(rsInvoice, "JOBCOST")), gsAmtFmt)
    
    
    wlLineNo = 1
    rsInvoice.MoveFirst
    With waResult
         .ReDim 0, -1, GLINENO, GPTJID
         Do While Not rsInvoice.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), GLINENO) = wlLineNo
           '  waResult(.UpperBound(1), GLNTYPE) = ReadRs(rsInvoice, "SOPTJLNTYPE")
           '  waResult(.UpperBound(1), GITMCODE) = IIf(To_Value(ReadRs(rsInvoice, "SOPTJITEMID")) = 0, "", Get_TableInfo("MSTITEM", "ITMID = " & To_Value(ReadRs(rsInvoice, "SOPTJITEMID")), "ITMCODE"))
             waResult(.UpperBound(1), GDESC1) = ReadRs(rsInvoice, "SOPTJDESC1")
             waResult(.UpperBound(1), GQTY) = Format(ReadRs(rsInvoice, "SOPTJQTY"), gsQtyFmt)
             waResult(.UpperBound(1), GPRICE) = Format(ReadRs(rsInvoice, "SOPTJDPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), GDISPER) = Format(ReadRs(rsInvoice, "SOPTJDISPER"), gsAmtFmt)
             waResult(.UpperBound(1), GMARKUP) = Format(ReadRs(rsInvoice, "SOPTJMARKUP"), gsAmtFmt)
             waResult(.UpperBound(1), GUOM) = ReadRs(rsInvoice, "SOPTJUOM")
             
             waResult(.UpperBound(1), GUCST) = Format(ReadRs(rsInvoice, "SOPTJUCST"), gsAmtFmt)
             waResult(.UpperBound(1), GAMT) = Format(ReadRs(rsInvoice, "SOPTJUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), GNET) = Format(ReadRs(rsInvoice, "SOPTJNET"), gsAmtFmt)
             waResult(.UpperBound(1), GCST) = Format(ReadRs(rsInvoice, "SOPTJCST"), gsAmtFmt)
             
             
             If LoadQTItem(wlKey, To_Value(ReadRs(rsInvoice, "SOPTJID")), wlLineNo) = True Then
             waResult(.UpperBound(1), GBOM) = "Y"
             Else
             waResult(.UpperBound(1), GBOM) = "N"
             End If
             waResult(.UpperBound(1), GDRMKID) = To_Value(ReadRs(rsInvoice, "SOPTJDRMKID"))
             waResult(.UpperBound(1), GMORE) = IIf(To_Value(ReadRs(rsInvoice, "SOPTJDRMKID")) <> 0, "Y", "N")
             waResult(.UpperBound(1), GPTJID) = To_Value(ReadRs(rsInvoice, "SOPTJID"))
             
             waResult(.UpperBound(1), GPRINT) = IIf(ReadRs(rsInvoice, "SOPTJPRTFLG") = "Y", "-1", "0")
             
             wlLineNo = wlLineNo + 1
             rsInvoice.MoveNext
         Loop
         'wlLineNo = waResult(.UpperBound(1), GLINENO) + 1

    End With
    tblDetail.ReBind
    tblDetail.FirstRow = 0
    rsInvoice.Close
    
    Set rsInvoice = Nothing
    
   
    LoadRecord = True
    
End Function

Private Sub Ini_Caption()

On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP_M", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
    lblRevNo.Caption = Get_Caption(waScrItm, "REVNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblCusCode.Caption = Get_Caption(waScrItm, "CUSCODE")
    lblCusName.Caption = Get_Caption(waScrItm, "CUSNAME")
    lblCusTel.Caption = Get_Caption(waScrItm, "CUSTEL")
    lblCusFax.Caption = Get_Caption(waScrItm, "CUSFAX")
    lblCusEMail.Caption = Get_Caption(waScrItm, "CUSEMAIL")
    
    lblCurr.Caption = Get_Caption(waScrItm, "CURR")
    lblExcr.Caption = Get_Caption(waScrItm, "EXCR")
    
    lblSaleCode.Caption = Get_Caption(waScrItm, "SALECODE")
    lblPayCode.Caption = Get_Caption(waScrItm, "PAYCODE")
    lblPrcCode.Caption = Get_Caption(waScrItm, "PRCCODE")
    lblMlCode.Caption = Get_Caption(waScrItm, "MLCODE")
    lblCRMl.Caption = Get_Caption(waScrItm, "CRML")
    

    lblReserveDate.Caption = Get_Caption(waScrItm, "ReserveDate")
    lblExpiryDate.Caption = Get_Caption(waScrItm, "ExpiryDate")
    lblSpecDis.Caption = Get_Caption(waScrItm, "SPECDIS")
    lblDisAmt.Caption = Get_Caption(waScrItm, "DISAMTORG")
    
    lblGrsAmtOrg.Caption = Get_Caption(waScrItm, "GRSAMTORG")
    lblDisAmtOrg.Caption = Get_Caption(waScrItm, "DISAMTORG")
    lblNetAmtOrg.Caption = Get_Caption(waScrItm, "NETAMTORG")
  '  lblCstAmtOrg.Caption = Get_Caption(waScrItm, "CSTAMTORG")
    
    lblTotalQty.Caption = Get_Caption(waScrItm, "TOTALQTY")
    
   
    lblCol(0).Caption = Get_Caption(waScrItm, "GLINENO")
    lblCol(1).Caption = Get_Caption(waScrItm, "GMARKUP")
    lblCol(2).Caption = Get_Caption(waScrItm, "GUOM")
    lblCol(3).Caption = Get_Caption(waScrItm, "GQTY")
    lblCol(4).Caption = Get_Caption(waScrItm, "GPRICE")
    lblCol(5).Caption = Get_Caption(waScrItm, "GDISPER")
    lblCol(6).Caption = Get_Caption(waScrItm, "GAMT")
    lblCol(7).Caption = Get_Caption(waScrItm, "GNET")
    lblCol(8).Caption = Get_Caption(waScrItm, "GUCST")
    lblCol(9).Caption = Get_Caption(waScrItm, "GMORE")
    lblCol(10).Caption = Get_Caption(waScrItm, "GBOM")
    
    
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
    btnGetDisAmt.Caption = Get_Caption(waScrItm, "GETDISAMT")
    lblRefNo.Caption = Get_Caption(waScrItm, "REFNO")
    
    lblJobNo.Caption = Get_Caption(waScrItm, "JOBNO")
    lblJobCost.Caption = Get_Caption(waScrItm, "JOBCOST")
    
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
    
    lblKeyDesc = Get_Caption(waScrToolTip, "KEYDESC")
    lblComboPrompt = Get_Caption(waScrToolTip, "COMBOPROMPT")
    lblInsertLine = Get_Caption(waScrToolTip, "INSERTLINE")
    lblDeleteLine = Get_Caption(waScrToolTip, "DELETELINE")
    
    wsActNam(1) = Get_Caption(waScrItm, "SOADD")
    wsActNam(2) = Get_Caption(waScrItm, "SOEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "SODELETE")
    wgsTitle = Get_Caption(waScrItm, "TITLE")
    
    Call Ini_PopMenu(mnuPopUpSub, "POPUP_T", waPopUpSub)
    
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
    Set waItem = Nothing
    Set waScrToolTip = Nothing
    Set waScrItm = Nothing
    Set waPopUpSub = Nothing
'    Set waPgmItm = Nothing
    Set frmSO001 = Nothing

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


Private Sub medReserveDate_GotFocus()
    
  FocusMe medReserveDate
    
End Sub


Private Sub medReserveDate_LostFocus()

    FocusMe medReserveDate, True
    
End Sub


Private Sub medReserveDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medReserveDate Then
        tabDetailInfo.Tab = 0
        medExpiryDate.SetFocus
        End If
    End If
End Sub

Private Function Chk_medReserveDate() As Boolean

    
    Chk_medReserveDate = False
    
    If Trim(medReserveDate.Text) = "/  /" Then
        Chk_medReserveDate = True
        Exit Function
    End If
    
    If Chk_Date(medReserveDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medReserveDate.SetFocus
        Exit Function
    End If
    
    
    Chk_medReserveDate = True

End Function


Private Sub medExpiryDate_GotFocus()
    
  FocusMe medExpiryDate
    
End Sub


Private Sub medExpiryDate_LostFocus()

    FocusMe medExpiryDate, True
    
End Sub


Private Sub medExpiryDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Chk_medExpiryDate Then
            tabDetailInfo.Tab = 0
                txtSpecDis.SetFocus
            
        End If
    End If
End Sub

Private Function Chk_medExpiryDate() As Boolean

    
    Chk_medExpiryDate = False
    
    If Trim(medExpiryDate.Text) = "/  /" Then
        Chk_medExpiryDate = True
        Exit Function
    End If
    
    If Chk_Date(medExpiryDate) = False Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medExpiryDate.SetFocus
        Exit Function
    End If
    
    
    Chk_medExpiryDate = True

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
          Case GDESC1
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
              Case GDESC1
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
    
    Dim rsSOASOHD As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT SOHDSTATUS FROM SOASOHD WHERE SOHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rsSOASOHD.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsSOASOHD.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rsSOASOHD.Close
    Set rsSOASOHD = Nothing
    

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
    Dim wlRowCtr As Long
    Dim wsCtlPrd As String
    Dim wiSLN As Integer
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
       If wiAction = RevRec Then wiAction = CorRec
       Exit Function
    End If
    
    
 '   If wiAction = CorRec Then
 '   If DelValidation(wlKey) = False Then
 '      MousePointer = vbDefault
 '      Exit Function
 '   End If
 '   End If
    
    
    
    If wiAction = AddRec Then
        If Chk_KeyExist() = True Then
            Call GetNewKey
        End If
    End If
    
   
     
    
'    If lblDspNetAmtOrg.Caption > Get_CreditLimit(wlCusID, Trim(medDocDate.Text)) Then
'       gsMsg = "已超過信貸額!"
'       MsgBox gsMsg, vbOKOnly, gsTitle
'       MousePointer = vbDefault
'       Exit Function
 '   End If
    
    
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    cnCon.BeginTrans
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SO001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlCusID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, lblDspRevNo.Caption)
    Call SetSPPara(adcmdSave, 8, cboCurr.Text)
    Call SetSPPara(adcmdSave, 9, txtExcr.Text)
    Call SetSPPara(adcmdSave, 10, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medReserveDate.Text))
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medExpiryDate.Text))
    
    Call SetSPPara(adcmdSave, 13, wlSaleID)
    
    Call SetSPPara(adcmdSave, 14, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 15, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 16, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 17, cboCRML.Text)
    Call SetSPPara(adcmdSave, 18, cboShipCode.Text)
    Call SetSPPara(adcmdSave, 19, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 20, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 21, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 22, txtPortNo.Text)
    
    Call SetSPPara(adcmdSave, 23, txtShipFrom.Text)
    Call SetSPPara(adcmdSave, 24, txtShipTo.Text)
    Call SetSPPara(adcmdSave, 25, txtShipVia.Text)
    Call SetSPPara(adcmdSave, 26, txtShipPer.Text)
    Call SetSPPara(adcmdSave, 27, txtShipName.Text)
    Call SetSPPara(adcmdSave, 28, txtShipAdr1.Text)
    Call SetSPPara(adcmdSave, 29, txtShipAdr2.Text)
    Call SetSPPara(adcmdSave, 30, txtShipAdr3.Text)
    Call SetSPPara(adcmdSave, 31, txtShipAdr4.Text)
    
    For i = 1 To 10
    Call SetSPPara(adcmdSave, 32 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdSave, 42, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 43, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 44, lblDspNetAmtOrg)
    Call SetSPPara(adcmdSave, 45, txtSpecDis.Text)
    Call SetSPPara(adcmdSave, 46, cboJobNo.Text)
    Call SetSPPara(adcmdSave, 47, txtJobCost.Text)
    
    Call SetSPPara(adcmdSave, 48, wsFormID)
    
    Call SetSPPara(adcmdSave, 49, gsUserID)
    Call SetSPPara(adcmdSave, 50, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 51)
    wsDocNo = GetSPPara(adcmdSave, 52)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_SO001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, GQTY)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, GLINENO))
                Call SetSPPara(adcmdSave, 4, "D")
                Call SetSPPara(adcmdSave, 5, "")
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, GDESC1))
                Call SetSPPara(adcmdSave, 7, "")
                Call SetSPPara(adcmdSave, 8, "")
                Call SetSPPara(adcmdSave, 9, "")
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, GQTY))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, GPRICE))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, GAMT))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, GDISPER))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, GMARKUP))
                Call SetSPPara(adcmdSave, 15, waResult(wiCtr, GUOM))
                
                Call SetSPPara(adcmdSave, 16, waResult(wiCtr, GNET))
                Call SetSPPara(adcmdSave, 17, waResult(wiCtr, GNET))
                Call SetSPPara(adcmdSave, 18, waResult(wiCtr, GUCST))
                Call SetSPPara(adcmdSave, 19, waResult(wiCtr, GCST))
                Call SetSPPara(adcmdSave, 20, waResult(wiCtr, GDRMKID))
                Call SetSPPara(adcmdSave, 21, waResult(wiCtr, GPTJID))
                Call SetSPPara(adcmdSave, 22, IIf(waResult(wiCtr, GPRINT) = "-1", "Y", "N"))
                Call SetSPPara(adcmdSave, 23, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 24, gsUserID)
                Call SetSPPara(adcmdSave, 25, wsGenDte)
                adcmdSave.Execute
            End If
        Next
    End If
      
    wiSLN = 0
    For wiCtr = 0 To waItem.UpperBound(1)
          If Trim(waItem(wiCtr, SLN)) <> "0" And Trim(waItem(wiCtr, SITMCODE)) <> "" And Trim(waItem(wiCtr, SITMTYPE)) <> "" Then
                    wiSLN = wiSLN + 1
          End If
    Next
      
    i = 1
    
    If waItem.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_SO001C"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
        
        For wiCtr = 0 To waItem.UpperBound(1)
            If Trim(waItem(wiCtr, SLN)) <> "0" And Trim(waItem(wiCtr, SITMCODE)) <> "" And Trim(waItem(wiCtr, SITMTYPE)) <> "" Then
            
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waItem(wiCtr, SLN))
                Call SetSPPara(adcmdSave, 4, waItem(wiCtr, SITMID))
                Call SetSPPara(adcmdSave, 5, i)
                Call SetSPPara(adcmdSave, 6, waItem(wiCtr, SVDRID))
                Call SetSPPara(adcmdSave, 7, waItem(wiCtr, SITMNAME))
                Call SetSPPara(adcmdSave, 8, waItem(wiCtr, SQTY))
                Call SetSPPara(adcmdSave, 9, waItem(wiCtr, SUNITPRICE))
                Call SetSPPara(adcmdSave, 10, waItem(wiCtr, SUCST))
                Call SetSPPara(adcmdSave, 11, waItem(wiCtr, SDISPER))
                Call SetSPPara(adcmdSave, 12, waItem(wiCtr, SAMT))
                Call SetSPPara(adcmdSave, 13, waItem(wiCtr, SNET))
                Call SetSPPara(adcmdSave, 14, waItem(wiCtr, SCST))
                Call SetSPPara(adcmdSave, 15, IIf(waItem(wiCtr, SINDENT) = "-1", "Y", "N"))
                Call SetSPPara(adcmdSave, 16, IIf(wiSLN = i, "Y", "N"))
                Call SetSPPara(adcmdSave, 17, waItem(wiCtr, SDTID))

                
                
                adcmdSave.Execute
                i = i + 1
            
            End If
        Next
    End If
    
    
  '  adcmdSave.CommandText = "USP_UPDJOB"
  '  adcmdSave.CommandType = adCmdStoredProc
  '  adcmdSave.Parameters.Refresh
      
  '  Call SetSPPara(adcmdSave, 1, wiAction)
  '  Call SetSPPara(adcmdSave, 2, cboJobNo.Text)
  '  Call SetSPPara(adcmdSave, 3, 0)
  '  Call SetSPPara(adcmdSave, 4, txtJobCost.Text)
  '  Call SetSPPara(adcmdSave, 5, gsUserID)
  '  Call SetSPPara(adcmdSave, 6, wsGenDte)
  '  adcmdSave.Execute
    
    
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
    
    
    
 '   If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboCusCode() Then Exit Function
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    If Not Chk_cboPayCode Then Exit Function
    If Not Chk_cboPrcCode Then Exit Function
    If Not Chk_cboMLCode Then Exit Function
    If Not Chk_cboCRML Then Exit Function
    
    

    If Not Chk_medReserveDate Then Exit Function
    If Not Chk_medExpiryDate Then Exit Function
    If Not Chk_txtSpecDis Then Exit Function
    If Not chk_txtDisAmt Then Exit Function
    
    If Not Chk_cboJobNo Then Exit Function
    If Not chk_txtJobCost Then Exit Function
    
    If Not Chk_cboShipCode Then Exit Function
    If Not Chk_cboRmkCode Then Exit Function
    
    
    Dim wiEmptyGrid As Boolean
    Dim wlCtr As Long
    
    wiEmptyGrid = True
    With waResult
        For wlCtr = 0 To .UpperBound(1)
            If Trim(waResult(wlCtr, GDESC1)) <> "" Then
                wiEmptyGrid = False
                If Chk_GrdRow(wlCtr) = False Then
                    tabDetailInfo.Tab = 1
                    tblDetail.Col = GDESC1
                    tblDetail.SetFocus
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
            tblDetail.Col = GQTY
            tblDetail.SetFocus
        End If
        Exit Function
    End If
    
    
    If ChkSoDetail = False Then
        tabDetailInfo.Tab = 1
        tblDetail.Col = GDESC1
        tblDetail.SetFocus
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

    Dim newForm As New frmSO001
    
    newForm.Top = Me.Top + 200
    newForm.Left = Me.Left + 200
    
    newForm.Show

End Sub

Private Sub cmdOpen()

    Dim newForm As New frmSO001
    
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
    wsFormID = "SO001"
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "SO"
    
    


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
Dim wsPrtDocNo As String

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
        
          If MsgBox("你是否確定儲存現時之變更而列印?", vbYesNo, gsTitle) = vbYes Then
                wsPrtDocNo = cboDocNo.Text
                If cmdSave = False Then Exit Sub
                cboDocNo.Text = wsPrtDocNo
                Call Ini_Scr_AfrKey
           End If
        
            Call cmdPrint
        Case tcRevise
            Call cmdRevise
        Case tcExit
            Unload Me
        Case tcAppendix
            Call cmdAppendix
            
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


Private Sub cboCusCode_DropDown()
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboCusCode
    
    wsSQL = "SELECT CUSCODE, CUSNAME FROM mstCUSTOMER "
    wsSQL = wsSQL & "WHERE CUSCODE LIKE '%" & IIf(cboCusCode.SelLength > 0, "", Set_Quote(cboCusCode.Text)) & "%' "
    wsSQL = wsSQL & "AND CUSSTATUS = '1' "
    wsSQL = wsSQL & " AND CusInactive = 'N' "
    wsSQL = wsSQL & "ORDER BY CUSCODE "
    Call Ini_Combo(2, wsSQL, cboCusCode.Left + tabDetailInfo.Left, cboCusCode.Top + cboCusCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLCUSNO", Me.Width, Me.Height)
    
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
    Dim wsEMail As String
    
    chk_cboCusCode = False
    
    
    If Trim(cboCusCode) = "" Then
        gsMsg = "必需輸入客戶編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboCusCode.SetFocus
        Exit Function
    End If
    
    If Chk_CusCode(cboCusCode, wlID, wsName, wsTel, wsFax, wsEMail) Then
        wlCusID = wlID
        lblDspCusName.Caption = wsName
        lblDspCusTel.Caption = wsTel
        lblDspCusFax.Caption = wsFax
        lblDspCusEMail.Caption = wsEMail
    Else
        wlCusID = 0
        lblDspCusName.Caption = ""
        lblDspCusTel.Caption = ""
        lblDspCusFax.Caption = ""
        lblDspCusEMail.Caption = ""
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
        cboCurr.Text = ReadRs(rsDefVal, "CUSCURR")
        cboPayCode.Text = ReadRs(rsDefVal, "CUSPAYCODE")
        cboMLCode.Text = ReadRs(rsDefVal, "CUSMLCODE")
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
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    


End Sub



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
        .ColumnHeaders = False
        For wiCtr = GLINENO To GPTJID
            .Columns(wiCtr).AllowSizing = True
            .Columns(wiCtr).Visible = True
            .Columns(wiCtr).Locked = False
            .Columns(wiCtr).Button = False
            .Columns(wiCtr).Alignment = dbgLeft
            .Columns(wiCtr).HeadAlignment = dbgLeft
            
            Select Case wiCtr
                Case GLINENO
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Locked = True
            '    Case GLNTYPE
            '        .Columns(wiCtr).Width = 1000
            '        .Columns(wiCtr).DataWidth = 1
            '        .Columns(wiCtr).Button = True
            '    Case GITMCODE
            '        .Columns(wiCtr).Width = 1500
            '        .Columns(wiCtr).DataWidth = 30
            '        .Columns(wiCtr).Button = True
                Case GQTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                Case GPRICE
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Locked = True
                Case GDISPER
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                 Case GMARKUP
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case GUOM
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Button = True
                    
                Case GAMT
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                   ' .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case GNET
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                Case GUCST
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Visible = False
                Case GCST
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).NumberFormat = gsAmtFmt
                    .Columns(wiCtr).Visible = False
                Case GBOM
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 2
                    .Columns(wiCtr).Button = True

                Case GDESC1
                    .Columns(wiCtr).Width = 3500
                    .Columns(wiCtr).DataWidth = 60
                Case GMORE
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 2
                    .Columns(wiCtr).Button = True
                Case GDRMKID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 10
                Case GPTJID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 10
                Case GPRINT
                    .Columns(wiCtr).DataWidth = 1
                    .Columns(wiCtr).Width = 300
                    .Columns(wiCtr).ValueItems.Presentation = dbgCheckBox
                    
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


    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
            
             
           Case GDESC1
            
                If Chk_grdDesc(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
            If .Columns(GLINENO).Text = "" Then
                .Columns(GLINENO).Text = wlLineNo
                 wlLineNo = wlLineNo + 1
                 .Columns(GUCST).Text = "0"
                 .Columns(GCST).Text = "0"
                 .Columns(GDISPER).Text = "1"
                 .Columns(GPRICE).Text = "0"
                 .Columns(GAMT).Text = "0"
                 .Columns(GNET).Text = "0"
                 .Columns(GMARKUP).Text = "1"
                 .Columns(GUOM).Text = ""
                 .Columns(GBOM).Text = "N"
                 .Columns(GMORE).Text = "N"
                 .Columns(GQTY).Text = "1"

            End If
            
                
                
           Case GQTY, GPRICE, GDISPER, GMARKUP
                
               If ColIndex = GQTY Then
                
                If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
               
               End If
               
              ' If ColIndex = GPRICE Then
               
              '  If Chk_grdPrice(.Columns(ColIndex).Text) = False Then
              '      GoTo Tbl_BeforeColUpdate_Err
              '  End If
                
              ' End If
                    
               If ColIndex = GDISPER Then
               
                If Chk_grdDisPer(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
               End If
               
               If ColIndex = GMARKUP Then
               
                If Chk_grdMarkUp(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
               End If
                
                
                If Trim(.Columns(GPRICE).Text) <> "" And Trim(.Columns(GQTY).Text) <> "" Then
                .Columns(GAMT).Text = Format(To_Value(.Columns(GPRICE).Text) * To_Value(.Columns(GDISPER).Text) / To_Value(.Columns(GMARKUP).Text), gsAmtFmt)
                End If
                
                If Trim(.Columns(GPRICE).Text) <> "" And Trim(.Columns(GDISPER).Text) <> "" And Trim(.Columns(GMARKUP).Text) <> "" And Trim(.Columns(GQTY).Text) <> "" Then
                .Columns(GNET).Text = Format(To_Value(.Columns(GPRICE).Text) * To_Value(.Columns(GQTY).Text) * To_Value(.Columns(GDISPER).Text) / To_Value(.Columns(GMARKUP).Text), gsAmtFmt)
                End If
                
                If Trim(.Columns(GUCST).Text) <> "" And Trim(.Columns(GQTY).Text) <> "" Then
                .Columns(GCST).Text = Format(To_Value(.Columns(GUCST).Text) * To_Value(.Columns(GQTY).Text), gsAmtFmt)
                End If
                
                
                Case GAMT
                
                 If Chk_grdPrice(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                 End If
                
                If Trim(.Columns(GAMT).Text) <> "" And Trim(.Columns(GQTY).Text) <> "" Then
                .Columns(GNET).Text = Format(To_Value(.Columns(GAMT).Text) * To_Value(.Columns(GQTY).Text), gsAmtFmt)
                End If
 
              
            End Select
            
             If .Columns(ColIndex).Text <> OldValue Then
                wbUpdate = True
             End If
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
    Dim wiCtr As Long
    Dim wbRtnResult As Boolean
    Dim wsRmkID As String
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
           
           Case GUOM
                
                wsSQL = "SELECT UOMCODE FROM MSTUOM "
                wsSQL = wsSQL & " WHERE UOMSTATUS <> '2'"
                
                Call Ini_Combo(1, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top, .Top + .RowTop(.Row) + .RowHeight, tblCommon, wsFormID, "TBLUOMCODE", Me.Width, Me.Height)
                tblCommon.Visible = True
                tblCommon.SetFocus
                Set wcCombo = tblDetail
            
                
            Case GBOM
                
                If wiAction = DelRec Or Trim(.Columns(GLINENO).Text) = "" Then Exit Sub
                    
                
                If waItem.UpperBound(1) >= 0 Then
                    frmQTN002.InvDoc.ReDim 0, waItem.UpperBound(1), SLINENO, SSOID
                End If
                    
                    wiCtr = IIf(.Columns(GLINENO).Text = "", wlLineNo, .Columns(GLINENO).Text)
                    frmQTN002.InTrnCd = "SO"
                    frmQTN002.inLineNo = wiCtr
                    frmQTN002.inLineDesc = .Columns(GDESC1).Text
                    frmQTN002.InvDoc = waItem
                    frmQTN002.InCurr = cboCurr.Text
                    frmQTN002.InExcr = txtExcr.Text
                    frmQTN002.Show vbModal
                    waItem.ReDim 0, frmQTN002.InvDoc.UpperBound(1), SLINENO, SSOID
                    Set waItem = frmQTN002.InvDoc
                    wbRtnResult = frmQTN002.Result
                    Unload frmQTN002
                    
                    If wbRtnResult = True Then
                    Call cmdCstRefresh(wiCtr)
                    End If
                    
            Case GMORE
                
                 
                    frmDocRemark.RmkID = IIf(.Columns(GDRMKID).Text = "", "0", .Columns(GDRMKID).Text)
                    frmDocRemark.RmkType = "ST"
                    frmDocRemark.Show vbModal
                    wsRmkID = frmDocRemark.RmkID
                    Unload frmDocRemark
                    
                    
                    Call cmdRmkID(.Bookmark, wsRmkID)
                    
               
               
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
            If Chk_SoExistSp(.Bookmark) = False Then Exit Sub
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
                Case GNET
                    KeyCode = vbKeyDown
                    .Col = GDESC1
                Case GBOM
                     KeyCode = vbDefault
                    .Col = GPRICE
                Case Else
                    KeyCode = vbDefault
                    .Col = .Col + 1
            End Select
            
        Case vbKeyLeft
            KeyCode = vbDefault
            If .Col <> GLINENO Then
                    .Col = .Col - 1
            End If
            
        Case vbKeyRight
            KeyCode = vbDefault
            If .Col <> GNET Then
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
        
        Case GQTY
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
        Case GPRICE, GDISPER, GMARKUP, GAMT
            Call Chk_InpNum(KeyAscii, tblDetail.Text, False, True)
            
       
    End Select

End Sub

Private Sub tblDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    wbErr = False
    On Error GoTo RowColChange_Err
    
    If ActiveControl.Name <> tblDetail.Name Then Exit Sub
    
    With tblDetail
        If IsEmptyRow() Then
           .Col = GDESC1
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case GDESC1
                    
                    Call Chk_grdDesc(.Columns(GDESC1).Text)
                
                Case GQTY
                    
                    Call Chk_grdQty(.Columns(GQTY).Text)
                    
                Case GPRICE
                    
                 '   Call Chk_grdPrice(.Columns(GPRICE).Text)
            
                
                Case GDISPER
                    
                    Call Chk_grdPrice(.Columns(GDISPER).Text)
                    
                 Case GMARKUP
                    
                    Call Chk_grdMarkUp(.Columns(GMARKUP).Text)
                    
            End Select
        End If
    End With
        
    Exit Sub

RowColChange_Err:
    
    MsgBox "Check tblDeiail RowColChange"
    wbErr = True
    
End Sub


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


Private Function Chk_grdDesc(inCode As String) As Boolean
    
    Chk_grdDesc = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入內容!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdDesc = False
        Exit Function
    End If

       
    
End Function
Private Function Chk_grdPrice(inCode As String) As Boolean
    
    Chk_grdPrice = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入售價!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdPrice = False
        Exit Function
    End If

    If To_Value(inCode) < 0 Then
        gsMsg = "售價必需大為零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdPrice = False
        Exit Function
    End If
    
End Function

Private Function Chk_grdDisPer(inCode As String) As Boolean
    
    Chk_grdDisPer = True
    

    If To_Value(inCode) < 0 Then
        gsMsg = "折扣必需大為零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdDisPer = False
        Exit Function
    End If
    
End Function


Private Function Chk_grdMarkUp(inCode As String) As Boolean
    
    Chk_grdMarkUp = True
    

    If To_Value(inCode) < 0 Then
        gsMsg = "M/U必需大為零!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdMarkUp = False
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
                If Trim(.Columns(GDESC1)) = "" Then
                    Exit Function
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, GLINENO)) = "" And _
                   Trim(waResult(inRow, GDESC1)) = "" And _
                   Trim(waResult(inRow, GMORE)) = "" And _
                   Trim(waResult(inRow, GQTY)) = "" And _
                   Trim(waResult(inRow, GPRICE)) = "" And _
                   Trim(waResult(inRow, GDISPER)) = "" And _
                   Trim(waResult(inRow, GMARKUP)) = "" And _
                   Trim(waResult(inRow, GUOM)) = "" And _
                   Trim(waResult(inRow, GUCST)) = "" And _
                   Trim(waResult(inRow, GCST)) = "" And _
                   Trim(waResult(inRow, GAMT)) = "" And _
                   Trim(waResult(inRow, GNET)) = "" Then
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
        
        
          
        
        If Chk_grdDesc(waResult(LastRow, GDESC1)) = False Then
                .Col = GDESC1
                Exit Function
        End If
        
        If Chk_grdQty(waResult(LastRow, GQTY)) = False Then
                .Col = GQTY
                Exit Function
        End If
        
      '  If Chk_grdPrice(waResult(LastRow, GPRICE)) = False Then
      '          .Col = GPRICE
      '          Exit Function
      '  End If
        
        If Chk_grdDisPer(waResult(LastRow, GDISPER)) = False Then
                .Col = GDISPER
                Exit Function
        End If
        
        If Chk_grdMarkUp(waResult(LastRow, GMARKUP)) = False Then
                .Col = GMARKUP
                Exit Function
        End If
        
        If Chk_Amount(waResult(LastRow, GAMT)) = False Then
            .Col = GAMT
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
    Dim wiTotalCst As Double
    Dim wiTotalQty As Double
    
    Dim wiRowCtr As Integer
    
    Calc_Total = False
    For wiRowCtr = 0 To waResult.UpperBound(1)
        wiTotalGrs = wiTotalGrs + To_Value(waResult(wiRowCtr, GNET))
        'wiTotalDis = wiTotalDis + To_Value(waResult(wiRowCtr, GNET)) * To_Value(txtSpecDis.Text)
        wiTotalNet = wiTotalNet + To_Value(waResult(wiRowCtr, GNET))
        wiTotalQty = wiTotalQty + To_Value(waResult(wiRowCtr, GQTY))
        wiTotalCst = wiTotalCst + To_Value(waResult(wiRowCtr, GCST))
        
    Next
    
    lblDspGrsAmtOrg.Caption = Format(CStr(wiTotalGrs), gsAmtFmt)
    lblDspDisAmtOrg.Caption = Format(CStr(wiTotalDis), gsAmtFmt)
    lblDspNetAmtOrg.Caption = Format(CStr(wiTotalNet), gsAmtFmt)
    lblDspCstAmtOrg.Caption = Format(CStr(wiTotalCst), gsAmtFmt)
    
    lblDspTotalQty.Caption = Format(CStr(wiTotalQty), gsQtyFmt)
    
    btnGetDisAmt_Click
    
    Calc_Total = True

End Function




Private Function cmdDel() As Boolean

    Dim wsGenDte As String
    Dim adcmdSave As New ADODB.Command
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
    Set adcmdSave.ActiveConnection = cnCon
        
    adcmdSave.CommandText = "USP_SO001A"
    adcmdSave.CommandType = adCmdStoredProc
    adcmdSave.Parameters.Refresh
      
    Call SetSPPara(adcmdSave, 1, wiAction)
    Call SetSPPara(adcmdSave, 2, wsTrnCd)
    Call SetSPPara(adcmdSave, 3, wlKey)
    Call SetSPPara(adcmdSave, 4, Trim(cboDocNo.Text))
    Call SetSPPara(adcmdSave, 5, wlCusID)
    Call SetSPPara(adcmdSave, 6, medDocDate.Text)
    Call SetSPPara(adcmdSave, 7, 0)
    Call SetSPPara(adcmdSave, 8, cboCurr.Text)
    Call SetSPPara(adcmdSave, 9, txtExcr.Text)
    Call SetSPPara(adcmdSave, 10, "")
    
    Call SetSPPara(adcmdSave, 11, Set_MedDate(medReserveDate.Text))
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medExpiryDate.Text))
    
    Call SetSPPara(adcmdSave, 13, wlSaleID)
    
    Call SetSPPara(adcmdSave, 14, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 15, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 16, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 17, cboCRML.Text)
    Call SetSPPara(adcmdSave, 18, cboShipCode.Text)
    Call SetSPPara(adcmdSave, 19, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 20, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 21, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 22, txtPortNo.Text)
    
    Call SetSPPara(adcmdSave, 23, txtShipFrom.Text)
    Call SetSPPara(adcmdSave, 24, txtShipTo.Text)
    Call SetSPPara(adcmdSave, 25, txtShipVia.Text)
    Call SetSPPara(adcmdSave, 26, txtShipPer.Text)
    Call SetSPPara(adcmdSave, 27, txtShipName.Text)
    Call SetSPPara(adcmdSave, 28, txtShipAdr1.Text)
    Call SetSPPara(adcmdSave, 29, txtShipAdr2.Text)
    Call SetSPPara(adcmdSave, 30, txtShipAdr3.Text)
    Call SetSPPara(adcmdSave, 31, txtShipAdr4.Text)
    
    For i = 1 To 10
    Call SetSPPara(adcmdSave, 32 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdSave, 42, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 43, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 44, lblDspNetAmtOrg)
    Call SetSPPara(adcmdSave, 45, txtSpecDis.Text)
    Call SetSPPara(adcmdSave, 46, cboJobNo.Text)
    Call SetSPPara(adcmdSave, 47, txtJobCost.Text)
    Call SetSPPara(adcmdSave, 48, wsFormID)
    
    Call SetSPPara(adcmdSave, 49, gsUserID)
    Call SetSPPara(adcmdSave, 50, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 51)
    wsDocNo = GetSPPara(adcmdSave, 52)
    

    
    
    cnCon.CommitTrans
    
    gsMsg = wsDocNo & " 檔案已刪除!"
    MsgBox gsMsg, vbOKOnly, gsTitle
    Call cmdCancel
    MousePointer = vbDefault
    
    Set adcmdSave = Nothing
    cmdDel = True
    
    Exit Function
    
cmdDelete_Err:
    MsgBox "Check cmdDel"
    MousePointer = vbDefault
    cnCon.RollbackTrans
    Set adcmdSave = Nothing

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
                .Buttons(tcAppendix).Enabled = False
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
                .Buttons(tcAppendix).Enabled = False
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
                .Buttons(tcAppendix).Enabled = True
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
                .Buttons(tcAppendix).Enabled = False
            End With
            
       
    
    End Select
End Sub



'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
    Select Case sStatus
        Case "Default"
        
            Me.cboDocNo.Enabled = False
            Me.cboCusCode.Enabled = False
        '    Me.txtRevNo.Enabled = False
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            
            Me.medReserveDate.Enabled = False
            Me.medExpiryDate.Enabled = False
            Me.cboSaleCode.Enabled = False
            Me.cboPayCode.Enabled = False
            Me.cboPrcCode.Enabled = False
            Me.cboMLCode.Enabled = False
            Me.cboCRML.Enabled = False
            
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
            Me.txtSpecDis.Enabled = False
            Me.txtDisAmt.Enabled = False
            
            Me.cboJobNo.Enabled = False
            Me.txtJobCost.Enabled = False
            
            
            Me.btnGetDisAmt.Enabled = False
            
            
            Me.picRmk.Enabled = False
            
            Me.tblDetail.Enabled = False
            
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
       
        Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
        
            Me.cboDocNo.Enabled = False
            
            Me.cboCusCode.Enabled = True
          '  Me.txtRevNo.Enabled = True
            Me.medDocDate.Enabled = True
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.medReserveDate.Enabled = True
            Me.medExpiryDate.Enabled = True
            Me.cboSaleCode.Enabled = True
            Me.cboPayCode.Enabled = True
            Me.cboPrcCode.Enabled = True
            Me.cboMLCode.Enabled = True
            Me.cboCRML.Enabled = True
            
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
            
            Me.cboJobNo.Enabled = True
            Me.txtJobCost.Enabled = True
            
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
        .TableKey = "SOHDDocNo"
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
    vFilterAry(1, 2) = "SOHDDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "SOHDDocDate"
    
    vFilterAry(3, 1) = "Customer #"
    vFilterAry(3, 2) = "CusCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "SOHDDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "SOHDDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Customer#"
    vAry(3, 2) = "CusCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Customer Name"
    vAry(4, 2) = "CusName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT SOASOHD.SOHDDocNo, SOASOHD.SOHDDocDate, mstCustomer.CusCode,  mstCustomer.CusName "
        wsSQL = wsSQL + "FROM MstCustomer, SOASOHD "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE SOASOHD.SOHDStatus = '1' And SOASOHD.SOHDCusID = MstCustomer.CusID "
        .sBindOrderSQL = "ORDER BY SOASOHD.SOHDDocNo"
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
    wsSQL = wsSQL & "AND SaleStatus = '1' "
    wsSQL = wsSQL & "AND SaleType = 'S' "
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
        
        txtPortNo = Get_TableInfo("MstPriceTerm", "PrcCode = '" & Set_Quote(cboPrcCode.Text) & "'", "PricePort")
        
        tabDetailInfo.Tab = 0
        cboMLCode.SetFocus
       
    End If
    
End Sub

Private Sub cboPrcCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboPrcCode
    
    wsSQL = "SELECT PrcCode, PRCDESC FROM mstPriceTerm WHERE PrcCode LIKE '%" & IIf(cboPrcCode.SelLength > 0, "", Set_Quote(cboPrcCode.Text)) & "%' "
    wsSQL = wsSQL & "AND PRCSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY PrcCode "
    Call Ini_Combo(2, wsSQL, cboPrcCode.Left + tabDetailInfo.Left, cboPrcCode.Top + cboPrcCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLPRCCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboPrcCode() As Boolean
Dim wsDesc As String

    Chk_cboPrcCode = False
     
    If Trim(cboPrcCode.Text) = "" Then
        lblDspPrcDesc.Caption = ""
        Chk_cboPrcCode = True
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
        cboCRML.SetFocus
       
    End If
    
End Sub

Private Sub cboMLCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboMLCode
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass WHERE MLCode LIKE '%" & IIf(cboMLCode.SelLength > 0, "", Set_Quote(cboMLCode.Text)) & "%' "
    wsSQL = wsSQL & "AND MLSTATUS = '1' "
    wsSQL = wsSQL & "AND MLTYPE = 'A' "
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
    
    
    If Chk_MClass(cboMLCode, "A", wsDesc) = False Then
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
    
    Call chk_InpLen(txtShipFrom, 1000, KeyAscii)
    
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
    
    Call chk_InpLen(txtShipTo, 1000, KeyAscii)
    
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
    
    Call chk_InpLen(txtShipVia, 1000, KeyAscii)
    
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
    
    Call chk_InpLen(txtCusPo, 50, KeyAscii)
    
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
    
    Call chk_InpLen(txtLcNo, 50, KeyAscii)
    
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
    
    Call chk_InpLen(txtPortNo, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_KeyFld = True Then
        tabDetailInfo.Tab = 1
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
    wsSQL = wsSQL & "AND ShipCardID = " & wlCusID & " "
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
    
    Call chk_InpLen(txtShipPer, 50, KeyAscii)
    
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
    Dim wlRow As Long
    
    wsAct = inArray(inMnuIdx, 0)
    
    With tblDetail
    Select Case wsAct
        Case "DELETE"
           
           If IsNull(.Bookmark) Then Exit Sub
            If .EditActive = True Then Exit Sub
            If Chk_SoExistSp(.Bookmark) = False Then Exit Sub
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
            
        Case "COPY"
           
           If IsNull(.Bookmark) Then Exit Sub
            If .EditActive = True Then Exit Sub
            gsMsg = "你是否確定要複製此列?"
            If MsgBox(gsMsg, vbOKCancel, gsTitle) = vbCancel Then Exit Sub
            wlRow = cmdCopyLine(.Bookmark)
            '.Update
            'If .Row = -1 Then
            '    .Row = 0
            'End If
            .ReBind
     '       .Row = wlRow - 1
            .SetFocus
            
                
            Call Calc_Total
            
        Case Else
            Exit Sub
                    
            
    End Select
    
    End With
             
    
End Sub

Private Sub cmdRefresh()
Dim wiCtr As Integer


   Me.MousePointer = vbHourglass
  If waResult.UpperBound(1) >= 0 Then
        
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, GDESC1)) <> "" Then
            
            
              waResult(wiCtr, GAMT) = Format(To_Value(waResult(wiCtr, GPRICE)) * To_Value(waResult(wiCtr, GDISPER)) / To_Value(waResult(wiCtr, GMARKUP)), gsAmtFmt)
              waResult(wiCtr, GNET) = Format(To_Value(waResult(wiCtr, GPRICE)) * To_Value(waResult(wiCtr, GQTY)) * To_Value(waResult(wiCtr, GDISPER)) / To_Value(waResult(wiCtr, GMARKUP)), gsAmtFmt)
              waResult(wiCtr, GCST) = Format(To_Value(waResult(wiCtr, GUCST)) * To_Value(waResult(wiCtr, GQTY)), gsAmtFmt)
                 
            End If
        Next
   
   
   
   tblDetail.ReBind
   tblDetail.FirstRow = 0
    
   Call Calc_Total
   
   End If
    
     Me.MousePointer = vbDefault
    
    
End Sub



Private Function LoadQTItem(InDocID As Long, inPtjID As Long, inLineNo As Long) As Boolean

    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wiRow As Long
    
    
    On Error GoTo LoadQTItem_Err
    
    LoadQTItem = False
    
    wsSQL = " SELECT SODTID, SODTDOCLINE, SODTLNTYPE, SODTITEMID, SODTVDRID, ITMCODE, ITMITMTYPECODE, VDRCODE, SODTITEMDESC, SODTUPRICE, "
    wsSQL = wsSQL & " SODTDISPER, SODTUCST, SODTQTY, SODTAMT, SODTNET, SODTCST, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITMNAME "
    wsSQL = wsSQL & " FROM soaSODT, MstItem, MstVendor "
    wsSQL = wsSQL & " WHERE SODTDocID = " & InDocID
    wsSQL = wsSQL & " AND ItmID = SODTItemID "
    wsSQL = wsSQL & " AND VdrID = SODTVdrID "
    wsSQL = wsSQL & " AND SODTPtjID = " & inPtjID
    wsSQL = wsSQL & " ORDER BY SODTDOCLINE "
     
     
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    wiRow = 1
  '  waItem.ReDim 0, waResult.UpperBound(1), SLINENO, SITMID
    
    If rsRcd.RecordCount = 0 Then
        rsRcd.Close
        Set rsRcd = Nothing
        Exit Function
    End If
    With waItem
    Do Until rsRcd.EOF
        .AppendRows
        waItem(.UpperBound(1), SLINENO) = wiRow
        waItem(.UpperBound(1), SLN) = inLineNo
        '' Tom 20090203
  '      waItem(.UpperBound(1), SINDENT) = IIf(ReadRs(rsRcd, "SODTLNTYPE") = "T", "-1", "0")
        waItem(.UpperBound(1), SINDENT) = ReadRs(rsRcd, "SODTLNTYPE")
  '      waItem(.UpperBound(1), SINDENT) = "0"
        waItem(.UpperBound(1), SITMTYPE) = ReadRs(rsRcd, "ITMITMTYPECODE")
        waItem(.UpperBound(1), SITMCODE) = ReadRs(rsRcd, "ITMCODE")
        waItem(.UpperBound(1), SVDRCODE) = ReadRs(rsRcd, "VDRCODE")
        '' Tom 20090203
'        waItem(.UpperBound(1), SITMNAME) = IIf(ReadRs(rsRcd, "SODTLNTYPE") = "T", ReadRs(rsRcd, "SODTITEMDESC"), ReadRs(rsRcd, "ITMNAME"))
        waItem(.UpperBound(1), SITMNAME) = IIf(ReadRs(rsRcd, "SODTLNTYPE") = "S", ReadRs(rsRcd, "SODTITEMDESC"), ReadRs(rsRcd, "ITMNAME"))
        
        waItem(.UpperBound(1), SUNITPRICE) = Format(ReadRs(rsRcd, "SODTUPRICE"), gsUprFmt)
        waItem(.UpperBound(1), SUCST) = Format(ReadRs(rsRcd, "SODTUCST"), gsUprFmt)
        
        waItem(.UpperBound(1), SDISPER) = Format(ReadRs(rsRcd, "SODTDISPER"), gsAmtFmt)
        
        waItem(.UpperBound(1), SQTY) = Format(ReadRs(rsRcd, "SODTQTY"), gsQtyFmt)
        waItem(.UpperBound(1), SAMT) = Format(ReadRs(rsRcd, "SODTAMT"), gsAmtFmt)
        waItem(.UpperBound(1), SNET) = Format(ReadRs(rsRcd, "SODTNET"), gsAmtFmt)
        waItem(.UpperBound(1), SCST) = Format(ReadRs(rsRcd, "SODTCST"), gsAmtFmt)
        
        waItem(.UpperBound(1), SDTID) = To_Value(ReadRs(rsRcd, "SODTID"))
        waItem(.UpperBound(1), SITMID) = To_Value(ReadRs(rsRcd, "SODTITEMID"))
        waItem(.UpperBound(1), SVDRID) = To_Value(ReadRs(rsRcd, "SODTVDRID"))
        
        wiRow = wiRow + 1
        rsRcd.MoveNext
    Loop
    End With

    rsRcd.Close
    Set rsRcd = Nothing
    
    LoadQTItem = True
    Exit Function
        
LoadQTItem_Err:
    MsgBox "LoadQTItem Err! " & Err.Description
    
    rsRcd.Close
    Set rsRcd = Nothing
    
End Function

Private Sub cmdCstRefresh(ByVal inLine As Long)
Dim wiCtr As Integer
Dim wiCtr1 As Integer
Dim wdTotCst As Double
Dim wdTotNet As Double
Dim wbBOM As Boolean


   Me.MousePointer = vbHourglass
  
  If waItem.UpperBound(1) >= 0 Then
        
        For wiCtr = 0 To waResult.UpperBound(1)
        wdTotCst = 0
        wdTotNet = 0
        wbBOM = False
        
            For wiCtr1 = 0 To waItem.UpperBound(1)
            If Trim(waResult(wiCtr, GLINENO)) = Trim(waItem(wiCtr1, SLN)) Then
                
                wdTotCst = wdTotCst + To_Value(waItem(wiCtr1, SCST))
                wdTotNet = wdTotNet + To_Value(waItem(wiCtr1, SNET))
                wbBOM = True
            End If
            
            Next wiCtr1
            
              If wbBOM = True Then
              If inLine = To_Value(waResult(wiCtr, GLINENO)) Then
              waResult(wiCtr, GPRICE) = Format(wdTotNet, gsAmtFmt)
              
              If To_Value(waResult(wiCtr, GAMT)) = 0 Then
              waResult(wiCtr, GAMT) = Format(wdTotNet * To_Value(waResult(wiCtr, GDISPER)) / To_Value(waResult(wiCtr, GMARKUP)), gsAmtFmt)
              waResult(wiCtr, GNET) = Format(wdTotNet * To_Value(waResult(wiCtr, GQTY)) * To_Value(waResult(wiCtr, GDISPER)) / To_Value(waResult(wiCtr, GMARKUP)), gsAmtFmt)
              End If
              
              waResult(wiCtr, GUCST) = Format(wdTotCst, gsAmtFmt)
              waResult(wiCtr, GCST) = Format(wdTotCst * To_Value(waResult(wiCtr, GQTY)), gsAmtFmt)
              End If
              waResult(wiCtr, GBOM) = "Y"
              Else
              waResult(wiCtr, GBOM) = "N"
              End If
              
        Next wiCtr
   
   tblDetail.ReBind
   tblDetail.Col = GDISPER
   tblDetail.SetFocus
   End If
   
   Call Calc_Total
    
     Me.MousePointer = vbDefault
    
    
End Sub


Private Function Chk_grdItmClass(ByVal inAccNo As String) As Boolean
    
     Chk_grdItmClass = False
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入物料分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Exit Function
    End If
    
    
    If UCase(inAccNo) <> "A" And UCase(inAccNo) <> "D" Then
        gsMsg = "沒有此物料分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdItmClass = False
        Exit Function
    End If
    
    Chk_grdItmClass = True
    
    
End Function





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

Private Function Chk_txtSpecDis() As Boolean
    
    Chk_txtSpecDis = False
    
    
    If To_Value(txtSpecDis.Text) > 1 Then
        gsMsg = "對換率超出範圍!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtSpecDis.SetFocus
        Exit Function
    End If
    txtSpecDis.Text = Format(txtSpecDis.Text, gsAmtFmt)
    
    Chk_txtSpecDis = True
    
End Function

Private Sub txtSpecDis_LostFocus()
FocusMe txtSpecDis, True
End Sub

Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsDetail As String
    Dim wsByItem As String
    Dim wsTitleCS As String
    
    
    
    'If InputValidation = False Then Exit Sub
    
    gsMsg = "你要否列印工程單之總額?"
    If MsgBox(gsMsg, vbYesNo, gsTitle) = vbYes Then
    wsDetail = "Y"
    Else
    wsDetail = "N"
    End If
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = ""
    wsSelection(2) = ""
    wsSelection(3) = ""
    wsSelection(4) = ""
    
    
     
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTSO002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'SO', "
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
    wsRptName = "C" + "RPTSO002"
    Else
    wsRptName = "RPTSO002"
    End If
    
    If wsDetail = "Y" Then wsRptName = wsRptName + "A"
    
    
    NewfrmPrint.ReportID = "SO002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SO002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    
    Set NewfrmPrint = Nothing
    
    ' Tom 20090203
    gsMsg = "你要否列印整份BOM由行數排序?"
    If MsgBox(gsMsg, vbYesNo, gsTitle) = vbYes Then
    wsByItem = "N"
    Else
    wsByItem = "Y"
    End If
    
    
    wsSQL = "EXEC usp_RPTSO002D '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wgsTitle & "', "
    wsSQL = wsSQL & "'', "
    wsSQL = wsSQL & "'SO', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & String(10, "z") & "', "
    wsSQL = wsSQL & "'" & "000000" & "', "
    wsSQL = wsSQL & "'" & "999999" & "', "
    wsSQL = wsSQL & "'" & "%" & "', "
    wsSQL = wsSQL & "'" & wsByItem & "', "
    wsSQL = wsSQL & gsLangID
    
    
    If wsByItem = "Y" Then
        wsRptName = "RPTSO002S"
    Else
        wsRptName = "RPTSO002N"
    End If
    
    If gsLangID = "2" Then
        wsRptName = "C" + wsRptName
    End If
    
    
    NewfrmPrint.ReportID = "SO002D"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "SO002D"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    Set NewfrmPrint = Nothing
    
    
    '' Print Appendix
    If Chk_DocAppendix(wlKey) = True Then
    
    wsTitleCS = "CONDITION OF SALES"
    
    
    wsSQL = "EXEC usp_RPTDA002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wsTitleCS & "', "
    wsSQL = wsSQL & "'SO', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & Set_Quote(cboDocNo.Text) & "', "
    wsSQL = wsSQL & "'" & "" & "', "
    wsSQL = wsSQL & "'" & String(10, "z") & "', "
    wsSQL = wsSQL & "'" & "000000" & "', "
    wsSQL = wsSQL & "'" & "999999" & "', "
    wsSQL = wsSQL & gsLangID
    
    If gsLangID = "2" Then
        wsRptName = "CRPTDA002"
    Else
        wsRptName = "RPTDA002"
    End If
    
    
    NewfrmPrint.ReportID = "DA002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "DA002"
    NewfrmPrint.RptDteTim = wsDteTim
    NewfrmPrint.StoreP = wsSQL
    NewfrmPrint.Selection = wsSelection
    NewfrmPrint.RptName = wsRptName
    NewfrmPrint.Show vbModal
    Set NewfrmPrint = Nothing
        
    
    
    End If
    
    
    
    
    
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdRevise()

     
    On Error GoTo cmdRevise_Err
    
    
    If DelValidation(wlKey) = False Then
       wiAction = CorRec
       MousePointer = vbDefault
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

Private Function cmdCopyLine(ByVal CRow As Long) As Long
    Dim wsLineNo As String
    Dim wiCtr As Long
    Dim wiLn As Integer
    
    On Error GoTo cmdCopyLine_Err
    
    wsLineNo = waResult(CRow, GLINENO)
    
    With waResult
    .AppendRows
    waResult(.UpperBound(1), GLINENO) = wlLineNo
    waResult(.UpperBound(1), GDESC1) = waResult(CRow, GDESC1)
    waResult(.UpperBound(1), GMORE) = waResult(CRow, GMORE)
    waResult(.UpperBound(1), GQTY) = Format(waResult(CRow, GQTY), gsQtyFmt)
    waResult(.UpperBound(1), GPRICE) = Format(waResult(CRow, GPRICE), gsAmtFmt)
    waResult(.UpperBound(1), GDISPER) = Format(waResult(CRow, GDISPER), gsAmtFmt)
    waResult(.UpperBound(1), GMARKUP) = Format(waResult(CRow, GMARKUP), gsAmtFmt)
    waResult(.UpperBound(1), GUOM) = waResult(CRow, GUOM)
    waResult(.UpperBound(1), GUCST) = Format(waResult(CRow, GUCST), gsAmtFmt)
    waResult(.UpperBound(1), GAMT) = Format(waResult(CRow, GAMT), gsAmtFmt)
    waResult(.UpperBound(1), GNET) = Format(waResult(CRow, GNET), gsAmtFmt)
    waResult(.UpperBound(1), GCST) = Format(waResult(CRow, GCST), gsAmtFmt)
    waResult(.UpperBound(1), GBOM) = waResult(CRow, GBOM)
    If To_Value(waResult(CRow, GDRMKID)) = 0 Then
    waResult(.UpperBound(1), GDRMKID) = "0"
    Else
    waResult(.UpperBound(1), GDRMKID) = Get_DRmkID("QT", waResult(CRow, GDRMKID))
    End If
    waResult(.UpperBound(1), GPTJID) = "0"
    cmdCopyLine = .UpperBound(1)
    
    End With
    
    
    wiLn = 1
    With waItem
    If .UpperBound(1) >= 0 Then
    For wiCtr = 0 To .UpperBound(1)
         If To_Value(waItem(wiCtr, SLN)) = To_Value(wsLineNo) Then
             .AppendRows
             waItem(.UpperBound(1), SLINENO) = wiLn
             waItem(.UpperBound(1), SINDENT) = waItem(wiCtr, SINDENT)
             waItem(.UpperBound(1), SITMTYPE) = waItem(wiCtr, SITMTYPE)
             waItem(.UpperBound(1), SITMCODE) = waItem(wiCtr, SITMCODE)
             waItem(.UpperBound(1), SVDRCODE) = waItem(wiCtr, SVDRCODE)
             waItem(.UpperBound(1), SITMNAME) = waItem(wiCtr, SITMNAME)
             waItem(.UpperBound(1), SUNITPRICE) = Format(waItem(wiCtr, SUNITPRICE), gsUprFmt)
             waItem(.UpperBound(1), SUCST) = Format(waItem(wiCtr, SUCST), gsUprFmt)
             
             waItem(.UpperBound(1), SDISPER) = Format(waItem(wiCtr, SDISPER), gsAmtFmt)
             
             waItem(.UpperBound(1), SQTY) = Format(waItem(wiCtr, SQTY), gsQtyFmt)
             waItem(.UpperBound(1), SAMT) = Format(waItem(wiCtr, SAMT), gsAmtFmt)
             waItem(.UpperBound(1), SNET) = Format(waItem(wiCtr, SNET), gsAmtFmt)
             waItem(.UpperBound(1), SCST) = Format(waItem(wiCtr, SCST), gsAmtFmt)
             
             waItem(.UpperBound(1), SLN) = wlLineNo
             waItem(.UpperBound(1), SDTID) = "0"
             waItem(.UpperBound(1), SITMID) = To_Value(waItem(wiCtr, SITMID))
             waItem(.UpperBound(1), SVDRID) = To_Value(waItem(wiCtr, SVDRID))
             wiLn = wiLn + 1
         End If
    Next wiCtr
    End If
    End With
    
    wlLineNo = wlLineNo + 1
    
    
    Exit Function
    
cmdCopyLine_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
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
            cboJobNo.SetFocus
            
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



Private Function Chk_CusCode(ByVal InCusNo As String, ByRef OutID As Long, ByRef OutName As String, ByRef OutTel As String, ByRef OutFax As String, ByRef OutEMail As String) As Boolean
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String

    
    wsSQL = "SELECT CusID, CusName, CusTel, CusFax, CusEMail FROM mstCustomer WHERE CusCode = '" & Set_Quote(InCusNo) & "' "
    wsSQL = wsSQL & "And CusStatus = '1' "
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rsRcd.RecordCount > 0 Then
    
        OutID = ReadRs(rsRcd, "CusID")
        OutName = ReadRs(rsRcd, "CusName")
        OutTel = ReadRs(rsRcd, "CusTel")
        OutFax = ReadRs(rsRcd, "CusFax")
        OutEMail = ReadRs(rsRcd, "CusEMail")
        
        Chk_CusCode = True
        
    Else
    
        OutID = 0
        OutName = ""
        OutTel = ""
        OutFax = ""
        OutEMail = ""
       
        Chk_CusCode = False
        
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
End Function


Private Function DelValidation(ByVal InDocID As Long) As Boolean
Dim OutTrnCd As String
Dim OutDocNo As String

    
    
    DelValidation = False
    
    On Error GoTo DelValidation_Err
    
    
    
 '   If Not chk_txtRevNo Then Exit Function
    If Chk_SoRefDoc(InDocID, OutTrnCd, OutDocNo) = True Then
        
        Select Case OutTrnCd
        Case "SP"
        gsMsg = "配貨單 : " & OutDocNo & " 是以此銷售轉為!不能刪除或改正"
        Case "SW"
        gsMsg = "發貨單 : " & OutDocNo & " 是以此銷售轉為!不能刪除或改正"
        Case "SR"
        gsMsg = "退貨單 : " & OutDocNo & " 是以此銷售轉為!不能刪除或改正"
        Case "IV"
        gsMsg = "發票 : " & OutDocNo & " 是以此銷售轉為!不能刪除或改正"
        Case "PO"
        gsMsg = "採購單 : " & OutDocNo & " 是以此銷售轉為!不能刪除或改正"
        End Select
        MsgBox gsMsg, vbOKOnly, gsTitle
        
        Exit Function
    
    End If
    
    DelValidation = True
    
    Exit Function
    
DelValidation_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Private Sub cboCRML_GotFocus()
    FocusMe cboCRML
End Sub

Private Sub cboCRML_LostFocus()
    FocusMe cboCRML, True
End Sub


Private Sub cboCRML_KeyPress(KeyAscii As Integer)
    Dim wsDesc As String
    
    Call chk_InpLen(cboCRML, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboCRML = False Then
                Exit Sub
        End If
        
        tabDetailInfo.Tab = 0
        medReserveDate.SetFocus
       
    End If
    
End Sub

Private Sub cboCRML_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboCRML
    
    wsSQL = "SELECT MLCode, MLDESC FROM mstMerchClass WHERE MLCode LIKE '%" & IIf(cboCRML.SelLength > 0, "", Set_Quote(cboCRML.Text)) & "%' "
    wsSQL = wsSQL & "AND MLSTATUS = '1' "
    wsSQL = wsSQL & "AND MLTYPE = 'S' "
    wsSQL = wsSQL & "ORDER BY MLCode "
    Call Ini_Combo(2, wsSQL, cboCRML.Left + tabDetailInfo.Left, cboCRML.Top + cboCRML.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLMLCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboCRML() As Boolean
Dim wsDesc As String

    Chk_cboCRML = False
     
    If Trim(cboCRML.Text) = "" Then
        gsMsg = "必需輸入會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboCRML.SetFocus
        Exit Function
    End If
    
    
    If Chk_MClass(cboCRML, "S", wsDesc) = False Then
        gsMsg = "沒有此會計分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboCRML.SetFocus
        lblDspMLDesc = ""
       Exit Function
    End If
    
    lblDspCRMLDesc = wsDesc
    
    Chk_cboCRML = True
    
End Function

Private Sub cmdRmkID(wiRow As Integer, wsRmkID As String)

  Me.MousePointer = vbHourglass
  
  If wiRow >= 0 Then
        
            waResult(wiRow, GDRMKID) = wsRmkID
            If To_Value(wsRmkID) = 0 Then
            waResult(wiRow, GMORE) = "N"
            Else
            waResult(wiRow, GMORE) = "Y"
            End If
            
            tblDetail.ReBind
            tblDetail.Col = GQTY
            tblDetail.SetFocus
            
            
  End If
   
  Me.MousePointer = vbDefault
    
    
End Sub


Private Function Chk_SoExistSp(ByVal CRow As Long) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsPtjID As String
    
    On Error GoTo Chk_SoExistSp_Err
    
    
    If wiAction <> CorRec Then
        Chk_SoExistSp = True
        Exit Function
    End If
    
    
    wsPtjID = waResult(CRow, GPTJID)
    
    
    If wsPtjID = "0" Then
        Chk_SoExistSp = True
        Exit Function
    End If
    
    
    wsSQL = "SELECT * FROM SoaSpHd, SoaSpDt, SoaSoDt "
    wsSQL = wsSQL & " WHERE SpDtSoDtID = SoDtID "
    wsSQL = wsSQL & " AND SoDtPtjID = " & To_Value(wsPtjID) & " "
    wsSQL = wsSQL & " AND SpHDDocID = SpDtDocID "
    wsSQL = wsSQL & " AND SpHDStatus In ('1','4') "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        gsMsg = "不能刪除!物料已轉移到(B)倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Chk_SoExistSp = False
        Exit Function
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    wsSQL = "SELECT * FROM SoaSwHd, SoaSwDt, SoaSoDt "
    wsSQL = wsSQL & " WHERE SwDtSoDtID = SoDtID "
    wsSQL = wsSQL & " AND SoDtPtjID = " & To_Value(wsPtjID) & " "
    wsSQL = wsSQL & " AND SwHDDocID = SwDtDocID "
    wsSQL = wsSQL & " AND SwHDStatus In ('1','4') "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        gsMsg = "不能刪除!物料已轉移到(C)倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Chk_SoExistSp = False
        Exit Function
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_SoExistSp = True
    
    
    Exit Function
    
Chk_SoExistSp_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function


Private Function Chk_SoExistSpDtQty(ByVal wsDtlID As String, ByVal wsLineNo As String, ByVal wsItemNo As String, ByVal InHDQty As Double, ByVal InQty As Double) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    On Error GoTo Chk_SoExistSpDtQty_Err
    
    
    If wsDtlID = "0" Then
        Chk_SoExistSpDtQty = True
        Exit Function
    End If
    
        
    wsSQL = "SELECT Sum(SpDtQty) Qty FROM SoaSpHd, SoaSpDt "
    wsSQL = wsSQL & " WHERE SpDtSoDtID = " & To_Value(wsDtlID) & " "
    wsSQL = wsSQL & " AND SpHDDocID = SpDtDocID "
    wsSQL = wsSQL & " AND SpHDStatus In ('1','4') "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
    If To_Value(InHDQty) * To_Value(InQty) < To_Value(ReadRs(rsRcd, "Qty")) Then
        gsMsg = wsLineNo & " : " & wsItemNo & " 數量不足!物料已轉移到(B)倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Chk_SoExistSpDtQty = False
        Exit Function
    End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_SoExistSpDtQty = True
    
    
    Exit Function
    
Chk_SoExistSpDtQty_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function


Private Function ChkSoDetail() As Boolean
Dim wiCtr As Integer
Dim wiCtr1 As Integer
Dim wdHdQty As Double

Dim wsDtlID As String
Dim wsItemNo As String
Dim wdQty As Double

ChkSoDetail = True

  Me.MousePointer = vbHourglass
  
  If waItem.UpperBound(1) >= 0 Then
        
        For wiCtr = 0 To waResult.UpperBound(1)
        wdHdQty = To_Value(waResult(wiCtr, GQTY))
        For wiCtr1 = 0 To waItem.UpperBound(1)
            If Trim(waResult(wiCtr, GLINENO)) = Trim(waItem(wiCtr1, SLN)) Then
                
                wsDtlID = To_Value(waItem(wiCtr1, SDTID))
                wsItemNo = waItem(wiCtr1, SITMCODE)
                wdQty = To_Value(waItem(wiCtr1, SQTY))
                
                If Chk_SoExistSpDtQty(wsDtlID, waResult(wiCtr, GLINENO), wsItemNo, wdHdQty, wdQty) = False Then
                ChkSoDetail = False
                Me.MousePointer = vbDefault
                Exit Function
                End If
                
            End If
            
        Next wiCtr1
        Next wiCtr
   
   End If
   
    
    Me.MousePointer = vbDefault
    
    
End Function


Private Sub cboJobNo_GotFocus()
    FocusMe cboJobNo
End Sub

Private Sub cboJobNo_LostFocus()
    FocusMe cboJobNo, True
End Sub


Private Sub cboJobNo_KeyPress(KeyAscii As Integer)
    Dim wdCost As Double
    
    Call chk_InpLen(cboJobNo, 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboJobNo = False Then
                Exit Sub
        End If
        
        tabDetailInfo.Tab = 0
        
        wdCost = Get_JobCost(cboJobNo.Text)
        txtJobCost.Text = Format(wdCost, gsAmtFmt)
        txtJobCost.SetFocus
       
    End If
    
End Sub

Private Sub cboJobNo_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboJobNo
    
    wsSQL = "SELECT SOHDDOCNO FROM SOASOHD WHERE SOHDDOCNO LIKE '%" & IIf(cboJobNo.SelLength > 0, "", Set_Quote(cboJobNo.Text)) & "%' "
    wsSQL = wsSQL & "AND SOHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & "AND SOHDCUSID = " & wlCusID & " "
    wsSQL = wsSQL & "ORDER BY SOHDDOCNO "
    
    Call Ini_Combo(1, wsSQL, cboJobNo.Left + tabDetailInfo.Left, cboJobNo.Top + cboJobNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLJOBNO", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboJobNo() As Boolean


    Chk_cboJobNo = False
     
    If Trim(cboJobNo.Text) = "" Then
        gsMsg = "必需輸入工程編號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboJobNo.SetFocus
        Exit Function
    End If
    
    
    

    
    Chk_cboJobNo = True
    
End Function



Private Sub txtJobCost_GotFocus()

    FocusMe txtJobCost
    
End Sub

Private Sub txtJobCost_KeyPress(KeyAscii As Integer)
    
    Call Chk_InpNum(KeyAscii, txtJobCost.Text, False, True)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        If chk_txtJobCost Then
            
            tabDetailInfo.Tab = 0
            txtShipFrom.SetFocus
            
        End If
    End If

End Sub

Private Function chk_txtJobCost() As Boolean
    
    chk_txtJobCost = False
    
    
    If To_Value(txtJobCost.Text) < 0 Then
        gsMsg = "錯誤!一定大於零"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        txtJobCost.SetFocus
        Exit Function
    End If
    
    txtJobCost.Text = Format(txtJobCost.Text, gsAmtFmt)
    
    chk_txtJobCost = True
    
End Function

Private Sub txtJobCost_LostFocus()
txtJobCost.Text = Format(txtJobCost.Text, gsAmtFmt)
FocusMe txtJobCost, True
End Sub


Private Sub cmdAppendix()

Dim newForm As Form

Set newForm = New frmDocAppendix

     With newForm
       .DocID = wlKey
       .RmkID = "0"
       .RmkType = wsTrnCd
       .Show vbModal
     End With
     
Unload newForm


End Sub



