VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPO001 
   Caption         =   "訂貨單"
   ClientHeight    =   8595
   ClientLeft      =   1.96650e5
   ClientTop       =   270
   ClientWidth     =   11880
   Icon            =   "frmPO001.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  '螢幕中央
   Begin TrueDBGrid60.TDBGrid tblCommon 
      Height          =   2070
      Left            =   11400
      OleObjectBlob   =   "frmPO001.frx":030A
      TabIndex        =   54
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
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":66AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPO001.frx":69C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrProcess 
      Align           =   1  '對齊表單上方
      Height          =   360
      Left            =   0
      TabIndex        =   55
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
      TabIndex        =   56
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
      TabPicture(0)   =   "frmPO001.frx":6CE5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboRefDocNo"
      Tab(0).Control(1)=   "cboVdrCode"
      Tab(0).Control(2)=   "cboDelCode"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "cboCurr"
      Tab(0).Control(5)=   "cboDocNo"
      Tab(0).Control(6)=   "cboPayCode"
      Tab(0).Control(7)=   "cboPrcCode"
      Tab(0).Control(8)=   "cboMLCode"
      Tab(0).Control(9)=   "cboSaleCode"
      Tab(0).Control(10)=   "FraDate"
      Tab(0).Control(11)=   "fraInfo"
      Tab(0).Control(12)=   "fraCode"
      Tab(0).Control(13)=   "fraKey"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Shipment "
      TabPicture(1)   =   "frmPO001.frx":6D01
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
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Item Information"
      TabPicture(2)   =   "frmPO001.frx":6D1D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboShipCode(1)"
      Tab(2).Control(1)=   "fraShip(1)"
      Tab(2).Control(2)=   "cboShipCode(0)"
      Tab(2).Control(3)=   "fraShip(0)"
      Tab(2).Control(4)=   "cboRmkCode"
      Tab(2).Control(5)=   "fraRmk"
      Tab(2).ControlCount=   6
      Begin VB.ComboBox cboRefDocNo 
         Height          =   300
         Left            =   -69720
         TabIndex        =   2
         Top             =   780
         Width           =   2055
      End
      Begin VB.ComboBox cboVdrCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   1
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox cboDelCode 
         Height          =   300
         Left            =   -68760
         TabIndex        =   15
         Top             =   4320
         Width           =   1770
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Index           =   1
         Left            =   -67320
         TabIndex        =   34
         Top             =   360
         Width           =   2010
      End
      Begin VB.Frame fraShip 
         Height          =   3495
         Index           =   1
         Left            =   -69000
         TabIndex        =   115
         Top             =   120
         Width           =   5655
         Begin VB.TextBox txtShipPer 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1680
            TabIndex        =   35
            Text            =   "01234567890123457890"
            Top             =   600
            Width           =   3585
         End
         Begin VB.TextBox txtShipName 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1680
            TabIndex        =   36
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   960
            Width           =   3585
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1335
            Index           =   1
            Left            =   1680
            ScaleHeight     =   1275
            ScaleWidth      =   3555
            TabIndex        =   116
            Top             =   1320
            Width           =   3615
            Begin VB.TextBox txtShipAdr1 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   37
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   3465
            End
            Begin VB.TextBox txtShipAdr2 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   38
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   3465
            End
            Begin VB.TextBox txtShipAdr3 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   39
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   720
               Width           =   3465
            End
            Begin VB.TextBox txtShipAdr4 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   40
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1080
               Width           =   3465
            End
         End
         Begin VB.TextBox txtShipFaxNo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1680
            TabIndex        =   42
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   3120
            Width           =   1905
         End
         Begin VB.TextBox txtShipTelNo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1680
            TabIndex        =   41
            Text            =   "01234567890123457890"
            Top             =   2760
            Width           =   1875
         End
         Begin VB.Label lblShipAdr 
            Caption         =   "SHIPADR"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   122
            Top             =   1320
            Width           =   1500
         End
         Begin VB.Label lblShipPer 
            Caption         =   "SHIPPER"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   121
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label lblShipName 
            Caption         =   "SHIPNAME"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   120
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label lblShipCode 
            Caption         =   "SHIPCODE"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   119
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label lblShipFaxNo 
            Caption         =   "SHIPNAME"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   118
            Top             =   3120
            Width           =   1380
         End
         Begin VB.Label lblShipTelNo 
            Caption         =   "SHIPPER"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   117
            Top             =   2760
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2280
         Left            =   -74880
         TabIndex        =   111
         Top             =   5050
         Width           =   3975
         Begin VB.CommandButton btnGetDisAmt 
            Caption         =   "Command1"
            Height          =   375
            Left            =   1800
            Picture         =   "frmPO001.frx":6D39
            TabIndex        =   14
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtDisAmt 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   13
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtSpecDis 
            Alignment       =   1  '靠右對齊
            Height          =   288
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblDisAmt 
            Caption         =   "EXCR"
            Height          =   495
            Left            =   240
            TabIndex        =   124
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label lblSpecDis 
            Caption         =   "SPECDIS"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   660
            Width           =   1545
         End
      End
      Begin VB.ComboBox cboCurr 
         Height          =   300
         Left            =   -65520
         TabIndex        =   4
         Top             =   1135
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   450
         Left            =   120
         TabIndex        =   106
         Top             =   7200
         Width           =   6135
         Begin VB.Label lblDeleteLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   4800
            TabIndex        =   110
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblInsertLine 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   3360
            TabIndex        =   109
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblComboPrompt 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   1920
            TabIndex        =   108
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label lblKeyDesc 
            Caption         =   "REMARK"
            Height          =   225
            Left            =   360
            TabIndex        =   107
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboShipCode 
         Height          =   300
         Index           =   0
         Left            =   -73200
         TabIndex        =   25
         Top             =   360
         Width           =   2010
      End
      Begin VB.Frame fraShip 
         Caption         =   "CCC"
         Height          =   3495
         Index           =   0
         Left            =   -74880
         TabIndex        =   84
         Top             =   120
         Width           =   5775
         Begin VB.TextBox txtShipTelNo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1680
            TabIndex        =   32
            Text            =   "01234567890123457890"
            Top             =   2760
            Width           =   1875
         End
         Begin VB.TextBox txtShipFaxNo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1680
            TabIndex        =   33
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   3120
            Width           =   1905
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1335
            Index           =   0
            Left            =   1680
            ScaleHeight     =   1275
            ScaleWidth      =   3555
            TabIndex        =   85
            Top             =   1320
            Width           =   3615
            Begin VB.TextBox txtShipAdr4 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   0
               TabIndex        =   31
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1080
               Width           =   3465
            End
            Begin VB.TextBox txtShipAdr3 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   0
               TabIndex        =   30
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   720
               Width           =   3465
            End
            Begin VB.TextBox txtShipAdr2 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   0
               TabIndex        =   29
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   3465
            End
            Begin VB.TextBox txtShipAdr1 
               BorderStyle     =   0  '沒有框線
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   0
               TabIndex        =   28
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   3465
            End
         End
         Begin VB.TextBox txtShipName 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1680
            TabIndex        =   27
            Text            =   "012345678901234578901234567890123457890123456789"
            Top             =   960
            Width           =   3585
         End
         Begin VB.TextBox txtShipPer 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1680
            TabIndex        =   26
            Text            =   "01234567890123457890"
            Top             =   600
            Width           =   3585
         End
         Begin VB.Label lblShipTelNo 
            Caption         =   "SHIPPER"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   114
            Top             =   2760
            Width           =   1500
         End
         Begin VB.Label lblShipFaxNo 
            Caption         =   "SHIPNAME"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   113
            Top             =   3120
            Width           =   1380
         End
         Begin VB.Label lblShipCode 
            Caption         =   "SHIPCODE"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label lblShipName 
            Caption         =   "SHIPNAME"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   1380
         End
         Begin VB.Label lblShipPer 
            Caption         =   "SHIPPER"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   600
            Width           =   1500
         End
         Begin VB.Label lblShipAdr 
            Caption         =   "SHIPADR"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   1320
            Width           =   1500
         End
      End
      Begin VB.ComboBox cboRmkCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   43
         Top             =   3840
         Width           =   1890
      End
      Begin VB.ComboBox cboDocNo 
         Height          =   300
         Left            =   -73200
         TabIndex        =   0
         Top             =   420
         Width           =   1935
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
         Top             =   3180
         Width           =   2370
      End
      Begin VB.ComboBox cboMLCode 
         Height          =   300
         Left            =   -73200
         TabIndex        =   9
         Top             =   3540
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
         Height          =   975
         Left            =   -74880
         TabIndex        =   64
         Top             =   4080
         Width           =   3975
         Begin MSMask.MaskEdBox medDueDate 
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   240
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
         Begin VB.Label lblExpiryDate 
            Caption         =   "ETADATE"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   660
            Width           =   1485
         End
         Begin VB.Label lblDueDate 
            Caption         =   "DUEDATE"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   300
            Width           =   1545
         End
      End
      Begin VB.Frame fraInfo 
         Height          =   3615
         Left            =   -70800
         TabIndex        =   57
         Top             =   4080
         Width           =   7575
         Begin VB.TextBox txtDelAdr4 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   20
            Text            =   "0123456789012345789"
            Top             =   2040
            Width           =   5265
         End
         Begin VB.TextBox txtDelAdr3 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   19
            Text            =   "0123456789012345789"
            Top             =   1680
            Width           =   5265
         End
         Begin VB.TextBox txtLcNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   22
            Text            =   "0123456789012345789"
            Top             =   2880
            Width           =   5265
         End
         Begin VB.TextBox txtPortNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   23
            Text            =   "0123456789012345789"
            Top             =   3240
            Width           =   5265
         End
         Begin VB.TextBox txtCusPo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   21
            Text            =   "0123456789012345789"
            Top             =   2520
            Width           =   5265
         End
         Begin VB.TextBox txtDelAdr1 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   17
            Text            =   "0123456789012345789"
            Top             =   960
            Width           =   5265
         End
         Begin VB.TextBox txtDelAdr2 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   18
            Text            =   "0123456789012345789"
            Top             =   1320
            Width           =   5265
         End
         Begin VB.TextBox txtDelName 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   16
            Text            =   "0123456789012345789"
            Top             =   600
            Width           =   5265
         End
         Begin VB.Label lblLcNo 
            Caption         =   "LCNO"
            Height          =   240
            Left            =   120
            TabIndex        =   63
            Top             =   2940
            Width           =   1905
         End
         Begin VB.Label lblPortNo 
            Caption         =   "PORTNO"
            Height          =   240
            Left            =   120
            TabIndex        =   62
            Top             =   3300
            Width           =   1860
         End
         Begin VB.Label lblCusPo 
            Caption         =   "CUSPO"
            Height          =   240
            Left            =   120
            TabIndex        =   61
            Top             =   2580
            Width           =   1860
         End
         Begin VB.Label lblDelAdr1 
            Caption         =   "SHIPTO"
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   1020
            Width           =   1860
         End
         Begin VB.Label lblDelCode 
            Caption         =   "SHIPVIA"
            Height          =   240
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1860
         End
         Begin VB.Label lblDelName 
            Caption         =   "SHIPFROM"
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   660
            Width           =   1860
         End
      End
      Begin TrueDBGrid60.TDBGrid tblDetail 
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "frmPO001.frx":717B
         TabIndex        =   24
         Top             =   780
         Width           =   11535
      End
      Begin VB.Frame fraCode 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   75
         Top             =   1980
         Width           =   11655
         Begin VB.Label lblMlCode 
            Caption         =   "MLCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   83
            Top             =   1620
            Width           =   1545
         End
         Begin VB.Label lblDspMLDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   82
            Top             =   1560
            Width           =   7335
         End
         Begin VB.Label lblPrcCode 
            Caption         =   "PRCCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   81
            Top             =   1260
            Width           =   1545
         End
         Begin VB.Label lblDspPrcDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   80
            Top             =   1200
            Width           =   7335
         End
         Begin VB.Label lblPayCode 
            Caption         =   "PAYCODE"
            Height          =   240
            Left            =   120
            TabIndex        =   79
            Top             =   900
            Width           =   1545
         End
         Begin VB.Label lblDspPayDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   78
            Top             =   840
            Width           =   7335
         End
         Begin VB.Label lblSaleCode 
            Caption         =   "SALECODE"
            Height          =   240
            Left            =   120
            TabIndex        =   77
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label lblDspSaleDesc 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   4080
            TabIndex        =   76
            Top             =   480
            Width           =   7335
         End
      End
      Begin VB.Frame fraRmk 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   90
         Top             =   3600
         Width           =   11535
         Begin VB.PictureBox picRmk 
            BackColor       =   &H80000009&
            Height          =   3375
            Left            =   1680
            ScaleHeight     =   3315
            ScaleWidth      =   9435
            TabIndex        =   91
            Top             =   600
            Width           =   9495
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   45
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   360
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   44
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   0
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   46
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   690
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   49
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1740
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   47
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1035
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   48
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   1395
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   50
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2085
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   8
               Left            =   0
               TabIndex        =   51
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2430
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   9
               Left            =   0
               TabIndex        =   52
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   2775
               Width           =   8985
            End
            Begin VB.TextBox txtRmk 
               BorderStyle     =   0  '沒有框線
               Height          =   300
               Index           =   10
               Left            =   0
               TabIndex        =   53
               Text            =   "012345678901234578901234567890123457890123456789"
               Top             =   3120
               Width           =   8985
            End
         End
         Begin VB.Label lblRmkCode 
            Caption         =   "RMKCODE"
            Height          =   360
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lblRmk 
            Caption         =   "RMK"
            Height          =   345
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Frame fraKey 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   94
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chkWorkOrder 
            Alignment       =   1  '靠右對齊
            Caption         =   "WORKORDER"
            Height          =   180
            Left            =   7320
            TabIndex        =   125
            Top             =   360
            Width           =   2295
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
            TabIndex        =   3
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
            Left            =   3840
            TabIndex        =   123
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblVdrCode 
            Caption         =   "VDRCODE"
            Height          =   255
            Left            =   120
            TabIndex        =   105
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
            TabIndex        =   104
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblDocDate 
            Caption         =   "DOCDATE"
            Height          =   255
            Left            =   7365
            TabIndex        =   103
            Top             =   720
            Width           =   1680
         End
         Begin VB.Label lblDspVdrName 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   102
            Top             =   1020
            Width           =   5535
         End
         Begin VB.Label LblCurr 
            Caption         =   "CURR"
            Height          =   255
            Left            =   7365
            TabIndex        =   101
            Top             =   1080
            Width           =   1680
         End
         Begin VB.Label lblExcr 
            Caption         =   "EXCR"
            Height          =   255
            Left            =   7365
            TabIndex        =   100
            Top             =   1440
            Width           =   1800
         End
         Begin VB.Label lblDspVdrTel 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   1680
            TabIndex        =   99
            Top             =   1380
            Width           =   1935
         End
         Begin VB.Label lblVdrName 
            Caption         =   "VDRNAME"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblDspVdrFax 
            BorderStyle     =   1  '單線固定
            Height          =   300
            Left            =   5160
            TabIndex        =   97
            Top             =   1380
            Width           =   2055
         End
         Begin VB.Label lblVdrFax 
            Caption         =   "VDRFAX"
            Height          =   255
            Left            =   3840
            TabIndex        =   96
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblVdrTel 
            Caption         =   "VDRTEL"
            Height          =   255
            Left            =   120
            TabIndex        =   95
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
         Top             =   60
         Width           =   3315
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
Private wsOldRmkCd As String
Private wsOldPayCd As String

Private wsOldDelCd As String
Private wsOldShipCd(2) As String

Private wbReadOnly As Boolean
Private wgsTitle As String

Private Const LINENO = 0
Private Const SOID = 1
Private Const ITMTYPE = 2
Private Const ITMCODE = 3
Private Const WHSCODE = 4
Private Const LOTNO = 5
Private Const ITMNAME = 6
Private Const GMORE = 7
Private Const WANTED = 8
Private Const PUBLISHER = 9
Private Const QTY = 10
Private Const PRICE = 11
Private Const DisPer = 12
Private Const Dis = 13
Private Const Amt = 14
Private Const NET = 15
Private Const Netl = 16
Private Const Disl = 17
Private Const Amtl = 18
Private Const ITMID = 19
Private Const GDRMKID = 20


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
Private wlLineNo As Long
Private wlRefDocID As Long


Private wlKey As Long
Private wsActNam(4) As String

Private wsConnTime As String
Private Const wsKeyType = "popPOHD"
Private wsFormID As String
Private wsUsrId As String
Private wsTrnCd As String
Private wsDocNo As String

Private wbErr As Boolean
Private wsBaseCurCd As String

Private wsFormCaption As String

Private Sub Ini_Scr()

    Dim MyControl As Control
    
    waResult.ReDim 0, -1, LINENO, GDRMKID
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
    Call SetDateMask(medExpiryDate)
    'Call SetPasswordChar(txtPassword, "*")
    
    
    wsOldVdrNo = ""
    wsOldCurCd = ""
    wsOldRmkCd = ""
    wsOldPayCd = ""
    
    wsOldDelCd = ""
    wsOldShipCd(0) = ""
    wsOldShipCd(1) = ""
    wsOldShipCd(2) = ""

    
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
    
    FocusMe cboDocNo
    tabDetailInfo.Tab = 0
End Sub

Private Sub cboCurr_GotFocus()
    FocusMe cboCurr
End Sub

Private Sub cboCurr_LostFocus()
    FocusMe cboCurr, True
End Sub





Private Sub cboDelCode_GotFocus()
    FocusMe cboDelCode
End Sub

Private Sub cboDelCode_LostFocus()
    FocusMe cboDelCode, True
End Sub

Private Sub cboDelCode_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(cboDelCode, 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboDelCode = False Then
                Exit Sub
        End If
        
        If wsOldDelCd <> cboDelCode.Text Then
            Get_DelName
            wsOldDelCd = cboDelCode.Text
        End If
        
        tabDetailInfo.Tab = 0
        txtDelName.SetFocus
       
    End If
    
End Sub

Private Sub cboDelCode_DropDown()
    
    Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboDelCode
    
    wsSQL = "SELECT WHSCode, WHSDESC FROM mstWarehouse WHERE WhsCode LIKE '%" & IIf(cboDelCode.SelLength > 0, "", Set_Quote(cboDelCode.Text)) & "%' "
    wsSQL = wsSQL & "AND WhsSTATUS = '1' "
    wsSQL = wsSQL & "ORDER BY WhsCode "
    Call Ini_Combo(2, wsSQL, cboDelCode.Left + tabDetailInfo.Left, cboDelCode.Top + cboDelCode.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLWHSCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault

End Sub

Private Function Chk_cboDelCode() As Boolean


    Chk_cboDelCode = False
     
    If Trim(cboDelCode.Text) = "" Then
        Chk_cboDelCode = True
        Exit Function
    End If
    
    If Chk_Whs(cboDelCode, "") = False Then
        gsMsg = "沒有此收貨貨倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboDelCode.SetFocus
       Exit Function
    End If
    
    Chk_cboDelCode = True
    
End Function




Private Sub cboRefDocNo_DropDown()
   
    Dim wsSQL As String

    Me.MousePointer = vbHourglass
    
    Set wcCombo = cboRefDocNo
    
    wsSQL = "SELECT SOHDDOCNO, SOHDDOCDATE , CUSCODE, CUSNAME FROM SOASOHD, mstCUSTOMER "
    wsSQL = wsSQL & " WHERE SOHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND SOHDCUSID = CUSID "
    wsSQL = wsSQL & " AND SOHDDOCNO LIKE '%" & IIf(cboRefDocNo.SelLength > 0, "", Set_Quote(cboRefDocNo.Text)) & "%' "
    wsSQL = wsSQL & " ORDER BY SOHDDOCNO "
                
    Call Ini_Combo(4, wsSQL, cboRefDocNo.Left + tabDetailInfo.Left, cboRefDocNo.Top + cboRefDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSONO", Me.Width, Me.Height)
            
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
        
        
         tabDetailInfo.Tab = 0
         cboCurr.SetFocus
         
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
        
      '  If wsStatus = "4" Then
      '      gsMsg = "文件已入數!"
      '      MsgBox gsMsg, vbOKOnly, gsTitle
      '      Exit Function
      '  End If
     
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
    Else
        gsMsg = "沒有此文件!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        wlRefDocID = 0
        Chk_cboRefDocNo = False
        Exit Function
    End If
    
    wlRefDocID = Get_TableInfo("SOASOHD", "SOHDDOCNO = '" & Set_Quote(cboRefDocNo.Text) & "'", "SOHDDOCID")
    
    Chk_cboRefDocNo = True

End Function


Private Sub cboShipCode_DropDown(Index As Integer)
Dim wsSQL As String
    
    Me.MousePointer = vbHourglass

    Set wcCombo = cboShipCode(Index)
    
    wsSQL = "SELECT ShipCode, ShipName, ShipPer FROM mstShip WHERE ShipCode LIKE '%" & IIf(cboShipCode(Index).SelLength > 0, "", Set_Quote(cboShipCode(Index).Text)) & "%' "
    wsSQL = wsSQL & "AND ShipSTATUS = '1' "
    wsSQL = wsSQL & "AND ShipCardID = " & wlVdrID & " "
    wsSQL = wsSQL & "ORDER BY ShipCode "
    Call Ini_Combo(3, wsSQL, cboShipCode(Index).Left + tabDetailInfo.Left, cboShipCode(Index).Top + cboShipCode(Index).Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLSHIPCOD", Me.Width, Me.Height)
    
    tblCommon.Visible = True
    tblCommon.SetFocus
    Me.MousePointer = vbDefault
End Sub

Private Sub cboShipCode_GotFocus(Index As Integer)
    FocusMe cboShipCode(Index)
End Sub

Private Sub cboShipCode_KeyPress(Index As Integer, KeyAscii As Integer)

   Call chk_InpLen(cboShipCode(Index), 10, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        If Chk_cboShipCode(Index) = False Then
                Exit Sub
        End If
        
        If wsOldShipCd(Index) <> cboShipCode(Index).Text Then
            Call Get_ShipMark(Index)
            wsOldShipCd(Index) = cboShipCode(Index).Text
        End If
        
        tabDetailInfo.Tab = 2
        txtShipPer(Index).SetFocus
       
    End If

End Sub

Private Sub cboShipCode_LostFocus(Index As Integer)
        FocusMe cboShipCode(Index), True
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
  
    If wsFormID = "PN001" Then
  
    wsSQL = "SELECT POHDDOCNO, VDRCODE, VDRNAME, POHDDOCDATE "
    wsSQL = wsSQL & " FROM popPOHD, MstVendor "
    wsSQL = wsSQL & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND POHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND POHDPGMNO  = '" & wsFormID & "' "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO DESC "
    
    Else
    
    wsSQL = "SELECT POHDDOCNO, VDRCODE, VDRNAME, POHDDOCDATE "
    wsSQL = wsSQL & " FROM popPOHD, MstVendor "
    wsSQL = wsSQL & " WHERE POHDDOCNO LIKE '%" & IIf(cboDocNo.SelLength > 0, "", Set_Quote(cboDocNo.Text)) & "%' "
    wsSQL = wsSQL & " AND POHDVDRID  = VDRID "
    wsSQL = wsSQL & " AND POHDSTATUS IN ('1','4') "
    wsSQL = wsSQL & " AND POHDPGMNO  <> 'PN001' "
    wsSQL = wsSQL & " ORDER BY POHDDOCNO DESC "
    
    End If
    
    Call Ini_Combo(4, wsSQL, cboDocNo.Left + tabDetailInfo.Left, cboDocNo.Top + cboDocNo.Height + tabDetailInfo.Top, tblCommon, wsFormID, "TBLDOCNO", Me.Width, Me.Height)
    
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
    Dim wsPgmNo As String
    
    Chk_cboDocNo = False
    
    If Trim(cboDocNo.Text) = "" And Chk_AutoGen(wsTrnCd) = "N" Then
        gsMsg = "必需輸入文件號!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        cboDocNo.SetFocus
        Exit Function
    End If
        
    If Chk_PoHdDocNo(cboDocNo, wsStatus, wsPgmNo) = True Then
        
        If wsStatus = "4" Then
            gsMsg = "文件已入數, 祇可以更新基本資料!"
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
        
        If wsPgmNo = "PN001" Then
        
        If wsFormID <> wsPgmNo Then
            gsMsg = "文件類別不同!不能開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboDocNo.SetFocus
            Exit Function
        End If
        
        Else
        
        If wsFormID = "PN001" Then
            gsMsg = "文件類別不同!不能開啟!"
            MsgBox gsMsg, vbOKOnly, gsTitle
            tabDetailInfo.Tab = 0
            cboDocNo.SetFocus
            Exit Function
        End If
        
        End If
        
        
    End If
    
    Chk_cboDocNo = True
End Function

Private Sub Ini_Scr_AfrKey()
    
    If LoadRecord() = False Then
        wiAction = AddRec
        wiRevNo = Format(0, "##0")
        medDocDate.Text = Dsp_Date(Now)
        'medReserveDate.Text = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(medDocDate.Text))), "yyyy/mm/dd")
        medDueDate = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(medDocDate.Text))), "yyyy/mm/dd")
        medExpiryDate = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(medDocDate.Text))), "yyyy/mm/dd")

        Call SetButtonStatus("AfrKeyAdd")
    Else
        wiAction = CorRec
        If RowLock(wsConnTime, wsKeyType, cboDocNo.Text, wsFormID, wsUsrId) = False Then
            gsMsg = "記錄已被以下使用者鎖定 : " & wsUsrId
            MsgBox gsMsg, vbOKOnly, gsTitle
            tblDetail.ReBind
        End If
        
        wsOldVdrNo = cboVdrCode.Text
        wsOldCurCd = cboCurr.Text
        wsOldRmkCd = cboRmkCode.Text
        wsOldPayCd = cboPayCode.Text
        
        wsOldShipCd(0) = cboShipCode(0).Text
        wsOldShipCd(1) = cboShipCode(1).Text

        wsOldDelCd = cboDelCode.Text
        
        
        
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
    
    If wsFormID = "PN001" Then
    
    
        wsSQL = "SELECT POHDDOCID, POHDDOCNO, POHDREFDOCID,  POHDVDRID, VDRID, VDRCODE, VDRNAME, VDRTEL, VDRFAX, PODTDOCLINE,"
        wsSQL = wsSQL & "POHDDOCDATE, POHDREVNO, POHDCURR, POHDEXCR, POHDSPECDIS, POHDETADATE, "
        wsSQL = wsSQL & "POHDDUEDATE, POHDPAYCODE, POHDPRCCODE, POHDSALEID, POHDMLCODE, "
        
        wsSQL = wsSQL & "POHDLCNO, POHDPORTNO, POHDCUSPO, "
        wsSQL = wsSQL & "POHDDELCODE, POHDDELNAME,  POHDDELADR1, POHDDELADR2, POHDDELADR3,  POHDDELADR4, "
        wsSQL = wsSQL & "POHDSHPFRCODE, POHDSHPFRNAME,  POHDSHPFRADR1, POHDSHPFRADR2, POHDSHPFRADR3,  POHDSHPFRADR4, POHDSHPFRTELNO, POHDSHPFRFAXNO, POHDSHPFRPER, "
        wsSQL = wsSQL & "POHDSHPTOCODE, POHDSHPTONAME,  POHDSHPTOADR1, POHDSHPTOADR2, POHDSHPTOADR3,  POHDSHPTOADR4, POHDSHPTOTELNO, POHDSHPTOFAXNO, POHDSHPTOPER, "
        wsSQL = wsSQL & "POHDSHPVIACODE, POHDSHPVIANAME,  POHDSHPVIAADR1, POHDSHPVIAADR2, POHDSHPVIAADR3,  POHDSHPVIAADR4, POHDSHPVIATELNO, POHDSHPVIAFAXNO, POHDSHPVIAPER, "
        
        wsSQL = wsSQL & "POHDRMKCODE, POHDRMK1,  POHDRMK2,  POHDRMK3,  POHDRMK4, POHDRMK5, "
        wsSQL = wsSQL & "POHDRMK6,  POHDRMK7,  POHDRMK8,  POHDRMK9, POHDRMK10, "
        wsSQL = wsSQL & "POHDGRSAMT , POHDGRSAMTL, POHDDISAMT, POHDDISAMTL, POHDNETAMT, POHDNETAMTL, "
        wsSQL = wsSQL & "PODTITEMID, '' ITMITMTYPECODE, PODTWHSCODE, PODTLOTNO, 'NONSTOCK' ITMCODE, PODTITEMDESC ITNAME, PODTWANTED, PODTQTY, PODTUPRICE, PODTDISPER, PODTAMT, PODTAMTL, PODTDIS, PODTDISL, PODTNET, PODTNETL, PODTDRMKID "
        wsSQL = wsSQL & "FROM  popPOHD, popPODT, MstVendor "
        wsSQL = wsSQL & "WHERE POHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND POHDDOCID = PODTDOCID "
        wsSQL = wsSQL & "AND POHDVDRID = VDRID "
        wsSQL = wsSQL & "AND POHDPGMNO = '" & wsFormID & "' "
        wsSQL = wsSQL & "ORDER BY PODTDOCLINE "
    
    Else
    
        wsSQL = "SELECT POHDDOCID, POHDDOCNO, POHDREFDOCID,  POHDVDRID, VDRID, VDRCODE, VDRNAME, VDRTEL, VDRFAX, PODTDOCLINE,"
        wsSQL = wsSQL & "POHDDOCDATE, POHDREVNO, POHDCURR, POHDEXCR, POHDSPECDIS, POHDETADATE, "
        wsSQL = wsSQL & "POHDDUEDATE, POHDPAYCODE, POHDPRCCODE, POHDSALEID, POHDMLCODE, "
        
        wsSQL = wsSQL & "POHDLCNO, POHDPORTNO, POHDCUSPO, "
        wsSQL = wsSQL & "POHDDELCODE, POHDDELNAME,  POHDDELADR1, POHDDELADR2, POHDDELADR3,  POHDDELADR4, "
        wsSQL = wsSQL & "POHDSHPFRCODE, POHDSHPFRNAME,  POHDSHPFRADR1, POHDSHPFRADR2, POHDSHPFRADR3,  POHDSHPFRADR4, POHDSHPFRTELNO, POHDSHPFRFAXNO, POHDSHPFRPER, "
        wsSQL = wsSQL & "POHDSHPTOCODE, POHDSHPTONAME,  POHDSHPTOADR1, POHDSHPTOADR2, POHDSHPTOADR3,  POHDSHPTOADR4, POHDSHPTOTELNO, POHDSHPTOFAXNO, POHDSHPTOPER, "
        wsSQL = wsSQL & "POHDSHPVIACODE, POHDSHPVIANAME,  POHDSHPVIAADR1, POHDSHPVIAADR2, POHDSHPVIAADR3,  POHDSHPVIAADR4, POHDSHPVIATELNO, POHDSHPVIAFAXNO, POHDSHPVIAPER, "
        
        wsSQL = wsSQL & "POHDRMKCODE, POHDRMK1,  POHDRMK2,  POHDRMK3,  POHDRMK4, POHDRMK5, "
        wsSQL = wsSQL & "POHDRMK6,  POHDRMK7,  POHDRMK8,  POHDRMK9, POHDRMK10, "
        wsSQL = wsSQL & "POHDGRSAMT , POHDGRSAMTL, POHDDISAMT, POHDDISAMTL, POHDNETAMT, POHDNETAMTL, "
        wsSQL = wsSQL & "PODTITEMID, ITMITMTYPECODE, PODTWHSCODE, PODTLOTNO, ITMCODE, PODTITEMDESC ITNAME, PODTWANTED, PODTQTY, PODTUPRICE, PODTDISPER, PODTAMT, PODTAMTL, PODTDIS, PODTDISL, PODTNET, PODTNETL, PODTDRMKID "
        wsSQL = wsSQL & "FROM  popPOHD, popPODT, MstVendor, mstITEM "
        wsSQL = wsSQL & "WHERE POHDDOCNO = '" & Set_Quote(cboDocNo) & "' "
        wsSQL = wsSQL & "AND POHDDOCID = PODTDOCID "
        wsSQL = wsSQL & "AND POHDVDRID = VDRID "
        wsSQL = wsSQL & "AND PODTITEMID = ITMID "
        wsSQL = wsSQL & "AND POHDPGMNO <> 'PN001' "
        wsSQL = wsSQL & "ORDER BY PODTDOCLINE "
    
    End If
    
    rsInvoice.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic

    If rsInvoice.RecordCount <= 0 Then
        rsInvoice.Close
        Set rsInvoice = Nothing
        Exit Function
    End If
    wlKey = ReadRs(rsInvoice, "POHDDOCID")
   ' txtRevNo.Text = Format(ReadRs(rsInvoice, "POHDREVNO") + 1, "##0")
    wiRevNo = To_Value(ReadRs(rsInvoice, "POHDREVNO"))
    medDocDate.Text = ReadRs(rsInvoice, "POHDDOCDATE")
    wlVdrID = ReadRs(rsInvoice, "VDRID")
    cboVdrCode.Text = ReadRs(rsInvoice, "VDRCODE")
    lblDspVdrName.Caption = ReadRs(rsInvoice, "VDRNAME")
    
    wlRefDocID = ReadRs(rsInvoice, "POHDREFDOCID")
    cboRefDocNo.Text = Get_TableInfo("soaSOHD", "SOHDDOCID =" & wlRefDocID, "SOHDDOCNO")
    
    lblDspVdrTel.Caption = ReadRs(rsInvoice, "VDRTEL")
    lblDspVdrFax.Caption = ReadRs(rsInvoice, "VDRFAX")
    cboCurr.Text = ReadRs(rsInvoice, "POHDCURR")

    txtExcr.Text = Format(ReadRs(rsInvoice, "POHDEXCR"), gsExrFmt)
    
    medDueDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "POHDDUEDATE"))
    medExpiryDate.Text = Dsp_MedDate(ReadRs(rsInvoice, "POHDETADATE"))
    
    wlSaleID = To_Value(ReadRs(rsInvoice, "POHDSALEID"))
    
    cboPayCode = ReadRs(rsInvoice, "POHDPAYCODE")
    cboPrcCode = ReadRs(rsInvoice, "POHDPRCCODE")
    cboMLCode = ReadRs(rsInvoice, "POHDMLCODE")
    cboRmkCode = ReadRs(rsInvoice, "POHDRMKCODE")
    
    txtCusPo = ReadRs(rsInvoice, "POHDCUSPO")
    txtLcNo = ReadRs(rsInvoice, "POHDLCNO")
    txtPortNo = ReadRs(rsInvoice, "POHDPORTNO")
    
    txtSpecDis = Format(ReadRs(rsInvoice, "POHDSPECDIS"), gsExrFmt)
    txtDisAmt.Text = Format(To_Value(ReadRs(rsInvoice, "POHDDISAMT")), gsAmtFmt)
    
    
    cboDelCode = ReadRs(rsInvoice, "POHDDELCODE")
    txtDelName = ReadRs(rsInvoice, "POHDDELName")
    txtDelAdr1 = ReadRs(rsInvoice, "POHDDELADR1")
    txtDelAdr2 = ReadRs(rsInvoice, "POHDDELADR2")
    txtDelAdr3 = ReadRs(rsInvoice, "POHDDELADR3")
    txtDelAdr4 = ReadRs(rsInvoice, "POHDDELADR4")
    
    cboShipCode(0) = ReadRs(rsInvoice, "POHDSHPFRCODE")
    txtShipName(0) = ReadRs(rsInvoice, "POHDSHPFRNAME")
    txtShipPer(0) = ReadRs(rsInvoice, "POHDSHPFRPER")
    txtShipAdr1(0) = ReadRs(rsInvoice, "POHDSHPFRADR1")
    txtShipAdr2(0) = ReadRs(rsInvoice, "POHDSHPFRADR2")
    txtShipAdr3(0) = ReadRs(rsInvoice, "POHDSHPFRADR3")
    txtShipAdr4(0) = ReadRs(rsInvoice, "POHDSHPFRADR4")
    txtShipTelNo(0) = ReadRs(rsInvoice, "POHDSHPFRTELNO")
    txtShipFaxNo(0) = ReadRs(rsInvoice, "POHDSHPFRFAXNO")
    
    cboShipCode(1) = ReadRs(rsInvoice, "POHDSHPTOCODE")
    txtShipName(1) = ReadRs(rsInvoice, "POHDSHPTONAME")
    txtShipPer(1) = ReadRs(rsInvoice, "POHDSHPTOPER")
    txtShipAdr1(1) = ReadRs(rsInvoice, "POHDSHPTOADR1")
    txtShipAdr2(1) = ReadRs(rsInvoice, "POHDSHPTOADR2")
    txtShipAdr3(1) = ReadRs(rsInvoice, "POHDSHPTOADR3")
    txtShipAdr4(1) = ReadRs(rsInvoice, "POHDSHPTOADR4")
    txtShipTelNo(1) = ReadRs(rsInvoice, "POHDSHPTOTELNO")
    txtShipFaxNo(1) = ReadRs(rsInvoice, "POHDSHPTOFAXNO")
    

    
    
    Dim i As Integer
    
    For i = 1 To 10
        txtRmk(i) = ReadRs(rsInvoice, "POHDRMK" & i)
    Next i
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspPrcDesc = Get_TableInfo("mstPriceTerm", "PrcCode ='" & Set_Quote(cboPrcCode.Text) & "'", "PRCDESC")
    lblDspMLDesc = Get_TableInfo("mstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    'lblDspNatureDesc = Get_TableInfo("mstNature", "NatureCode ='" & Set_Quote(cboNatureCode.Text) & "'", "NatureDESC")
    
    rsInvoice.MoveFirst
    With waResult
         .ReDim 0, -1, LINENO, GDRMKID
         Do While Not rsInvoice.EOF
             wiCtr = wiCtr + 1
             .AppendRows
             waResult(.UpperBound(1), LINENO) = ReadRs(rsInvoice, "PODTDOCLINE")
             waResult(.UpperBound(1), ITMCODE) = ReadRs(rsInvoice, "ITMCODE")
             waResult(.UpperBound(1), ITMTYPE) = ReadRs(rsInvoice, "ITMITMTYPECODE")
             waResult(.UpperBound(1), ITMNAME) = ReadRs(rsInvoice, "ITNAME")
             waResult(.UpperBound(1), WHSCODE) = ReadRs(rsInvoice, "PODTWHSCODE")
             waResult(.UpperBound(1), LOTNO) = ReadRs(rsInvoice, "PODTLOTNO")
             waResult(.UpperBound(1), PUBLISHER) = ""
             waResult(.UpperBound(1), WANTED) = Dsp_MedDate(ReadRs(rsInvoice, "PODTWANTED"))
             ' Tom 20090203
          '   waResult(.UpperBound(1), Qty) = Format(ReadRs(rsInvoice, "PODTQTY"), gsQtyFmt)
             waResult(.UpperBound(1), QTY) = Format(ReadRs(rsInvoice, "PODTQTY"), gsAmtFmt)
             waResult(.UpperBound(1), PRICE) = Format(ReadRs(rsInvoice, "PODTUPRICE"), gsAmtFmt)
             waResult(.UpperBound(1), DisPer) = Format(ReadRs(rsInvoice, "PODTDISPER"), gsAmtFmt)
             waResult(.UpperBound(1), Amt) = Format(ReadRs(rsInvoice, "PODTAMT"), gsAmtFmt)
             waResult(.UpperBound(1), Amtl) = Format(ReadRs(rsInvoice, "PODTAMTL"), gsAmtFmt)
             waResult(.UpperBound(1), Dis) = Format(ReadRs(rsInvoice, "PODTDIS"), gsAmtFmt)
             waResult(.UpperBound(1), Disl) = Format(ReadRs(rsInvoice, "PODTDISL"), gsAmtFmt)
             waResult(.UpperBound(1), NET) = Format(ReadRs(rsInvoice, "PODTNET"), gsAmtFmt)
             waResult(.UpperBound(1), Netl) = Format(ReadRs(rsInvoice, "PODTNETL"), gsAmtFmt)
             waResult(.UpperBound(1), ITMID) = ReadRs(rsInvoice, "PODTITEMID")
             waResult(.UpperBound(1), GDRMKID) = To_Value(ReadRs(rsInvoice, "PODTDRMKID"))
             waResult(.UpperBound(1), GMORE) = IIf(To_Value(ReadRs(rsInvoice, "PODTDRMKID")) <> 0, "Y", "N")
            
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
Dim wiCtr As Integer


On Error GoTo Ini_Caption_Err

    Call Get_Scr_Item(wsFormID, waScrItm)
    Call Get_Scr_Item("TOOLTIP_M", waScrToolTip)
    
    wsFormCaption = Get_Caption(waScrItm, "SCRHDR")
        
    lblDocNo.Caption = Get_Caption(waScrItm, "DOCNO")
   'lblRevNo.Caption = Get_Caption(waScrItm, "REVNO")
    lblDocDate.Caption = Get_Caption(waScrItm, "DOCDATE")
    lblVdrCode.Caption = Get_Caption(waScrItm, "VDRCODE")
    lblRefDocNo.Caption = Get_Caption(waScrItm, "REFDOCNO")
    
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
    lblExpiryDate.Caption = Get_Caption(waScrItm, "ETADATE")
    
    lblGrsAmtOrg.Caption = Get_Caption(waScrItm, "GRSAMTORG")
    lblNetAmtOrg.Caption = Get_Caption(waScrItm, "NETAMTORG")
    lblDisAmtOrg.Caption = Get_Caption(waScrItm, "DISAMTORG")
    lblTotalQty.Caption = Get_Caption(waScrItm, "TOTALQTY")
    lblSpecDis.Caption = Get_Caption(waScrItm, "SPECDIS")
    lblDisAmt.Caption = Get_Caption(waScrItm, "DISAMTORG")
    'lblPercent.Caption = Get_Caption(waScrItm, "PERCENT")
    
    btnGetDisAmt.Caption = Get_Caption(waScrItm, "GETDISAMT")
    
    
    With tblDetail
        .Columns(LINENO).Caption = Get_Caption(waScrItm, "LINENO")
        .Columns(SOID).Caption = Get_Caption(waScrItm, "SOID")
        .Columns(ITMCODE).Caption = Get_Caption(waScrItm, "ITMCODE")
        .Columns(ITMTYPE).Caption = Get_Caption(waScrItm, "ITMTYPE")
        .Columns(WHSCODE).Caption = Get_Caption(waScrItm, "WHSCODE")
        .Columns(LOTNO).Caption = Get_Caption(waScrItm, "LOTNO")
        .Columns(ITMNAME).Caption = Get_Caption(waScrItm, "ITMNAME")
        .Columns(WANTED).Caption = Get_Caption(waScrItm, "WANTED")
        .Columns(PUBLISHER).Caption = Get_Caption(waScrItm, "PUBLISHER")
        .Columns(QTY).Caption = Get_Caption(waScrItm, "QTY")
        .Columns(PRICE).Caption = Get_Caption(waScrItm, "PRICE")
        .Columns(DisPer).Caption = Get_Caption(waScrItm, "DISPER")
        .Columns(Dis).Caption = Get_Caption(waScrItm, "DIS")
        .Columns(NET).Caption = Get_Caption(waScrItm, "NET")
        .Columns(Amt).Caption = Get_Caption(waScrItm, "AMT")
        .Columns(GMORE).Caption = Get_Caption(waScrItm, "GMORE")
    End With
    
    tabDetailInfo.TabCaption(0) = Get_Caption(waScrItm, "TABDETAILINFO01")
    tabDetailInfo.TabCaption(1) = Get_Caption(waScrItm, "TABDETAILINFO02")
    tabDetailInfo.TabCaption(2) = Get_Caption(waScrItm, "TABDETAILINFO03")
    
    lblDelCode.Caption = Get_Caption(waScrItm, "DelCode")
    lblDelName.Caption = Get_Caption(waScrItm, "DelName")
    lblDelAdr1.Caption = Get_Caption(waScrItm, "DelAdr1")
    
    fraShip(0).Caption = Get_Caption(waScrItm, "SHIPFROM")
    fraShip(1).Caption = Get_Caption(waScrItm, "SHIPTO")

    
    For wiCtr = 0 To 1
    
    lblShipCode(wiCtr).Caption = Get_Caption(waScrItm, "SHIPCODE")
    lblShipName(wiCtr).Caption = Get_Caption(waScrItm, "SHIPNAME")
    lblShipPer(wiCtr).Caption = Get_Caption(waScrItm, "SHIPPER")
    lblShipAdr(wiCtr).Caption = Get_Caption(waScrItm, "SHIPADR")
    lblShipTelNo(wiCtr).Caption = Get_Caption(waScrItm, "SHIPTELNO")
    lblShipFaxNo(wiCtr).Caption = Get_Caption(waScrItm, "SHIPFAXNO")
    
    Next wiCtr
    
    lblCusPo.Caption = Get_Caption(waScrItm, "CUSPO")
    lblLcNo.Caption = Get_Caption(waScrItm, "LCNO")
    lblPortNo.Caption = Get_Caption(waScrItm, "PORTNO")
    
    lblRmkCode.Caption = Get_Caption(waScrItm, "RMKCODE")
    lblRmk.Caption = Get_Caption(waScrItm, "RMK")
    
    chkWorkOrder.Caption = Get_Caption(waScrItm, "WORKORDER")
    
 '   btnITMLST.Caption = Get_Caption(waScrItm, "ITMLIST")
    
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
    
    wsActNam(1) = Get_Caption(waScrItm, "POADD")
    wsActNam(2) = Get_Caption(waScrItm, "POEDIT")
    wsActNam(3) = Get_Caption(waScrItm, "PODELETE")
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
            medExpiryDate.SetFocus
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

Private Function Chk_medExpiryDate() As Boolean
    Chk_medExpiryDate = False
    
    If Trim(medExpiryDate.Text) = "/  /" Then
        gsMsg = "日期錯誤!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 0
        medExpiryDate.SetFocus
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


Private Function Chk_medReserveDate() As Boolean
    
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

Private Sub tabDetailInfo_Click(PreviousTab As Integer)
    If tabDetailInfo.Tab = 0 Then
        
        If cboVdrCode.Enabled Then
            cboVdrCode.SetFocus
        End If
    
    ElseIf tabDetailInfo.Tab = 1 Then
        
        If tblDetail.Enabled Then
            tblDetail.Col = IIf(wsFormID = "PN001", ITMNAME, ITMTYPE)
            tblDetail.SetFocus
        End If
        
    ElseIf tabDetailInfo.Tab = 2 Then
    
        If cboShipCode(0).Enabled Then
            cboShipCode(0).SetFocus
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
    
    Dim rspopPOHD As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT POHDSTATUS FROM popPOHD WHERE POHDDOCNO = '" & Set_Quote(cboDocNo) & "'"
    rspopPOHD.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    If rspopPOHD.RecordCount > 0 Then
        
        Chk_KeyExist = True
    
    Else
        
        Chk_KeyExist = False
    
    End If
    
    rspopPOHD.Close
    Set rspopPOHD = Nothing
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
    '   gsMsg = "已超過信貸額!"
    '   MsgBox gsMsg, vbOKOnly, gsTitle
    '   MousePointer = vbDefault
    '   Exit Function
    'End If
    
    wlRowCtr = waResult.UpperBound(1)
    wsCtlPrd = Left(medDocDate, 4) & Mid(medDocDate, 6, 2)
    
    If wbReadOnly = True Then
    wiAction = CorRO
    End If
    
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
    Call SetSPPara(adcmdSave, 6, wlRefDocID)
    
    Call SetSPPara(adcmdSave, 7, medDocDate.Text)
    Call SetSPPara(adcmdSave, 8, wiRevNo)
    Call SetSPPara(adcmdSave, 9, cboCurr.Text)
    Call SetSPPara(adcmdSave, 10, txtExcr.Text)
    Call SetSPPara(adcmdSave, 11, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medDueDate.Text))
    Call SetSPPara(adcmdSave, 13, Set_MedDate(medExpiryDate.Text))
    
    Call SetSPPara(adcmdSave, 14, wlSaleID)
    
    Call SetSPPara(adcmdSave, 15, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 16, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 17, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 18, txtSpecDis.Text)
    Call SetSPPara(adcmdSave, 19, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 20, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 21, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 22, txtPortNo.Text)
    
    Call SetSPPara(adcmdSave, 23, cboDelCode.Text)
    Call SetSPPara(adcmdSave, 24, txtDelName.Text)
    Call SetSPPara(adcmdSave, 25, txtDelAdr1.Text)
    Call SetSPPara(adcmdSave, 26, txtDelAdr2.Text)
    Call SetSPPara(adcmdSave, 27, txtDelAdr3.Text)
    Call SetSPPara(adcmdSave, 28, txtDelAdr4.Text)
    
    For i = 0 To 1
    Call SetSPPara(adcmdSave, 29 + i * 9, cboShipCode(i).Text)
    Call SetSPPara(adcmdSave, 30 + i * 9, txtShipName(i).Text)
    Call SetSPPara(adcmdSave, 31 + i * 9, txtShipPer(i).Text)
    Call SetSPPara(adcmdSave, 32 + i * 9, txtShipAdr1(i).Text)
    Call SetSPPara(adcmdSave, 33 + i * 9, txtShipAdr2(i).Text)
    Call SetSPPara(adcmdSave, 34 + i * 9, txtShipAdr3(i).Text)
    Call SetSPPara(adcmdSave, 35 + i * 9, txtShipAdr4(i).Text)
    Call SetSPPara(adcmdSave, 36 + i * 9, txtShipTelNo(i).Text)
    Call SetSPPara(adcmdSave, 37 + i * 9, txtShipFaxNo(i).Text)
    Next i
    
    For i = 1 To 10
    Call SetSPPara(adcmdSave, 46 + i, "")
    Next i
    
    
    For i = 1 To 10
    Call SetSPPara(adcmdSave, 56 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdSave, 66, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 67, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 68, lblDspNetAmtOrg)
    
    Call SetSPPara(adcmdSave, 69, wsFormID)
    
    Call SetSPPara(adcmdSave, 70, gsUserID)
    Call SetSPPara(adcmdSave, 71, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 72)
    wsDocNo = GetSPPara(adcmdSave, 73)
    
    If wiAction = AddRec And Trim(cboDocNo.Text) = "" Then cboDocNo.Text = wsDocNo
    
    If wbReadOnly = False Then
    
    If waResult.UpperBound(1) >= 0 Then
        adcmdSave.CommandText = "USP_PO001B"
        adcmdSave.CommandType = adCmdStoredProc
        adcmdSave.Parameters.Refresh
     
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                Call SetSPPara(adcmdSave, 1, wiAction)
                Call SetSPPara(adcmdSave, 2, wlKey)
                Call SetSPPara(adcmdSave, 3, waResult(wiCtr, ITMID))
                Call SetSPPara(adcmdSave, 4, wiCtr + 1)
                Call SetSPPara(adcmdSave, 5, waResult(wiCtr, ITMNAME))
                Call SetSPPara(adcmdSave, 6, waResult(wiCtr, QTY))
                Call SetSPPara(adcmdSave, 7, waResult(wiCtr, PRICE))
                Call SetSPPara(adcmdSave, 8, waResult(wiCtr, DisPer))
                Call SetSPPara(adcmdSave, 9, Set_MedDate(waResult(wiCtr, WANTED)))
                Call SetSPPara(adcmdSave, 10, waResult(wiCtr, WHSCODE))
                Call SetSPPara(adcmdSave, 11, waResult(wiCtr, LOTNO))
                Call SetSPPara(adcmdSave, 12, waResult(wiCtr, Amt))
                Call SetSPPara(adcmdSave, 13, waResult(wiCtr, Amtl))
                Call SetSPPara(adcmdSave, 14, waResult(wiCtr, Dis))
                Call SetSPPara(adcmdSave, 15, waResult(wiCtr, Disl))
                Call SetSPPara(adcmdSave, 16, waResult(wiCtr, NET))
                Call SetSPPara(adcmdSave, 17, waResult(wiCtr, Netl))
                Call SetSPPara(adcmdSave, 18, wlRefDocID)           'SOID
                Call SetSPPara(adcmdSave, 19, waResult(wiCtr, GDRMKID))
                Call SetSPPara(adcmdSave, 20, IIf(wlRowCtr = wiCtr, "Y", "N"))
                Call SetSPPara(adcmdSave, 21, gsUserID)
                Call SetSPPara(adcmdSave, 22, wsGenDte)
                Call SetSPPara(adcmdSave, 23, wsFormID)
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
    Dim i As Integer
    
    InputValidation = False
    
    On Error GoTo InputValidation_Err
    
    
    
   ' If Not chk_txtRevNo Then Exit Function
    If Not Chk_medDocDate Then Exit Function
    If Not chk_cboVdrCode() Then Exit Function
    If Not Chk_cboRefDocNo() Then Exit Function
    
    If Not getExcRate(cboCurr.Text, medDocDate.Text, wsExcRate, wsExcDesc) Then Exit Function
    If Not chk_txtExcr Then Exit Function
    
    If Not Chk_cboSaleCode Then Exit Function
    If Not Chk_cboPayCode Then Exit Function
    If Not Chk_cboPrcCode Then Exit Function
    If Not Chk_cboMLCode Then Exit Function
    
    If Not Chk_medDueDate Then Exit Function
    If Not Chk_medExpiryDate Then Exit Function
    
    If Not Chk_txtSpecDis Then Exit Function
    If Not chk_txtDisAmt Then Exit Function
    
    For i = 0 To 1
    If Not Chk_cboShipCode(i) Then Exit Function
    Next i
    
    If Not Chk_cboDelCode Then Exit Function
    
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
                    tblDetail.Col = IIf(wsFormID = "PN001", ITMNAME, ITMCODE)
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
        gsMsg = "銷售單沒有詳細資料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        If tblDetail.Enabled Then
            tabDetailInfo.Tab = 1
            tblDetail.Col = IIf(wsFormID = "PN001", ITMNAME, ITMCODE)
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
    wsBaseCurCd = Get_CompanyFlag("CMPCURR")
    wsTrnCd = "PO"

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
        '    If Chk_GrdRow(To_Value(.Bookmark)) = False Then
        '        Cancel = True
        '        Exit Sub
        '    End If
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
    End Select
    
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
            cboDelCode.SetFocus
            
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
            cboRefDocNo.SetFocus
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
        gsMsg = "必需輸入客戶編碼!"
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
    wsSQL = wsSQL & "FROM  MstVENDOR "
    wsSQL = wsSQL & "WHERE VDRID = " & wlVdrID
    rsDefVal.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDefVal.RecordCount > 0 Then
        cboCurr.Text = ReadRs(rsDefVal, "VDRCURR")
        cboPayCode.Text = ReadRs(rsDefVal, "VDRPAYCODE")
        cboMLCode.Text = ReadRs(rsDefVal, "VDRMLCODE")
        wlSaleID = ReadRs(rsDefVal, "VDRSALEID")
        txtShipName(0) = ReadRs(rsDefVal, "VDRSHIPNAME")
        txtShipPer(0) = ReadRs(rsDefVal, "VDRSHIPCONTACTPERSON")
        txtShipAdr1(0) = ReadRs(rsDefVal, "VDRSHIPADD1")
        txtShipAdr2(0) = ReadRs(rsDefVal, "VDRSHIPADD2")
        txtShipAdr3(0) = ReadRs(rsDefVal, "VDRSHIPADD3")
        txtShipAdr4(0) = ReadRs(rsDefVal, "VDRSHIPADD4")
     Else
        cboCurr.Text = ""
        cboPayCode.Text = ""
        cboMLCode.Text = ""
        wlSaleID = 0
        txtShipName(0) = ""
        txtShipPer(0) = ""
        txtShipAdr1(0) = ""
        txtShipAdr2(0) = ""
        txtShipAdr3(0) = ""
        txtShipAdr4(0) = ""
        
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
    
    cboDelCode.Text = Get_WorkStation_Info("WSWHSCODE")
    Call Get_DelName
    
    cboSaleCode.Text = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALECODE")
    lblDspSaleDesc = Get_TableInfo("mstSalesman", "SaleID =" & wlSaleID, "SALENAME")
    lblDspPayDesc = Get_TableInfo("mstPayTerm", "PayCode ='" & Set_Quote(cboPayCode.Text) & "'", "PAYDESC")
    lblDspMLDesc = Get_TableInfo("MstMerchClass", "MLCode ='" & Set_Quote(cboMLCode.Text) & "'", "MLDESC")
    
    'get Due Date Payment Term
    medDueDate = Dsp_Date(Get_DueDte(cboPayCode, medDocDate))
    medExpiryDate = Dsp_Date(Get_DueDte(cboPayCode, medDocDate))
End Sub

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
        
        
        For wiCtr = LINENO To GDRMKID
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
                    .Columns(wiCtr).Width = 2000
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 30
                    .Columns(wiCtr).Visible = IIf(wsFormID = "PN001", False, True)
                Case ITMTYPE
                    .Columns(wiCtr).Width = 1500
                    .Columns(wiCtr).DataWidth = 13
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).Visible = IIf(wsFormID = "PN001", False, True)
                Case WHSCODE
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).Button = True
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).Visible = False
                Case LOTNO
                    .Columns(wiCtr).Width = 1000
                    .Columns(wiCtr).DataWidth = 20
                    .Columns(wiCtr).Visible = False
                Case ITMNAME
                    .Columns(wiCtr).Width = IIf(wsFormID = "PN001", 6500, 3500)
                   ' .Columns(wiCtr).DataWidth = 60
                    .Columns(wiCtr).DataWidth = 1000
                    .Columns(wiCtr).Locked = IIf(wsFormID = "PN001", False, True)
                Case WANTED
                    .Columns(wiCtr).Width = 1200
                    .Columns(wiCtr).DataWidth = 10
                    .Columns(wiCtr).EditMask = "####/##/##"
                    .Columns(wiCtr).Visible = False
                Case PUBLISHER
                    .Columns(wiCtr).Width = 3000
                    .Columns(wiCtr).DataWidth = 50
                    .Columns(wiCtr).Locked = True
                    .Columns(wiCtr).Visible = False
                Case QTY
                    .Columns(wiCtr).Width = 800
                    .Columns(wiCtr).HeadAlignment = dbgRight
                    .Columns(wiCtr).Alignment = dbgRight
                    .Columns(wiCtr).DataWidth = 15
                    .Columns(wiCtr).NumberFormat = gsAmtFmt ' Tom 20090203
                Case PRICE
                    .Columns(wiCtr).Width = 800
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
                  '  .Columns(wiCtr).Locked = True
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
                Case SOID
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 5
                    .Columns(wiCtr).Visible = False
                Case ITMID
                    .Columns(wiCtr).DataWidth = 4
                    .Columns(wiCtr).Visible = False
                Case GMORE
                    .Columns(wiCtr).Width = 500
                    .Columns(wiCtr).DataWidth = 2
                    .Columns(wiCtr).Button = True
                Case GDRMKID
                    .Columns(wiCtr).Visible = False
                    .Columns(wiCtr).DataWidth = 10
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

    On Error GoTo tblDetail_BeforeColUpdate_Err
    
    If tblCommon.Visible = True Then
        Cancel = False
        tblDetail.Columns(ColIndex).Text = OldValue
        Exit Sub
    End If
       
    With tblDetail
        Select Case ColIndex
        
            Case ITMTYPE
                If Chk_grdITMTYPE(.Columns(ColIndex).Text) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Trim(.Columns(ITMCODE).Text) = "" Then
                    Call tblDetail_ButtonClick(ITMCODE)
                End If
                
            Case ITMCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_PoExistGrDt(To_Value(.Bookmark)) = False Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
                If Chk_grdITMCODE(.Columns(ColIndex).Text, .Columns(ITMTYPE).Text, wsITMID, wsITMCODE, wsITMNAME, wsPub, wdPrice, wdDisPer, wsLotNo) = False Then
                   GoTo Tbl_BeforeColUpdate_Err
                End If
                
                .Columns(LINENO).Text = wlLineNo
                .Columns(ITMID).Text = wsITMID
                .Columns(ITMNAME).Text = wsITMNAME
                .Columns(PUBLISHER).Text = wsPub
                .Columns(LOTNO).Text = wsLotNo
                .Columns(PRICE).Text = Format(wdPrice, gsAmtFmt)
                .Columns(QTY).Text = ""
                .Columns(DisPer).Text = Format(wdDisPer, gsAmtFmt)
                .Columns(WANTED).Text = medDueDate
                .Columns(WHSCODE).Text = Get_WorkStation_Info("WSWHSCODE")
                .Columns(GMORE).Text = "N"
                
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
        
        
        Case ITMNAME
                
                If Trim(.Columns(LINENO).Text) = "" Then
                .Columns(LINENO).Text = wlLineNo
                .Columns(ITMID).Text = "0"
                .Columns(PUBLISHER).Text = ""
                .Columns(LOTNO).Text = ""
                .Columns(ITMCODE).Text = "NONSTOCK"
                .Columns(PRICE).Text = Format("0", gsAmtFmt)
                .Columns(DisPer).Text = Format("0", gsAmtFmt)
                .Columns(WANTED).Text = medDueDate
                .Columns(GMORE).Text = "N"
                wlLineNo = wlLineNo + 1
                
                End If
        
             Case WHSCODE
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
              '  If Chk_grdWhsCode(.Columns(ColIndex).Text) = False Then
              '          GoTo Tbl_BeforeColUpdate_Err
              '  End If
                
             Case LOTNO
                If Not Chk_NoDup(.Row + To_Value(.FirstRow)) Then
                    GoTo Tbl_BeforeColUpdate_Err
                End If
                
             '   If Chk_grdLotNo(.Columns(ColIndex).Text) = False Then
             '           GoTo Tbl_BeforeColUpdate_Err
             '   End If
            
            Case WANTED
                If Chk_grdWantedDate(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                End If
                
            Case QTY, PRICE, DisPer
            
                If ColIndex = QTY Then
                        If Chk_grdQty(.Columns(ColIndex).Text) = False Then
                        GoTo Tbl_BeforeColUpdate_Err
                        End If
                      
                        If Chk_PoExistGrDtQty(To_Value(.Bookmark), .Columns(ColIndex).Text) = False Then
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
    Dim wsRmkID As String
    
    
    On Error GoTo tblDetail_ButtonClick_Err
    

    With tblDetail
        Select Case ColIndex
        
        
            Case ITMTYPE
            
            wsSQL = "SELECT ITMTYPECODE, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC") & " ITNAME "
            wsSQL = wsSQL & " FROM MSTITEMTYPE,mstITEM, mstVdrItem "
            wsSQL = wsSQL & " WHERE ITMTYPECODE LIKE '%" & Set_Quote(.Columns(ITMTYPE).Text) & "%' "
            wsSQL = wsSQL & " AND ITMTYPESTATUS  <> '2' "
            wsSQL = wsSQL & " AND ITMINACTIVE = 'N' "
            wsSQL = wsSQL & " AND ITMID = VDRITEMITMID "
            wsSQL = wsSQL & " AND ITMTYPECODE = ITMITMTYPECODE "
            wsSQL = wsSQL & " AND VDRITEMCURR = '" & Set_Quote(cboCurr.Text) & "' "
            wsSQL = wsSQL & " AND VDRITEMVDRID = " & To_Value(wlVdrID) & " "
            wsSQL = wsSQL & " GROUP BY ITMTYPECODE, " & IIf(gsLangID = "1", "ITMTYPEENGDESC", "ITMTYPECHIDESC") & " "
            
             Call Ini_Combo(2, wsSQL, .Columns(ColIndex).Left + .Left + .Columns(ColIndex).Top + tabDetailInfo.Left, .Top + .RowTop(.Row) + .RowHeight + tabDetailInfo.Top, tblCommon, wsFormID, "TBLITMTYPE", Me.Width, Me.Height)
             tblCommon.Visible = True
             tblCommon.SetFocus
             Set wcCombo = tblDetail
    
        
            Case ITMCODE
                
                wsSQL = "SELECT ITMCODE, ITMBARCODE, ITMENGNAME, ITMCHINAME "
                wsSQL = wsSQL & " FROM mstITEM, mstVdrItem "
                wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' "
                wsSQL = wsSQL & " AND VDRITEMSTATUS <> '2' "
                wsSQL = wsSQL & " AND ITMINACTIVE = 'N' "
                wsSQL = wsSQL & " AND ITMCODE LIKE '%" & Set_Quote(.Columns(ITMCODE).Text) & "%' "
                wsSQL = wsSQL & " AND ITMITMTYPECODE = '" & Set_Quote(.Columns(ITMTYPE).Text) & "' "
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
                
            Case GMORE
                
                 
                    frmDocRemark.RmkID = IIf(.Columns(GDRMKID).Text = "", "0", .Columns(GDRMKID).Text)
                    frmDocRemark.RmkType = "PO"
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
            If Chk_PoExistGrDt(To_Value(.Bookmark)) = False Then Exit Sub
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
                    .Col = IIf(wsFormID = "PN001", ITMNAME, ITMTYPE)
                Case ITMTYPE
                    KeyCode = vbDefault
                    .Col = ITMCODE
                Case ITMCODE
                    KeyCode = vbDefault
                    .Col = GMORE
                Case GMORE
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
                Case ITMNAME
                    KeyCode = vbDefault
                    .Col = GMORE
                Case GMORE
                     KeyCode = vbDefault
                    .Col = QTY
                Case NET
                    KeyCode = vbKeyDown
                    .Col = ITMCODE
            End Select
        Case vbKeyLeft
            KeyCode = vbDefault
            Select Case .Col
                Case NET
                    .Col = DisPer
                Case DisPer
                    .Col = PRICE
                Case PRICE
                    .Col = QTY
                Case QTY
                    .Col = GMORE
                Case GMORE
                    .Col = ITMNAME
                Case ITMNAME
                    .Col = ITMCODE
                Case ITMCODE
                    .Col = ITMTYPE
            End Select
            
        Case vbKeyRight
            KeyCode = vbDefault
            Select Case .Col
                Case DisPer
                    .Col = NET
                Case PRICE
                    .Col = DisPer
                Case QTY
                    .Col = PRICE
                Case ITMNAME
                    .Col = GMORE
                Case ITMNAME
                    .Col = GMORE
                Case GMORE
                    .Col = QTY
                Case LINENO
                    .Col = IIf(wsFormID = "PN001", ITMNAME, ITMTYPE)
                Case ITMTYPE
                    .Col = ITMCODE
                    
            End Select
        End Select
    End With

    Exit Sub
    
tblDetail_KeyDown_Err:
    MsgBox "Check tblDeiail KeyDown"

End Sub

Private Sub tblDetail_KeyPress(KeyAscii As Integer)
    
    Select Case tblDetail.Col
       ' Tom 20090203
       ' Case Qty
        '    Call Chk_InpNum(KeyAscii, tblDetail.Text, False, False)
        
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
           .Col = IIf(wsFormID = "PN001", ITMNAME, ITMTYPE)
        End If
        
        Call Calc_Total
        
        If Trim(.Columns(.Col).Text) <> "" Then
            Select Case .Col
                Case ITMTYPE
                    Call Chk_grdITMTYPE(.Columns(ITMTYPE))
                Case ITMCODE
                    Call Chk_grdITMCODE(.Columns(ITMCODE).Text, .Columns(ITMTYPE).Text, "", "", "", "", 0, 0, "")
                Case WHSCODE
                  '  Call Chk_grdWhsCode(.Columns(WHSCODE).Text)
                 Case LOTNO
                  '  Call Chk_grdLotNo(.Columns(LOTNO).Text)
                Case WANTED
                    Call Chk_grdWantedDate(.Columns(WANTED).Text)
                Case QTY
               '     Call Chk_grdQty(.Columns(Qty).Text)
                Case PRICE
                '    Call Chk_grdUPrice(.Columns(Price).Text)
                 
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

Private Function Chk_grdITMCODE(inAccNo As String, inItmType As String, outAccID As String, outAccNo As String, OutName As String, outPub As String, outPrice As Double, outDisPer As Double, outLotNo As String) As Boolean
    
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    
    
    If wsFormID = "PN001" Then
        Chk_grdITMCODE = True
        Exit Function
    End If
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
        Exit Function
    End If
    

   wsSQL = "SELECT ITMID, ITMCODE, " & IIf(gsLangID = "1", "ITMENGNAME", "ITMCHINAME") & " ITMNAME, VDRITEMCOST "
   wsSQL = wsSQL & " FROM mstITEM, mstVdrItem "
   wsSQL = wsSQL & " WHERE ITMSTATUS <> '2' "
   wsSQL = wsSQL & " AND VDRITEMSTATUS <> '2' "
   wsSQL = wsSQL & " AND ITMINACTIVE = 'N' "
   wsSQL = wsSQL & " AND ITMCODE = '" & Set_Quote(inAccNo) & "' "
   wsSQL = wsSQL & " AND ITMITMTYPECODE = '" & Set_Quote(inItmType) & "' "
   wsSQL = wsSQL & " AND ITMID = VDRITEMITMID "
   wsSQL = wsSQL & " AND VDRITEMCURR = '" & Set_Quote(cboCurr.Text) & "' "
   wsSQL = wsSQL & " AND VDRITEMVDRID = " & To_Value(wlVdrID) & " "

    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        outAccID = ReadRs(rsDes, "ITMID")
        outAccNo = ReadRs(rsDes, "ITMCODE")
        OutName = ReadRs(rsDes, "ITMNAME")
        outPub = ""
        outPrice = To_Value(ReadRs(rsDes, "VDRITEMCOST"))
        outLotNo = ""
        outDisPer = 0
        
        Chk_grdITMCODE = True
    Else
        outAccID = ""
        OutName = ""
        outPub = ""
        outPrice = 0
        outDisPer = 0
        outLotNo = ""
        gsMsg = "沒有此物料!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMCODE = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing

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
    
    Dim wsSQL As String
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


Private Function Chk_grdUPrice(inCode As String) As Boolean
    
    Chk_grdUPrice = True
    
    If Trim(inCode) = "" Then
        gsMsg = "必需輸入單價!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdUPrice = False
        Exit Function
    End If

  ' Tom 20090203
  '  If To_Value(inCode) = 0 Then
  '      gsMsg = "單價必需大於零!"
  '      MsgBox gsMsg, vbOKOnly, gsTitle
  '      Chk_grdUPrice = False
  '      Exit Function
  '  End If

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
                If wsFormID = "PN001" Then
                If Trim(.Columns(ITMNAME)) = "" Then
                    Exit Function
                End If
                Else
                If Trim(.Columns(ITMTYPE)) = "" Then
                    Exit Function
                End If
                End If
            End With
        Else
            If waResult.UpperBound(1) >= 0 Then
                If Trim(waResult(inRow, ITMTYPE)) = "" And _
                   Trim(waResult(inRow, ITMCODE)) = "" And _
                   Trim(waResult(inRow, ITMNAME)) = "" And _
                   Trim(waResult(inRow, GMORE)) = "" And _
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
        
        If Chk_grdITMTYPE(waResult(LastRow, ITMTYPE)) = False Then
            .Col = ITMTYPE
            .Row = LastRow
            Exit Function
        End If
        
        
        If Chk_grdITMCODE(waResult(LastRow, ITMCODE), waResult(LastRow, ITMTYPE), "", "", "", "", 0, 0, "") = False Then
            .Col = ITMCODE
            .Row = LastRow
            Exit Function
        End If
        
      '  If Chk_grdWhsCode(waResult(LastRow, WHSCODE)) = False Then
      '          .Col = WHSCODE
      '          .Row = LastRow
      '          Exit Function
      '  End If
        
      '  If Chk_grdLotNo(waResult(LastRow, LOTNO)) = False Then
      '          .Col = LOTNO
      '          .Row = LastRow
      '          Exit Function
      '  End If
        
        If Chk_grdWantedDate(waResult(LastRow, WANTED)) = False Then
                .Col = WANTED
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_grdQty(waResult(LastRow, QTY)) = False Then
                .Col = QTY
                .Row = LastRow
                Exit Function
        End If
        
        If Chk_PoExistGrDtQty(LastRow, waResult(LastRow, QTY)) = False Then
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
    Dim adcmdSave As New ADODB.Command

    Dim i As Integer
    Dim wsCtlPrd As String
    
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
    Call SetSPPara(adcmdSave, 6, wlRefDocID)
    
    Call SetSPPara(adcmdSave, 7, medDocDate.Text)
    Call SetSPPara(adcmdSave, 8, wiRevNo)
    Call SetSPPara(adcmdSave, 9, cboCurr.Text)
    Call SetSPPara(adcmdSave, 10, txtExcr.Text)
    Call SetSPPara(adcmdSave, 11, wsCtlPrd)
    
    Call SetSPPara(adcmdSave, 12, Set_MedDate(medDueDate.Text))
    Call SetSPPara(adcmdSave, 13, Set_MedDate(medExpiryDate.Text))
    
    Call SetSPPara(adcmdSave, 14, wlSaleID)
    
    Call SetSPPara(adcmdSave, 15, cboPayCode.Text)
    Call SetSPPara(adcmdSave, 16, cboPrcCode.Text)
    Call SetSPPara(adcmdSave, 17, cboMLCode.Text)
    Call SetSPPara(adcmdSave, 18, txtSpecDis.Text)
    Call SetSPPara(adcmdSave, 19, cboRmkCode.Text)
    
    Call SetSPPara(adcmdSave, 20, txtCusPo.Text)
    Call SetSPPara(adcmdSave, 21, txtLcNo.Text)
    Call SetSPPara(adcmdSave, 22, txtPortNo.Text)
    
    Call SetSPPara(adcmdSave, 23, cboDelCode.Text)
    Call SetSPPara(adcmdSave, 24, txtDelName.Text)
    Call SetSPPara(adcmdSave, 25, txtDelAdr1.Text)
    Call SetSPPara(adcmdSave, 26, txtDelAdr2.Text)
    Call SetSPPara(adcmdSave, 27, txtDelAdr3.Text)
    Call SetSPPara(adcmdSave, 28, txtDelAdr4.Text)
    
    For i = 0 To 1
    Call SetSPPara(adcmdSave, 29 + i * 9, cboShipCode(i).Text)
    Call SetSPPara(adcmdSave, 30 + i * 9, txtShipName(i).Text)
    Call SetSPPara(adcmdSave, 31 + i * 9, txtShipPer(i).Text)
    Call SetSPPara(adcmdSave, 32 + i * 9, txtShipAdr1(i).Text)
    Call SetSPPara(adcmdSave, 33 + i * 9, txtShipAdr2(i).Text)
    Call SetSPPara(adcmdSave, 34 + i * 9, txtShipAdr3(i).Text)
    Call SetSPPara(adcmdSave, 35 + i * 9, txtShipAdr4(i).Text)
    Call SetSPPara(adcmdSave, 36 + i * 9, txtShipTelNo(i).Text)
    Call SetSPPara(adcmdSave, 37 + i * 9, txtShipFaxNo(i).Text)
    Next i
    
    For i = 1 To 10
    Call SetSPPara(adcmdSave, 46 + i, "")
    Next i
    
    For i = 1 To 10
    Call SetSPPara(adcmdSave, 56 + i - 1, txtRmk(i).Text)
    Next
    
    Call SetSPPara(adcmdSave, 66, lblDspGrsAmtOrg)
    Call SetSPPara(adcmdSave, 67, lblDspDisAmtOrg)
    Call SetSPPara(adcmdSave, 68, lblDspNetAmtOrg)
    
    Call SetSPPara(adcmdSave, 69, wsFormID)
    
    Call SetSPPara(adcmdSave, 70, gsUserID)
    Call SetSPPara(adcmdSave, 71, wsGenDte)
    adcmdSave.Execute
    wlKey = GetSPPara(adcmdSave, 72)
    wsDocNo = GetSPPara(adcmdSave, 73)
      
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
                .Buttons(tcEdit).Enabled = False
                .Buttons(tcDelete).Enabled = False
                .Buttons(tcSave).Enabled = True
                .Buttons(tcCancel).Enabled = True
                .Buttons(tcFind).Enabled = False
                .Buttons(tcExit).Enabled = True
                .Buttons(tcRefresh).Enabled = False
                .Buttons(tcPrint).Enabled = True
                .Buttons(tcRevise).Enabled = False
            End With
    End Select
End Sub

'-- Set field status, Default, Add, Edit.
Public Sub SetFieldStatus(ByVal sStatus As String)
Dim i As Integer

    Select Case sStatus
        Case "Default"
        
            Me.cboDocNo.Enabled = False
            Me.cboVdrCode.Enabled = False
            Me.cboRefDocNo.Enabled = False
            
            Me.medDocDate.Enabled = False
            Me.cboCurr.Enabled = False
            Me.txtExcr.Enabled = False
            
            Me.medDueDate.Enabled = False
            Me.medExpiryDate.Enabled = False
            Me.cboSaleCode.Enabled = False
            Me.cboPayCode.Enabled = False
            Me.cboPrcCode.Enabled = False
            Me.cboMLCode.Enabled = False
            Me.cboRmkCode.Enabled = False
            
            Me.cboDelCode.Enabled = False
            Me.txtDelName.Enabled = False
            Me.txtDelAdr1.Enabled = False
            Me.txtDelAdr2.Enabled = False
            Me.txtDelAdr3.Enabled = False
            Me.txtDelAdr4.Enabled = False
            
            For i = 0 To 1
            Me.cboShipCode(i).Enabled = False
            Me.txtShipName(i).Enabled = False
            Me.txtShipAdr1(i).Enabled = False
            Me.txtShipAdr2(i).Enabled = False
            Me.txtShipAdr3(i).Enabled = False
            Me.txtShipAdr4(i).Enabled = False
            Me.txtShipPer(i).Enabled = False
            Me.txtShipTelNo(i).Enabled = False
            Me.txtShipFaxNo(i).Enabled = False
            Next i
            
            Me.txtCusPo.Enabled = False
            Me.txtLcNo.Enabled = False
            Me.txtPortNo.Enabled = False
            
            Me.picRmk.Enabled = False
            
            Me.tblDetail.Enabled = False
            Me.txtSpecDis.Enabled = False
            Me.txtDisAmt.Enabled = False
            Me.btnGetDisAmt.Enabled = False
            
            Me.chkWorkOrder.Enabled = False
            
             
            
             
        Case "AfrActAdd"
        
            Me.cboDocNo.Enabled = True
       
       Case "AfrActEdit"
       
            Me.cboDocNo.Enabled = True
        
        Case "AfrKey"
            Me.cboDocNo.Enabled = False
            
            Me.cboVdrCode.Enabled = True
            Me.cboRefDocNo.Enabled = True
            
           Me.medDocDate.Enabled = True
            Me.cboCurr.Enabled = True
            Me.txtExcr.Enabled = True
            
            Me.medDueDate.Enabled = True
            Me.medExpiryDate.Enabled = True
            Me.cboSaleCode.Enabled = True
            Me.cboPayCode.Enabled = True
            Me.cboPrcCode.Enabled = True
            Me.cboMLCode.Enabled = True
            Me.cboRmkCode.Enabled = True
            
            Me.cboDelCode.Enabled = True
            Me.txtDelName.Enabled = True
            Me.txtDelAdr1.Enabled = True
            Me.txtDelAdr2.Enabled = True
            Me.txtDelAdr3.Enabled = True
            Me.txtDelAdr4.Enabled = True
            
            For i = 0 To 1
            Me.cboShipCode(i).Enabled = True
            Me.txtShipName(i).Enabled = True
            Me.txtShipAdr1(i).Enabled = True
            Me.txtShipAdr2(i).Enabled = True
            Me.txtShipAdr3(i).Enabled = True
            Me.txtShipAdr4(i).Enabled = True
            Me.txtShipPer(i).Enabled = True
            Me.txtShipTelNo(i).Enabled = True
            Me.txtShipFaxNo(i).Enabled = True
            Next i
            
            Me.txtCusPo.Enabled = True
            Me.txtLcNo.Enabled = True
            Me.txtPortNo.Enabled = True
            
            Me.picRmk.Enabled = True
            Me.txtSpecDis.Enabled = True
            Me.txtDisAmt.Enabled = True
            Me.btnGetDisAmt.Enabled = True
            
            Me.chkWorkOrder.Enabled = True
            
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
    Dim wsSQL As String
    
    ReDim vFilterAry(3, 2)
    vFilterAry(1, 1) = "Doc No."
    vFilterAry(1, 2) = "PoHdDocNo"
    
    vFilterAry(2, 1) = "Doc. Date"
    vFilterAry(2, 2) = "PoHdDocDate"
    
    vFilterAry(3, 1) = "Vendor #"
    vFilterAry(3, 2) = "VdrCode"
    
    ReDim vAry(4, 3)
    vAry(1, 1) = "Doc No."
    vAry(1, 2) = "PoHdDocNo"
    vAry(1, 3) = "1500"
    
    vAry(2, 1) = "Date"
    vAry(2, 2) = "PoHdDocDate"
    vAry(2, 3) = "1500"
    
    vAry(3, 1) = "Vendor#"
    vAry(3, 2) = "VdrCode"
    vAry(3, 3) = "2000"
    
    vAry(4, 1) = "Vendor Name"
    vAry(4, 2) = "VdrName"
    vAry(4, 3) = "5000"
    
    
    Me.MousePointer = vbHourglass
    With frmShareSearch
        wsSQL = "SELECT popPOHD.PoHdDocNo, popPOHD.PoHdDocDate, MstVendor.VdrCode,  MstVendor.VdrName "
        wsSQL = wsSQL + "FROM MstVendor, popPOHD "
        .sBindSQL = wsSQL
        .sBindWhereSQL = "WHERE popPOHD.PoHdStatus = '1' And popPOHD.PoHdVdrID = MstVendor.VdrID "
        .sBindOrderSQL = "ORDER BY popPOHD.PoHdDocNo"
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
        medDueDate.SetFocus
       
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

Private Sub txtDelName_GotFocus()
    FocusMe txtDelName
End Sub

Private Sub txtDelName_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtDelName, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 0
        txtDelAdr1.SetFocus
       
    End If
    
End Sub

Private Sub txtDelName_LostFocus()
    FocusMe txtDelName, True
End Sub

Private Sub txtDelAdr1_GotFocus()
    FocusMe txtDelAdr1
End Sub

Private Sub txtDelAdr1_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtDelAdr1, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtDelAdr2.SetFocus
       
    End If
    
End Sub

Private Sub txtDelAdr1_LostFocus()
    FocusMe txtDelAdr1, True
End Sub


Private Sub txtDelAdr2_GotFocus()
    FocusMe txtDelAdr2
End Sub

Private Sub txtDelAdr2_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtDelAdr2, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtDelAdr3.SetFocus
       
    End If
    
End Sub

Private Sub txtDelAdr2_LostFocus()
    FocusMe txtDelAdr2, True
End Sub


Private Sub txtDelAdr3_GotFocus()
    FocusMe txtDelAdr3
End Sub

Private Sub txtDelAdr3_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtDelAdr3, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtDelAdr4.SetFocus
       
    End If
    
End Sub

Private Sub txtDelAdr3_LostFocus()
    FocusMe txtDelAdr3, True
End Sub

Private Sub txtDelAdr4_GotFocus()
    FocusMe txtDelAdr4
End Sub

Private Sub txtDelAdr4_KeyPress(KeyAscii As Integer)
    
    Call chk_InpLen(txtDelAdr4, 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 0
        txtCusPo.SetFocus
       
    End If
    
End Sub

Private Sub txtDelAdr4_LostFocus()
    FocusMe txtDelAdr4, True
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
            tblDetail.Col = IIf(wsFormID = "PN001", ITMNAME, ITMCODE)
            tblDetail.SetFocus
        End If
    End If
    
End Sub

Private Sub txtPortNo_LostFocus()
    FocusMe txtPortNo, True
End Sub


Private Function Chk_cboShipCode(inIndex As Integer) As Boolean

    Chk_cboShipCode = False
     
    If Trim(cboShipCode(inIndex).Text) = "" Then
        Chk_cboShipCode = True
        Exit Function
    End If
    
    If Chk_Ship(cboShipCode(inIndex)) = False Then
        gsMsg = "沒有此貨運編碼!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        tabDetailInfo.Tab = 2
        cboShipCode(inIndex).SetFocus
       Exit Function
    End If
    
    Chk_cboShipCode = True
    
End Function


Private Sub txtShipName_GotFocus(Index As Integer)
        FocusMe txtShipName(Index)
End Sub

Private Sub txtShipName_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipName(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 2
        txtShipAdr1(Index).SetFocus
        
    End If
    
End Sub

Private Sub txtShipName_LostFocus(Index As Integer)
    FocusMe txtShipName(Index), True
End Sub

Private Sub txtShipPer_GotFocus(Index As Integer)
    FocusMe txtShipPer(Index)
End Sub

Private Sub txtShipPer_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipPer(Index), 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtShipName(Index).SetFocus
       
    End If
    
End Sub

Private Sub txtShipPer_LostFocus(Index As Integer)
    FocusMe txtShipPer(Index), True
End Sub

Private Sub txtShipAdr1_GotFocus(Index As Integer)
    FocusMe txtShipAdr1(Index)
End Sub

Private Sub txtShipAdr1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipAdr1(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtShipAdr2(Index).SetFocus
       
    End If
    
End Sub

Private Sub txtShipAdr1_LostFocus(Index As Integer)
    FocusMe txtShipAdr1(Index), True
End Sub

Private Sub txtShipAdr2_GotFocus(Index As Integer)
    FocusMe txtShipAdr2(Index)
End Sub

Private Sub txtShipAdr2_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipAdr2(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtShipAdr3(Index).SetFocus
       
    End If
    
End Sub

Private Sub txtShipAdr2_LostFocus(Index As Integer)
    FocusMe txtShipAdr2(Index), True
End Sub

Private Sub txtShipAdr3_GotFocus(Index As Integer)
    FocusMe txtShipAdr3(Index)
End Sub

Private Sub txtShipAdr3_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipAdr3(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        tabDetailInfo.Tab = 2
        txtShipAdr4(Index).SetFocus
       
    End If
    
End Sub

Private Sub txtShipAdr3_LostFocus(Index As Integer)
    FocusMe txtShipAdr3(Index), True
End Sub

Private Sub txtShipAdr4_GotFocus(Index As Integer)
    FocusMe txtShipAdr4(Index)
End Sub

Private Sub txtShipAdr4_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipAdr4(Index), 50, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 2
        txtShipTelNo(Index).SetFocus
        
    End If
    
End Sub

Private Sub txtShipAdr4_LostFocus(Index As Integer)
    FocusMe txtShipAdr4(Index), True
End Sub




Private Sub txtShipTelNo_GotFocus(Index As Integer)
    FocusMe txtShipTelNo(Index)
End Sub

Private Sub txtShipTelNo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipTelNo(Index), 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 2
        txtShipFaxNo(Index).SetFocus
        
    End If
    
End Sub

Private Sub txtShipTelNo_LostFocus(Index As Integer)
    FocusMe txtShipTelNo(Index), True
End Sub




Private Sub txtShipFaxNo_GotFocus(Index As Integer)
    FocusMe txtShipFaxNo(Index)
End Sub

Private Sub txtShipFaxNo_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Call chk_InpLen(txtShipFaxNo(Index), 20, KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = vbDefault
        
        
        tabDetailInfo.Tab = 2
        If Index = 1 Then
        cboRmkCode.SetFocus
        Else
        cboShipCode(Index + 1).SetFocus
        End If
        
    End If
    
End Sub

Private Sub txtShipFaxNo_LostFocus(Index As Integer)
    FocusMe txtShipFaxNo(Index), True
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

Private Sub Get_ShipMark(inIndex As Integer)
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM  mstShip "
    wsSQL = wsSQL & "WHERE ShipCode = '" & Set_Quote(cboShipCode(inIndex)) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        
        txtShipName(inIndex) = ReadRs(rsRcd, "SHIPNAME")
        txtShipPer(inIndex) = ReadRs(rsRcd, "SHIPPER")
        txtShipAdr1(inIndex) = ReadRs(rsRcd, "SHIPADR1")
        txtShipAdr2(inIndex) = ReadRs(rsRcd, "SHIPADR2")
        txtShipAdr3(inIndex) = ReadRs(rsRcd, "SHIPADR3")
        txtShipAdr4(inIndex) = ReadRs(rsRcd, "SHIPADR4")
        txtShipTelNo(inIndex) = ReadRs(rsRcd, "SHIPTELNO")
        txtShipFaxNo(inIndex) = ReadRs(rsRcd, "SHIPFAXNO")
        
    Else
        txtShipName(inIndex) = ""
        txtShipPer(inIndex) = ""
        txtShipAdr1(inIndex) = ""
        txtShipAdr2(inIndex) = ""
        txtShipAdr3(inIndex) = ""
        txtShipAdr4(inIndex) = ""
        txtShipTelNo(inIndex) = ""
        txtShipFaxNo(inIndex) = ""
        
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

Private Sub Get_DelName()
    
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim i As Integer
    
    wsSQL = "SELECT * "
    wsSQL = wsSQL & "FROM  mstWarehouse "
    wsSQL = wsSQL & "WHERE WhsCode = '" & Set_Quote(cboDelCode) & "'"
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsRcd.RecordCount > 0 Then
        
        txtDelName = ReadRs(rsRcd, "WHSDESC")
    Else
        txtDelName = ""
        
    End If
    rsRcd.Close
    Set rsRcd = Nothing
End Sub

Private Function Chk_NoDup(inRow As Long) As Boolean
    
    Dim wlCtr As Long
    Dim wsCurRec As String
    Dim wsCurRecLn As String
    Dim wsCurRecLn2 As String
    
    Chk_NoDup = False
    
    wsCurRec = tblDetail.Columns(ITMCODE)
    wsCurRecLn = tblDetail.Columns(WHSCODE)
    wsCurRecLn2 = tblDetail.Columns(LOTNO)
   
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           If wsCurRec = waResult(wlCtr, ITMCODE) And _
              wsCurRecLn = waResult(wlCtr, WHSCODE) And _
              wsCurRecLn2 = waResult(wlCtr, LOTNO) Then
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
            If .EditActive = True Then Exit Sub
            If Chk_PoExistGrDt(To_Value(.Bookmark)) = False Then Exit Sub
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

Private Sub cmdPrint()
    Dim wsDteTim As String
    Dim wsSQL As String
    Dim wsSelection() As String
    Dim NewfrmPrint As New frmPrint
    Dim wsRptName As String
    Dim wsTitle As String
    
    
    'If InputValidation = False Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    'Create Selection Criteria
    ReDim wsSelection(4)
    wsSelection(1) = ""
    wsSelection(2) = ""
    wsSelection(3) = ""
    wsSelection(4) = ""
    
    If chkWorkOrder.Value = 1 Then
        If gsLangID = "2" Then
        wsTitle = "工作單"
        Else
        wsTitle = "WORK ORDER"
        End If
    Else
        wsTitle = wgsTitle
    End If
    
    
     
    'Create Stored Procedure String
    wsDteTim = Now
    wsSQL = "EXEC usp_RPTPO002 '" & Set_Quote(gsUserID) & "', "
    wsSQL = wsSQL & "'" & Change_SQLDate(wsDteTim) & "', "
    wsSQL = wsSQL & "'" & wsTitle & "', "
    wsSQL = wsSQL & "'" & wsTitle & "', "
    wsSQL = wsSQL & "'" & wsTrnCd & "', "
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
    wsRptName = "C" + "RPTPO002"
    Else
    wsRptName = "RPTPO002"
    End If
    
    NewfrmPrint.ReportID = "PO002"
    NewfrmPrint.RptTitle = Me.Caption
    NewfrmPrint.TableID = "PO002"
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

    Me.MousePointer = vbHourglass
    
    If waResult.UpperBound(1) >= 0 Then
        
        For wiCtr = 0 To waResult.UpperBound(1)
            If Trim(waResult(wiCtr, ITMCODE)) <> "" Then
                wsITMID = waResult(wiCtr, ITMID)
                wdDisPer = waResult(wiCtr, DisPer)
                'wdNewDisPer = Get_SaleDiscount(cboNatureCode.Text, wlVdrID, wsITMID)
                wdNewDisPer = To_Value(txtSpecDis)
            '    If wdDisPer <> wdNewDisPer Then
                waResult(wiCtr, DisPer) = Format(wdNewDisPer, gsAmtFmt)
                waResult(wiCtr, Dis) = Format(To_Value(waResult(wiCtr, Amt)) * To_Value(waResult(wiCtr, DisPer)) / 100, gsAmtFmt)
                waResult(wiCtr, Disl) = Format(To_Value(waResult(wiCtr, Amtl)) * To_Value(waResult(wiCtr, DisPer)) / 100, gsAmtFmt)
                waResult(wiCtr, NET) = Format(To_Value(waResult(wiCtr, Amt)) * (1 - (To_Value(waResult(wiCtr, DisPer)) / 100)), gsAmtFmt)
                waResult(wiCtr, Netl) = Format(To_Value(waResult(wiCtr, Amtl)) * (1 - (To_Value(waResult(wiCtr, DisPer)) / 100)), gsAmtFmt)
            '    End If
            End If
        Next
   
        tblDetail.ReBind
        tblDetail.FirstRow = 0
         
        Call Calc_Total
    End If
    
    Me.MousePointer = vbDefault
End Sub







Private Function Chk_NoDup2(ByRef inRow As Long, ByVal wsCurRec As String, ByVal wsCurRecLn As String, ByVal wsCurRecLn2 As String) As Boolean
    
    Dim wlCtr As Long
     
    Chk_NoDup2 = False
   
   
    If wsFormID = "PN001" Then
    Chk_NoDup2 = True
    Exit Function
    End If
   
   
    For wlCtr = 0 To waResult.UpperBound(1)
        If inRow <> wlCtr Then
           If wsCurRec = waResult(wlCtr, ITMCODE) And _
              wsCurRecLn = waResult(wlCtr, WHSCODE) And _
              wsCurRecLn2 = waResult(wlCtr, LOTNO) Then
              gsMsg = "重覆物料於第 " & waResult(wlCtr, LINENO) & " 行!"
              MsgBox gsMsg, vbInformation + vbOKOnly, gsTitle
              inRow = To_Value(waResult(wlCtr, LINENO))
              Exit Function
           End If
        End If
    Next
    
    Chk_NoDup2 = True

End Function


Private Function Chk_grdITMTYPE(inAccNo As String) As Boolean
    Dim wsSQL As String
    Dim rsDes As New ADODB.Recordset
    
    If wsFormID = "PN001" Then
    Chk_grdITMTYPE = True
    Exit Function
    End If
    
    
    If Trim(inAccNo) = "" Then
        gsMsg = "沒有輸入!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMTYPE = False
        Exit Function
    End If
    
    
    wsSQL = "SELECT * FROM MSTITEMTYPE"
    wsSQL = wsSQL & " WHERE ITMTYPECODE = '" & Set_Quote(inAccNo) & "'"
    
    rsDes.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
    
    If rsDes.RecordCount > 0 Then
        Chk_grdITMTYPE = True
    Else
        gsMsg = "沒有此分類!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        Chk_grdITMTYPE = False
    End If
    
    rsDes.Close
    Set rsDes = Nothing
End Function
Private Sub cmdRevise()

     
    On Error GoTo cmdRevise_Err
    
    
'    If DelValidation(wlKey) = False Then
'       wiAction = CorRec
'       MousePointer = vbDefault
'       Exit Sub
'    End If
    
    
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

Private Function DelValidation(ByVal InDocID As Long) As Boolean
Dim OutTrnCd As String
Dim OutDocNo As String

    
    
    DelValidation = False
    
    On Error GoTo DelValidation_Err
    
    
    
 '   If Not chk_txtRevNo Then Exit Function
    If Chk_PoRefDoc(InDocID, OutTrnCd, OutDocNo) = True Then
        
        Select Case OutTrnCd
        Case "GR"
        gsMsg = "進貨單 : " & OutDocNo & " 是以此採購轉為!不能刪除或改正"
        Case "PV"
        gsMsg = "供應商發票 : " & OutDocNo & " 是以此採購轉為!不能刪除或改正"
        Case "PR"
        gsMsg = "採購退貨單 : " & OutDocNo & " 是以此採購轉為!不能刪除或改正"
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
            tblDetail.Col = QTY
            tblDetail.SetFocus
            
            
  End If
   
  Me.MousePointer = vbDefault
    
    
End Sub


Private Function Chk_PoExistGrDt(ByVal CRow As Long) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsDtID As String
    
    On Error GoTo Chk_PoExistGrDt_Err
    
    
    
    wsDtID = waResult(CRow, ITMID)
    
    If wsDtID = "0" Then
        Chk_PoExistGrDt = True
        Exit Function
    End If
    
    
    wsSQL = "SELECT * FROM PopGrHd, PopGrDt "
    wsSQL = wsSQL & " WHERE GrDtPoID = " & To_Value(wlKey) & " "
    wsSQL = wsSQL & " AND GrDtItemID = " & To_Value(wsDtID) & " "
    wsSQL = wsSQL & " AND GrHDDocID = GrDtDocID "
    wsSQL = wsSQL & " AND GrHDStatus In ('1','4') "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        gsMsg = "不能更改或刪除!物料已進倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Chk_PoExistGrDt = False
        Exit Function
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    
    wsSQL = "SELECT * FROM PopPvHd, PopPvDt "
    wsSQL = wsSQL & " WHERE PvDtPoID = " & To_Value(wlKey) & " "
    wsSQL = wsSQL & " AND PvDtItemID = " & To_Value(wsDtID) & " "
    wsSQL = wsSQL & " AND PvHDDocID = PvDtDocID "
    wsSQL = wsSQL & " AND PvHDStatus In ('1','4') "
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
        gsMsg = "不能更改或刪除!物料出發票!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Chk_PoExistGrDt = False
        Exit Function
    End If
        
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_PoExistGrDt = True
    
    
    Exit Function
    
Chk_PoExistGrDt_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function
Private Function Chk_PoExistGrDtQty(ByVal CRow As Long, ByVal InQty As Double) As Boolean
    Dim rsRcd As New ADODB.Recordset
    Dim wsSQL As String
    Dim wsDtID As String
    
    On Error GoTo Chk_PoExistGrDtQty_Err
    
    
    
    wsDtID = waResult(CRow, ITMID)
    
        
    If wsDtID = "0" Then
        Chk_PoExistGrDtQty = True
        Exit Function
    End If
    
    
    wsSQL = "SELECT Sum(GrDtQty) QTY FROM PopGrHd, PopGrDt "
    wsSQL = wsSQL & " WHERE GrDtPoID = " & To_Value(wlKey) & " "
    wsSQL = wsSQL & " AND GrDtItemID = " & To_Value(wsDtID) & " "
    wsSQL = wsSQL & " AND GrHDDocID = GrDtDocID "
    wsSQL = wsSQL & " AND GrHDStatus In ('1','4') "
    
    
    rsRcd.Open wsSQL, cnCon, adOpenStatic, adLockOptimistic
     
    If rsRcd.RecordCount > 0 Then
    If To_Value(InQty) < To_Value(ReadRs(rsRcd, "QTY")) Then
        gsMsg = "數量不足!物料已進倉!"
        MsgBox gsMsg, vbOKOnly, gsTitle
        rsRcd.Close
        Set rsRcd = Nothing
        Chk_PoExistGrDtQty = False
        Exit Function
    End If
    End If
    
    rsRcd.Close
    Set rsRcd = Nothing
    
    Chk_PoExistGrDtQty = True
    
    
    Exit Function
    
Chk_PoExistGrDtQty_Err:
        gsMsg = Err.Description
        MsgBox gsMsg, vbOKOnly, gsTitle
    
End Function

Public Property Let FormID(InFormID As String)
    wsFormID = InFormID
End Property
